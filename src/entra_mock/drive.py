"""Mock Microsoft Graph — la surface « fichiers » (driveItem).

Un site, sa bibliothèque par défaut (« Documents »), et rien dedans : le
lecteur est VIDE au démarrage. Un test le remplit par le plan de contrôle
(`PUT /__admin/drive/files/{chemin}`, cf. `admin.py`) avec les fichiers dont
il a besoin — le mock n'embarque aucun jeu de fichiers.

┌─ CE QUE LE VRAI GRAPH FAIT, ET QU'UN MOCK GENTIL CACHERAIT ─────────────────┐
│ 1. une application en `Sites.Selected` sans droit sur LE site reçoit 403    │
│    `accessDenied`, sur le site comme sur tout ce qu'il contient —           │
│    `ENTRA_MOCK_SITE_GRANT=none` le reproduit ;                              │
│ 2. `/children` PAGINE, par curseur `$skiptoken` opaque — deux par page ici, │
│    pour que le chemin de pagination soit exercé par construction ;          │
│ 3. sans `$select`, chaque élément porte `createdBy`/`lastModifiedBy` : un   │
│    nom et un courriel — un client qui n'en a pas besoin doit le DIRE ;      │
│ 4. les bibliothèques de documents ne servent QUE `quickXorHash` — jamais    │
│    `sha1Hash` ni `sha256Hash` ;                                             │
│ 5. `/content` rend 302 vers une URL PRÉ-AUTHENTIFIÉE. Y renvoyer le Bearer  │
│    est une faute : 401 ici.                                                 │
└─────────────────────────────────────────────────────────────────────────────┘

Chaque route Graph passe par `_refus` (jeton, puis étranglement) — le MÊME
préambule que `/v1.0/groups` : un 429 injecté frappe les fichiers comme
l'annuaire.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
import uuid
from collections import Counter
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from typing import Any
from urllib.parse import parse_qs, quote, urlencode

from fastapi import APIRouter, Request, Response
from fastapi.responses import JSONResponse, RedirectResponse

from .app import _erreur_graph, _lien_suivant, _projeter, _refus, _SelectInconnu

# ┌─ LE SITE SERVI ────────────────────────────────────────────────────────────┐
# │ Un hôte et un chemin relatif au serveur — l'adressage qu'un client         │
# │ configure. Les deux sont réglables, pour coller à la configuration du      │
# │ client testé sans toucher au mock.                                         │
# └────────────────────────────────────────────────────────────────────────────┘
SITE_HOSTNAME = os.environ.get("ENTRA_MOCK_SITE_HOSTNAME", "contoso.sharepoint.com").lower()
SITE_PATH = "/" + os.environ.get("ENTRA_MOCK_SITE_PATH", "/sites/documents").strip("/")

# ┌─ LE DROIT DE L'APPLICATION SUR LE SITE ────────────────────────────────────┐
# │ Une application en `Sites.Selected` ne voit RIEN tant qu'un administrateur │
# │ ne lui a pas accordé un rôle sur le site précis. Sans ce geste, Graph rend │
# │ 403 `accessDenied` — le jeton est valide, c'est le site qui est fermé.     │
# │ C'est LA panne de première mise en service, et elle doit se distinguer     │
# │ d'un 401 (jeton) comme d'un 404 (chemin) dans le diagnostic du client.     │
# └────────────────────────────────────────────────────────────────────────────┘
SITE_GRANT = os.environ.get("ENTRA_MOCK_SITE_GRANT", "read")

# Deux par page, pour la même raison qu'un membre par page sur les groupes :
# Graph pagine `/children` (200 par page), et un dossier qui tiendrait en une
# page rendrait invisible un client qui ne suit pas `@odata.nextLink`.
DRIVE_PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_DRIVE_PAGE_SIZE", "2"))

#: Durée de vie d'une URL de téléchargement pré-authentifiée.
TEMPAUTH_SECONDS = int(os.environ.get("ENTRA_MOCK_TEMPAUTH_SECONDS", "3600"))

# ┌─ L'HÔTE DES URL DE TÉLÉCHARGEMENT ─────────────────────────────────────────┐
# │ En production, la redirection QUITTE graph.microsoft.com pour l'hôte du    │
# │ site. httpx comme requests retirent alors d'eux-mêmes l'en-tête            │
# │ Authorization — c'est un changement d'origine. Le mock, lui, redirige vers │
# │ LUI-MÊME : un client qui suit la redirection garde son Bearer, et reçoit   │
# │ 401 là où la production aurait répondu.                                    │
# │                                                                            │
# │ Défaut assumé : il impose au client de suivre la redirection à la main,    │
# │ sans Authorization — ce qui marche partout. Pour reproduire le changement  │
# │ d'origine, donner ici un second nom du même conteneur (alias réseau        │
# │ compose, par exemple `http://files-mock:8000`).                            │
# └────────────────────────────────────────────────────────────────────────────┘
DOWNLOAD_BASE_URL = os.environ.get("ENTRA_MOCK_DOWNLOAD_BASE_URL", "").rstrip("/")

BIBLIOTHEQUE = "Documents"


def _guid(graine: str) -> uuid.UUID:
    return uuid.uuid5(uuid.NAMESPACE_URL, f"https://{SITE_HOSTNAME}{SITE_PATH}#{graine}")


# Les identifiants ont la FORME de ceux du vrai service, parce qu'un client
# finit toujours par en loguer, en stocker ou en découper un :
#   • site  : `<hôte>,<GUID de collection>,<GUID du web>` — un identifiant
#             COMPOSITE, avec des virgules ;
#   • drive : `b!` + base64url des trois GUID (collection, web, liste), dans
#             l'ordre d'octets .NET — 66 caractères, `-` et `_` compris ;
#   • item  : `01` + 32 caractères base32 majuscules.
_COLLECTION, _WEB, _LISTE = _guid("collection"), _guid("web"), _guid("liste")
SITE_ID = f"{SITE_HOSTNAME},{_COLLECTION},{_WEB}"
DRIVE_ID = "b!" + base64.urlsafe_b64encode(
    _COLLECTION.bytes_le + _WEB.bytes_le + _LISTE.bytes_le
).decode().rstrip("=")
URL_SITE = f"https://{SITE_HOSTNAME}{SITE_PATH}"
URL_BIBLIOTHEQUE = f"{URL_SITE}/Shared%20Documents"
_CREATION_SITE = "2026-01-01T00:00:00Z"

# ┌─ L'HORLOGE DU MOCK, ET POURQUOI CE N'EST PAS L'HORLOGE MURALE ─────────────┐
# │ Une écriture par le plan de contrôle doit faire AVANCER                    │
# │ `lastModifiedDateTime` — c'est ce qu'un client incrémental compare. Mais   │
# │ l'heure murale rendrait chaque run différent du précédent, et un           │
# │ instantané pris par le client ne serait plus jamais stable.                │
# │                                                                            │
# │ D'où une horloge MONOTONE et DÉTERMINISTE : une base fixe, plus une minute │
# │ par écriture depuis la dernière remise à zéro. La même suite d'écritures   │
# │ rend les mêmes horodatages.                                                │
# └────────────────────────────────────────────────────────────────────────────┘
_BASE_HORLOGE = datetime(2026, 1, 1, tzinfo=UTC)

#: L'identité générique qui écrit les fichiers — aucune personne réelle.
UTILISATEUR = {
    "email": "user@example.invalid",
    "id": str(uuid.uuid5(uuid.NAMESPACE_DNS, "user@example.invalid")),
    "displayName": "Mock User",
}

#: Signe les `tempauth`. Tiré au démarrage : une URL émise par un autre
#: processus est refusée, comme une URL d'un autre locataire.
_SECRET = secrets.token_bytes(32)

#: Le type MIME par extension — explicite, pour ne pas dépendre du
#: `/etc/mime.types` de l'image (absent de `python:3.12-slim`).
_TYPES_MIME = {
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".json": "application/json",
    ".pdf": "application/pdf",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

#: Les caractères qu'un nom de fichier ne peut pas porter.
_INTERDITS = frozenset('"*:<>?\\|')


def _iso(instant: datetime) -> str:
    return instant.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def quick_xor_hash(donnees: bytes) -> str:
    """Le `quickXorHash` de Microsoft — la SEULE empreinte servie ici.

    Algorithme publié (OneDrive API, « QuickXorHash ») : chaque octet est XORé
    dans un registre de 160 bits, décalé de 11 bits de plus que le précédent
    (rotation modulo 160), puis la longueur sur 8 octets little-endian est
    XORée dans les 8 derniers octets. Vérifié contre l'implémentation de
    référence (paquet `quickxorhash`) — cf. tests.

    Les octets aux positions i, i+160, i+320… reçoivent le même décalage
    (160 * 11 ≡ 0 mod 160) : on les XORe d'abord, bloc par bloc, en entiers.
    """
    largeur, decalage = 160, 11
    masque = (1 << largeur) - 1
    cumul = 0
    for debut in range(0, len(donnees), largeur):
        cumul ^= int.from_bytes(donnees[debut : debut + largeur], "little")
    registre = 0
    for position, octet in enumerate(cumul.to_bytes(largeur, "little")):
        if octet:
            valeur = octet << ((position * decalage) % largeur)
            registre ^= (valeur & masque) | (valeur >> largeur)
    resultat = bytearray(registre.to_bytes(largeur // 8, "little"))
    for i, octet in enumerate(len(donnees).to_bytes(8, "little")):
        resultat[largeur // 8 - 8 + i] ^= octet
    return base64.b64encode(bytes(resultat)).decode()


@dataclass
class Element:
    """Un `driveItem` — dossier quand `contenu` est None."""

    #: Depuis la racine du lecteur, casse d'origine ; "" pour la racine.
    chemin: str
    identifiant: str
    unique_id: str
    #: L'ID d'élément de liste — il figure dans le `$skiptoken`.
    numero: int
    cree: str
    modifie: str
    auteur: dict[str, str]
    modificateur: dict[str, str]
    contenu: bytes | None = None
    version: int = 1
    empreinte: str = ""
    type_mime: str = ""

    @property
    def nom(self) -> str:
        return self.chemin.rpartition("/")[2] if self.chemin else "root"

    @property
    def parent(self) -> str:
        return self.chemin.rpartition("/")[0]

    @property
    def dossier(self) -> bool:
        return self.contenu is None


def _cle(chemin: str) -> str:
    """Les chemins ne distinguent pas la casse : `Reports` et `reports`
    désignent le même dossier. Un mock sensible à la casse ferait échouer un
    client correct."""
    return "/".join(p for p in chemin.split("/") if p).casefold()


class Lecteur:
    """L'état du lecteur — VIDE (la seule racine) après `reinitialiser()`."""

    def __init__(self) -> None:
        self.verrou = threading.RLock()
        self.elements: dict[str, Element] = {}
        self.compteurs: Counter[str] = Counter()
        self.naissances: Counter[str] = Counter()
        self.ecritures = 0
        self.dernier_numero = 0
        self.droit = SITE_GRANT
        self.reinitialiser()

    def reinitialiser(self) -> None:
        with self.verrou:
            self.elements = {}
            self.compteurs = Counter()
            self.naissances = Counter()
            self.ecritures = 0
            self.dernier_numero = 0
            # Relu à chaque remise à zéro : un test peut changer SITE_GRANT et
            # retrouver l'état que l'environnement aurait produit au démarrage.
            self.droit = SITE_GRANT
            systeme = {"displayName": "System Account"}
            self._creer("", None, _CREATION_SITE, _CREATION_SITE, systeme)

    def _creer(
        self, chemin: str, contenu: bytes | None, cree: str, modifie: str, auteur: dict[str, str]
    ) -> Element:
        cle = _cle(chemin)
        # Un fichier supprimé puis réécrit est un NOUVEL élément : nouvel
        # identifiant. La naissance entre donc dans la graine — déterministe
        # pour une même suite d'opérations.
        graine = f"{cle}#{self.naissances[cle]}"
        self.naissances[cle] += 1
        self.dernier_numero += 1
        empreinte = hashlib.sha1(graine.encode(), usedforsecurity=False).digest()
        element = Element(
            chemin=chemin.strip("/"),
            identifiant="01" + base64.b32encode(empreinte).decode()[:32],
            unique_id=str(_guid(graine)),
            numero=self.dernier_numero,
            cree=cree,
            modifie=modifie,
            auteur=auteur,
            modificateur=auteur,
        )
        if contenu is not None:
            self._remplir(element, contenu)
        self.elements[cle] = element
        return element

    @staticmethod
    def _remplir(element: Element, contenu: bytes) -> None:
        element.contenu = contenu
        element.empreinte = quick_xor_hash(contenu)
        extension = "." + element.nom.rpartition(".")[2].lower() if "." in element.nom else ""
        element.type_mime = _TYPES_MIME.get(extension, "application/octet-stream")

    def instant(self) -> str:
        self.ecritures += 1
        return _iso(_BASE_HORLOGE + timedelta(minutes=self.ecritures))

    def racine(self) -> Element:
        return self.elements[""]

    def trouver(self, chemin: str) -> Element | None:
        return self.elements.get(_cle(chemin))

    def par_identifiant(self, identifiant: str) -> Element | None:
        return next((e for e in self.elements.values() if e.identifiant == identifiant), None)

    def par_unique_id(self, unique_id: str) -> Element | None:
        unique_id = unique_id.strip("{}").lower()
        return next((e for e in self.elements.values() if e.unique_id == unique_id), None)

    def enfants(self, dossier: Element) -> list[Element]:
        """Les enfants DIRECTS, triés par nom sans tenir compte de la casse."""
        if not dossier.dossier:
            return []
        cle = _cle(dossier.chemin)
        return sorted(
            (e for e in self.elements.values() if e.chemin and _cle(e.parent) == cle),
            key=lambda e: e.nom.casefold(),
        )

    def taille(self, element: Element) -> int:
        if element.contenu is not None:
            return len(element.contenu)
        prefixe = _cle(element.chemin) + "/" if element.chemin else ""
        return sum(
            len(e.contenu)
            for cle, e in self.elements.items()
            if e.contenu is not None and cle.startswith(prefixe)
        )

    def ecrire(self, chemin: str, contenu: bytes) -> tuple[Element, bool]:
        """Crée ou ÉCRASE un fichier.

        L'écrasement GARDE l'identifiant (c'est le même élément), incrémente le
        `n` de `eTag`/`cTag` et avance `lastModifiedDateTime` sur l'horloge du
        mock. Les dossiers intermédiaires sont créés.
        """
        chemin = "/".join(p for p in chemin.split("/") if p)
        with self.verrou:
            existant = self.trouver(chemin)
            if existant is not None:
                if existant.dossier:
                    raise IsADirectoryError(chemin)
                self._remplir(existant, contenu)
                existant.version += 1
                existant.modifie = self.instant()
                existant.modificateur = dict(UTILISATEUR)
                return existant, False
            morceaux = chemin.split("/")
            for profondeur in range(1, len(morceaux)):
                parent = "/".join(morceaux[:profondeur])
                if (dossier := self.trouver(parent)) is None:
                    instant = self.instant()
                    self._creer(parent, None, instant, instant, dict(UTILISATEUR))
                elif not dossier.dossier:
                    raise NotADirectoryError(parent)
            instant = self.instant()
            return self._creer(chemin, contenu, instant, instant, dict(UTILISATEUR)), True

    def retirer(self, chemin: str) -> bool:
        """Supprime un fichier — ou un dossier et tout ce qu'il contient."""
        cle = _cle(chemin)
        with self.verrou:
            if not cle or cle not in self.elements:
                return False
            for autre in [c for c in self.elements if c == cle or c.startswith(cle + "/")]:
                del self.elements[autre]
            return True


lecteur_mock = Lecteur()

routeur: APIRouter = APIRouter()


# ── Les URL de téléchargement pré-authentifiées ──────────────────────────────


def _b64(donnees: bytes) -> str:
    return base64.urlsafe_b64encode(donnees).decode().rstrip("=")


def _signer(charge: str) -> str:
    return _b64(hmac.new(_SECRET, charge.encode(), hashlib.sha256).digest())


def emettre_tempauth(unique_id: str) -> str:
    """Un `tempauth` opaque, signé, qui EXPIRE — rien n'est stocké côté mock."""
    expiration = int(time.time()) + TEMPAUTH_SECONDS
    charge = _b64(json.dumps({"uid": unique_id, "exp": expiration}).encode())
    return f"v1.{charge}.{_signer(charge)}"


def tempauth_valide(jeton: str, unique_id: str) -> bool:
    try:
        version, charge, signature = jeton.split(".")
        contenu = json.loads(base64.urlsafe_b64decode(charge + "=" * (-len(charge) % 4)))
    except ValueError:
        return False
    if version != "v1" or not hmac.compare_digest(signature, _signer(charge)):
        return False
    if not isinstance(contenu, dict):
        return False
    return bool(contenu.get("uid") == unique_id and contenu.get("exp", 0) > time.time())


def _url_telechargement(request: Request, element: Element) -> str:
    base = DOWNLOAD_BASE_URL or str(request.base_url).rstrip("/")
    parametres = urlencode(
        {
            "UniqueId": element.unique_id,
            "Translate": "false",
            "tempauth": emettre_tempauth(element.unique_id),
            "ApiVersion": "2.0",
        }
    )
    return f"{base}{SITE_PATH}/_layouts/15/download.aspx?{parametres}"


# ── Le rendu ─────────────────────────────────────────────────────────────────


def _etiquette(element: Element, prefixe: str = "") -> str:
    """`"{GUID},n"` — et `"c:{GUID},n"` pour le cTag. Les guillemets font partie
    de la valeur, comme chez Graph."""
    return f'"{prefixe}{{{element.unique_id.upper()}}},{element.version}"'


def _reference_parent(element: Element) -> dict[str, Any]:
    reference: dict[str, Any] = {"driveType": "documentLibrary", "driveId": DRIVE_ID}
    if not element.chemin:
        return reference
    parent = lecteur_mock.trouver(element.parent)
    if parent is not None:
        reference["id"] = parent.identifiant
        if parent.chemin:
            reference["name"] = parent.nom
        reference["path"] = f"/drives/{DRIVE_ID}/root:" + (
            f"/{parent.chemin}" if parent.chemin else ""
        )
    reference["siteId"] = SITE_ID
    return reference


def _rendu(element: Element, request: Request) -> dict[str, Any]:
    """Le jeu de propriétés PAR DÉFAUT d'un `driveItem` — large, comme Graph.

    `createdBy` et `lastModifiedBy` y portent une identité. Un client qui n'a
    besoin que du nom, de la taille et du `cTag` doit le DIRE par `$select`.
    """
    objet: dict[str, Any] = {}
    if not element.dossier:
        # Rendue dans le jeu par défaut, comme chez Graph — mais PAS avec un
        # `$select` qui ne la nomme pas.
        objet["@microsoft.graph.downloadUrl"] = _url_telechargement(request, element)
    objet |= {
        "createdBy": {"user": dict(element.auteur)},
        "createdDateTime": element.cree,
        "eTag": _etiquette(element),
        "id": element.identifiant,
        "lastModifiedBy": {"user": dict(element.modificateur)},
        "lastModifiedDateTime": element.modifie,
        "name": element.nom,
        "parentReference": _reference_parent(element),
        "webUrl": f"{URL_BIBLIOTHEQUE}/{quote(element.chemin)}".rstrip("/"),
        "cTag": _etiquette(element, "c:"),
        "fileSystemInfo": {
            "createdDateTime": element.cree,
            "lastModifiedDateTime": element.modifie,
        },
        "size": lecteur_mock.taille(element),
    }
    if element.dossier:
        objet["folder"] = {"childCount": len(lecteur_mock.enfants(element))}
        if not element.chemin:
            objet["root"] = {}
    else:
        # `quickXorHash` SEUL : les bibliothèques de documents ne calculent ni
        # sha1 ni sha256. Un client qui compare des sha256 tombe sur une clé
        # absente — ici comme en production.
        objet["file"] = {
            "mimeType": element.type_mime,
            "hashes": {"quickXorHash": element.empreinte},
        }
    return objet


#: L'univers de `$select` — les propriétés que le TYPE `driveItem` déclare
#: (baseItem + driveItem au CSDL), pas celles que la page contient.
PROPRIETES_ELEMENT = frozenset(
    {
        "@microsoft.graph.downloadUrl",
        "audio",
        "bundle",
        "cTag",
        "createdBy",
        "createdDateTime",
        "deleted",
        "description",
        "eTag",
        "file",
        "fileSystemInfo",
        "folder",
        "id",
        "image",
        "lastModifiedBy",
        "lastModifiedDateTime",
        "location",
        "malware",
        "name",
        "package",
        "parentReference",
        "photo",
        "publication",
        "remoteItem",
        "root",
        "searchResult",
        "shared",
        "sharepointIds",
        "size",
        "specialFolder",
        "video",
        "webDavUrl",
        "webUrl",
    }
)


def _derniere_modification() -> str:
    return max(e.modifie for e in lecteur_mock.elements.values())


def _site() -> dict[str, Any]:
    nom = SITE_PATH.rpartition("/")[2]
    return {
        "createdDateTime": _CREATION_SITE,
        "description": "",
        "id": SITE_ID,
        "lastModifiedDateTime": _derniere_modification(),
        "name": nom,
        "webUrl": URL_SITE,
        "displayName": nom,
        "root": {},
        "siteCollection": {"hostname": SITE_HOSTNAME},
    }


def _lecteur() -> dict[str, Any]:
    utilise = sum(len(e.contenu) for e in lecteur_mock.elements.values() if e.contenu is not None)
    total = 27_487_790_694_400
    return {
        "createdDateTime": _CREATION_SITE,
        "description": "",
        "id": DRIVE_ID,
        "lastModifiedDateTime": _derniere_modification(),
        "name": BIBLIOTHEQUE,
        "webUrl": URL_BIBLIOTHEQUE,
        "driveType": "documentLibrary",
        "createdBy": {"user": {"displayName": "System Account"}},
        "owner": {
            "group": {
                "id": str(_guid("groupe-proprietaires")),
                "displayName": f"{SITE_PATH.rpartition('/')[2]} Owners",
            }
        },
        "quota": {
            "deleted": 0,
            "remaining": total - utilise,
            "state": "normal",
            "total": total,
            "used": utilise,
        },
    }


# ── Le préambule et les erreurs ──────────────────────────────────────────────


def _prelude(request: Request, operation: str) -> JSONResponse | None:
    """Jeton, étranglement (le préambule commun de `app.py`), puis droit sur le
    site. Le compteur ne retient que les appels authentifiés et non étranglés :
    un 429 rejoué par le client ne fausse pas le décompte de ses
    téléchargements."""
    if (refus := _refus(request)) is not None:
        return refus
    with lecteur_mock.verrou:
        lecteur_mock.compteurs[operation] += 1
    if lecteur_mock.droit == "none":
        return _erreur_graph(403, "accessDenied", "Access denied")
    return None


def _introuvable() -> JSONResponse:
    return _erreur_graph(404, "itemNotFound", "The resource could not be found.")


def _site_introuvable() -> JSONResponse:
    return _erreur_graph(404, "itemNotFound", "Requested site could not be found")


def _select_inconnu(exc: _SelectInconnu, type_: str) -> JSONResponse:
    return _erreur_graph(
        400,
        "BadRequest",
        "Parsing OData Select and Expand failed: Could not find a property named "
        f"'{exc.champs[0]}' on type 'microsoft.graph.{type_}'.",
    )


def _projeter_un(
    objet: dict[str, Any], request: Request, univers: frozenset[str], contexte: str
) -> dict[str, Any]:
    (projete,) = _projeter([objet], request.query_params.get("$select"), univers)
    return {"@odata.context": f"https://graph.microsoft.com/v1.0/$metadata#{contexte}", **projete}


_REPONSES: dict[int | str, dict[str, Any]] = {
    401: {"description": "`InvalidAuthenticationToken` — Bearer absent, expiré, autre audience"},
    403: {"description": "`accessDenied` — aucun rôle `Sites.Selected` accordé sur le site"},
    404: {"description": "`itemNotFound`"},
    429: {"description": "`TooManyRequests` + `Retry-After` (injectable)"},
}


# ── Le site ──────────────────────────────────────────────────────────────────


def _servir_site(request: Request) -> Response:
    try:
        return JSONResponse(_projeter_un(_site(), request, frozenset(_site()), "sites/$entity"))
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "site")


@routeur.get("/v1.0/sites/{hostname}:/{server_relative_path:path}", responses=_REPONSES)
def site_par_chemin(hostname: str, server_relative_path: str, request: Request) -> Response:
    """Le site par son CHEMIN — `/sites/{hôte}:/{chemin relatif au serveur}`.

    C'est l'adressage qu'un client configure (un hôte et un chemin lisibles),
    et il en tire l'identifiant COMPOSITE du site pour la suite. Le
    deux-points final optionnel (`…:/sites/documents:`) est accepté, comme
    chez Graph.
    """
    if (refus := _prelude(request, "site")) is not None:
        return refus
    if hostname.casefold() != SITE_HOSTNAME:
        return _erreur_graph(400, "invalidRequest", "Invalid hostname for this tenancy")
    if _cle(server_relative_path.rstrip(":")) != _cle(SITE_PATH):
        return _site_introuvable()
    return _servir_site(request)


@routeur.get("/v1.0/sites/{site_id}", responses=_REPONSES)
def site(site_id: str, request: Request) -> Response:
    """Le site par son identifiant composite."""
    if (refus := _prelude(request, "site")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _site_introuvable()
    return _servir_site(request)


def _servir_lecteur(request: Request) -> Response:
    try:
        return JSONResponse(
            _projeter_un(_lecteur(), request, frozenset(_lecteur()), "drives/$entity")
        )
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "drive")


@routeur.get("/v1.0/sites/{site_id}/drive", responses=_REPONSES)
def lecteur_par_defaut(site_id: str, request: Request) -> Response:
    """La bibliothèque PAR DÉFAUT du site — « Documents »."""
    if (refus := _prelude(request, "drive")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _site_introuvable()
    return _servir_lecteur(request)


@routeur.get("/v1.0/sites/{site_id}/drives", responses=_REPONSES)
def lecteurs(site_id: str, request: Request) -> Response:
    """Les bibliothèques du site, en collection `value[]`."""
    if (refus := _prelude(request, "drives")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _site_introuvable()
    try:
        valeurs = _projeter(
            [_lecteur()], request.query_params.get("$select"), frozenset(_lecteur())
        )
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "drive")
    return JSONResponse(
        {"@odata.context": "https://graph.microsoft.com/v1.0/$metadata#drives", "value": valeurs}
    )


@routeur.get("/v1.0/drives/{drive_id}", responses=_REPONSES)
def lecteur(drive_id: str, request: Request) -> Response:
    """Un lecteur par son identifiant `b!…`."""
    if (refus := _prelude(request, "drive")) is not None:
        return refus
    if drive_id != DRIVE_ID:
        return _introuvable()
    return _servir_lecteur(request)


# ── Les éléments ─────────────────────────────────────────────────────────────


def _servir_element(request: Request, element: Element) -> Response:
    try:
        return JSONResponse(
            _projeter_un(
                _rendu(element, request),
                request,
                PROPRIETES_ELEMENT,
                f"drives('{DRIVE_ID}')/items/$entity",
            )
        )
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "driveItem")


def _curseur(element: Element) -> str:
    """Le `$skiptoken` — opaque pour le client : la clé de tri du DERNIER
    élément rendu, pas un rang. Un fichier ajouté ou retiré pendant la
    pagination ne décale donc rien."""
    return _b64(
        f"Paged=TRUE&p_SortBehavior=0&p_FileLeafRef={quote(element.nom)}"
        f"&p_ID={element.numero}".encode()
    )


def _lire_curseur(jeton: str) -> str:
    try:
        texte = base64.urlsafe_b64decode(jeton + "=" * (-len(jeton) % 4)).decode()
        return parse_qs(texte, strict_parsing=True)["p_FileLeafRef"][0].casefold()
    except (ValueError, KeyError) as exc:
        raise ValueError("Invalid skip token.") from exc


def _taille_de_page(request: Request) -> int:
    """`$top` est un PLAFOND, jamais une promesse : OData autorise le serveur à
    rendre moins, et Graph le fait. Un `$top=999` ne court-circuite donc pas
    la pagination du mock — le client doit suivre `@odata.nextLink` quoi qu'il
    ait demandé."""
    taille = max(1, DRIVE_PAGE_SIZE)
    if (top := request.query_params.get("$top")) is not None:
        if not top.isdigit() or int(top) < 1:
            raise ValueError(f"Invalid value '{top}' for $top query option.")
        taille = min(taille, int(top))
    return taille


def _servir_enfants(request: Request, dossier: Element) -> Response:
    try:
        taille = _taille_de_page(request)
        jeton = request.query_params.get("$skiptoken")
        apres = _lire_curseur(jeton) if jeton is not None else None
    except ValueError as exc:
        return _erreur_graph(400, "invalidRequest", str(exc))

    with lecteur_mock.verrou:
        enfants = lecteur_mock.enfants(dossier)
        if apres is not None:
            enfants = [e for e in enfants if e.nom.casefold() > apres]
        page, reste = enfants[:taille], enfants[taille:]
        try:
            valeurs = _projeter(
                [_rendu(e, request) for e in page],
                request.query_params.get("$select"),
                PROPRIETES_ELEMENT,
            )
        except _SelectInconnu as exc:
            return _select_inconnu(exc, "driveItem")

    corps: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(driveItem)",
        "value": valeurs,
    }
    if reste:
        # Le lien suivant reprend le chemin de la requête — par chemin ou par
        # identifiant, comme elle est venue — et recopie `$select` et `$top`.
        chemin = quote(request.url.path, safe="/:!,$'()*+;=@")
        corps["@odata.nextLink"] = _lien_suivant(request, chemin, _curseur(page[-1]))
    return JSONResponse(corps)


def _servir_contenu(request: Request, element: Element) -> Response:
    """302 vers une URL PRÉ-AUTHENTIFIÉE — jamais les octets en direct."""
    if element.dossier:
        return _introuvable()
    return RedirectResponse(_url_telechargement(request, element), status_code=302)


def _decouper(chemin: str) -> tuple[str, str]:
    """`reports:/children` → (`reports`, `children`) ;
    `reports/report-1.txt:` → (`reports/report-1.txt`, ``)."""
    if ":/" in chemin:
        cible, _, segment = chemin.rpartition(":/")
        return cible, segment.strip("/")
    return chemin.rstrip(":/"), ""


_REPONSES_CONTENU: dict[int | str, dict[str, Any]] = {
    **_REPONSES,
    302: {
        "description": "Redirection vers une URL de téléchargement PRÉ-AUTHENTIFIÉE (`Location`, "
        "`tempauth` à durée limitée). À suivre SANS en-tête Authorization : l'URL "
        "rend 401 si on lui en présente un."
    },
}


@routeur.get("/v1.0/drives/{drive_id}/root", responses=_REPONSES)
def racine(drive_id: str, request: Request) -> Response:
    """La racine du lecteur."""
    if (refus := _prelude(request, "item")) is not None:
        return refus
    if drive_id != DRIVE_ID:
        return _introuvable()
    return _servir_element(request, lecteur_mock.racine())


@routeur.get("/v1.0/drives/{drive_id}/root/children", responses=_REPONSES)
def enfants_de_la_racine(drive_id: str, request: Request) -> Response:
    """Les enfants de la racine — vides tant que rien n'a été écrit."""
    if (refus := _prelude(request, "children")) is not None:
        return refus
    if drive_id != DRIVE_ID:
        return _introuvable()
    return _servir_enfants(request, lecteur_mock.racine())


@routeur.get(
    "/v1.0/drives/{drive_id}/root:/{chemin:path}",
    responses=_REPONSES_CONTENU,
)
def par_chemin(drive_id: str, chemin: str, request: Request) -> Response:
    """L'adressage PAR CHEMIN, sous ses trois formes :

    • `root:/reports:/children` — les enfants d'un dossier, PAGINÉS
      (`@odata.nextLink`, `$skiptoken` opaque), `$select` honoré ;
    • `root:/reports/report-1.txt` (deux-points final facultatif) — un
      élément ;
    • `root:/reports/report-1.txt:/content` — 302 vers l'URL de
      téléchargement.

    La casse du chemin est indifférente.
    """
    cible, segment = _decouper(chemin)
    servir = {"": _servir_element, "children": _servir_enfants, "content": _servir_contenu}
    if (
        refus := _prelude(request, segment if segment in servir and segment else "item")
    ) is not None:
        return refus
    if segment not in servir:
        return _erreur_graph(400, "BadRequest", f"Resource not found for the segment '{segment}'.")
    element = lecteur_mock.trouver(cible) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else servir[segment](request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}", responses=_REPONSES)
def element_par_identifiant(drive_id: str, item_id: str, request: Request) -> Response:
    """Un élément par son identifiant `01…`."""
    if (refus := _prelude(request, "item")) is not None:
        return refus
    element = lecteur_mock.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_element(request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}/children", responses=_REPONSES)
def enfants_par_identifiant(drive_id: str, item_id: str, request: Request) -> Response:
    """Les enfants d'un dossier désigné par son identifiant — même pagination."""
    if (refus := _prelude(request, "children")) is not None:
        return refus
    element = lecteur_mock.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_enfants(request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}/content", responses=_REPONSES_CONTENU)
def contenu(drive_id: str, item_id: str, request: Request) -> Response:
    """Le contenu d'un fichier : 302 vers une URL PRÉ-AUTHENTIFIÉE.

    C'est ce que Graph rend, et il faut le suivre SANS le Bearer — l'URL porte
    sa propre autorisation (`tempauth`), limitée dans le temps.
    """
    if (refus := _prelude(request, "content")) is not None:
        return refus
    element = lecteur_mock.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_contenu(request, element)


# ── L'hôte du site : le téléchargement ───────────────────────────────────────


def _refus_telechargement(description: str) -> JSONResponse:
    # Le STATUT est ce qui fait foi (401) ; le corps exact du vrai service sur
    # ces refus n'a pas été relevé — un client ne doit pas s'y adosser.
    return JSONResponse(
        status_code=401, content={"error": "invalid_request", "error_description": description}
    )


@routeur.get(f"{SITE_PATH}/_layouts/15/download.aspx", include_in_schema=False)
def telecharger(request: Request) -> Response:
    """L'URL PRÉ-AUTHENTIFIÉE — hôte du site, hors de Graph, hors contrat.

    ┌─ UN BEARER ICI EST UNE FAUTE ──────────────────────────────────────────┐
    │ L'autorisation est DANS l'URL (`tempauth`). Un client qui y renvoie le │
    │ jeton Graph expose ce jeton à un autre hôte — et se fait refuser : 401.│
    │ Un `tempauth` inconnu, falsifié ou expiré : 401 aussi.                 │
    └────────────────────────────────────────────────────────────────────────┘
    """
    if "authorization" in request.headers:
        return _refus_telechargement(
            "Pre-authenticated download URL: the request must not carry an Authorization header."
        )
    unique_id = request.query_params.get("UniqueId", "").strip("{}").lower()
    if not tempauth_valide(request.query_params.get("tempauth", ""), unique_id):
        return _refus_telechargement("Invalid or expired tempauth.")
    with lecteur_mock.verrou:
        lecteur_mock.compteurs["download"] += 1
        element = lecteur_mock.par_unique_id(unique_id)
    if element is None or element.contenu is None:
        return JSONResponse(status_code=404, content={"error": "File Not Found."})
    return Response(
        element.contenu,
        media_type=element.type_mime,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(element.nom)}",
            "ETag": _etiquette(element),
            "CTag": _etiquette(element, "c:"),
        },
    )


def nom_valide(chemin: str) -> bool:
    """Un chemin que le vrai service accepterait.

    Règles publiées (« Restrictions and limitations in OneDrive and
    SharePoint », support.microsoft.com) : ni `" * : < > ? / \\ |`, ni espace
    en tête ou en fin, ni `_vti_` nulle part dans le nom.

    Exception ASSUMÉE : les noms en `~$…`, que la même page déclare invalides.
    Un fichier de verrou doit pouvoir être écrit pour éprouver le filtre de
    fichiers temporaires d'un client.
    """
    morceaux = [p for p in chemin.split("/") if p]
    return bool(morceaux) and all(
        not (_INTERDITS & set(p)) and p == p.strip() and "_vti_" not in p.lower() for p in morceaux
    )
