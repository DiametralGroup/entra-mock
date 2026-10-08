"""Mock Microsoft Graph — la surface SharePoint : un site, sa bibliothèque, un dossier.

┌─ POURQUOI CETTE SURFACE VIT DANS CE MOCK, ET PAS DANS UN AUTRE ─────────────┐
│ Le consommateur (insights360, source `depot_finance`) lit le dépôt Finance  │
│ avec la MÊME application Entra que l'appartenance aux groupes : même        │
│ autorité, même jeton client credentials, même scope                         │
│ `https://graph.microsoft.com/.default`. Un second mock aurait son propre    │
│ jeton — et laisserait passer un client qui en demande deux.                 │
└─────────────────────────────────────────────────────────────────────────────┘

┌─ CE QUE LE VRAI GRAPH FAIT, ET QU'UN MOCK GENTIL CACHERAIT ─────────────────┐
│ 1. `Sites.Selected` : une application sans droit sur LE site reçoit 403     │
│    `accessDenied`, sur le site comme sur tout ce qu'il contient —           │
│    `ENTRA_MOCK_SITE_GRANT=none` le reproduit ;                              │
│ 2. `/children` PAGINE, par curseur `$skiptoken` opaque — deux par page ici, │
│    pour que le chemin de pagination soit exercé par construction ;          │
│ 3. sans `$select`, chaque élément porte `createdBy`/`lastModifiedBy` : le   │
│    nom et le courriel d'une PERSONNE : sans `$select`, on les collecte ;    │
│ 4. SharePoint ne sert QUE `quickXorHash` — jamais `sha1Hash` ni             │
│    `sha256Hash`, qui n'existent que sur OneDrive grand public ;             │
│ 5. `/content` rend 302 vers une URL PRÉ-AUTHENTIFIÉE. Y renvoyer le Bearer  │
│    est une faute : 401 ici.                                                 │
└─────────────────────────────────────────────────────────────────────────────┘

Chaque route Graph passe par `_refus` (jeton, puis étranglement) — le MÊME
préambule que `/v1.0/groups` : un 429 injecté frappe le dépôt comme l'annuaire.
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
from datetime import timedelta
from typing import Any
from urllib.parse import parse_qs, quote, urlencode

from fastapi import APIRouter, Request, Response
from fastapi.responses import JSONResponse, RedirectResponse

from .app import _erreur_graph, _lien_suivant, _projeter, _refus, _SelectInconnu
from .dataset import depot_finance as jeu

# ┌─ LE DROIT DE L'APPLICATION SUR LE SITE ────────────────────────────────────┐
# │ Une application en `Sites.Selected` ne voit RIEN tant qu'un administrateur │
# │ ne lui a pas accordé un rôle (read, write…) sur le site précis. Sans ce    │
# │ geste, Graph rend 403 `accessDenied` — le jeton est valide, c'est le site  │
# │ qui est fermé. C'est LA panne de première mise en service, et elle doit    │
# │ se distinguer d'un 401 (jeton) comme d'un 404 (chemin) dans le diagnostic  │
# │ du consommateur.                                                           │
# └────────────────────────────────────────────────────────────────────────────┘
SITE_GRANT = os.environ.get("ENTRA_MOCK_SITE_GRANT", "read")

# Deux par page, pour la même raison qu'un membre par page sur les groupes :
# Graph pagine `/children` (200 par page), et un dossier qui tiendrait en une
# page rendrait invisible un client qui ne suit pas `@odata.nextLink`.
DRIVE_PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_DRIVE_PAGE_SIZE", "2"))

#: Durée de vie d'une URL de téléchargement pré-authentifiée.
TEMPAUTH_SECONDS = int(os.environ.get("ENTRA_MOCK_TEMPAUTH_SECONDS", "3600"))

# ┌─ L'HÔTE DES URL DE TÉLÉCHARGEMENT ─────────────────────────────────────────┐
# │ En production, la redirection QUITTE graph.microsoft.com pour l'hôte       │
# │ SharePoint du locataire. httpx comme requests retirent alors d'eux-mêmes   │
# │ l'en-tête Authorization — c'est un changement d'origine. Le mock, lui,     │
# │ redirige vers LUI-MÊME : un client qui suit la redirection garde son       │
# │ Bearer, et reçoit 401 là où la production aurait répondu.                  │
# │                                                                            │
# │ Défaut assumé : il impose au consommateur de suivre la redirection à la    │
# │ main, sans Authorization — ce qui marche partout. Pour reproduire le       │
# │ changement d'origine, donner ici un second nom du même conteneur (alias    │
# │ réseau compose, par exemple `http://sharepoint-mock:8000`).                │
# └────────────────────────────────────────────────────────────────────────────┘
DOWNLOAD_BASE_URL = os.environ.get("ENTRA_MOCK_DOWNLOAD_BASE_URL", "").rstrip("/")


def _guid(graine: str) -> uuid.UUID:
    return uuid.uuid5(uuid.NAMESPACE_URL, f"https://{jeu.HOTE}{jeu.CHEMIN_SITE}#{graine}")


# Les identifiants ont la FORME de ceux de SharePoint, parce qu'un consommateur
# finit toujours par en loguer, en stocker ou en découper un :
#   • site  : `<hôte>,<GUID de collection>,<GUID du web>` — un identifiant
#             COMPOSITE, avec des virgules ;
#   • drive : `b!` + base64url des trois GUID (collection, web, liste), dans
#             l'ordre d'octets .NET — 66 caractères, `-` et `_` compris ;
#   • item  : `01` + 32 caractères base32 majuscules.
_COLLECTION, _WEB, _LISTE = _guid("collection"), _guid("web"), _guid("liste")
SITE_ID = f"{jeu.HOTE},{_COLLECTION},{_WEB}"
DRIVE_ID = "b!" + base64.urlsafe_b64encode(
    _COLLECTION.bytes_le + _WEB.bytes_le + _LISTE.bytes_le
).decode().rstrip("=")
URL_SITE = f"https://{jeu.HOTE}{jeu.CHEMIN_SITE}"
URL_BIBLIOTHEQUE = f"{URL_SITE}/Shared%20Documents"
_CREATION_SITE = "2024-03-11T09:00:00Z"

# ┌─ L'HORLOGE DU MOCK, ET POURQUOI CE N'EST PAS L'HORLOGE MURALE ─────────────┐
# │ Un dépôt par le plan de contrôle doit faire AVANCER `lastModifiedDateTime` │
# │ — c'est ce qu'un consommateur incrémental compare. Mais l'heure murale     │
# │ rendrait chaque run différent du précédent, et l'instantané d'idempotence  │
# │ du consommateur ne serait plus jamais stable.                              │
# │                                                                            │
# │ D'où une horloge MONOTONE et DÉTERMINISTE : une base fixe, postérieure à   │
# │ tout le jeu par défaut, plus une minute par écriture depuis la dernière    │
# │ remise à zéro. La même suite d'écritures rend les mêmes horodatages.       │
# └────────────────────────────────────────────────────────────────────────────┘
_BASE_HORLOGE = jeu._instant("2026-07-15T08:00:00Z")

#: Signe les `tempauth`. Tiré au démarrage : une URL émise par un autre
#: processus est refusée, comme une URL d'un autre locataire.
_SECRET = secrets.token_bytes(32)

#: Le type MIME par extension — explicite, pour ne pas dépendre du
#: `/etc/mime.types` de l'image (absent de `python:3.12-slim`).
_TYPES_MIME = {
    ".xlsx": jeu.MIME_XLSX,
    ".xls": "application/vnd.ms-excel",
    ".csv": "text/csv",
    ".txt": "text/plain",
    ".pdf": "application/pdf",
}

#: Les caractères qu'un nom SharePoint ne peut pas porter.
_INTERDITS = frozenset('"*:<>?\\|')


def quick_xor_hash(donnees: bytes) -> str:
    """Le `quickXorHash` de Microsoft — la SEULE empreinte que SharePoint sert.

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
    #: L'ID d'élément de liste SharePoint — il figure dans le `$skiptoken`.
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
    """SharePoint ne distingue pas la casse des chemins : `insights360` et
    `Insights360` désignent le même dossier. Un mock sensible à la casse ferait
    échouer un consommateur correct."""
    return "/".join(p for p in chemin.split("/") if p).casefold()


def _identite(code: str) -> dict[str, str]:
    nom, courriel = jeu.DEPOSITAIRES[code]
    return {
        "email": courriel,
        "id": str(uuid.uuid5(uuid.NAMESPACE_DNS, courriel)),
        "displayName": nom,
    }


def _depositaire(nom: str) -> str:
    """Le déposant d'un fichier écrit par le plan de contrôle : l'entité que
    nomme le fichier si elle est connue, la Finance de Nantes sinon."""
    morceaux = nom.split("_")
    return morceaux[1] if len(morceaux) > 2 and morceaux[1] in jeu.DEPOSITAIRES else "NTE"


class Depot:
    """L'état du lecteur — remis au jeu par défaut par `reinitialiser()`."""

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
            for chemin, cree in jeu.DOSSIERS:
                self._creer(chemin, None, cree, cree, _identite("NTE"))
            for f in jeu.jeu_par_defaut():
                element = self._creer(
                    f.chemin, f.contenu, f.cree, f.modifie, _identite(f.depositaire)
                )
                element.version = f.version

    def _creer(
        self, chemin: str, contenu: bytes | None, cree: str, modifie: str, auteur: dict[str, str]
    ) -> Element:
        cle = _cle(chemin)
        # Un fichier supprimé puis re-déposé est un NOUVEL élément pour
        # SharePoint : nouvel identifiant. La naissance entre donc dans la
        # graine — déterministe pour une même suite d'opérations.
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
        return jeu.iso(_BASE_HORLOGE + timedelta(minutes=self.ecritures))

    def trouver(self, chemin: str) -> Element | None:
        return self.elements.get(_cle(chemin))

    def par_identifiant(self, identifiant: str) -> Element | None:
        return next((e for e in self.elements.values() if e.identifiant == identifiant), None)

    def par_unique_id(self, unique_id: str) -> Element | None:
        unique_id = unique_id.strip("{}").lower()
        return next((e for e in self.elements.values() if e.unique_id == unique_id), None)

    def enfants(self, dossier: Element) -> list[Element]:
        """Les enfants DIRECTS, triés par nom — le sous-dossier `modeles` tombe
        donc au milieu des fichiers, pas en tête : un consommateur doit
        l'écarter en cours de flux, pas en sautant la première ligne."""
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

    def deposer(self, chemin: str, contenu: bytes) -> tuple[Element, bool]:
        """Crée ou ÉCRASE un fichier — le geste de la Finance qui re-dépose.

        L'écrasement GARDE l'identifiant (c'est le même élément SharePoint),
        incrémente le `n` de `eTag`/`cTag` et avance `lastModifiedDateTime`
        sur l'horloge du mock. Les dossiers intermédiaires sont créés.
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
                existant.modificateur = _identite(_depositaire(existant.nom))
                return existant, False
            morceaux = chemin.split("/")
            for profondeur in range(1, len(morceaux)):
                parent = "/".join(morceaux[:profondeur])
                if (dossier := self.trouver(parent)) is None:
                    instant = self.instant()
                    self._creer(parent, None, instant, instant, _identite("NTE"))
                elif not dossier.dossier:
                    raise NotADirectoryError(parent)
            instant = self.instant()
            auteur = _identite(_depositaire(morceaux[-1]))
            return self._creer(chemin, contenu, instant, instant, auteur), True

    def retirer(self, chemin: str) -> bool:
        """Supprime un fichier — ou un dossier et tout ce qu'il contient."""
        cle = _cle(chemin)
        with self.verrou:
            if not cle or cle not in self.elements:
                return False
            for autre in [c for c in self.elements if c == cle or c.startswith(cle + "/")]:
                del self.elements[autre]
            return True


depot = Depot()

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
    return f"{base}{jeu.CHEMIN_SITE}/_layouts/15/download.aspx?{parametres}"


# ── Le rendu ─────────────────────────────────────────────────────────────────


def _etiquette(element: Element, prefixe: str = "") -> str:
    """`"{GUID},n"` — et `"c:{GUID},n"` pour le cTag. Les guillemets font partie
    de la valeur, comme chez Graph."""
    return f'"{prefixe}{{{element.unique_id.upper()}}},{element.version}"'


def _reference_parent(element: Element) -> dict[str, Any]:
    reference: dict[str, Any] = {"driveType": "documentLibrary", "driveId": DRIVE_ID}
    if not element.chemin:
        return reference
    parent = depot.trouver(element.parent)
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

    `createdBy` et `lastModifiedBy` y portent une personne. Un consommateur qui
    n'a besoin que du nom, de la taille et du `cTag` doit le DIRE par
    `$select` ; sinon il collecte un nom et un courriel par fichier.
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
        "size": depot.taille(element),
    }
    if element.dossier:
        objet["folder"] = {"childCount": len(depot.enfants(element))}
        if not element.chemin:
            objet["root"] = {}
    else:
        # `quickXorHash` SEUL : SharePoint et OneDrive Entreprise ne calculent
        # ni sha1 ni sha256. Un consommateur qui compare des sha256 tombe sur
        # une clé absente — ici comme en production.
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


def _site() -> dict[str, Any]:
    return {
        "createdDateTime": _CREATION_SITE,
        "description": "Balances mensuelles et taux de change déposés par la Finance",
        "id": SITE_ID,
        "lastModifiedDateTime": max(e.modifie for e in depot.elements.values()),
        "name": jeu.CHEMIN_SITE.rpartition("/")[2],
        "webUrl": URL_SITE,
        "displayName": jeu.TITRE_SITE,
        "root": {},
        "siteCollection": {"hostname": jeu.HOTE},
    }


def _lecteur() -> dict[str, Any]:
    utilise = sum(len(e.contenu) for e in depot.elements.values() if e.contenu is not None)
    total = 27_487_790_694_400
    return {
        "createdDateTime": _CREATION_SITE,
        "description": "",
        "id": DRIVE_ID,
        "lastModifiedDateTime": max(e.modifie for e in depot.elements.values()),
        "name": jeu.BIBLIOTHEQUE,
        "webUrl": URL_BIBLIOTHEQUE,
        "driveType": "documentLibrary",
        "createdBy": {"user": {"displayName": "System Account"}},
        "owner": {
            "group": {
                "email": f"depot-finance@{jeu.DEPOSITAIRES['NTE'][1].partition('@')[2]}",
                "id": str(_guid("groupe-proprietaires")),
                "displayName": f"{jeu.TITRE_SITE} Owners",
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
    un 429 rejoué par le consommateur ne fausse pas le décompte de ses
    téléchargements."""
    if (refus := _refus(request)) is not None:
        return refus
    with depot.verrou:
        depot.compteurs[operation] += 1
    if depot.droit == "none":
        return _erreur_graph(403, "accessDenied", "Access denied")
    return None


def _introuvable() -> JSONResponse:
    return _erreur_graph(404, "itemNotFound", "The resource could not be found.")


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


@routeur.get("/v1.0/sites/{hostname}:/{server_relative_path:path}", responses=_REPONSES)
def site_par_chemin(hostname: str, server_relative_path: str, request: Request) -> Response:
    """Le site par son CHEMIN — `/sites/{hôte}:/{chemin relatif au serveur}`.

    C'est l'adressage qu'un consommateur configure (un hôte et un chemin
    lisibles), et il en tire l'identifiant COMPOSITE du site pour la suite.
    Le deux-points final optionnel (`…:/sites/depot-finance:`) est accepté,
    comme chez Graph.
    """
    if (refus := _prelude(request, "site")) is not None:
        return refus
    if hostname.casefold() != jeu.HOTE:
        return _erreur_graph(400, "invalidRequest", "Invalid hostname for this tenancy")
    if _cle(server_relative_path.rstrip(":")) != _cle(jeu.CHEMIN_SITE):
        return _erreur_graph(404, "itemNotFound", "Requested site could not be found")
    try:
        return JSONResponse(_projeter_un(_site(), request, frozenset(_site()), "sites/$entity"))
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "site")


@routeur.get("/v1.0/sites/{site_id}", responses=_REPONSES)
def site(site_id: str, request: Request) -> Response:
    """Le site par son identifiant composite."""
    if (refus := _prelude(request, "site")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _erreur_graph(404, "itemNotFound", "Requested site could not be found")
    try:
        return JSONResponse(_projeter_un(_site(), request, frozenset(_site()), "sites/$entity"))
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "site")


@routeur.get("/v1.0/sites/{site_id}/drive", responses=_REPONSES)
def lecteur_par_defaut(site_id: str, request: Request) -> Response:
    """La bibliothèque PAR DÉFAUT du site — « Documents »."""
    if (refus := _prelude(request, "drive")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _erreur_graph(404, "itemNotFound", "Requested site could not be found")
    try:
        return JSONResponse(
            _projeter_un(_lecteur(), request, frozenset(_lecteur()), "drives/$entity")
        )
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "drive")


@routeur.get("/v1.0/sites/{site_id}/drives", responses=_REPONSES)
def lecteurs(site_id: str, request: Request) -> Response:
    """Les bibliothèques du site, en collection `value[]`."""
    if (refus := _prelude(request, "drives")) is not None:
        return refus
    if site_id.casefold() != SITE_ID.casefold():
        return _erreur_graph(404, "itemNotFound", "Requested site could not be found")
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
    try:
        return JSONResponse(
            _projeter_un(_lecteur(), request, frozenset(_lecteur()), "drives/$entity")
        )
    except _SelectInconnu as exc:
        return _select_inconnu(exc, "drive")


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
    """Le `$skiptoken` — opaque pour le client, à la façon de SharePoint : la
    clé de tri du DERNIER élément rendu, pas un rang. Un fichier ajouté ou
    retiré pendant la pagination ne décale donc rien."""
    return _b64(
        f"Paged=TRUE&p_SortBehavior=0&p_FileLeafRef={quote(element.nom)}&p_ID={element.numero}".encode()
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

    with depot.verrou:
        enfants = depot.enfants(dossier)
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
    """`Insights360:/children` → (`Insights360`, `children`) ;
    `Insights360/a.xlsx:` → (`Insights360/a.xlsx`, ``)."""
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
    return _servir_element(request, depot.elements[""])


@routeur.get("/v1.0/drives/{drive_id}/root/children", responses=_REPONSES)
def enfants_de_la_racine(drive_id: str, request: Request) -> Response:
    """Les enfants de la racine — le dossier `Insights360`."""
    if (refus := _prelude(request, "children")) is not None:
        return refus
    if drive_id != DRIVE_ID:
        return _introuvable()
    return _servir_enfants(request, depot.elements[""])


@routeur.get(
    "/v1.0/drives/{drive_id}/root:/{chemin:path}",
    responses=_REPONSES_CONTENU,
)
def par_chemin(drive_id: str, chemin: str, request: Request) -> Response:
    """L'adressage PAR CHEMIN, sous ses trois formes :

    • `root:/Insights360:/children` — les enfants d'un dossier, PAGINÉS
      (`@odata.nextLink`, `$skiptoken` opaque), `$select` honoré ;
    • `root:/Insights360/balance_NTE_2026-01.xlsx` (deux-points final
      facultatif) — un élément ;
    • `root:/Insights360/balance_NTE_2026-01.xlsx:/content` — 302 vers l'URL
      de téléchargement.

    La casse du chemin est indifférente, comme dans SharePoint.
    """
    cible, segment = _decouper(chemin)
    servir = {"": _servir_element, "children": _servir_enfants, "content": _servir_contenu}
    if (
        refus := _prelude(request, segment if segment in servir and segment else "item")
    ) is not None:
        return refus
    if segment not in servir:
        return _erreur_graph(400, "BadRequest", f"Resource not found for the segment '{segment}'.")
    element = depot.trouver(cible) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else servir[segment](request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}", responses=_REPONSES)
def element_par_identifiant(drive_id: str, item_id: str, request: Request) -> Response:
    """Un élément par son identifiant `01…`."""
    if (refus := _prelude(request, "item")) is not None:
        return refus
    element = depot.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_element(request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}/children", responses=_REPONSES)
def enfants_par_identifiant(drive_id: str, item_id: str, request: Request) -> Response:
    """Les enfants d'un dossier désigné par son identifiant — même pagination."""
    if (refus := _prelude(request, "children")) is not None:
        return refus
    element = depot.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_enfants(request, element)


@routeur.get("/v1.0/drives/{drive_id}/items/{item_id}/content", responses=_REPONSES_CONTENU)
def contenu(drive_id: str, item_id: str, request: Request) -> Response:
    """Le contenu d'un fichier : 302 vers une URL PRÉ-AUTHENTIFIÉE.

    C'est ce que Graph rend, et il faut le suivre SANS le Bearer — l'URL porte
    sa propre autorisation (`tempauth`), limitée dans le temps.
    """
    if (refus := _prelude(request, "content")) is not None:
        return refus
    element = depot.par_identifiant(item_id) if drive_id == DRIVE_ID else None
    return _introuvable() if element is None else _servir_contenu(request, element)


# ── L'hôte SharePoint : le téléchargement ────────────────────────────────────


def _refus_sharepoint(description: str) -> JSONResponse:
    # Le STATUT est ce qui fait foi (401) ; le corps exact de SharePoint sur ces
    # refus n'a pas été relevé — un consommateur ne doit pas s'y adosser.
    return JSONResponse(
        status_code=401, content={"error": "invalid_request", "error_description": description}
    )


@routeur.get(f"{jeu.CHEMIN_SITE}/_layouts/15/download.aspx", include_in_schema=False)
def telecharger(request: Request) -> Response:
    """L'URL PRÉ-AUTHENTIFIÉE — hôte SharePoint, hors de Graph, hors contrat.

    ┌─ UN BEARER ICI EST UNE FAUTE ──────────────────────────────────────────┐
    │ L'autorisation est DANS l'URL (`tempauth`). Un client qui y renvoie le │
    │ jeton Graph expose ce jeton à un autre hôte — et se fait refuser : 401.│
    │ Un `tempauth` inconnu, falsifié ou expiré : 401 aussi.                 │
    └────────────────────────────────────────────────────────────────────────┘
    """
    if "authorization" in request.headers:
        return _refus_sharepoint(
            "Pre-authenticated download URL: the request must not carry an Authorization header."
        )
    unique_id = request.query_params.get("UniqueId", "").strip("{}").lower()
    if not tempauth_valide(request.query_params.get("tempauth", ""), unique_id):
        return _refus_sharepoint("Invalid or expired tempauth.")
    with depot.verrou:
        depot.compteurs["download"] += 1
        element = depot.par_unique_id(unique_id)
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


# ── Les variantes de fichiers, à déposer par le plan de contrôle ─────────────


@routeur.get("/__fixtures/depot_finance", include_in_schema=False)
def fixtures(request: Request) -> JSONResponse:
    """L'index des variantes — sans authentification : ce sont des données de
    TEST fictives, utiles sans le plan de contrôle (un test unitaire du
    consommateur peut les télécharger et les parser directement)."""
    base = str(request.base_url).rstrip("/")
    return JSONResponse(
        {
            "fixtures": [
                {
                    "name": f.nom,
                    "valid": f.valide,
                    "defect": f.defaut,
                    "url": f"{base}/__fixtures/depot_finance/{quote(f.nom)}",
                }
                for f in jeu.FIXTURES
            ]
        }
    )


@routeur.get("/__fixtures/depot_finance/{nom}", include_in_schema=False)
def fixture(nom: str) -> Response:
    donnees = jeu.fixture(nom)
    if donnees is None:
        return JSONResponse(status_code=404, content={"error": f"unknown fixture '{nom}'"})
    extension = "." + nom.rpartition(".")[2].lower()
    return Response(donnees, media_type=_TYPES_MIME.get(extension, "application/octet-stream"))


def nom_valide(chemin: str) -> bool:
    """Un chemin que SharePoint accepterait.

    Règles relevées le 2026-10-08 dans « Restrictions and limitations in
    OneDrive and SharePoint » (support.microsoft.com) : ni `" * : < > ? / \\ |`,
    ni espace en tête ou en fin, ni `_vti_` nulle part dans le nom.

    Exception ASSUMÉE : les noms en `~$…`, que la même page déclare invalides.
    La variante de verrou Excel doit pouvoir être déposée pour éprouver le
    filtre du consommateur — défense en profondeur s'il lit un jour un dossier
    synchronisé.
    """
    morceaux = [p for p in chemin.split("/") if p]
    return bool(morceaux) and all(
        not (_INTERDITS & set(p)) and p == p.strip() and "_vti_" not in p.lower() for p in morceaux
    )
