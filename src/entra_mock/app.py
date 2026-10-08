"""Mock Microsoft Graph — l'appartenance aux groupes Entra.

┌─ POURQUOI CE MOCK EST SÉPARÉ DE CELUI DE BOONDMANAGER ──────────────────────┐
│ La spec range « Entra group membership » sous §4.2, avec les ressources     │
│ BoondManager. C'est une erreur : ce n'est pas la même API.                  │
│                                                                             │
│                    BoondManager              Microsoft Graph                │
│   auth             JWT HS256 statique        OAuth2 client credentials      │
│                    en en-tête custom         → jeton Bearer expirant        │
│   enveloppe        data[] + meta.totals      value[] + @odata.nextLink      │
│   pagination       page / maxResults         curseur opaque dans une URL    │
│                                                                             │
│ Les mélanger dans un même client produirait un client qui ment sur ce qu'il │
│ parle. Cf. docs/SPEC-DEVIATIONS.md #4.                                      │
└─────────────────────────────────────────────────────────────────────────────┘

┌─ 0.3.0 — CE QUE LA SONDE A TROUVÉ, ET QUI EST CORRIGÉ ICI ──────────────────┐
│ `insights360:scripts/relever_dialecte_entra.py` confronte ce mock au vrai   │
│ service, le 2026-09-03, SANS aucun credential : le CSDL officiel de Graph   │
│ (`/v1.0/$metadata`) et les refus des deux hôtes suffisent. Cinq divergences │
│ en sont sorties, toutes du même genre — le mock était plus GENTIL que le    │
│ fournisseur, donc un client vert ici pouvait échouer là-bas :               │
│                                                                             │
│   1. l'autorité AAD rend une enveloppe OAuth2 PLATE et un code `AADSTS…` ;  │
│      le mock rendait l'enveloppe imbriquée de Graph, sans code ;            │
│   2. Graph loge `innerError.request-id` dans ses erreurs — l'identifiant    │
│      que le support Microsoft réclame ; le mock ne le rendait pas ;         │
│   3. `/v1.0/groups` PAGINE chez Microsoft ; le mock rendait tout d'un coup, │
│      ce qui rendait invisible un client ne suivant pas le curseur ;         │
│   4. sans `$select`, Graph rend le jeu de propriétés PAR DÉFAUT — pour      │
│      `user`, l'essentiel de 79 propriétés ; le mock en rendait quatre, donc │
│      une collecte non minimisée ne se voyait pas ;                          │
│   5. Graph ÉTRANGLE (429 + `Retry-After`) ; le mock ne l'a jamais fait.     │
│                                                                             │
│ Règle qui en découle, et qui vaut pour les quatre mocks de l'écosystème :   │
│ **un mock doit être au moins aussi SÉVÈRE que le fournisseur.** Un mock     │
│ permissif ne rend pas la CI plus verte, il la rend moins informative.       │
└─────────────────────────────────────────────────────────────────────────────┘

Le jeu de données est aligné sur le profil `insights360` de boondmanager-mock :
les mêmes UPN, et surtout les deux cas limites qui n'existent QUE dans la
jonction des deux sources —

  • `ext.consultant@boreal-conseil.example` : dans un groupe, ABSENT du SIRH ;
  • `kevin.silva@boreal-conseil.example`     : au SIRH, dans AUCUN groupe.

Depuis 0.5.0, le mock sert aussi la surface « fichiers » de Graph (driveItem :
un site, sa bibliothèque par défaut, VIDE au démarrage). Elle vit dans
`drive.py` et passe par le MÊME préambule (`_refus`) ; le plan de contrôle
`/__admin`, qui la remplit, vit dans `admin.py`.
"""

from __future__ import annotations

import base64
import json
import os
import time
import uuid
from datetime import UTC, datetime
from typing import Any

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

# Le domaine du jeu `realiste` de boondmanager-mock. L'UPN est la SEULE clé de
# jointure entre l'annuaire et le SIRH : un domaine qui diverge ne produit pas
# d'erreur, il produit zéro appartenance — donc zéro visibilité, donc des tests
# verts par vacuité.
UPN_DOMAIN = os.environ.get("ENTRA_MOCK_UPN_DOMAIN", "boreal-conseil.example")
TENANT_ID = os.environ.get("ENTRA_TENANT_ID", "00000000-0000-0000-0000-000000000000")
CLIENT_ID = os.environ.get("ENTRA_CLIENT_ID", "00000000-0000-0000-0000-000000000000")
CLIENT_SECRET = os.environ.get("ENTRA_CLIENT_SECRET", "change-me-entra")

# Taille de page volontairement MINIMALE : Graph pagine avec un curseur opaque,
# et un pipeline qui ne suivrait pas `@odata.nextLink` ne verrait que les
# premiers membres — en silence, sans la moindre erreur. Avec une page de 1,
# TOUT groupe de plus d'un membre est paginé : le chemin de pagination est donc
# exercé par construction, pas par chance.
PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_PAGE_SIZE", "1"))

# ┌─ ET `/v1.0/groups` PAGINE AUSSI ───────────────────────────────────────────┐
# │ Microsoft y sert 100 groupes par page. Le mock en rendait quatre d'un      │
# │ coup : un client qui suivait le curseur sur `/members` mais pas sur        │
# │ `/groups` — le cas réel, trouvé dans insights360 le 2026-09-03 — était     │
# │ vert en CI et tronquait au 101ᵉ groupe d'un annuaire réel. Sans erreur :   │
# │ les membres des groupes non listés n'obtenaient simplement aucun           │
# │ périmètre.                                                                 │
# │                                                                            │
# │ Deux par page, donc, pour la même raison qu'un membre par page : que le    │
# │ chemin soit exercé par CONSTRUCTION, pas par chance.                       │
# └────────────────────────────────────────────────────────────────────────────┘
GROUPS_PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_GROUPS_PAGE_SIZE", "2"))

# ┌─ L'ÉTRANGLEMENT, SUR DEMANDE ──────────────────────────────────────────────┐
# │ Graph étrangle par locataire ET par application. Le profil d'appels d'une  │
# │ extraction d'appartenance — une liste, puis un appel paginé PAR GROUPE —   │
# │ est exactement celui qui le déclenche : sur un annuaire réel, le 429 est   │
# │ le régime nominal, pas le cas rare.                                        │
# │                                                                            │
# │ Éteint par défaut (0) : un mock qui étrangle sans qu'on le lui demande     │
# │ rendrait la CI intermittente. Les tests de reprise le posent explicitement.│
# └────────────────────────────────────────────────────────────────────────────┘
THROTTLE_EVERY = int(os.environ.get("ENTRA_MOCK_THROTTLE_EVERY", "0"))
RETRY_AFTER = os.environ.get("ENTRA_MOCK_RETRY_AFTER", "1")

# Plan de contrôle /__admin. Fermé par défaut : il n'a de sens qu'en test — et
# fermé veut dire NON MONTÉ.
ADMIN_ENABLED = os.environ.get("ENTRA_MOCK_ADMIN_ENABLED", "false").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}


def _d(upn: str) -> str:
    return upn.format(d=UPN_DOMAIN)


def _guid(graine: str) -> str:
    """Un GUID STABLE dérivé d'une chaîne.

    Les objets d'annuaire réels sont identifiés par un GUID, pas par une clé
    lisible : le mock rendait `user-<upn>`, ce qui donnait à un consommateur
    l'illusion que l'identifiant portait l'UPN. Déterministe pour que le jeu de
    données reste reproductible d'un run à l'autre.
    """
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, graine))


# ── L'appartenance aux groupes ───────────────────────────────────────────────
#
# Les quatre groupes correspondent aux quatre règles d'autorisation :
#   grp-bi-rh / grp-bi-sales / grp-bi-direction → règle 3 (portée du groupe)
#   grp-bi-comex                                → règle 4 (visibilité totale)
#
# ⚠️ `grp-bi-comex`, et non `grp-comex` — renommé le 2026-09-15. Le locataire
#    réel porte les DEUX : l'ancien est vide et mort, le vrai groupe des
#    associés est `grp-bi-comex`, et c'est celui sur lequel le consommateur
#    (insights360:int_acl_regles.sql) keye désormais sa règle 4.
#
#    Ce que ce mock a coûté en restant sur l'ancien nom : côté consommateur,
#    la règle 4 ne correspondait plus à personne dans le jeu de test, et la
#    suite restait VERTE — `arthur.ivanov` est aussi membre de
#    `grp-bi-direction`, qui lui donnait assez de visibilité pour que les
#    assertions tiennent. Une règle d'autorisation entière avait cessé d'être
#    éprouvée sans qu'un seul test ne bouge. C'est précisément le mode de
#    panne qu'un mock existe pour empêcher.
GROUPES: dict[str, dict[str, Any]] = {
    "11111111-0000-0000-0000-000000000001": {
        "displayName": "grp-bi-rh",
        "membres": [
            # Manager de niveau 2, agence Paris. Sa portée de groupe (règle 3)
            # lui donne les rattachés de son périmètre.
            _d("yuki.lambert@{d}"),
            # Membre d'un groupe, INCONNU du SIRH. Le pipeline doit le laisser
            # tomber à la jointure sur dim_collaborateur, jamais produire une
            # ligne à clé inconnue — ce qui ferait dégénérer le prédicat interne
            # en « pas de filtre ».
            _d("ext.consultant@{d}"),
        ],
    },
    "11111111-0000-0000-0000-000000000002": {
        "displayName": "grp-bi-sales",
        "membres": [_d("lars.marchand@{d}")],
    },
    "11111111-0000-0000-0000-000000000003": {
        "displayName": "grp-bi-direction",
        "membres": [_d("arthur.ivanov@{d}")],
    },
    "11111111-0000-0000-0000-000000000004": {
        "displayName": "grp-bi-comex",
        # Le sommet de la hiérarchie, seul sans manager. Le Comex lui donne
        # TOUS les collaborateurs, Nantes compris — alors que le périmètre RLS
        # externe de bi_rh et bi_sales exclut Nantes. C'est ce couple qui fait
        # travailler l'invariant `inner ⊆ outer` : il tient PARCE QUE le RLS
        # rogne, et échouerait bruyamment si la requête interne était un jour
        # exécutée hors du rôle. Sans lui, l'invariant serait vrai par vacuité.
        "membres": [_d("arthur.ivanov@{d}")],
    },
    # ┌─ LE GROUPE IMBRIQUÉ — UNE PERTE DE DROITS QUI NE LEVAIT AUCUNE ERREUR ─┐
    # │ `/members` rend les membres DIRECTS. Un groupe qui en contient un      │
    # │ autre y apparaît comme un objet `#microsoft.graph.group`, et les       │
    # │ personnes qui sont dedans n'apparaissent NULLE PART — alors qu'Entra   │
    # │ leur accorde bien l'appartenance effective.                            │
    # │                                                                        │
    # │ Le mock ne rendait que des `user` : ce cas n'existait pas en CI, et un │
    # │ consommateur le confondait avec un principal de service — les deux     │
    # │ n'ont pas d'UPN, les deux tombaient dans le même filtre. L'un est      │
    # │ correct, l'autre fait disparaître des ayants droit.                    │
    # │                                                                        │
    # │ Ce groupe existe pour que le consommateur ait à distinguer les deux.   │
    # │ Il ne produit AUCUNE ligne d'appartenance : il n'ajoute donc rien au   │
    # │ modèle d'autorisation aval, il en éprouve seulement le diagnostic.     │
    # └────────────────────────────────────────────────────────────────────────┘
    "11111111-0000-0000-0000-000000000005": {
        "displayName": "grp-bi-imbrique",
        "membres": [
            {"type": "#microsoft.graph.group", "id": "11111111-0000-0000-0000-000000000002"},
            # Sans UPN par nature, et parfaitement attendu : c'est le cas qu'un
            # consommateur DOIT ignorer en silence.
            {"type": "#microsoft.graph.servicePrincipal", "id": _guid("sp-powerbi")},
        ],
    },
}

# `kevin.silva@…` n'apparaît DÉLIBÉRÉMENT nulle part : présent au SIRH, membre
# d'aucun groupe. Il ne doit voir que lui-même (règle 1).
#
# ⚠️ Ces UPN sont ceux du jeu `realiste` de boondmanager-mock. Ils ne sont PAS
#    arbitraires : chacun est choisi pour la propriété qu'il rend testable
#    (sommet sans manager, agence hors périmètre, membre hors SIRH). Un
#    changement de jeu de données côté mock BoondManager casse ce fichier — et
#    c'est voulu : mieux vaut un test rouge qu'un modèle d'autorisation dont les
#    cas limites ont silencieusement cessé d'exister.


# ┌─ LE JEU PAR DÉFAUT EST LARGE, ET C'EST LE POINT ───────────────────────────┐
# │ Sans `$select`, Graph rend le jeu de propriétés PAR DÉFAUT du type. Le     │
# │ CSDL officiel en déclare 79 sur `user` ; le mock en rendait quatre.        │
# │                                                                            │
# │ Conséquence : un client qui OUBLIAIT `$select` collectait, en production,  │
# │ le poste, le téléphone mobile, le service, le matricule — et rien, en CI,  │
# │ ne le lui disait. Or c'est la COLLECTE qui fonde l'obligation, pas le      │
# │ stockage : le tri à l'arrivée arrive trop tard.                            │
# │                                                                            │
# │ Le mock rend donc un jeu par défaut LARGE — un échantillon représentatif   │
# │ des propriétés sensibles du type — et honore `$select` quand on le lui     │
# │ demande. Oublier `$select` devient visible ici, gratuitement.              │
# └────────────────────────────────────────────────────────────────────────────┘
def _membre_utilisateur(upn: str) -> dict[str, Any]:
    """Le jeu par défaut d'un `user`, dans l'esprit de ce que rend Graph.

    Les valeurs sont fictives — c'est un mock — mais les NOMS sont ceux du CSDL
    officiel, et ce sont eux qui décident si un `$select` est correct.
    """
    prenom, _, nom = upn.split("@", maxsplit=1)[0].partition(".")
    return {
        "@odata.type": "#microsoft.graph.user",
        "id": _guid(upn),
        "userPrincipalName": upn,
        "mail": upn,
        "displayName": f"{prenom.capitalize()} {nom.capitalize()}",
        "givenName": prenom.capitalize(),
        "surname": nom.capitalize(),
        "jobTitle": "Consultant",
        "department": "Delivery",
        "officeLocation": "Paris",
        "mobilePhone": "+33 6 00 00 00 00",
        "businessPhones": ["+33 1 00 00 00 00"],
        "employeeId": "E-0000",
        "accountEnabled": True,
        "userType": "Member",
    }


def _membre_autre(brut: dict[str, Any]) -> dict[str, Any]:
    """Un objet d'annuaire qui n'est PAS une personne — donc sans UPN."""
    objet: dict[str, Any] = {"@odata.type": brut["type"], "id": brut["id"]}
    if brut["type"] == "#microsoft.graph.group":
        objet["displayName"] = GROUPES[brut["id"]]["displayName"]
    else:
        objet["displayName"] = "Power BI Service"
    return objet


def _membres_transitifs(group_id: str, vus: set[str] | None = None) -> list[dict[str, Any]]:
    """L'appartenance EFFECTIVE — les groupes imbriqués aplatis.

    ┌─ POURQUOI CET ENDPOINT EXISTE, MESURÉ CHEZ UN VRAI FOURNISSEUR ────────┐
    │ `/members` ne rend que les membres DIRECTS. Sondé le 2026-09-03 contre │
    │ un annuaire d'entreprise RÉEL : un groupe transverse y rend UNE         │
    │ personne en direct et une TRENTAINE en transitif — il contient dix-neuf │
    │ groupes imbriqués. Quatre groupes de direction régionaux en perdent     │
    │ trois chacun.                                                            │
    │                                                                         │
    │ Autrement dit, un consommateur qui lit `/members` pour construire une   │
    │ ACL prive 97 % des ayants droit de leur périmètre sur ce groupe-là —    │
    │ sans une erreur, sans un avertissement.                                 │
    │                                                                         │
    │ Le mock ne servait pas cet endpoint : le consommateur ne POUVAIT donc   │
    │ pas se corriger sans casser sa stack locale. C'est une absence qui      │
    │ enferme dans le défaut, pas seulement qui le masque.                    │
    └─────────────────────────────────────────────────────────────────────────┘

    Graph rend les groupes imbriqués EN PLUS de leurs membres : l'objet groupe
    reste dans la collection. On fait pareil — un consommateur qui filtre sur
    l'UPN les écarte de toute façon, et les omettre cacherait la structure.

    `vus` coupe les cycles : Entra autorise A ∋ B ∋ A, et une récursion naïve
    y tournerait indéfiniment.
    """
    vus = set() if vus is None else vus
    if group_id in vus:
        return []
    vus.add(group_id)

    rendus: list[dict[str, Any]] = []
    for membre in GROUPES[group_id]["membres"]:
        if isinstance(membre, dict):
            rendus.append(_membre_autre(membre))
            if membre["type"] == "#microsoft.graph.group" and membre["id"] in GROUPES:
                rendus.extend(_membres_transitifs(membre["id"], vus))
        else:
            rendus.append(_membre_utilisateur(membre))

    # Dédoublonnage : une personne membre de deux sous-groupes n'apparaît
    # qu'une fois, comme chez Graph.
    uniques: dict[str, dict[str, Any]] = {}
    for objet in rendus:
        uniques.setdefault(objet["id"], objet)
    return list(uniques.values())


def _membres_rendus(group_id: str) -> list[dict[str, Any]]:
    return [
        _membre_autre(m) if isinstance(m, dict) else _membre_utilisateur(m)
        for m in GROUPES[group_id]["membres"]
    ]


def _groupe_rendu(group_id: str) -> dict[str, Any]:
    """Le jeu par défaut d'un `group` — large, pour la même raison."""
    groupe = GROUPES[group_id]
    return {
        "@odata.type": "#microsoft.graph.group",
        "id": group_id,
        "displayName": groupe["displayName"],
        "description": f"Groupe {groupe['displayName']}",
        "mail": None,
        "mailEnabled": False,
        "mailNickname": groupe["displayName"],
        "securityEnabled": True,
        "groupTypes": [],
        "createdDateTime": "2024-01-01T00:00:00Z",
        "visibility": "Private",
    }


app = FastAPI(title="Microsoft Graph mock (Entra groups, drive files)", version="0.5.0")

#: Compteur d'appels Graph authentifiés — sert UNIQUEMENT à l'étranglement
#: déterministe. Un mock mono-processus : pas de course à craindre.
_appels = 0


def reinitialiser_appels() -> None:
    """Remet la cadence de l'étranglement à zéro — `POST /__admin/reset`."""
    global _appels  # noqa: PLW0603 — un compteur de mock, mono-processus
    _appels = 0


def _erreur_graph(status: int, code: str, message: str) -> JSONResponse:
    """L'enveloppe d'erreur de l'API de RESSOURCES — imbriquée, avec innerError.

    Relevé contre `graph.microsoft.com` le 2026-09-03 : `error.innerError` porte
    `date`, `request-id` et `client-request-id`. Ce `request-id` est
    l'identifiant que le support Microsoft réclame sur un ticket ; un mock qui
    ne le rendait pas laissait écrire des diagnostics qui le perdaient.
    """
    identifiant = str(uuid.uuid4())
    return JSONResponse(
        status_code=status,
        headers={"request-id": identifiant, "client-request-id": identifiant},
        content={
            "error": {
                "code": code,
                "message": message,
                "innerError": {
                    "date": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S"),
                    "request-id": identifiant,
                    "client-request-id": identifiant,
                },
            }
        },
    )


def _erreur_aad(status: int, code: str, aadsts: int, message: str) -> JSONResponse:
    """L'enveloppe de l'AUTORITÉ — OAuth2, PLATE, avec un code AADSTS.

    ┌─ LES DEUX HÔTES NE PARLENT PAS LE MÊME DIALECTE D'ERREUR ──────────────┐
    │ Relevé le 2026-09-03 contre `login.microsoftonline.com` :              │
    │   {"error": "invalid_client", "error_description": "AADSTS7000215: …", │
    │    "error_codes": [7000215], "timestamp", "trace_id", "correlation_id"}│
    │                                                                        │
    │ Rien à voir avec `{"error": {"code", "message"}}` de Graph. Le mock     │
    │ rendait l'enveloppe de Graph SUR LE CHEMIN DU JETON : un diagnostic     │
    │ écrit contre lui lisait `error.code`, et trouvait `None` en production. │
    │                                                                        │
    │ Le code AADSTS est ce qui rend une panne d'authentification            │
    │ diagnosticable — 7000215 (secret faux) et 700038 (application inconnue) │
    │ mènent à deux gestes différents. Le perdre, c'est ne plus savoir lequel.│
    │                                                                        │
    │ Le STATUT n'est pas ce qui distingue les deux hôtes : l'autorité rend   │
    │ 401 sur un secret faux et 400 sur un grant_type inconnu. C'est          │
    │ l'ENVELOPPE. Les statuts ci-dessous sont ceux qui ont été observés.     │
    └────────────────────────────────────────────────────────────────────────┘
    """
    trace = str(uuid.uuid4())
    return JSONResponse(
        status_code=status,
        content={
            "error": code,
            "error_description": (
                f"AADSTS{aadsts}: {message} "
                f"Trace ID: {trace} Correlation ID: {trace} "
                f"Timestamp: {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%SZ')}"
            ),
            "error_codes": [aadsts],
            "timestamp": datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%SZ"),
            "trace_id": trace,
            "correlation_id": trace,
            "error_uri": f"https://login.microsoftonline.com/error?code={aadsts}",
        },
    )


def _bearer_valide(request: Request) -> bool:
    header = request.headers.get("Authorization", "")
    if not header.startswith("Bearer "):
        return False
    jeton = header.removeprefix("Bearer ")
    try:
        charge = json.loads(base64.urlsafe_b64decode(jeton + "=" * (-len(jeton) % 4)))
    except Exception:
        return False
    # L'expiration est vérifiée pour de bon : c'est ce qui permet de tester le
    # renouvellement côté client, la principale différence opérationnelle avec
    # le credential statique de BoondManager.
    return bool(charge.get("aud") == CLIENT_ID and charge.get("exp", 0) > time.time())


def _refus(request: Request) -> JSONResponse | None:
    """Le préambule commun à tout appel Graph : jeton, puis étranglement."""
    if not _bearer_valide(request):
        return _erreur_graph(401, "InvalidAuthenticationToken", "Access token is empty or expired.")
    global _appels  # noqa: PLW0603 — un compteur de mock, mono-processus
    _appels += 1
    if THROTTLE_EVERY and _appels % THROTTLE_EVERY == 0:
        reponse = _erreur_graph(
            429, "TooManyRequests", "Too many requests. Please retry after some time."
        )
        reponse.headers["Retry-After"] = RETRY_AFTER
        return reponse
    return None


def _projeter(
    objets: list[dict[str, Any]], select: str | None, univers: frozenset[str]
) -> list[dict[str, Any]]:
    """Applique `$select` — et refuse une propriété que le TYPE ne porte pas.

    Refuser est le point : Graph rend 400 sur un `$select` inconnu, et le fait
    pour la COLLECTION ENTIÈRE. Un mock qui ignorerait le paramètre laisserait
    passer une faute de frappe jusqu'au premier run réel, où elle n'emporterait
    pas une ligne mais toute la source.

    ┌─ CONTRE LE TYPE, PAS CONTRE LA PAGE ───────────────────────────────────┐
    │ Première version : l'univers des propriétés valides était déduit des    │
    │ objets de la page en cours. Faux, et attrapé au premier run bout-en-bout │
    │ d'insights360 — `$select=id,userPrincipalName` sur un groupe dont AUCUN  │
    │ membre n'est une personne rendait 400, alors que Graph l'accepte : il    │
    │ valide contre le TYPE de la collection, indépendamment de ce que la      │
    │ page contient.                                                          │
    │                                                                          │
    │ Un mock plus sévère que le fournisseur est un défaut au même titre qu'un │
    │ mock plus permissif : il fait échouer du code correct, ce qui pousse à   │
    │ écrire du code faux pour le contenter.                                   │
    └──────────────────────────────────────────────────────────────────────────┘

    Les annotations OData (`@odata.type`) ne se sélectionnent pas et sont
    toujours rendues : Graph les traite comme des métadonnées, pas comme des
    propriétés. Une propriété valide mais absente de l'objet est simplement
    omise — Graph ne rend pas de clé nulle pour un champ qu'un type dérivé ne
    porte pas.
    """
    if not select:
        return objets
    demandes = [c.strip() for c in select.split(",") if c.strip()]
    if inconnues := [c for c in demandes if c not in univers]:
        raise _SelectInconnu(inconnues)
    return [
        {c: v for c, v in o.items() if c in demandes or c.startswith("@odata.type")} for o in objets
    ]


#: Les propriétés que chaque collection peut porter — l'univers contre lequel
#: `$select` est validé. Union des types dérivés pour `/members` : la collection
#: est déclarée `Collection(graph.directoryObject)`, et Graph accepte d'y
#: sélectionner une propriété de `user` même si la page n'en contient aucun.
PROPRIETES_GROUPE = frozenset(_groupe_rendu(next(iter(GROUPES))))
PROPRIETES_MEMBRE = frozenset(
    set(_membre_utilisateur("prenom.nom@exemple")) | {"id", "displayName", "deletedDateTime"}
)


class _SelectInconnu(Exception):
    def __init__(self, champs: list[str]) -> None:
        self.champs = champs


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "entra-mock"}


@app.post("/{tenant}/oauth2/v2.0/token")
async def token(tenant: str, request: Request) -> JSONResponse:
    """Flux client credentials — le vrai chemin d'authentification de Graph.

    Les quatre refus reproduisent des réponses OBSERVÉES le 2026-09-03 : statut,
    code OAuth2 et code AADSTS. Aucun n'est inventé.
    """
    formulaire = await request.form()
    if tenant != TENANT_ID:
        # AADSTS900021, observé sur un GUID de locataire inexistant.
        return _erreur_aad(
            400, "invalid_request", 900021, f"Requested tenant identifier '{tenant}' is not valid."
        )
    if formulaire.get("grant_type") is None:
        # AADSTS900144, observé sur un corps JSON — que FastAPI ne sait pas lire
        # comme un formulaire, exactement comme l'autorité ne le lit pas.
        return _erreur_aad(
            400,
            "invalid_request",
            900144,
            "The request body must contain the following parameter: 'grant_type'.",
        )
    if formulaire.get("grant_type") != "client_credentials":
        # AADSTS70003, observé.
        return _erreur_aad(
            400,
            "unsupported_grant_type",
            70003,
            f"The app requested an unsupported grant type '{formulaire.get('grant_type')}'.",
        )
    if formulaire.get("client_id") != CLIENT_ID:
        # AADSTS700038, observé sur une application inconnue du locataire — 400,
        # et NON 401 : ce n'est pas le secret qui est faux, c'est l'application
        # qui n'existe pas. Les deux pannes appellent deux gestes différents.
        return _erreur_aad(
            400,
            "unauthorized_client",
            700038,
            f"{formulaire.get('client_id')} is not a valid client id for this tenant.",
        )
    if formulaire.get("client_secret") != CLIENT_SECRET:
        # AADSTS7000215, observé — et un 401, contrairement aux trois autres.
        return _erreur_aad(
            401,
            "invalid_client",
            7000215,
            "Invalid client secret provided. Ensure the secret being sent in the request "
            "is the client secret value, not the client secret ID.",
        )

    duree = int(os.environ.get("ENTRA_MOCK_TOKEN_TTL", "3600"))
    charge = json.dumps({"aud": CLIENT_ID, "exp": int(time.time()) + duree}).encode()
    jeton = base64.urlsafe_b64encode(charge).rstrip(b"=").decode()
    return JSONResponse(
        {
            "token_type": "Bearer",
            "expires_in": duree,
            # Rendu par l'autorité réelle à côté d'`expires_in` : un client qui
            # le lirait ne doit pas trouver un mock qui l'ignore.
            "ext_expires_in": duree,
            "access_token": jeton,
        }
    )


@app.get("/v1.0/groups")
def groupes(request: Request) -> JSONResponse:
    """Les groupes, PAGINÉS — comme Graph, qui en sert 100 par page.

    Un client qui suivrait `@odata.nextLink` sur `/members` mais pas ici ne
    verrait que les `GROUPS_PAGE_SIZE` premiers groupes, sans la moindre erreur.
    """
    if (refus := _refus(request)) is not None:
        return refus

    depuis = int(request.query_params.get("$skiptoken", "0"))
    identifiants = list(GROUPES)[depuis : depuis + GROUPS_PAGE_SIZE]
    try:
        valeurs = _projeter(
            [_groupe_rendu(gid) for gid in identifiants],
            request.query_params.get("$select"),
            PROPRIETES_GROUPE,
        )
    except _SelectInconnu as exc:
        return _erreur_graph(
            400,
            "Request_UnsupportedQuery",
            f"Could not find a property named '{exc.champs[0]}' on type 'microsoft.graph.group'.",
        )

    corps: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
        "value": valeurs,
    }
    suivant = depuis + GROUPS_PAGE_SIZE
    if suivant < len(GROUPES):
        corps["@odata.nextLink"] = _lien_suivant(request, "/v1.0/groups", suivant)
    return JSONResponse(corps)


@app.get("/v1.0/groups/{group_id}/transitiveMembers")
def membres_transitifs(group_id: str, request: Request) -> JSONResponse:
    """L'appartenance EFFECTIVE, groupes imbriqués aplatis — comme Graph.

    Même enveloppe, même curseur opaque, même `$select` que `/members` : seule
    la composition de la collection change. C'est ce qui permet à un
    consommateur de basculer d'un chemin à l'autre sans réécrire sa pagination.
    """
    return _servir_membres(group_id, request, transitif=True)


@app.get("/v1.0/groups/{group_id}/members")
def membres(group_id: str, request: Request) -> JSONResponse:
    """Membres d'un groupe, paginés par curseur opaque — comme Graph.

    Le curseur est ici l'index de départ encodé dans `@odata.nextLink`. Un
    pipeline qui ne suivrait pas ce lien ne verrait que les PAGE_SIZE premiers
    membres, sans la moindre erreur. C'est pourquoi la page est petite.

    Les membres ne sont PAS tous des personnes : `members` est déclaré
    `Collection(graph.directoryObject)` dans le CSDL officiel, et rend aussi
    bien un groupe imbriqué qu'un principal de service. Les deux sont sans UPN,
    et un consommateur doit les traiter différemment.
    """
    return _servir_membres(group_id, request, transitif=False)


def _servir_membres(group_id: str, request: Request, *, transitif: bool) -> JSONResponse:
    """Le corps commun aux deux collections de membres."""
    if (refus := _refus(request)) is not None:
        return refus
    if group_id not in GROUPES:
        return _erreur_graph(404, "Request_ResourceNotFound", f"group {group_id} does not exist")

    depuis = int(request.query_params.get("$skiptoken", "0"))
    tous = _membres_transitifs(group_id) if transitif else _membres_rendus(group_id)
    chemin = "transitiveMembers" if transitif else "members"
    try:
        valeurs = _projeter(
            tous[depuis : depuis + PAGE_SIZE],
            request.query_params.get("$select"),
            PROPRIETES_MEMBRE,
        )
    except _SelectInconnu as exc:
        return _erreur_graph(
            400,
            "Request_UnsupportedQuery",
            f"Could not find a property named '{exc.champs[0]}' on type 'microsoft.graph.user'.",
        )

    corps: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects",
        "value": valeurs,
    }
    suivant = depuis + PAGE_SIZE
    if suivant < len(tous):
        corps["@odata.nextLink"] = _lien_suivant(
            request, f"/v1.0/groups/{group_id}/{chemin}", suivant
        )
    return JSONResponse(corps)


def _lien_suivant(request: Request, chemin: str, curseur: int | str) -> str:
    """Le lien suivant, en RECOPIANT les paramètres — comme Graph.

    Graph reporte `$select` (et les autres options) dans le `@odata.nextLink`.
    Un mock qui ne le ferait pas rendrait la deuxième page plus large que la
    première, et masquerait un client qui repasse ses paramètres à la main —
    lequel écraserait alors le curseur et boucherait sur la première page.

    Le curseur est un rang pour les groupes, un jeton opaque pour un dossier
    (`drive.py`) ; `$top`, quand il a été donné, suit le même chemin.
    """
    base = str(request.base_url).rstrip("/")
    parametres = [f"$skiptoken={curseur}"]
    for option in ("$select", "$top"):
        if valeur := request.query_params.get(option):
            parametres.append(f"{option}={valeur}")
    return f"{base}{chemin}?" + "&".join(parametres)


# ── La surface « fichiers », et le plan de contrôle ──────────────────────────
#
# Importés EN FIN de module, et c'est voulu : `drive.py` réutilise `_refus`,
# `_erreur_graph`, `_projeter` et `_lien_suivant` ci-dessus — le même jeton, le
# même étranglement, la même enveloppe d'erreur que l'annuaire. Les dupliquer
# laisserait les deux surfaces diverger au premier correctif.
from .drive import routeur as routeur_drive  # noqa: E402

app.include_router(routeur_drive)

# Le plan de contrôle n'est pas « monté puis interdit » : quand il est
# désactivé, la surface n'existe pas.
if ADMIN_ENABLED:
    from .admin import router as routeur_admin

    app.include_router(routeur_admin)
