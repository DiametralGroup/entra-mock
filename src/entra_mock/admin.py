"""Le plan de contrôle `/__admin` — hors de la surface Graph.

  • le routeur n'est MONTÉ que si `ENTRA_MOCK_ADMIN_ENABLED` est vrai — pas
    « monté puis interdit » : absent. Impossible donc de le laisser ouvert par
    accident ;
  • chaque appel présente `X-Mock-Admin-Token` (`ENTRA_MOCK_ADMIN_TOKEN`) ;
  • le préfixe `/__admin` ne peut collisionner avec aucun chemin Graph, qui
    vivent tous sous `/v1.0`, ni avec l'hôte du site (`/sites/…`).

Il existe parce que le mock tourne en CONTENEUR chez son client : un test ne
peut pas muter l'état en Python, il lui faut du HTTP — pour écrire un fichier,
en écraser un, en retirer un, ou compter ses téléchargements. Le lecteur est
VIDE au démarrage : c'est ici qu'un test le remplit.
"""

from __future__ import annotations

import hmac
import os
from typing import Any, Literal

from fastapi import APIRouter, Header, Request, Response
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from .app import _erreur_graph, reinitialiser_appels
from .drive import _rendu, lecteur_mock, nom_valide

ADMIN_TOKEN = os.environ.get("ENTRA_MOCK_ADMIN_TOKEN", "mock-admin-token")

router: APIRouter = APIRouter(
    prefix="/__admin", include_in_schema=False, tags=["mock control plane"]
)


def _refuse(jeton: str | None) -> JSONResponse | None:
    if not jeton or not hmac.compare_digest(jeton.encode(), ADMIN_TOKEN.encode()):
        return _erreur_graph(403, "Forbidden", "invalid or missing X-Mock-Admin-Token")
    return None


@router.post("/reset")
def reset(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Tout le mock remis à neuf : le lecteur (vide), ses compteurs, le droit
    sur le site, et le compteur d'appels qui cadence l'étranglement."""
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    lecteur_mock.reinitialiser()
    reinitialiser_appels()
    return JSONResponse({"status": "reset"})


@router.put("/drive/files/{chemin:path}")
async def ecrire(
    chemin: str, request: Request, x_mock_admin_token: str | None = Header(default=None)
) -> Response:
    """Écrit un fichier — le corps BRUT est son contenu.

    Le chemin part de la RACINE du lecteur : `reports/report-1.txt`. Les
    dossiers qui n'existent pas encore sont créés au passage.
    201 à la création, 200 à l'écrasement — qui garde l'identifiant, passe le
    `n` de `eTag`/`cTag` à n+1 et avance `lastModifiedDateTime` d'une minute
    sur l'horloge du mock (jamais l'horloge murale : cf. `drive.py`).
    """
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    if not nom_valide(chemin):
        return _erreur_graph(400, "invalidRequest", f"Invalid file name: '{chemin}'")
    try:
        element, cree = lecteur_mock.ecrire(chemin, await request.body())
    except (IsADirectoryError, NotADirectoryError) as exc:
        return _erreur_graph(409, "nameAlreadyExists", f"A folder is in the way: '{exc}'")
    return JSONResponse(_rendu(element, request), status_code=201 if cree else 200)


@router.delete("/drive/files/{chemin:path}")
def retirer(chemin: str, x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Retire un fichier — ou un dossier et tout son contenu. 204, ou 404."""
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    if not lecteur_mock.retirer(chemin):
        return _erreur_graph(404, "itemNotFound", "The resource could not be found.")
    return Response(status_code=204)


@router.post("/drive/reset")
def reset_lecteur(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Le lecteur seul redevient VIDE — compteurs et horloge compris."""
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    lecteur_mock.reinitialiser()
    return JSONResponse({"status": "reset"})


@router.get("/drive/counters")
def compteurs(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Les appels par opération depuis la dernière remise à zéro.

    Seuls comptent les appels AUTHENTIFIÉS et NON ÉTRANGLÉS : un 429 rejoué ne
    gonfle pas `content`. C'est ce qui permet de PROUVER qu'un second run
    idempotent ne retélécharge rien (`content` et `download` inchangés).
    """
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    operations = ("site", "drive", "drives", "item", "children", "content", "download")
    return JSONResponse({op: lecteur_mock.compteurs[op] for op in operations})


class DemandeDroit(BaseModel):
    grant: Literal["none", "read"]


@router.post("/drive/grant")
def droit(demande: DemandeDroit, x_mock_admin_token: str | None = Header(default=None)) -> Any:
    """Change le rôle de l'application sur le site, à chaud.

    `none` reproduit l'application `Sites.Selected` à qui personne n'a encore
    accordé le site : 403 `accessDenied` partout. Le pendant runtime de
    `ENTRA_MOCK_SITE_GRANT` — un conteneur de CI ne se redémarre pas entre
    deux tests.
    """
    if (refus := _refuse(x_mock_admin_token)) is not None:
        return refus
    lecteur_mock.droit = demande.grant
    return {"status": "updated", "grant": lecteur_mock.droit}
