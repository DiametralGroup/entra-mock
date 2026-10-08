"""Harnais de tests — l'application montée en process.

La propriété qu'il ne faut pas perdre : l'application que la stack interroge
EST celle que les tests exercent.

Le plan de contrôle `/__admin` est monté — le montage est conditionnel et la
configuration est lue à l'import, d'où le réglage posé AVANT l'import du paquet.
"""

from __future__ import annotations

import base64
import json
import os
import time
from importlib import import_module

os.environ.setdefault("ENTRA_MOCK_ADMIN_ENABLED", "true")

import pytest
from fastapi.testclient import TestClient

import entra_mock as mock

# `entra_mock.app` est l'INSTANCE FastAPI après le `from .app import app` du
# paquet : importer le module par son nom est le seul accès non ambigu à ses
# constantes (taille de page, étranglement) que les tests doivent piloter.
app_module = import_module("entra_mock.app")
drive_module = import_module("entra_mock.drive")

TENANT = "00000000-0000-0000-0000-000000000000"
CLIENT_ID = "00000000-0000-0000-0000-000000000000"
CLIENT_SECRET = "change-me-entra"
ADMIN = {"X-Mock-Admin-Token": "mock-admin-token"}


@pytest.fixture()
def client():
    return TestClient(mock.app)


@pytest.fixture()
def jeton(client) -> str:
    """Un jeton obtenu par le VRAI flux — jamais fabriqué à la main : c'est le
    chemin d'authentification que le consommateur doit emprunter."""
    r = client.post(
        f"/{TENANT}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials",
        },
    )
    assert r.status_code == 200, r.text
    return str(r.json()["access_token"])


@pytest.fixture()
def auth(jeton: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {jeton}"}


def tous_les_groupes(client, auth) -> list[dict]:
    """Les groupes en SUIVANT le curseur — `/v1.0/groups` pagine depuis 0.3.0.

    Le mock rendait ses groupes d'un coup, comme aucune instance réelle ne le
    fait. Un helper qui ne paginerait pas ferait passer les tests de ce dépôt
    pour verts tout en laissant le consommateur ignorer le curseur : exactement
    le défaut trouvé dans insights360 le 2026-09-03.
    """
    groupes: list[dict] = []
    url = "/v1.0/groups"
    while url:
        corps = client.get(url, headers=auth).json()
        groupes.extend(corps["value"])
        suivant = corps.get("@odata.nextLink")
        url = suivant.replace("http://testserver", "") if suivant else ""
    return groupes


def forger_jeton(*, aud: str = CLIENT_ID, expire_dans: int = 3600) -> str:
    """Un jeton fabriqué — sert UNIQUEMENT à éprouver les refus (expiration,
    mauvaise audience) que le flux normal ne sait pas produire."""
    charge = json.dumps({"aud": aud, "exp": int(time.time()) + expire_dans}).encode()
    return base64.urlsafe_b64encode(charge).rstrip(b"=").decode()
