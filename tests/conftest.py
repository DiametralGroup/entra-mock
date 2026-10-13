"""Test harness -- the application mounted in-process.

The property that must not be lost: the application the stack calls IS the
one the tests exercise.

The `/__admin` control plane is mounted -- mounting is conditional and the
configuration is read at import time, hence the setting is applied BEFORE the
package is imported.
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

# `entra_mock.app` is the FastAPI INSTANCE after the package's
# `from .app import app`: importing the module by name is the only
# unambiguous access to its constants (page size, throttling) the tests must
# drive.
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
def token(client) -> str:
    """A token obtained through the REAL flow -- never crafted by hand: that
    is the authentication path a consumer must take."""
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
def auth(token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"}


def all_groups(client, auth) -> list[dict]:
    """The groups, FOLLOWING the cursor -- `/v1.0/groups` paginates since
    0.3.0.

    The mock used to return all its groups at once, as no real instance
    does. A helper that didn't paginate would make this repo's tests pass
    while still letting the consumer ignore the cursor: exactly the defect
    found in insights360 on 2026-09-03.
    """
    groups: list[dict] = []
    url = "/v1.0/groups"
    while url:
        body = client.get(url, headers=auth).json()
        groups.extend(body["value"])
        next_link = body.get("@odata.nextLink")
        url = next_link.replace("http://testserver", "") if next_link else ""
    return groups


def forge_token(*, aud: str = CLIENT_ID, expires_in: int = 3600) -> str:
    """A crafted token -- used ONLY to exercise refusals (expiry, wrong
    audience) the normal flow can't produce."""
    payload = json.dumps({"aud": aud, "exp": int(time.time()) + expires_in}).encode()
    return base64.urlsafe_b64encode(payload).rstrip(b"=").decode()
