"""The `/__admin` control plane -- outside the Graph surface.

  • the router is MOUNTED only if `ENTRA_MOCK_ADMIN_ENABLED` is true -- not
    "mounted then forbidden": absent. So it can never be left open by
    accident;
  • every call presents `X-Mock-Admin-Token` (`ENTRA_MOCK_ADMIN_TOKEN`);
  • the `/__admin` prefix can't collide with any Graph path, which all live
    under `/v1.0`, nor with the site's host (`/sites/…`).

It exists because the mock runs in a CONTAINER at the client's premises: a
test can't mutate the state in Python, it needs HTTP -- to write a file,
overwrite one, remove one, or count its downloads. The drive is EMPTY at
startup: this is where a test fills it.
"""

from __future__ import annotations

import hmac
import os
import re
from typing import Any, Literal

from fastapi import APIRouter, Header, Request, Response
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from .app import _graph_error, reset_calls
from .drive import _render, mock_drive, valid_name

_ISO_UTC = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z")

ADMIN_TOKEN = os.environ.get("ENTRA_MOCK_ADMIN_TOKEN", "mock-admin-token")

router: APIRouter = APIRouter(
    prefix="/__admin", include_in_schema=False, tags=["mock control plane"]
)


def _reject(token: str | None) -> JSONResponse | None:
    if not token or not hmac.compare_digest(token.encode(), ADMIN_TOKEN.encode()):
        return _graph_error(403, "Forbidden", "invalid or missing X-Mock-Admin-Token")
    return None


@router.post("/reset")
def reset(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Resets the whole mock: the drive (empty), its counters, the site
    grant, and the call counter that paces throttling."""
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    mock_drive.reset()
    reset_calls()
    return JSONResponse({"status": "reset"})


@router.put("/drive/files/{path:path}")
async def write(
    path: str, request: Request, x_mock_admin_token: str | None = Header(default=None)
) -> Response:
    """Writes a file -- the RAW body is its content.

    The path starts from the drive's ROOT: `reports/report-1.txt`. Folders
    that don't exist yet are created along the way.
    201 on create, 200 on overwrite -- which keeps the id, bumps the `n` of
    `eTag`/`cTag` by one, and advances `lastModifiedDateTime` by one minute
    on the mock's clock (never the wall clock: see `drive.py`).

    `?lastModifiedDateTime=2026-09-05T08:00:00Z` sets the file's date instead
    of the mock clock.
    """
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    if not valid_name(path):
        return _graph_error(400, "invalidRequest", f"Invalid file name: '{path}'")
    modified = request.query_params.get("lastModifiedDateTime")
    if modified is not None and not _ISO_UTC.fullmatch(modified):
        return _graph_error(
            400, "invalidRequest", "lastModifiedDateTime: expected YYYY-MM-DDTHH:MM:SSZ"
        )
    try:
        item, created = mock_drive.write(path, await request.body(), modified)
    except (IsADirectoryError, NotADirectoryError) as exc:
        return _graph_error(409, "nameAlreadyExists", f"A folder is in the way: '{exc}'")
    return JSONResponse(_render(item, request), status_code=201 if created else 200)


@router.delete("/drive/files/{path:path}")
def remove(path: str, x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Removes a file -- or a folder and everything inside it. 204, or 404."""
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    if not mock_drive.remove(path):
        return _graph_error(404, "itemNotFound", "The resource could not be found.")
    return Response(status_code=204)


@router.post("/drive/reset")
def reset_drive(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """The drive alone goes back to EMPTY -- counters and clock included."""
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    mock_drive.reset()
    return JSONResponse({"status": "reset"})


@router.get("/drive/counters")
def counters(x_mock_admin_token: str | None = Header(default=None)) -> Response:
    """Calls per operation since the last reset.

    Only AUTHENTICATED and NON-THROTTLED calls count: a replayed 429 doesn't
    inflate `content`. That is what makes it possible to PROVE that a second,
    idempotent run doesn't redownload anything (`content` and `download`
    unchanged).
    """
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    operations = ("site", "drive", "drives", "item", "children", "content", "download")
    return JSONResponse({op: mock_drive.counters[op] for op in operations})


class GrantRequest(BaseModel):
    grant: Literal["none", "read"]


@router.post("/drive/grant")
def grant(body: GrantRequest, x_mock_admin_token: str | None = Header(default=None)) -> Any:
    """Changes the application's role on the site, at runtime.

    `none` reproduces the `Sites.Selected` application nobody has yet granted
    the site to: 403 `accessDenied` everywhere. The runtime counterpart of
    `ENTRA_MOCK_SITE_GRANT` -- a CI container doesn't restart between two
    tests.
    """
    if (refusal := _reject(x_mock_admin_token)) is not None:
        return refusal
    mock_drive.grant = body.grant
    return {"status": "updated", "grant": mock_drive.grant}
