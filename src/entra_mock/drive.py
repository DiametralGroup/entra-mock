"""Mock Microsoft Graph -- the "files" surface (driveItem).

A site, its default library ("Documents"), and nothing in it: the drive is
EMPTY at startup. A test fills it through the control plane
(`PUT /__admin/drive/files/{path}`, see `admin.py`) with the files it needs --
the mock ships no file dataset.

┌─ WHAT THE REAL GRAPH DOES, AND A NICE MOCK WOULD HIDE ──────────────────────┐
│ 1. an application in `Sites.Selected` without a grant on THE site gets 403  │
│    `accessDenied`, on the site as on everything it contains --              │
│    `ENTRA_MOCK_SITE_GRANT=none` reproduces it;                              │
│ 2. `/children` PAGINATES, by an opaque `$skiptoken` cursor -- two per page   │
│    here, so the pagination path is exercised by construction;              │
│ 3. without `$select`, each item carries `createdBy`/`lastModifiedBy`: a     │
│    name and an email -- a client that doesn't need it must SAY so;         │
│ 4. document libraries serve ONLY `quickXorHash` -- never `sha1Hash` nor     │
│    `sha256Hash`;                                                            │
│ 5. `/content` returns a 302 to a PRE-AUTHENTICATED URL. Forwarding the      │
│    Bearer there is a mistake: 401 here.                                    │
└─────────────────────────────────────────────────────────────────────────────┘

Every Graph route goes through `_denied` (token, then throttling) -- the SAME
preamble as `/v1.0/groups`: an injected 429 hits the files surface just as it
does the directory.
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

from .app import _denied, _graph_error, _next_link, _project, _UnknownSelect

# ┌─ THE SITE SERVED ───────────────────────────────────────────────────────────┐
# │ A host and a server-relative path -- the addressing a client configures.   │
# │ Both are configurable, to match the tested client's configuration without  │
# │ touching the mock.                                                        │
# └────────────────────────────────────────────────────────────────────────────┘
SITE_HOSTNAME = os.environ.get("ENTRA_MOCK_SITE_HOSTNAME", "contoso.sharepoint.com").lower()
SITE_PATH = "/" + os.environ.get("ENTRA_MOCK_SITE_PATH", "/sites/documents").strip("/")

# ┌─ THE APPLICATION'S GRANT ON THE SITE ──────────────────────────────────────┐
# │ An application in `Sites.Selected` sees NOTHING until an administrator     │
# │ has granted it a role on that specific site. Without that gesture, Graph   │
# │ returns 403 `accessDenied` -- the token is valid, it's the site that's     │
# │ closed. This is THE go-live failure, and it must be distinguishable from   │
# │ a 401 (token) and a 404 (path) in the client's diagnostic.                 │
# └────────────────────────────────────────────────────────────────────────────┘
SITE_GRANT = os.environ.get("ENTRA_MOCK_SITE_GRANT", "read")

# Two per page, for the same reason as one member per page on groups: Graph
# paginates `/children` (200 per page), and a folder that fit in one page
# would make a client not following `@odata.nextLink` invisible.
DRIVE_PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_DRIVE_PAGE_SIZE", "2"))

#: Lifetime of a pre-authenticated download URL.
TEMPAUTH_SECONDS = int(os.environ.get("ENTRA_MOCK_TEMPAUTH_SECONDS", "3600"))

# ┌─ THE DOWNLOAD URL'S HOST ───────────────────────────────────────────────────┐
# │ In production, the redirect LEAVES graph.microsoft.com for the site's      │
# │ host. Both httpx and requests then strip the Authorization header on      │
# │ their own -- that's an origin change. The mock, instead, redirects to      │
# │ ITSELF: a client following the redirect keeps its Bearer, and gets 401     │
# │ where production would have answered.                                     │
# │                                                                            │
# │ Assumed default: it forces the client to follow the redirect by hand,      │
# │ without Authorization -- which works everywhere. To reproduce the origin   │
# │ change, give it a second name of the same container here (e.g. a compose   │
# │ network alias, `http://files-mock:8000`).                                  │
# └────────────────────────────────────────────────────────────────────────────┘
DOWNLOAD_BASE_URL = os.environ.get("ENTRA_MOCK_DOWNLOAD_BASE_URL", "").rstrip("/")

LIBRARY = "Documents"


def _guid(seed: str) -> uuid.UUID:
    return uuid.uuid5(uuid.NAMESPACE_URL, f"https://{SITE_HOSTNAME}{SITE_PATH}#{seed}")


# Identifiers have the SHAPE of the real service's, because a client always
# ends up logging, storing or splitting one:
#   • site  : `<host>,<collection GUID>,<web GUID>` -- a COMPOSITE identifier,
#             with commas;
#   • drive : `b!` + base64url of the three GUIDs (collection, web, list), in
#             .NET byte order -- 66 characters, `-` and `_` included;
#   • item  : `01` + 32 uppercase base32 characters.
_COLLECTION, _WEB, _LIST = _guid("collection"), _guid("web"), _guid("list")
SITE_ID = f"{SITE_HOSTNAME},{_COLLECTION},{_WEB}"
DRIVE_ID = "b!" + base64.urlsafe_b64encode(
    _COLLECTION.bytes_le + _WEB.bytes_le + _LIST.bytes_le
).decode().rstrip("=")
SITE_URL = f"https://{SITE_HOSTNAME}{SITE_PATH}"
LIBRARY_URL = f"{SITE_URL}/Shared%20Documents"
_SITE_CREATED = "2026-01-01T00:00:00Z"

# ┌─ THE MOCK'S CLOCK, AND WHY IT ISN'T THE WALL CLOCK ────────────────────────┐
# │ A write through the control plane must make `lastModifiedDateTime`         │
# │ ADVANCE -- that's what an incremental client compares. But the wall clock  │
# │ would make each run different from the last, and a snapshot taken by the   │
# │ client would never be stable again.                                       │
# │                                                                            │
# │ Hence a MONOTONIC and DETERMINISTIC clock: a fixed base, plus one minute   │
# │ per write since the last reset. The same sequence of writes yields the     │
# │ same timestamps.                                                           │
# └────────────────────────────────────────────────────────────────────────────┘
_CLOCK_BASE = datetime(2026, 1, 1, tzinfo=UTC)

#: The generic identity that writes files -- no real person.
USER = {
    "email": "user@example.invalid",
    "id": str(uuid.uuid5(uuid.NAMESPACE_DNS, "user@example.invalid")),
    "displayName": "Mock User",
}

#: Signs `tempauth`. Drawn at startup: a URL issued by another process is
#: refused, like a URL from another tenant.
_SECRET = secrets.token_bytes(32)

#: The MIME type by extension -- explicit, so as not to depend on the
#: image's `/etc/mime.types` (absent from `python:3.12-slim`).
_MIME_TYPES = {
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".json": "application/json",
    ".pdf": "application/pdf",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

#: The characters a file name cannot carry.
_FORBIDDEN_CHARS = frozenset('"*:<>?\\|')


def _iso(instant: datetime) -> str:
    return instant.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def quick_xor_hash(data: bytes) -> str:
    """Microsoft's `quickXorHash` -- the ONLY hash served here.

    Published algorithm (OneDrive API, "QuickXorHash"): each byte is XORed
    into a 160-bit register, shifted 11 bits more than the previous one
    (rotation modulo 160), then the 8-byte little-endian length is XORed into
    the last 8 bytes. Verified against the reference implementation
    (`quickxorhash` package) -- see tests.

    Bytes at positions i, i+160, i+320… receive the same shift
    (160 * 11 ≡ 0 mod 160): they are XORed first, block by block, as
    integers.
    """
    width, shift = 160, 11
    mask = (1 << width) - 1
    acc = 0
    for start in range(0, len(data), width):
        acc ^= int.from_bytes(data[start : start + width], "little")
    register = 0
    for position, byte_ in enumerate(acc.to_bytes(width, "little")):
        if byte_:
            value = byte_ << ((position * shift) % width)
            register ^= (value & mask) | (value >> width)
    result = bytearray(register.to_bytes(width // 8, "little"))
    for i, byte_ in enumerate(len(data).to_bytes(8, "little")):
        result[width // 8 - 8 + i] ^= byte_
    return base64.b64encode(bytes(result)).decode()


@dataclass
class Element:
    """A `driveItem` -- a folder when `content` is None."""

    #: From the drive's root, original casing; "" for the root.
    path: str
    item_id: str
    unique_id: str
    #: The list item id -- it figures in the `$skiptoken`.
    number: int
    created: str
    modified: str
    author: dict[str, str]
    modifier: dict[str, str]
    content: bytes | None = None
    version: int = 1
    digest: str = ""
    mime_type: str = ""

    @property
    def name(self) -> str:
        return self.path.rpartition("/")[2] if self.path else "root"

    @property
    def parent(self) -> str:
        return self.path.rpartition("/")[0]

    @property
    def is_folder(self) -> bool:
        return self.content is None


def _key(path: str) -> str:
    """Paths are case-insensitive: `Reports` and `reports` designate the
    same folder. A case-sensitive mock would fail a correct client."""
    return "/".join(p for p in path.split("/") if p).casefold()


class Drive:
    """The drive's state -- EMPTY (the root alone) after `reset()`."""

    def __init__(self) -> None:
        self.lock = threading.RLock()
        self.items: dict[str, Element] = {}
        self.counters: Counter[str] = Counter()
        self.births: Counter[str] = Counter()
        self.writes = 0
        self.last_number = 0
        self.grant = SITE_GRANT
        self.reset()

    def reset(self) -> None:
        with self.lock:
            self.items = {}
            self.counters = Counter()
            self.births = Counter()
            self.writes = 0
            self.last_number = 0
            # Reread on every reset: a test can change SITE_GRANT and get
            # back the state the environment would have produced at startup.
            self.grant = SITE_GRANT
            system = {"displayName": "System Account"}
            self._create("", None, _SITE_CREATED, _SITE_CREATED, system)

    def _create(
        self, path: str, content: bytes | None, created: str, modified: str, author: dict[str, str]
    ) -> Element:
        key = _key(path)
        # A file deleted then rewritten is a NEW item: a new id. The birth
        # count therefore enters the seed -- deterministic for the same
        # sequence of operations.
        seed = f"{key}#{self.births[key]}"
        self.births[key] += 1
        self.last_number += 1
        digest = hashlib.sha1(seed.encode(), usedforsecurity=False).digest()
        element = Element(
            path=path.strip("/"),
            item_id="01" + base64.b32encode(digest).decode()[:32],
            unique_id=str(_guid(seed)),
            number=self.last_number,
            created=created,
            modified=modified,
            author=author,
            modifier=author,
        )
        if content is not None:
            self._fill(element, content)
        self.items[key] = element
        return element

    @staticmethod
    def _fill(element: Element, content: bytes) -> None:
        element.content = content
        element.digest = quick_xor_hash(content)
        extension = "." + element.name.rpartition(".")[2].lower() if "." in element.name else ""
        element.mime_type = _MIME_TYPES.get(extension, "application/octet-stream")

    def now(self) -> str:
        self.writes += 1
        return _iso(_CLOCK_BASE + timedelta(minutes=self.writes))

    def root(self) -> Element:
        return self.items[""]

    def find(self, path: str) -> Element | None:
        return self.items.get(_key(path))

    def by_id(self, item_id: str) -> Element | None:
        return next((e for e in self.items.values() if e.item_id == item_id), None)

    def by_unique_id(self, unique_id: str) -> Element | None:
        unique_id = unique_id.strip("{}").lower()
        return next((e for e in self.items.values() if e.unique_id == unique_id), None)

    def children(self, folder: Element) -> list[Element]:
        """DIRECT children, sorted by name case-insensitively."""
        if not folder.is_folder:
            return []
        key = _key(folder.path)
        return sorted(
            (e for e in self.items.values() if e.path and _key(e.parent) == key),
            key=lambda e: e.name.casefold(),
        )

    def size(self, element: Element) -> int:
        if element.content is not None:
            return len(element.content)
        prefix = _key(element.path) + "/" if element.path else ""
        return sum(
            len(e.content)
            for key, e in self.items.items()
            if e.content is not None and key.startswith(prefix)
        )

    def write(self, path: str, content: bytes, modified: str | None = None) -> tuple[Element, bool]:
        """Creates or OVERWRITES a file.

        Overwriting KEEPS the id (it's the same item), bumps the `n` of
        `eTag`/`cTag`, and advances `lastModifiedDateTime` on the mock's
        clock. Intermediate folders are created.

        `modified` (ISO 8601 in UTC) replaces the mock clock for THIS file: a
        client reasoning about a file's date -- "was this month closed when
        the file was written?" -- needs chosen dates, not a clock that starts
        on January 1st.
        """
        path = "/".join(p for p in path.split("/") if p)
        with self.lock:
            existing = self.find(path)
            if existing is not None:
                if existing.is_folder:
                    raise IsADirectoryError(path)
                self._fill(existing, content)
                existing.version += 1
                existing.modified = modified or self.now()
                existing.modifier = dict(USER)
                return existing, False
            parts = path.split("/")
            for depth in range(1, len(parts)):
                parent = "/".join(parts[:depth])
                if (folder := self.find(parent)) is None:
                    instant = self.now()
                    self._create(parent, None, instant, instant, dict(USER))
                elif not folder.is_folder:
                    raise NotADirectoryError(parent)
            instant = modified or self.now()
            return self._create(path, content, instant, instant, dict(USER)), True

    def remove(self, path: str) -> bool:
        """Removes a file -- or a folder and everything it contains."""
        key = _key(path)
        with self.lock:
            if not key or key not in self.items:
                return False
            for other in [k for k in self.items if k == key or k.startswith(key + "/")]:
                del self.items[other]
            return True


mock_drive = Drive()

router: APIRouter = APIRouter()


# ── Pre-authenticated download URLs ──────────────────────────────────────────


def _b64(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode().rstrip("=")


def _sign(payload: str) -> str:
    return _b64(hmac.new(_SECRET, payload.encode(), hashlib.sha256).digest())


def issue_tempauth(unique_id: str) -> str:
    """An opaque, signed `tempauth` that EXPIRES -- nothing is stored mock-side."""
    expiry = int(time.time()) + TEMPAUTH_SECONDS
    payload = _b64(json.dumps({"uid": unique_id, "exp": expiry}).encode())
    return f"v1.{payload}.{_sign(payload)}"


def tempauth_valid(token: str, unique_id: str) -> bool:
    try:
        version, payload, signature = token.split(".")
        content = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    except ValueError:
        return False
    if version != "v1" or not hmac.compare_digest(signature, _sign(payload)):
        return False
    if not isinstance(content, dict):
        return False
    return bool(content.get("uid") == unique_id and content.get("exp", 0) > time.time())


def _download_url(request: Request, element: Element) -> str:
    base = DOWNLOAD_BASE_URL or str(request.base_url).rstrip("/")
    params = urlencode(
        {
            "UniqueId": element.unique_id,
            "Translate": "false",
            "tempauth": issue_tempauth(element.unique_id),
            "ApiVersion": "2.0",
        }
    )
    return f"{base}{SITE_PATH}/_layouts/15/download.aspx?{params}"


# ── Rendering ─────────────────────────────────────────────────────────────────


def _tag(element: Element, prefix: str = "") -> str:
    """`"{GUID},n"` -- and `"c:{GUID},n"` for the cTag. The quotes are part
    of the value, as they are at Graph."""
    return f'"{prefix}{{{element.unique_id.upper()}}},{element.version}"'


def _parent_reference(element: Element) -> dict[str, Any]:
    reference: dict[str, Any] = {"driveType": "documentLibrary", "driveId": DRIVE_ID}
    if not element.path:
        return reference
    parent = mock_drive.find(element.parent)
    if parent is not None:
        reference["id"] = parent.item_id
        if parent.path:
            reference["name"] = parent.name
        reference["path"] = f"/drives/{DRIVE_ID}/root:" + (f"/{parent.path}" if parent.path else "")
    reference["siteId"] = SITE_ID
    return reference


def _render(element: Element, request: Request) -> dict[str, Any]:
    """The DEFAULT property set of a `driveItem` -- wide, like Graph.

    `createdBy` and `lastModifiedBy` carry an identity there. A client that
    only needs the name, the size and the `cTag` must SAY so through
    `$select`.
    """
    obj: dict[str, Any] = {}
    if not element.is_folder:
        # Returned in the default set, as Graph does -- but NOT with a
        # `$select` that doesn't name it.
        obj["@microsoft.graph.downloadUrl"] = _download_url(request, element)
    obj |= {
        "createdBy": {"user": dict(element.author)},
        "createdDateTime": element.created,
        "eTag": _tag(element),
        "id": element.item_id,
        "lastModifiedBy": {"user": dict(element.modifier)},
        "lastModifiedDateTime": element.modified,
        "name": element.name,
        "parentReference": _parent_reference(element),
        "webUrl": f"{LIBRARY_URL}/{quote(element.path)}".rstrip("/"),
        "cTag": _tag(element, "c:"),
        "fileSystemInfo": {
            "createdDateTime": element.created,
            "lastModifiedDateTime": element.modified,
        },
        "size": mock_drive.size(element),
    }
    if element.is_folder:
        obj["folder"] = {"childCount": len(mock_drive.children(element))}
        if not element.path:
            obj["root"] = {}
    else:
        # `quickXorHash` ONLY: document libraries compute neither sha1 nor
        # sha256. A client comparing sha256 hits a missing key -- here as in
        # production.
        obj["file"] = {
            "mimeType": element.mime_type,
            "hashes": {"quickXorHash": element.digest},
        }
    return obj


#: The `$select` universe -- the properties the `driveItem` TYPE declares
#: (baseItem + driveItem in the CSDL), not the ones the page contains.
ITEM_PROPERTIES = frozenset(
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


def _last_modified() -> str:
    return max(e.modified for e in mock_drive.items.values())


def _site() -> dict[str, Any]:
    name = SITE_PATH.rpartition("/")[2]
    return {
        "createdDateTime": _SITE_CREATED,
        "description": "",
        "id": SITE_ID,
        "lastModifiedDateTime": _last_modified(),
        "name": name,
        "webUrl": SITE_URL,
        "displayName": name,
        "root": {},
        "siteCollection": {"hostname": SITE_HOSTNAME},
    }


def _drive_payload() -> dict[str, Any]:
    used = sum(len(e.content) for e in mock_drive.items.values() if e.content is not None)
    total = 27_487_790_694_400
    return {
        "createdDateTime": _SITE_CREATED,
        "description": "",
        "id": DRIVE_ID,
        "lastModifiedDateTime": _last_modified(),
        "name": LIBRARY,
        "webUrl": LIBRARY_URL,
        "driveType": "documentLibrary",
        "createdBy": {"user": {"displayName": "System Account"}},
        "owner": {
            "group": {
                "id": str(_guid("group-owners")),
                "displayName": f"{SITE_PATH.rpartition('/')[2]} Owners",
            }
        },
        "quota": {
            "deleted": 0,
            "remaining": total - used,
            "state": "normal",
            "total": total,
            "used": used,
        },
    }


# ── The preamble and errors ──────────────────────────────────────────────────


def _prelude(request: Request, operation: str) -> JSONResponse | None:
    """Token, throttling (the common preamble from `app.py`), then the site
    grant. The counter only retains authenticated, non-throttled calls: a
    429 replayed by the client doesn't skew its download count."""
    if (denied := _denied(request)) is not None:
        return denied
    with mock_drive.lock:
        mock_drive.counters[operation] += 1
    if mock_drive.grant == "none":
        return _graph_error(403, "accessDenied", "Access denied")
    return None


def _not_found() -> JSONResponse:
    return _graph_error(404, "itemNotFound", "The resource could not be found.")


def _site_not_found() -> JSONResponse:
    return _graph_error(404, "itemNotFound", "Requested site could not be found")


def _unknown_select(exc: _UnknownSelect, type_: str) -> JSONResponse:
    return _graph_error(
        400,
        "BadRequest",
        "Parsing OData Select and Expand failed: Could not find a property named "
        f"'{exc.fields[0]}' on type 'microsoft.graph.{type_}'.",
    )


def _project_one(
    obj: dict[str, Any], request: Request, universe: frozenset[str], context: str
) -> dict[str, Any]:
    (projected,) = _project([obj], request.query_params.get("$select"), universe)
    return {"@odata.context": f"https://graph.microsoft.com/v1.0/$metadata#{context}", **projected}


_RESPONSES: dict[int | str, dict[str, Any]] = {
    401: {"description": "`InvalidAuthenticationToken` -- Bearer missing, expired, wrong audience"},
    403: {"description": "`accessDenied` -- no `Sites.Selected` role granted on the site"},
    404: {"description": "`itemNotFound`"},
    429: {"description": "`TooManyRequests` + `Retry-After` (injectable)"},
}


# ── The site ──────────────────────────────────────────────────────────────────


def _serve_site(request: Request) -> Response:
    try:
        return JSONResponse(_project_one(_site(), request, frozenset(_site()), "sites/$entity"))
    except _UnknownSelect as exc:
        return _unknown_select(exc, "site")


@router.get("/v1.0/sites/{hostname}:/{server_relative_path:path}", responses=_RESPONSES)
def site_by_path(hostname: str, server_relative_path: str, request: Request) -> Response:
    """The site by its PATH -- `/sites/{host}:/{server-relative path}`.

    That's the addressing a client configures (a readable host and path),
    and it derives the site's COMPOSITE identifier from it for the rest. The
    optional trailing colon (`…:/sites/documents:`) is accepted, as at Graph.
    """
    if (denied := _prelude(request, "site")) is not None:
        return denied
    if hostname.casefold() != SITE_HOSTNAME:
        return _graph_error(400, "invalidRequest", "Invalid hostname for this tenancy")
    if _key(server_relative_path.rstrip(":")) != _key(SITE_PATH):
        return _site_not_found()
    return _serve_site(request)


@router.get("/v1.0/sites/{site_id}", responses=_RESPONSES)
def site(site_id: str, request: Request) -> Response:
    """The site by its composite identifier."""
    if (denied := _prelude(request, "site")) is not None:
        return denied
    if site_id.casefold() != SITE_ID.casefold():
        return _site_not_found()
    return _serve_site(request)


def _serve_drive(request: Request) -> Response:
    try:
        return JSONResponse(
            _project_one(_drive_payload(), request, frozenset(_drive_payload()), "drives/$entity")
        )
    except _UnknownSelect as exc:
        return _unknown_select(exc, "drive")


@router.get("/v1.0/sites/{site_id}/drive", responses=_RESPONSES)
def default_drive(site_id: str, request: Request) -> Response:
    """The site's DEFAULT library -- "Documents"."""
    if (denied := _prelude(request, "drive")) is not None:
        return denied
    if site_id.casefold() != SITE_ID.casefold():
        return _site_not_found()
    return _serve_drive(request)


@router.get("/v1.0/sites/{site_id}/drives", responses=_RESPONSES)
def drives(site_id: str, request: Request) -> Response:
    """The site's libraries, as a `value[]` collection."""
    if (denied := _prelude(request, "drives")) is not None:
        return denied
    if site_id.casefold() != SITE_ID.casefold():
        return _site_not_found()
    try:
        values = _project(
            [_drive_payload()], request.query_params.get("$select"), frozenset(_drive_payload())
        )
    except _UnknownSelect as exc:
        return _unknown_select(exc, "drive")
    return JSONResponse(
        {"@odata.context": "https://graph.microsoft.com/v1.0/$metadata#drives", "value": values}
    )


@router.get("/v1.0/drives/{drive_id}", responses=_RESPONSES)
def drive(drive_id: str, request: Request) -> Response:
    """A drive by its `b!…` identifier."""
    if (denied := _prelude(request, "drive")) is not None:
        return denied
    if drive_id != DRIVE_ID:
        return _not_found()
    return _serve_drive(request)


# ── Items ─────────────────────────────────────────────────────────────────────


def _serve_item(request: Request, element: Element) -> Response:
    try:
        return JSONResponse(
            _project_one(
                _render(element, request),
                request,
                ITEM_PROPERTIES,
                f"drives('{DRIVE_ID}')/items/$entity",
            )
        )
    except _UnknownSelect as exc:
        return _unknown_select(exc, "driveItem")


def _cursor(element: Element) -> str:
    """The `$skiptoken` -- opaque to the client: the sort key of the LAST
    item returned, not a rank. A file added or removed during pagination
    therefore shifts nothing."""
    return _b64(
        f"Paged=TRUE&p_SortBehavior=0&p_FileLeafRef={quote(element.name)}"
        f"&p_ID={element.number}".encode()
    )


def _read_cursor(token: str) -> str:
    try:
        text = base64.urlsafe_b64decode(token + "=" * (-len(token) % 4)).decode()
        return parse_qs(text, strict_parsing=True)["p_FileLeafRef"][0].casefold()
    except (ValueError, KeyError) as exc:
        raise ValueError("Invalid skip token.") from exc


def _page_size(request: Request) -> int:
    """`$top` is a CEILING, never a promise: OData lets the server return
    less, and Graph does. A `$top=999` therefore does not short-circuit the
    mock's pagination -- the client must follow `@odata.nextLink` whatever it
    asked for."""
    size = max(1, DRIVE_PAGE_SIZE)
    if (top := request.query_params.get("$top")) is not None:
        if not top.isdigit() or int(top) < 1:
            raise ValueError(f"Invalid value '{top}' for $top query option.")
        size = min(size, int(top))
    return size


def _serve_children(request: Request, folder: Element) -> Response:
    try:
        size = _page_size(request)
        token = request.query_params.get("$skiptoken")
        after = _read_cursor(token) if token is not None else None
    except ValueError as exc:
        return _graph_error(400, "invalidRequest", str(exc))

    with mock_drive.lock:
        children = mock_drive.children(folder)
        if after is not None:
            children = [e for e in children if e.name.casefold() > after]
        page, rest = children[:size], children[size:]
        try:
            values = _project(
                [_render(e, request) for e in page],
                request.query_params.get("$select"),
                ITEM_PROPERTIES,
            )
        except _UnknownSelect as exc:
            return _unknown_select(exc, "driveItem")

    body: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(driveItem)",
        "value": values,
    }
    if rest:
        # The next link reuses the request's path -- by path or by id, as it
        # came in -- and carries `$select` and `$top` over.
        path = quote(request.url.path, safe="/:!,$'()*+;=@")
        body["@odata.nextLink"] = _next_link(request, path, _cursor(page[-1]))
    return JSONResponse(body)


def _serve_content(request: Request, element: Element) -> Response:
    """302 to a PRE-AUTHENTICATED URL -- never the raw bytes."""
    if element.is_folder:
        return _not_found()
    return RedirectResponse(_download_url(request, element), status_code=302)


def _split(path: str) -> tuple[str, str]:
    """`reports:/children` → (`reports`, `children`);
    `reports/report-1.txt:` → (`reports/report-1.txt`, ``)."""
    if ":/" in path:
        target, _, segment = path.rpartition(":/")
        return target, segment.strip("/")
    return path.rstrip(":/"), ""


_CONTENT_RESPONSES: dict[int | str, dict[str, Any]] = {
    **_RESPONSES,
    302: {
        "description": "Redirect to a PRE-AUTHENTICATED download URL (`Location`, short-lived "
        "`tempauth`). Follow it WITHOUT an Authorization header: the URL returns 401 if "
        "one is presented."
    },
}


@router.get("/v1.0/drives/{drive_id}/root", responses=_RESPONSES)
def root(drive_id: str, request: Request) -> Response:
    """The drive's root."""
    if (denied := _prelude(request, "item")) is not None:
        return denied
    if drive_id != DRIVE_ID:
        return _not_found()
    return _serve_item(request, mock_drive.root())


@router.get("/v1.0/drives/{drive_id}/root/children", responses=_RESPONSES)
def root_children(drive_id: str, request: Request) -> Response:
    """The root's children -- empty until something is written."""
    if (denied := _prelude(request, "children")) is not None:
        return denied
    if drive_id != DRIVE_ID:
        return _not_found()
    return _serve_children(request, mock_drive.root())


@router.get(
    "/v1.0/drives/{drive_id}/root:/{path:path}",
    responses=_CONTENT_RESPONSES,
)
def by_path(drive_id: str, path: str, request: Request) -> Response:
    """PATH addressing, in its three forms:

    • `root:/reports:/children` -- a folder's children, PAGINATED
      (`@odata.nextLink`, opaque `$skiptoken`), `$select` honoured;
    • `root:/reports/report-1.txt` (trailing colon optional) -- an item;
    • `root:/reports/report-1.txt:/content` -- 302 to the download URL.

    The path's case is immaterial.
    """
    target, segment = _split(path)
    handlers = {"": _serve_item, "children": _serve_children, "content": _serve_content}
    if (
        denied := _prelude(request, segment if segment in handlers and segment else "item")
    ) is not None:
        return denied
    if segment not in handlers:
        return _graph_error(400, "BadRequest", f"Resource not found for the segment '{segment}'.")
    element = mock_drive.find(target) if drive_id == DRIVE_ID else None
    return _not_found() if element is None else handlers[segment](request, element)


@router.get("/v1.0/drives/{drive_id}/items/{item_id}", responses=_RESPONSES)
def item_by_id(drive_id: str, item_id: str, request: Request) -> Response:
    """An item by its `01…` identifier."""
    if (denied := _prelude(request, "item")) is not None:
        return denied
    element = mock_drive.by_id(item_id) if drive_id == DRIVE_ID else None
    return _not_found() if element is None else _serve_item(request, element)


@router.get("/v1.0/drives/{drive_id}/items/{item_id}/children", responses=_RESPONSES)
def children_by_id(drive_id: str, item_id: str, request: Request) -> Response:
    """A folder's children by its identifier -- same pagination."""
    if (denied := _prelude(request, "children")) is not None:
        return denied
    element = mock_drive.by_id(item_id) if drive_id == DRIVE_ID else None
    return _not_found() if element is None else _serve_children(request, element)


@router.get("/v1.0/drives/{drive_id}/items/{item_id}/content", responses=_CONTENT_RESPONSES)
def content(drive_id: str, item_id: str, request: Request) -> Response:
    """A file's content: 302 to a PRE-AUTHENTICATED URL.

    That's what Graph returns, and it must be followed WITHOUT the Bearer --
    the URL carries its own authorization (`tempauth`), time-limited.
    """
    if (denied := _prelude(request, "content")) is not None:
        return denied
    element = mock_drive.by_id(item_id) if drive_id == DRIVE_ID else None
    return _not_found() if element is None else _serve_content(request, element)


# ── The site's host: downloading ─────────────────────────────────────────────


def _download_denied(description: str) -> JSONResponse:
    # The STATUS is what's load-bearing (401); the exact body of the real
    # service on these refusals hasn't been observed -- a client must not
    # rely on it.
    return JSONResponse(
        status_code=401, content={"error": "invalid_request", "error_description": description}
    )


@router.get(f"{SITE_PATH}/_layouts/15/download.aspx", include_in_schema=False)
def download(request: Request) -> Response:
    """The PRE-AUTHENTICATED URL -- the site's host, outside Graph, outside
    the contract.

    ┌─ A BEARER HERE IS A MISTAKE ───────────────────────────────────────────┐
    │ Authorization is IN the URL (`tempauth`). A client forwarding the      │
    │ Graph token there exposes that token to another host -- and gets      │
    │ refused: 401. An unknown, tampered with, or expired `tempauth`: 401    │
    │ too.                                                                   │
    └──────────────────────────────────────────────────────────────────────────┘
    """
    if "authorization" in request.headers:
        return _download_denied(
            "Pre-authenticated download URL: the request must not carry an Authorization header."
        )
    unique_id = request.query_params.get("UniqueId", "").strip("{}").lower()
    if not tempauth_valid(request.query_params.get("tempauth", ""), unique_id):
        return _download_denied("Invalid or expired tempauth.")
    with mock_drive.lock:
        mock_drive.counters["download"] += 1
        element = mock_drive.by_unique_id(unique_id)
    if element is None or element.content is None:
        return JSONResponse(status_code=404, content={"error": "File Not Found."})
    return Response(
        element.content,
        media_type=element.mime_type,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(element.name)}",
            "ETag": _tag(element),
            "CTag": _tag(element, "c:"),
        },
    )


def valid_name(path: str) -> bool:
    """A path the real service would accept.

    Published rules ("Restrictions and limitations in OneDrive and
    SharePoint", support.microsoft.com): no `" * : < > ? / \\ |`, no leading
    or trailing space, no `_vti_` anywhere in the name.

    ASSUMED exception: names starting with `~$…`, which the same page
    declares invalid. A lock file must be writable to exercise a client's
    temporary-file filter.
    """
    parts = [p for p in path.split("/") if p]
    return bool(parts) and all(
        not (_FORBIDDEN_CHARS & set(p)) and p == p.strip() and "_vti_" not in p.lower()
        for p in parts
    )
