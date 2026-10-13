"""Graph's "files" surface (driveItem).

Each test targets a behaviour of the REAL Graph a nice mock would hide: the
403 of a `Sites.Selected` application without a grant on the site, the
pagination of `/children`, the default property set carrying an identity,
the `quickXorHash` alone, the 302 redirect to a pre-authenticated URL to be
followed WITHOUT the Bearer.

The drive is EMPTY at startup: tests that need files write them themselves
through the control plane -- small, neutral files.
"""

from __future__ import annotations

import base64
import re
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlsplit

import pytest
import yaml
from conftest import ADMIN, app_module, drive_module

SITE = "/v1.0/sites/contoso.sharepoint.com:/sites/documents"
DRIVE = drive_module.DRIVE_ID
ROOT = f"/v1.0/drives/{DRIVE}/root/children"
REPORTS = f"/v1.0/drives/{DRIVE}/root:/reports:/children"

#: A small, neutral tree: six children under `reports/` (three pages of two),
#: one subfolder, one file at the root.
TREE = {
    "reports/report-1.txt": b"hello",
    "reports/report-2.txt": b"world",
    "reports/report-3.csv": b"a,b\n1,2\n",
    "reports/report-4.txt": b"four",
    "reports/report-5.txt": b"five",
    "reports/archive/old.txt": b"old",
    "notes.txt": b"notes",
}


@pytest.fixture(autouse=True)
def empty_drive():
    """A drive RESET TO EMPTY -- before and after, so a write or a revoked
    grant doesn't leak from one test to the next."""
    drive_module.mock_drive.reset()
    yield
    drive_module.mock_drive.reset()


def write(client, path: str, content: bytes):
    return client.put(f"/__admin/drive/files/{path}", content=content, headers=ADMIN)


@pytest.fixture()
def tree(client) -> None:
    for path, content in TREE.items():
        assert write(client, path, content).status_code == 201, path


def _relative(url: str) -> str:
    return url.replace("http://testserver", "")


def all_children(client, auth, url: str) -> list[dict[str, Any]]:
    """Children FOLLOWING `@odata.nextLink` to the end."""
    items: list[dict[str, Any]] = []
    while url:
        response = client.get(_relative(url), headers=auth)
        assert response.status_code == 200, response.text
        body = response.json()
        items.extend(body["value"])
        url = body.get("@odata.nextLink", "")
    return items


def whole_tree(client, auth, select: str = "") -> dict[str, dict[str, Any]]:
    """The whole tree, path → item: descend into EVERY folder (by id),
    following pagination at EVERY level."""
    suffix = f"?$select={select}" if select else ""
    rendered: dict[str, dict[str, Any]] = {}
    to_visit = [("", ROOT)]
    while to_visit:
        prefix, url = to_visit.pop()
        for item in all_children(client, auth, url + suffix):
            path = prefix + item["name"]
            rendered[path] = item
            if "folder" in item:
                children = f"/v1.0/drives/{DRIVE}/items/{item['id']}/children"
                to_visit.append((f"{path}/", children))
    return rendered


def item(client, auth, path: str) -> dict[str, Any]:
    response = client.get(f"/v1.0/drives/{DRIVE}/root:/{path}", headers=auth)
    assert response.status_code == 200, response.text
    return dict(response.json())


def pages(client, auth, url: str) -> list[list[str]]:
    """Names, PAGE by page, following the cursor."""
    rendered: list[list[str]] = []
    while url:
        body = client.get(_relative(url), headers=auth).json()
        rendered.append([e["name"] for e in body["value"]])
        url = body.get("@odata.nextLink", "")
    return rendered


def download(client, auth, item_id: str) -> bytes:
    """The CORRECT move: `/content` without following, then `Location`
    WITHOUT the token."""
    redirect = client.get(
        f"/v1.0/drives/{DRIVE}/items/{item_id}/content", headers=auth, follow_redirects=False
    )
    assert redirect.status_code == 302, redirect.text
    response = client.get(redirect.headers["location"])
    assert response.status_code == 200, response.text
    return bytes(response.content)


# ── The site and its libraries ───────────────────────────────────────────────


def test_the_site_by_its_path_returns_a_composite_id(client, auth):
    r = client.get(SITE, headers=auth)
    assert r.status_code == 200, r.text
    site = r.json()
    host, collection, web = site["id"].split(",")
    assert host == "contoso.sharepoint.com"
    assert re.fullmatch(r"[0-9a-f-]{36}", collection) and re.fullmatch(r"[0-9a-f-]{36}", web)
    assert site["webUrl"] == "https://contoso.sharepoint.com/sites/documents"
    assert {"displayName", "name"} <= set(site)
    # The same site by its identifier.
    assert client.get(f"/v1.0/sites/{site['id']}", headers=auth).json()["id"] == site["id"]


@pytest.mark.parametrize(
    "path",
    [
        "/v1.0/sites/contoso.sharepoint.com:/sites/documents:",
        "/v1.0/sites/contoso.sharepoint.com:/sites/Documents",
        "/v1.0/sites/contoso.sharepoint.com%3A/sites/documents",
    ],
)
def test_the_sites_addressing_variants(client, auth, path):
    """Trailing colon, case, encoded colon: Graph accepts all three."""
    assert client.get(path, headers=auth).status_code == 200


def test_unknown_site_404_itemNotFound(client, auth):
    r = client.get("/v1.0/sites/contoso.sharepoint.com:/sites/autre", headers=auth)
    assert r.status_code == 404
    assert r.json()["error"]["code"] == "itemNotFound"
    assert client.get("/v1.0/sites/unknown,a,b", headers=auth).status_code == 404


def test_host_of_another_tenant_400(client, auth):
    r = client.get("/v1.0/sites/fabrikam.sharepoint.com:/sites/documents", headers=auth)
    assert r.status_code == 400
    assert r.json()["error"]["code"] == "invalidRequest"


def test_no_grant_on_the_site_403_accessDenied_EVERYWHERE(client, auth, monkeypatch):
    """The `Sites.Selected` application nobody has granted the site to.

    The token is VALID -- it's the site that's closed. The 403 must fall on
    the site as on everything it contains, with Graph's envelope: that is
    what lets the client say "request the grant" rather than "renew the
    token" (401) or "fix the path" (404).
    """
    monkeypatch.setattr(drive_module, "SITE_GRANT", "none")
    drive_module.mock_drive.reset()
    write(client, "reports/report-1.txt", b"hello")
    for url in (
        SITE,
        f"/v1.0/sites/{drive_module.SITE_ID}/drive",
        ROOT,
        REPORTS,
        f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
    ):
        r = client.get(url, headers=auth, follow_redirects=False)
        assert r.status_code == 403, url
        error = r.json()["error"]
        assert error["code"] == "accessDenied"
        assert error["message"] == "Access denied"
        assert {"date", "request-id", "client-request-id"} <= set(error["innerError"])


def test_the_grant_is_revoked_and_restored_at_runtime(client, auth):
    """A CI container doesn't restart between two tests: the runtime
    counterpart of ENTRA_MOCK_SITE_GRANT goes through the control plane."""
    assert client.post("/__admin/drive/grant", json={"grant": "none"}, headers=ADMIN).is_success
    assert client.get(SITE, headers=auth).status_code == 403
    assert client.post("/__admin/drive/grant", json={"grant": "read"}, headers=ADMIN).is_success
    assert client.get(SITE, headers=auth).status_code == 200
    r = client.post("/__admin/drive/grant", json={"grant": "owner"}, headers=ADMIN)
    assert r.status_code == 422


def test_the_default_library_and_its_listing(client, auth):
    site_id = client.get(SITE, headers=auth).json()["id"]
    drive = client.get(f"/v1.0/sites/{site_id}/drive", headers=auth).json()
    assert drive["name"] == "Documents"
    assert drive["driveType"] == "documentLibrary"
    assert drive["id"].startswith("b!") and len(drive["id"]) == 66
    drives = client.get(f"/v1.0/sites/{site_id}/drives", headers=auth).json()
    assert [d["id"] for d in drives["value"]] == [drive["id"]]
    assert client.get(f"/v1.0/drives/{drive['id']}", headers=auth).json()["name"] == "Documents"


# ── The drive, EMPTY by default ──────────────────────────────────────────────


def test_the_drive_is_EMPTY_at_startup(client, auth):
    """No file, no folder: the root alone."""
    body = client.get(ROOT, headers=auth).json()
    assert body["value"] == []
    assert "@odata.nextLink" not in body
    root = client.get(f"/v1.0/drives/{DRIVE}/root", headers=auth).json()
    assert root["folder"] == {"childCount": 0}
    assert root["root"] == {}
    assert root["size"] == 0
    assert root["name"] == "root"


# ── A folder's children ───────────────────────────────────────────────────────


@pytest.mark.usefixtures("tree")
def test_a_folder_PAGINATES_by_opaque_cursor(client, auth):
    """Six children, two per page: a client ignoring `@odata.nextLink` sees
    only the first two -- without any error."""
    first = client.get(REPORTS, headers=auth).json()
    assert len(first["value"]) == 2
    next_link = first["@odata.nextLink"]
    assert next_link.startswith("http://testserver/v1.0/drives/"), "nextLink must be ABSOLUTE"
    token = parse_qs(urlsplit(next_link).query)["$skiptoken"][0]
    assert not token.isdigit(), "the cursor must be OPAQUE, not a rank"

    assert pages(client, auth, REPORTS) == [
        ["archive", "report-1.txt"],
        ["report-2.txt", "report-3.csv"],
        ["report-4.txt", "report-5.txt"],
    ]


@pytest.mark.usefixtures("tree")
def test_by_path_and_by_id_the_same_pagination(client, auth):
    """`root:/reports:/children` and `items/{id}/children`: same pages, and
    each next link keeps the shape of the request."""
    reports = item(client, auth, "reports")
    by_id = f"/v1.0/drives/{DRIVE}/items/{reports['id']}/children"
    assert pages(client, auth, REPORTS) == pages(client, auth, by_id)
    assert (
        "/root:/reports:/children?" in client.get(REPORTS, headers=auth).json()["@odata.nextLink"]
    )
    assert (
        f"/items/{reports['id']}/children?"
        in client.get(by_id, headers=auth).json()["@odata.nextLink"]
    )


@pytest.mark.usefixtures("tree")
def test_the_whole_tree(client, auth):
    rendered = whole_tree(client, auth)
    assert {c for c, e in rendered.items() if "file" in e} == set(TREE)
    assert {c for c, e in rendered.items() if "folder" in e} == {"reports", "reports/archive"}
    assert rendered["reports"]["folder"]["childCount"] == 6
    assert rendered["reports"]["size"] == sum(
        len(v) for k, v in TREE.items() if k.startswith("rep")
    )


@pytest.mark.usefixtures("tree")
def test_top_is_a_ceiling_not_a_promise(client, auth):
    """`$top=500` does NOT short-circuit pagination: the server returns less."""
    big = client.get(f"{REPORTS}?$top=500", headers=auth).json()
    assert len(big["value"]) == 2
    assert "$top=500" in big["@odata.nextLink"]
    assert len(client.get(f"{REPORTS}?$top=1", headers=auth).json()["value"]) == 1
    assert client.get(f"{REPORTS}?$top=zero", headers=auth).status_code == 400


def test_an_unreadable_cursor_400(client, auth):
    assert client.get(f"{ROOT}?$skiptoken=not-a-token", headers=auth).status_code == 400


@pytest.mark.usefixtures("tree")
def test_select_MINIMIZES_and_follows_the_link(client, auth):
    """`$select` returns EXACTLY the requested properties -- no free `id`,
    no identity -- and it is carried over into the next link."""
    requested = "id,name,size,cTag,lastModifiedDateTime,file,folder"
    rendered = whole_tree(client, auth, requested)
    assert len(rendered) == len(TREE) + 2
    for item_ in rendered.values():
        assert set(item_) <= set(requested.split(",")), item_
        assert "createdBy" not in item_ and "lastModifiedBy" not in item_
    first = client.get(f"{REPORTS}?$select=name", headers=auth).json()
    assert "$select=name" in first["@odata.nextLink"]
    assert all(
        set(item_) == {"name"} for item_ in all_children(client, auth, f"{REPORTS}?$select=name")
    )


def test_an_unknown_select_400(client, auth):
    r = client.get(f"{ROOT}?$select=name,sha256Hash", headers=auth)
    assert r.status_code == 400
    assert r.json()["error"]["code"] == "BadRequest"
    assert "sha256Hash" in r.json()["error"]["message"]
    assert client.get(f"{SITE}?$select=nope", headers=auth).status_code == 400


def test_without_select_the_default_set_exposes_an_identity(client, auth):
    """The default `$select` exists to avoid: a name and an email."""
    write(client, "reports/report-1.txt", b"hello")
    file_item = item(client, auth, "reports/report-1.txt")
    for key in ("createdBy", "lastModifiedBy"):
        assert file_item[key]["user"]["displayName"] == "Mock User"
        assert file_item[key]["user"]["email"] == "user@example.invalid"
    assert {"parentReference", "webUrl", "fileSystemInfo", "@microsoft.graph.downloadUrl"} <= set(
        file_item
    )
    assert file_item["parentReference"]["path"] == f"/drives/{DRIVE}/root:/reports"


@pytest.mark.usefixtures("tree")
def test_ONLY_quickXorHash_never_sha(client, auth):
    for item_ in whole_tree(client, auth).values():
        if "folder" in item_:
            assert "file" not in item_
            continue
        assert set(item_["file"]["hashes"]) == {"quickXorHash"}
        assert "folder" not in item_
    assert item(client, auth, "reports/report-3.csv")["file"]["mimeType"] == "text/csv"


def test_etag_and_ctag_have_the_expected_shape(client, auth):
    write(client, "reports/report-1.txt", b"hello")
    file_item = item(client, auth, "reports/report-1.txt")
    assert re.fullmatch(r'"\{[0-9A-F-]{36}\},1"', file_item["eTag"])
    assert re.fullmatch(r'"c:\{[0-9A-F-]{36}\},1"', file_item["cTag"])
    assert file_item["eTag"][2:38] == file_item["cTag"][4:40]


@pytest.mark.usefixtures("tree")
def test_item_by_id_equals_item_by_path(client, auth):
    by_path = item(client, auth, "reports/report-2.txt")
    by_id = client.get(f"/v1.0/drives/{DRIVE}/items/{by_path['id']}", headers=auth).json()
    for key in ("id", "name", "size", "eTag", "cTag", "file", "lastModifiedDateTime"):
        assert by_id[key] == by_path[key]
    assert re.fullmatch(r"01[A-Z2-7]{32}", by_id["id"])


@pytest.mark.usefixtures("tree")
def test_path_case_is_immaterial(client, auth):
    assert client.get(REPORTS.replace("reports", "REPORTS"), headers=auth).status_code == 200


def test_unknown_folder_404_itemNotFound(client, auth):
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/absent:/children", headers=auth)
    assert r.status_code == 404
    assert r.json()["error"]["code"] == "itemNotFound"
    assert client.get("/v1.0/drives/b!unknown/root/children", headers=auth).status_code == 404
    assert client.get(f"/v1.0/drives/{DRIVE}/items/01UNKNOWN", headers=auth).status_code == 404
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/absent:/unknown", headers=auth)
    assert r.status_code == 400


# ── Content: 302, then a pre-authenticated URL ───────────────────────────────


def test_content_returns_302_to_a_pre_authenticated_url(client, auth):
    write(client, "reports/report-1.txt", b"hello")
    file_item = item(client, auth, "reports/report-1.txt")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/items/{file_item['id']}/content",
        headers=auth,
        follow_redirects=False,
    )
    assert r.status_code == 302
    location = urlsplit(r.headers["location"])
    assert location.scheme == "http"
    assert location.netloc == "testserver", "Location must be ABSOLUTE"
    assert location.path == "/sites/documents/_layouts/15/download.aspx"
    assert {"UniqueId", "tempauth"} <= set(parse_qs(location.query))

    # Followed WITHOUT a token: the bytes, with the right type.
    downloaded = client.get(r.headers["location"])
    assert downloaded.status_code == 200
    assert downloaded.content == b"hello"
    assert downloaded.headers["content-type"].startswith("text/plain")
    assert len(downloaded.content) == file_item["size"]


def test_content_by_path_too_and_not_on_a_folder(client, auth):
    write(client, "reports/report-1.txt", b"hello")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
        headers=auth,
        follow_redirects=False,
    )
    assert r.status_code == 302
    r = client.get(f"/v1.0/drives/{DRIVE}/root:/reports:/content", headers=auth)
    assert r.status_code == 404


def test_forwarding_the_bearer_to_the_download_url_401(client, auth):
    """THE mistake: authorization is IN the URL. The Graph token has no
    business there -- a client forwarding it is refused."""
    write(client, "reports/report-1.txt", b"hello")
    file_item = item(client, auth, "reports/report-1.txt")
    r = client.get(
        f"/v1.0/drives/{DRIVE}/items/{file_item['id']}/content",
        headers=auth,
        follow_redirects=False,
    )
    assert client.get(r.headers["location"], headers=auth).status_code == 401
    # The real-world trap: following the redirect WITH the token (same origin here).
    followed = client.get(f"/v1.0/drives/{DRIVE}/items/{file_item['id']}/content", headers=auth)
    assert followed.status_code == 401


def test_tampered_unknown_or_expired_tempauth_401(client, auth, monkeypatch):
    write(client, "reports/report-1.txt", b"hello")
    write(client, "reports/report-2.txt", b"world")
    url = item(client, auth, "reports/report-1.txt")["@microsoft.graph.downloadUrl"]
    assert client.get(url).status_code == 200
    token = parse_qs(urlsplit(url).query)["tempauth"][0]
    tampered = url.replace(token, token[:-4] + ("AAAA" if not token.endswith("AAAA") else "BBBB"))
    assert client.get(tampered).status_code == 401
    assert client.get(url.replace(token, "unknown")).status_code == 401
    # A tempauth valid for ANOTHER file doesn't open this one.
    other = item(client, auth, "reports/report-2.txt")["@microsoft.graph.downloadUrl"]
    other_token = parse_qs(urlsplit(other).query)["tempauth"][0]
    assert client.get(url.replace(token, other_token)).status_code == 401

    monkeypatch.setattr(drive_module, "TEMPAUTH_SECONDS", -1)
    expired = item(client, auth, "reports/report-1.txt")["@microsoft.graph.downloadUrl"]
    assert client.get(expired).status_code == 401


def test_quick_xor_hash_reference_vectors():
    """Vectors measured against the reference implementation (PyPI package
    `quickxorhash` 1.0.5, on 2026-10-08)."""
    vectors = {
        b"": "AAAAAAAAAAAAAAAAAAAAAAAAAAA=",
        b"a": "YQAAAAAAAAAAAAAAAQAAAAAAAAA=",
        b"hello world": "aCgDG9jwBhDc4Q1yawMZAAAAAAA=",
        bytes(range(256)) * 3: "rxAOGe1RimTF/e+k/m0O5nnSZT8=",
    }
    for data, expected in vectors.items():
        assert drive_module.quick_xor_hash(data) == expected


def test_the_served_quickXorHash_matches_the_bytes(client, auth):
    content = bytes(range(256)) * 3
    write(client, "data/blob.bin", content)
    file_item = item(client, auth, "data/blob.bin")
    assert download(client, auth, file_item["id"]) == content
    assert file_item["file"]["hashes"]["quickXorHash"] == "rxAOGe1RimTF/e+k/m0O5nnSZT8="
    assert len(base64.b64decode(file_item["file"]["hashes"]["quickXorHash"])) == 20
    assert file_item["file"]["mimeType"] == "application/octet-stream"


# ── The control plane ─────────────────────────────────────────────────────────


def test_the_control_plane_requires_its_token(client):
    assert client.post("/__admin/drive/reset").status_code == 403
    assert (
        client.post("/__admin/drive/reset", headers={"X-Mock-Admin-Token": "wrong"}).status_code
        == 403
    )
    assert client.put("/__admin/drive/files/a.txt", content=b"x").status_code == 403
    assert client.get("/__admin/drive/counters").status_code == 403
    assert client.post("/__admin/drive/reset", headers=ADMIN).status_code == 200


def test_overwriting_KEEPS_the_id_and_advances_ctag_etag_and_the_clock(client, auth):
    """A file rewritten under the SAME name: same item (same id), fresh
    content, `n` goes to 2 in `eTag` and `cTag`, `lastModifiedDateTime`
    advances -- on the mock's clock, not the wall clock."""
    assert write(client, "reports/report-1.txt", b"hello").status_code == 201
    before = item(client, auth, "reports/report-1.txt")
    assert before["createdDateTime"] == "2026-01-01T00:02:00Z"  # the folder took 00:01
    r = write(client, "reports/report-1.txt", b"hello, again")
    assert r.status_code == 200, r.text
    after = item(client, auth, "reports/report-1.txt")

    assert after["id"] == before["id"]
    assert after["cTag"] == before["cTag"].replace(",1", ",2")
    assert after["eTag"] == before["eTag"].replace(",1", ",2")
    assert after["lastModifiedDateTime"] == "2026-01-01T00:03:00Z"
    assert after["createdDateTime"] == before["createdDateTime"]
    assert after["file"]["hashes"] != before["file"]["hashes"]
    assert after["size"] == len(b"hello, again")
    assert download(client, auth, after["id"]) == b"hello, again"


def test_writing_creates_intermediate_folders(client, auth):
    r = write(client, "a/b/c/deep.txt", b"x")
    assert r.status_code == 201, r.text
    assert r.json()["parentReference"]["path"].endswith("/root:/a/b/c")
    assert r.json()["lastModifiedDateTime"] == "2026-01-01T00:04:00Z"
    for i, folder in enumerate(("a", "a/b", "a/b/c"), start=1):
        folder_item = item(client, auth, folder)
        assert folder_item["folder"]["childCount"] == 1
        assert folder_item["createdDateTime"] == f"2026-01-01T00:0{i}:00Z"
    assert [e["name"] for e in all_children(client, auth, ROOT)] == ["a"]


def test_the_mock_uses_no_wall_clock(client):
    """The same sequence of writes yields the same timestamps AND the same
    ids, run after run."""

    def record() -> list[tuple[str, str]]:
        return [
            (r.json()["id"], r.json()["lastModifiedDateTime"])
            for r in (write(client, f"f{i}.txt", b"x") for i in range(3))
        ]

    first = record()
    assert [t for _, t in first] == [
        "2026-01-01T00:01:00Z",
        "2026-01-01T00:02:00Z",
        "2026-01-01T00:03:00Z",
    ]
    client.post("/__admin/drive/reset", headers=ADMIN)
    assert record() == first


def test_delete_then_rewrite_yields_a_NEW_item(client, auth):
    write(client, "reports/report-1.txt", b"hello")
    before = item(client, auth, "reports/report-1.txt")
    to_delete = "/__admin/drive/files/reports/report-1.txt"
    assert client.delete(to_delete, headers=ADMIN).status_code == 204
    assert client.get(f"/v1.0/drives/{DRIVE}/items/{before['id']}", headers=auth).status_code == 404
    assert client.delete(to_delete, headers=ADMIN).status_code == 404
    assert write(client, "reports/report-1.txt", b"hello").status_code == 201
    assert item(client, auth, "reports/report-1.txt")["id"] != before["id"]


@pytest.mark.usefixtures("tree")
def test_deleting_a_folder_takes_its_content_with_it(client, auth):
    assert client.delete("/__admin/drive/files/reports", headers=ADMIN).status_code == 204
    assert set(whole_tree(client, auth)) == {"notes.txt"}


@pytest.mark.usefixtures("tree")
def test_resetting_the_drive_makes_it_EMPTY(client, auth):
    assert client.post("/__admin/drive/reset", headers=ADMIN).status_code == 200
    assert client.get(ROOT, headers=auth).json()["value"] == []


def test_invalid_names_are_refused(client):
    write(client, "reports/report-1.txt", b"hello")
    assert write(client, "a:b.txt", b"x").status_code == 400
    assert write(client, "a_vti_b.txt", b"x").status_code == 400
    assert write(client, " space.txt", b"x").status_code == 400
    assert write(client, "reports", b"x").status_code == 409
    assert write(client, "reports/report-1.txt/inside.txt", b"x").status_code == 409


def test_counters_only_retain_served_calls(client, auth):
    """What's needed to PROVE that a second, idempotent run doesn't
    redownload anything: the counters, not the content."""
    write(client, "reports/report-1.txt", b"hello")
    client.get(REPORTS)  # 401: not counted
    file_item = client.get(f"{REPORTS}?$select=id,name", headers=auth).json()["value"][0]
    download(client, auth, file_item["id"])
    counters = client.get("/__admin/drive/counters", headers=ADMIN).json()
    assert counters == {
        "site": 0,
        "drive": 0,
        "drives": 0,
        "item": 0,
        "children": 1,
        "content": 1,
        "download": 1,
    }


def test_the_global_reset_also_resets_the_throttling_cadence(client, auth, monkeypatch):
    monkeypatch.setattr(app_module, "THROTTLE_EVERY", 2)
    monkeypatch.setattr(app_module, "_calls", 0)
    assert client.get(SITE, headers=auth).status_code == 200
    assert client.post("/__admin/reset", headers=ADMIN).status_code == 200
    assert client.get(SITE, headers=auth).status_code == 200  # 1st call after the reset
    assert client.get(SITE, headers=auth).status_code == 429


# ── The common preamble: token and throttling ────────────────────────────────

GRAPH_ROUTES = [
    SITE,
    f"/v1.0/sites/{drive_module.SITE_ID}",
    f"/v1.0/sites/{drive_module.SITE_ID}/drive",
    f"/v1.0/sites/{drive_module.SITE_ID}/drives",
    f"/v1.0/drives/{DRIVE}",
    f"/v1.0/drives/{DRIVE}/root",
    ROOT,
    REPORTS,
    f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt",
    f"/v1.0/drives/{DRIVE}/root:/reports/report-1.txt:/content",
    f"/v1.0/drives/{DRIVE}/items/01WHATEVER",
    f"/v1.0/drives/{DRIVE}/items/01WHATEVER/children",
    f"/v1.0/drives/{DRIVE}/items/01WHATEVER/content",
]


@pytest.mark.parametrize("url", GRAPH_ROUTES)
def test_the_bearer_is_required_on_every_route(client, url):
    r = client.get(url, follow_redirects=False)
    assert r.status_code == 401
    assert r.json()["error"]["code"] == "InvalidAuthenticationToken"


@pytest.mark.parametrize("url", GRAPH_ROUTES)
def test_injected_throttling_hits_every_route(client, auth, monkeypatch, url):
    monkeypatch.setattr(app_module, "THROTTLE_EVERY", 1)
    monkeypatch.setattr(app_module, "_calls", 0)
    r = client.get(url, headers=auth, follow_redirects=False)
    assert r.status_code == 429
    assert r.headers["Retry-After"] == app_module.RETRY_AFTER
    assert r.json()["error"]["code"] == "TooManyRequests"


# ── The contract ──────────────────────────────────────────────────────────────


def test_the_published_contract_is_up_to_date_and_covers_files():
    """`make contract` must have been re-run: a client pins a copy of it,
    and a stale contract would make it validate against a false surface."""
    from entra_mock import app

    published = yaml.safe_load(Path("contracts/msgraph.openapi.yaml").read_text())
    assert published == app.openapi()
    paths = set(published["paths"])
    assert {
        "/v1.0/sites/{hostname}:/{server_relative_path}",
        "/v1.0/sites/{site_id}/drive",
        "/v1.0/sites/{site_id}/drives",
        "/v1.0/drives/{drive_id}/root:/{path}",
        "/v1.0/drives/{drive_id}/items/{item_id}",
        "/v1.0/drives/{drive_id}/items/{item_id}/content",
    } <= paths
    assert not any(c.startswith(("/__admin", "/sites/")) for c in paths)
    assert (
        "302"
        in published["paths"]["/v1.0/drives/{drive_id}/items/{item_id}/content"]["get"]["responses"]
    )


def test_the_control_plane_can_date_a_file(client, auth):
    r = client.put(
        "/__admin/drive/files/reports/q3.txt?lastModifiedDateTime=2025-10-05T08:00:00Z",
        content=b"q3",
        headers=ADMIN,
    )
    assert r.status_code == 201
    assert item(client, auth, "reports/q3.txt")["lastModifiedDateTime"] == "2025-10-05T08:00:00Z"
    # without the parameter, the mock clock takes over again
    assert write(client, "reports/q3.txt", b"q3 bis").status_code == 200
    assert item(client, auth, "reports/q3.txt")["lastModifiedDateTime"] != "2025-10-05T08:00:00Z"


def test_a_malformed_date_is_refused(client):
    r = client.put(
        "/__admin/drive/files/reports/q3.txt?lastModifiedDateTime=05/10/2025",
        content=b"q3",
        headers=ADMIN,
    )
    assert r.status_code == 400 and r.json()["error"]["code"] == "invalidRequest"
