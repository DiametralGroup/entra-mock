"""The Microsoft Graph dialect -- what sets it apart from BoondManager.

Each test targets a difference that would break a client that confused the
two APIs: the OAuth2 flow (against a static credential), the `value[]`
envelope (against `data[]` + `meta.totals`), the opaque cursor
`@odata.nextLink` (against `page`/`maxResults`), and the error envelope
`{"error": {"code", "message"}}`.
"""

from __future__ import annotations

from conftest import (
    CLIENT_ID,
    CLIENT_SECRET,
    TENANT,
    all_groups,
    app_module,
    forge_token,
)

# ── The client credentials flow ──────────────────────────────────────────────


def test_the_token_is_an_expiring_bearer(client):
    r = client.post(
        f"/{TENANT}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials",
        },
    )
    assert r.status_code == 200
    body = r.json()
    assert body["token_type"] == "Bearer"
    assert body["expires_in"] > 0
    assert body["access_token"]


def test_the_token_body_is_form_encoded(client):
    """Graph posts a FORM-encoded body. Sending JSON must fail: a permissive
    mock would let through a client unable to talk to the real API."""
    r = client.post(
        f"/{TENANT}/oauth2/v2.0/token",
        json={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials",
        },
    )
    assert r.status_code >= 400


def test_the_authority_speaks_oauth2_not_graphs_dialect(client):
    """THE test from 0.3.0: the token path's envelope was WRONG.

    Observed against `login.microsoftonline.com` on 2026-09-03. The authority
    returns the OAuth2 envelope, FLAT, with an `AADSTS…` code -- not
    `{"error": {"code", "message"}}`, which is the RESOURCE API's.

    The mock returned Graph's. A consumer writing its diagnostic against it
    read `error.code`, and found `None` in production -- losing the AADSTS
    code, the only thing that separates "wrong secret" (7000215) from
    "unknown application" (700038), two failures calling for opposite
    actions.

    The statuses below are observed too: the authority returns 401 on a
    wrong secret, but 400 on the other three refusals. So it is NOT the
    status that distinguishes the two hosts, it's the envelope.
    """
    base = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
    }
    cases = [
        ({**base, "client_secret": "wrong"}, TENANT, 401, "invalid_client", 7000215),
        (
            {**base, "client_id": "11111111-2222-3333-4444-555555555555"},
            TENANT,
            400,
            "unauthorized_client",
            700038,
        ),
        (base, "11111111-2222-3333-4444-555555555555", 400, "invalid_request", 900021),
        ({**base, "grant_type": "password"}, TENANT, 400, "unsupported_grant_type", 70003),
    ]
    for sent_body, tenant, status, code, aadsts in cases:
        r = client.post(f"/{tenant}/oauth2/v2.0/token", data=sent_body)
        body = r.json()
        assert r.status_code == status, (code, r.text)
        # FLAT: `error` is a STRING, not an object. That's the distinction.
        assert body["error"] == code
        assert isinstance(body["error"], str)
        assert body["error_description"].startswith(f"AADSTS{aadsts}:")
        assert body["error_codes"] == [aadsts]
        assert {"timestamp", "trace_id", "correlation_id"} <= set(body)


def test_the_token_carries_ext_expires_in(client):
    """The real authority returns it alongside `expires_in`."""
    r = client.post(
        f"/{TENANT}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials",
        },
    )
    assert r.json()["ext_expires_in"] == r.json()["expires_in"]


# ── Authenticating Graph calls ───────────────────────────────────────────────


def test_no_token_401_graph_envelope(client):
    """Graph's error envelope, not BoondManager's."""
    r = client.get("/v1.0/groups")
    assert r.status_code == 401
    assert r.json()["error"]["code"] == "InvalidAuthenticationToken"
    assert "error" in r.json() and "errors" not in r.json()


def test_graph_error_carries_the_support_request_id(client):
    """`innerError.request-id` is the identifier a Microsoft ticket requires.

    Observed on 2026-09-03: every error from `graph.microsoft.com` carries
    `error.innerError` with `date`, `request-id` and `client-request-id`. The
    mock didn't return them -- a consumer could write a diagnostic that kept
    no trace of it, and find out the day of a real outage.
    """
    body = client.get("/v1.0/groups").json()
    inner = body["error"]["innerError"]
    assert {"date", "request-id", "client-request-id"} <= set(inner)
    assert inner["request-id"]


def test_expired_token_refused(client):
    """Expiry is checked FOR REAL -- that is what makes client-side renewal
    testable, the main operational difference with BoondManager's static
    credential."""
    expired = forge_token(expires_in=-60)
    r = client.get("/v1.0/groups", headers={"Authorization": f"Bearer {expired}"})
    assert r.status_code == 401


def test_token_for_another_audience_refused(client):
    other = forge_token(aud="99999999-9999-9999-9999-999999999999")
    r = client.get("/v1.0/groups", headers={"Authorization": f"Bearer {other}"})
    assert r.status_code == 401


def test_unreadable_token_refused(client):
    r = client.get("/v1.0/groups", headers={"Authorization": "Bearer not-base64!!"})
    assert r.status_code == 401


# ── Envelope and pagination ──────────────────────────────────────────────────


def test_value_envelope_and_odata_context(client, auth):
    r = client.get("/v1.0/groups", headers=auth)
    assert r.status_code == 200
    body = r.json()
    assert "value" in body and "@odata.context" in body
    # Markers of the other dialect must NOT appear.
    assert "data" not in body and "meta" not in body


def test_pagination_by_opaque_cursor(client, auth):
    """Group grp-bi-rh has two members and the page is 1: the second is only
    reachable through @odata.nextLink. A client ignoring that link would see
    only the first -- silently, without any error."""
    groups = client.get("/v1.0/groups", headers=auth).json()["value"]
    rh = next(g for g in groups if g["displayName"] == "grp-bi-rh")

    url = f"/v1.0/groups/{rh['id']}/members"
    seen: list[str] = []
    pages = 0
    while url and pages < 10:
        body = client.get(url, headers=auth).json()
        seen.extend(m["userPrincipalName"] for m in body["value"])
        next_link = body.get("@odata.nextLink")
        url = next_link.replace("http://testserver", "") if next_link else ""
        pages += 1

    assert pages >= 2, "pagination was not exercised"
    assert len(seen) == 2
    # The last page carries NO next link: that's the stop signal.
    assert (
        "@odata.nextLink"
        not in client.get(f"/v1.0/groups/{rh['id']}/members?$skiptoken=1", headers=auth).json()
    )


def test_member_carries_upn_and_mail(client, auth):
    groups = client.get("/v1.0/groups", headers=auth).json()["value"]
    body = client.get(f"/v1.0/groups/{groups[0]['id']}/members", headers=auth).json()
    member = body["value"][0]
    # `userPrincipalName` is THE key of the downstream authorization model --
    # `mail` also exists and the two diverge in a real organization.
    assert member["userPrincipalName"]
    assert member["@odata.type"] == "#microsoft.graph.user"


def test_unknown_group_404_graph_envelope(client, auth):
    r = client.get("/v1.0/groups/11111111-0000-0000-0000-999999999999/members", headers=auth)
    assert r.status_code == 404
    assert r.json()["error"]["code"] == "Request_ResourceNotFound"


def test_health_unauthenticated(client):
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json() == {"status": "ok", "service": "entra-mock"}


# ── What 0.3.0 adds: the three remaining gaps ───────────────────────────────


def test_groups_are_PAGINATED_like_microsoft(client, auth):
    """Microsoft serves 100 groups per page; the mock used to return
    everything at once.

    A consumer that followed `@odata.nextLink` on `/members` but not here
    was green in CI and truncated at the 101st group of a real directory --
    without error: the members of unlisted groups got no perimeter at all.
    That's the defect found in insights360 on 2026-09-03.
    """
    first = client.get("/v1.0/groups", headers=auth).json()
    assert "@odata.nextLink" in first, "group pagination is not exercised"
    assert len(first["value"]) == app_module.GROUPS_PAGE_SIZE

    everything = all_groups(client, auth)
    assert len(everything) == len(first["value"]) + 1 or len(everything) > len(first["value"])
    assert len({g["id"] for g in everything}) == len(everything), "a group was returned twice"


def test_without_select_the_default_set_is_WIDE(client, auth):
    """It's the COLLECTION that founds the obligation, not the storage.

    The official CSDL declares 79 properties on `user`. The mock used to
    return four: a consumer that forgot `$select` collected, in production,
    the job title, mobile, department and employee id -- and nothing, in CI,
    told it so. So the default set is wide HERE too, so forgetting it shows
    up for free.
    """
    groups = all_groups(client, auth)
    member = client.get(f"/v1.0/groups/{groups[0]['id']}/members", headers=auth).json()["value"][0]
    assert {"jobTitle", "mobilePhone", "department", "employeeId"} <= set(member)


def test_select_is_HONOURED_and_carried_into_the_next_link(client, auth):
    """Asking for two fields must return two -- on EVERY page.

    Graph copies `$select` into `@odata.nextLink`. A mock that didn't would
    make the second page wider than the first, and would hide a consumer
    passing its parameters back by hand -- which would then overwrite the
    cursor.
    """
    groups = all_groups(client, auth)
    rh = next(g for g in groups if g["displayName"] == "grp-bi-rh")

    url = f"/v1.0/groups/{rh['id']}/members?$select=id,userPrincipalName"
    pages = 0
    while url and pages < 10:
        body = client.get(url, headers=auth).json()
        for member in body["value"]:
            # `@odata.type` is an ANNOTATION, not a property: Graph always
            # returns it, it can't be selected.
            assert {c for c in member if not c.startswith("@")} == {"id", "userPrincipalName"}
        next_link = body.get("@odata.nextLink")
        url = next_link.replace("http://testserver", "") if next_link else ""
        pages += 1
    assert pages >= 2, "carrying $select over to the next page was not exercised"


def test_an_unknown_select_returns_400_on_the_WHOLE_collection(client, auth):
    """Graph 400s on a property the type doesn't carry.

    A mock that ignored the parameter would let a typo through to the first
    real run -- where it would cost not one row, but the entire source.
    """
    groups = all_groups(client, auth)
    r = client.get(
        f"/v1.0/groups/{groups[0]['id']}/members?$select=id,userPrincipalNam", headers=auth
    )
    assert r.status_code == 400
    assert r.json()["error"]["code"] == "Request_UnsupportedQuery"


def test_throttling_is_injectable_and_carries_retry_after(client, auth, monkeypatch):
    """Graph throttles; the mock never had.

    The call profile of a membership extraction -- a list, then one
    paginated call PER GROUP -- is exactly what triggers it. Without an
    injectable 429, no consumer can exercise its recovery, and the first
    encounter happens in production.

    Off by default: a mock that throttles spontaneously would make CI
    intermittent.
    """
    monkeypatch.setattr(app_module, "THROTTLE_EVERY", 1)
    monkeypatch.setattr(app_module, "_calls", 0)

    r = client.get("/v1.0/groups", headers=auth)
    assert r.status_code == 429
    assert r.headers["Retry-After"] == "1"
    assert r.json()["error"]["code"] == "TooManyRequests"


def test_member_ids_are_GUIDs(client, auth):
    """A real directory object is identified by a GUID, not by its UPN.

    The mock used to return `user-<upn>`: a consumer could infer that the
    identifier carried the join key. It doesn't.
    """
    groups = all_groups(client, auth)
    member = client.get(f"/v1.0/groups/{groups[0]['id']}/members", headers=auth).json()["value"][0]
    assert len(member["id"]) == 36 and member["id"].count("-") == 4
    assert member["userPrincipalName"] not in member["id"]


def test_select_is_validated_against_the_TYPE_not_the_page(client, auth):
    """`$select=userPrincipalName` must pass even on a page with no person.

    Defect caught on insights360's first end-to-end run, on 2026-09-03: the
    mock validated `$select` against the page's objects, so it returned 400
    on `grp-bi-imbrique`, none of whose members is a `user`. Graph validates
    against the collection's TYPE --
    `Collection(graph.directoryObject)` -- and accepts it.

    A mock STRICTER than the provider is a defect just as much as a more
    permissive one: it fails correct code, and pushes people to write wrong
    code to appease it.
    """
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    r = client.get(
        f"/v1.0/groups/{groups['grp-bi-imbrique']}/members?$select=id,userPrincipalName",
        headers=auth,
    )
    assert r.status_code == 200, r.text
    # The property is valid for the type but absent from these objects:
    # Graph omits it, it doesn't return a null key.
    assert all("userPrincipalName" not in m for m in r.json()["value"])
    assert all(set(m) <= {"id", "@odata.type"} for m in r.json()["value"])


def test_transitiveMembers_flattens_what_members_loses(client, auth):
    """THE test that justifies the endpoint, measured against a real provider.

    Probed on 2026-09-03 against a real corporate directory: a
    cross-cutting group returns ONE person direct and about THIRTY
    transitively -- nineteen nested groups. Four regional leadership groups
    each lose three.

    A consumer reading `/members` to build an ACL therefore deprives nearly
    all rightful holders of their perimeter, without an error and without a
    warning. The mock not serving this endpoint, it COULDN'T correct itself
    without breaking its local stack: an absence that locks you into the
    defect, not merely one that hides it.

    `grp-bi-imbrique` reproduces the shape in miniature: no person direct,
    one transitively.
    """
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    gid = groups["grp-bi-imbrique"]

    def people(path: str) -> set[str]:
        seen: set[str] = set()
        url = f"/v1.0/groups/{gid}/{path}"
        pages = 0
        while url and pages < 20:
            body = client.get(url, headers=auth).json()
            seen.update(m["userPrincipalName"] for m in body["value"] if m.get("userPrincipalName"))
            next_link = body.get("@odata.nextLink")
            url = next_link.replace("http://testserver", "") if next_link else ""
            pages += 1
        return seen

    direct, transitive = people("members"), people("transitiveMembers")
    assert direct == set(), "the nested group must have NO direct member"
    assert transitive == {"lars.marchand@boreal-conseil.example"}, transitive


def test_transitiveMembers_ALSO_returns_nested_groups(client, auth):
    """Graph keeps the group object in the collection, alongside its members.

    Omitting them would hide the directory's structure from a consumer
    trying to diagnose it -- and would make the mock diverge from the
    provider on an observable point.
    """
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    client.get(
        f"/v1.0/groups/{groups['grp-bi-imbrique']}/transitiveMembers?$skiptoken=0", headers=auth
    ).json()
    # The page is 1: we walk it to see every type.
    types = set()
    url = f"/v1.0/groups/{groups['grp-bi-imbrique']}/transitiveMembers"
    pages = 0
    while url and pages < 20:
        body = client.get(url, headers=auth).json()
        types.update(m["@odata.type"] for m in body["value"])
        next_link = body.get("@odata.nextLink")
        url = next_link.replace("http://testserver", "") if next_link else ""
        pages += 1
    assert "#microsoft.graph.group" in types
    assert "#microsoft.graph.user" in types


def test_transitiveMembers_does_not_loop_on_a_cycle(client, auth, monkeypatch):
    """Entra allows A ∋ B ∋ A; a naive recursion would loop on it forever."""
    groups = {g["displayName"]: g["id"] for g in all_groups(client, auth)}
    nested, sales = groups["grp-bi-imbrique"], groups["grp-bi-sales"]

    # Close the cycle: grp-bi-sales contains grp-bi-imbrique, which contains it.
    original = app_module.GROUPS[sales]["members"]
    monkeypatch.setitem(
        app_module.GROUPS[sales],
        "members",
        [*original, {"type": "#microsoft.graph.group", "id": nested}],
    )

    body = client.get(f"/v1.0/groups/{nested}/transitiveMembers", headers=auth)
    assert body.status_code == 200
