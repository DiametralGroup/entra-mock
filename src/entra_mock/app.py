"""Mock Microsoft Graph -- Entra group membership.

┌─ WHY THIS MOCK IS SEPARATE FROM BOONDMANAGER'S ─────────────────────────────┐
│ The spec files "Entra group membership" under §4.2, with the BoondManager   │
│ resources. That's a mistake: it isn't the same API.                        │
│                                                                             │
│                    BoondManager              Microsoft Graph                │
│   auth             static HS256 JWT          OAuth2 client credentials      │
│                    in a custom header        → expiring Bearer token        │
│   envelope         data[] + meta.totals      value[] + @odata.nextLink      │
│   pagination       page / maxResults         opaque cursor in a URL         │
│                                                                             │
│ Mixing them in one client would produce a client that lies about what it    │
│ speaks. See docs/SPEC-DEVIATIONS.md #4.                                     │
└─────────────────────────────────────────────────────────────────────────────┘

┌─ 0.3.0 -- WHAT THE PROBE FOUND, AND WHAT IS FIXED HERE ─────────────────────┐
│ `insights360:scripts/relever_dialecte_entra.py` challenges this mock        │
│ against the real service, on 2026-09-03, WITHOUT any credential: Graph's    │
│ official CSDL (`/v1.0/$metadata`) and both hosts' refusals are enough. Five │
│ divergences came out of it, all of the same kind -- the mock was NICER than │
│ the provider, so a client green here could fail there:                     │
│                                                                             │
│   1. the AAD authority returns a FLAT OAuth2 envelope and an `AADSTS…`      │
│      code; the mock returned Graph's nested envelope, with no code;        │
│   2. Graph logs `innerError.request-id` in its errors -- the identifier     │
│      Microsoft support asks for; the mock didn't return it;                │
│   3. `/v1.0/groups` PAGINATES at Microsoft; the mock returned everything    │
│      at once, which made a client not following the cursor invisible;      │
│   4. without `$select`, Graph returns the DEFAULT property set -- for      │
│      `user`, the bulk of 79 properties; the mock returned four, so a       │
│      non-minimized collection went unnoticed;                              │
│   5. Graph THROTTLES (429 + `Retry-After`); the mock never did.            │
│                                                                             │
│ Rule that follows, and that holds for the ecosystem's four mocks:           │
│ **a mock must be at least as STRICT as the provider.** A permissive mock    │
│ doesn't make CI greener, it makes it less informative.                     │
└─────────────────────────────────────────────────────────────────────────────┘

The dataset is aligned with boondmanager-mock's `insights360` profile: the
same UPNs, and above all the two edge cases that exist ONLY at the junction of
the two sources --

  • `ext.consultant@boreal-conseil.example`: in a group, ABSENT from the HR system;
  • `kevin.silva@boreal-conseil.example`     : in the HR system, in NO group.

Since 0.5.0, the mock also serves the Graph "files" surface (driveItem: a
site, its default library, EMPTY at startup). It lives in `drive.py` and
shares the SAME preamble (`_denied`); the control plane `/__admin`, which
fills it, lives in `admin.py`.
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

# The domain of boondmanager-mock's `realiste` dataset. The UPN is the ONLY
# join key between the directory and the HR system: a domain that drifts does
# not produce an error, it produces zero memberships -- hence zero
# visibility, hence green tests by vacuity.
UPN_DOMAIN = os.environ.get("ENTRA_MOCK_UPN_DOMAIN", "boreal-conseil.example")
TENANT_ID = os.environ.get("ENTRA_TENANT_ID", "00000000-0000-0000-0000-000000000000")
CLIENT_ID = os.environ.get("ENTRA_CLIENT_ID", "00000000-0000-0000-0000-000000000000")
CLIENT_SECRET = os.environ.get("ENTRA_CLIENT_SECRET", "change-me-entra")

# Page size deliberately MINIMAL: Graph paginates with an opaque cursor, and a
# pipeline that didn't follow `@odata.nextLink` would see only the first
# members -- silently, without any error. With a page of 1, EVERY group of
# more than one member is paginated: the pagination path is exercised by
# construction, not by luck.
PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_PAGE_SIZE", "1"))

# ┌─ AND `/v1.0/groups` PAGINATES TOO ──────────────────────────────────────────┐
# │ Microsoft serves 100 groups per page there. The mock returned four at      │
# │ once: a client that followed the cursor on `/members` but not here --      │
# │ the real case, found in insights360 on 2026-09-03 -- was green in CI and    │
# │ truncated at the 101st group of a real directory. Without error: the       │
# │ members of unlisted groups simply got no perimeter at all.                 │
# │                                                                            │
# │ Two per page, then, for the same reason as one member per page: so the     │
# │ path is exercised by CONSTRUCTION, not by luck.                            │
# └────────────────────────────────────────────────────────────────────────────┘
GROUPS_PAGE_SIZE = int(os.environ.get("ENTRA_MOCK_GROUPS_PAGE_SIZE", "2"))

# ┌─ THROTTLING, ON DEMAND ─────────────────────────────────────────────────────┐
# │ Graph throttles per tenant AND per application. The call profile of a      │
# │ membership extraction -- a list, then one paginated call PER GROUP -- is   │
# │ exactly what triggers it: on a real directory, the 429 is the nominal      │
# │ regime, not the rare case.                                                 │
# │                                                                            │
# │ Off by default (0): a mock that throttles without being asked to would     │
# │ make CI intermittent. Recovery tests set it explicitly.                    │
# └────────────────────────────────────────────────────────────────────────────┘
THROTTLE_EVERY = int(os.environ.get("ENTRA_MOCK_THROTTLE_EVERY", "0"))
RETRY_AFTER = os.environ.get("ENTRA_MOCK_RETRY_AFTER", "1")

# Control plane /__admin. Off by default: it only makes sense in tests -- and
# off means NOT MOUNTED.
ADMIN_ENABLED = os.environ.get("ENTRA_MOCK_ADMIN_ENABLED", "false").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}


def _with_domain(upn: str) -> str:
    return upn.format(d=UPN_DOMAIN)


def _guid(seed: str) -> str:
    """A STABLE GUID derived from a string.

    Real directory objects are identified by a GUID, not a readable key: the
    mock returned `user-<upn>`, which gave a consumer the illusion that the
    identifier carried the UPN. Deterministic so the dataset stays
    reproducible run after run.
    """
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, seed))


# ── Group membership ─────────────────────────────────────────────────────────
#
# The four groups match the four authorization rules:
#   grp-bi-rh / grp-bi-sales / grp-bi-direction → rule 3 (group scope)
#   grp-bi-comex                                → rule 4 (full visibility)
#
# ⚠️ `grp-bi-comex`, not `grp-comex` -- renamed on 2026-09-15. The real tenant
#    carries BOTH: the old one is empty and dead, the real partners' group is
#    `grp-bi-comex`, and that is the one the consumer
#    (insights360:int_acl_regles.sql) now keys rule 4 on.
#
#    What staying on the old name cost this mock: on the consumer side, rule 4
#    no longer matched anyone in the test dataset, and the suite stayed
#    GREEN -- `arthur.ivanov` is also a member of `grp-bi-direction`, which
#    gave him enough visibility for the assertions to hold. An entire
#    authorization rule had stopped being exercised without a single test
#    moving. That is precisely the failure mode a mock exists to prevent.
GROUPS: dict[str, dict[str, Any]] = {
    "11111111-0000-0000-0000-000000000001": {
        "displayName": "grp-bi-rh",
        "members": [
            # Level-2 manager, Paris agency. Her group scope (rule 3) gives
            # her the reports in her perimeter.
            _with_domain("yuki.lambert@{d}"),
            # Member of a group, UNKNOWN to the HR system. The pipeline must
            # drop her at the join to dim_collaborateur, never emit a row
            # with an unknown key -- which would degrade the inner predicate
            # into "no filter".
            _with_domain("ext.consultant@{d}"),
        ],
    },
    "11111111-0000-0000-0000-000000000002": {
        "displayName": "grp-bi-sales",
        "members": [_with_domain("lars.marchand@{d}")],
    },
    "11111111-0000-0000-0000-000000000003": {
        "displayName": "grp-bi-direction",
        "members": [_with_domain("arthur.ivanov@{d}")],
    },
    "11111111-0000-0000-0000-000000000004": {
        "displayName": "grp-bi-comex",
        # The top of the hierarchy, the only one without a manager. The
        # Comex grants him EVERY collaborator, Nantes included -- while the
        # outer RLS perimeter of bi_rh and bi_sales excludes Nantes. That
        # pairing is what makes the `inner ⊆ outer` invariant do real work:
        # it holds BECAUSE RLS trims, and would fail loudly if the inner
        # query were ever run outside the role. Without it, the invariant
        # would be true by vacuity.
        "members": [_with_domain("arthur.ivanov@{d}")],
    },
    # ┌─ THE NESTED GROUP -- A LOSS OF RIGHTS THAT RAISED NO ERROR ────────────┐
    # │ `/members` returns DIRECT members. A group that contains another one  │
    # │ appears there as a `#microsoft.graph.group` object, and the people    │
    # │ inside it appear NOWHERE -- although Entra does grant them effective  │
    # │ membership.                                                           │
    # │                                                                       │
    # │ The mock only returned `user` objects: this case didn't exist in CI,  │
    # │ and a consumer confused it with a service principal -- neither has an │
    # │ UPN, both fell into the same filter. One is correct, the other makes  │
    # │ rightful holders disappear.                                          │
    # │                                                                       │
    # │ This group exists so the consumer has to distinguish the two. It     │
    # │ produces NO membership row: it adds nothing to the downstream         │
    # │ authorization model, it only exercises the diagnostic.               │
    # └────────────────────────────────────────────────────────────────────────┘
    "11111111-0000-0000-0000-000000000005": {
        "displayName": "grp-bi-imbrique",
        "members": [
            {"type": "#microsoft.graph.group", "id": "11111111-0000-0000-0000-000000000002"},
            # No UPN by nature, and perfectly expected: this is the case a
            # consumer MUST silently ignore.
            {"type": "#microsoft.graph.servicePrincipal", "id": _guid("sp-powerbi")},
        ],
    },
}

# `kevin.silva@…` DELIBERATELY appears nowhere: present in the HR system,
# member of no group. He must see only himself (rule 1).
#
# ⚠️ These UPNs are those of boondmanager-mock's `realiste` dataset. They are
#    NOT arbitrary: each is chosen for the property it makes testable (top of
#    the hierarchy without a manager, agency outside the perimeter, member
#    outside the HR system). A dataset change on the BoondManager mock side
#    breaks this file -- and that is intended: better a red test than an
#    authorization model whose edge cases have silently ceased to exist.


# ┌─ THE DEFAULT SET IS WIDE, AND THAT IS THE POINT ───────────────────────────┐
# │ Without `$select`, Graph returns the DEFAULT property set of the type.     │
# │ The official CSDL declares 79 on `user`; the mock returned four.          │
# │                                                                            │
# │ Consequence: a client that FORGOT `$select` collected, in production,      │
# │ job title, mobile phone, department, employee id -- and nothing, in CI,    │
# │ told it so. But it is the COLLECTION that founds the obligation, not the   │
# │ storage: sorting on arrival comes too late.                                │
# │                                                                            │
# │ The mock therefore returns a WIDE default set -- a representative sample   │
# │ of the type's sensitive properties -- and honours `$select` when asked     │
# │ for. Forgetting `$select` becomes visible here, for free.                  │
# └────────────────────────────────────────────────────────────────────────────┘
def _user_member(upn: str) -> dict[str, Any]:
    """The default set of a `user`, in the spirit of what Graph returns.

    The values are fictional -- it's a mock -- but the NAMES are those of the
    official CSDL, and they are what decides whether a `$select` is correct.
    """
    first_name, _, last_name = upn.split("@", maxsplit=1)[0].partition(".")
    return {
        "@odata.type": "#microsoft.graph.user",
        "id": _guid(upn),
        "userPrincipalName": upn,
        "mail": upn,
        "displayName": f"{first_name.capitalize()} {last_name.capitalize()}",
        "givenName": first_name.capitalize(),
        "surname": last_name.capitalize(),
        "jobTitle": "Consultant",
        "department": "Delivery",
        "officeLocation": "Paris",
        "mobilePhone": "+33 6 00 00 00 00",
        "businessPhones": ["+33 1 00 00 00 00"],
        "employeeId": "E-0000",
        "accountEnabled": True,
        "userType": "Member",
    }


def _other_member(raw: dict[str, Any]) -> dict[str, Any]:
    """A directory object that is NOT a person -- hence without a UPN."""
    obj: dict[str, Any] = {"@odata.type": raw["type"], "id": raw["id"]}
    if raw["type"] == "#microsoft.graph.group":
        obj["displayName"] = GROUPS[raw["id"]]["displayName"]
    else:
        obj["displayName"] = "Power BI Service"
    return obj


def _transitive_members(group_id: str, seen: set[str] | None = None) -> list[dict[str, Any]]:
    """EFFECTIVE membership -- nested groups flattened.

    ┌─ WHY THIS ENDPOINT EXISTS, MEASURED AGAINST A REAL PROVIDER ───────────┐
    │ `/members` only returns DIRECT members. Probed on 2026-09-03 against  │
    │ a REAL corporate directory: a cross-cutting group returns ONE person   │
    │ direct and about THIRTY transitively -- it contains nineteen nested    │
    │ groups. Four regional leadership groups each lose three.               │
    │                                                                        │
    │ In other words, a consumer reading `/members` to build an ACL         │
    │ deprives 97% of the rightful holders of their perimeter on that        │
    │ group -- without an error, without a warning.                         │
    │                                                                        │
    │ The mock didn't serve this endpoint: the consumer therefore COULDN'T  │
    │ correct itself without breaking its local stack. That's an absence    │
    │ that locks you into the defect, not merely one that hides it.          │
    └─────────────────────────────────────────────────────────────────────────┘

    Graph returns nested groups IN ADDITION TO their members: the group
    object stays in the collection. We do the same -- a consumer filtering on
    UPN discards them anyway, and omitting them would hide the structure.

    `seen` cuts cycles: Entra allows A ∋ B ∋ A, and a naive recursion would
    loop on it forever.
    """
    seen = set() if seen is None else seen
    if group_id in seen:
        return []
    seen.add(group_id)

    rendered: list[dict[str, Any]] = []
    for member in GROUPS[group_id]["members"]:
        if isinstance(member, dict):
            rendered.append(_other_member(member))
            if member["type"] == "#microsoft.graph.group" and member["id"] in GROUPS:
                rendered.extend(_transitive_members(member["id"], seen))
        else:
            rendered.append(_user_member(member))

    # Deduplication: a person who is a member of two subgroups appears only
    # once, as Graph does.
    unique: dict[str, dict[str, Any]] = {}
    for obj in rendered:
        unique.setdefault(obj["id"], obj)
    return list(unique.values())


def _rendered_members(group_id: str) -> list[dict[str, Any]]:
    return [
        _other_member(m) if isinstance(m, dict) else _user_member(m)
        for m in GROUPS[group_id]["members"]
    ]


def _rendered_group(group_id: str) -> dict[str, Any]:
    """The default set of a `group` -- wide, for the same reason."""
    group = GROUPS[group_id]
    return {
        "@odata.type": "#microsoft.graph.group",
        "id": group_id,
        "displayName": group["displayName"],
        "description": f"Group {group['displayName']}",
        "mail": None,
        "mailEnabled": False,
        "mailNickname": group["displayName"],
        "securityEnabled": True,
        "groupTypes": [],
        "createdDateTime": "2024-01-01T00:00:00Z",
        "visibility": "Private",
    }


app = FastAPI(title="Microsoft Graph mock (Entra groups, drive files)", version="0.5.0")

#: Authenticated Graph call counter -- used ONLY for deterministic
#: throttling. A single-process mock: no race to worry about.
_calls = 0


def reset_calls() -> None:
    """Resets the throttling cadence to zero -- `POST /__admin/reset`."""
    global _calls  # noqa: PLW0603 -- a mock counter, single-process
    _calls = 0


def _graph_error(status: int, code: str, message: str) -> JSONResponse:
    """The error envelope of the RESOURCE API -- nested, with innerError.

    Observed against `graph.microsoft.com` on 2026-09-03: `error.innerError`
    carries `date`, `request-id` and `client-request-id`. This `request-id`
    is the identifier Microsoft support asks for on a ticket; a mock that
    didn't return it let diagnostics be written that lost it.
    """
    identifier = str(uuid.uuid4())
    return JSONResponse(
        status_code=status,
        headers={"request-id": identifier, "client-request-id": identifier},
        content={
            "error": {
                "code": code,
                "message": message,
                "innerError": {
                    "date": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S"),
                    "request-id": identifier,
                    "client-request-id": identifier,
                },
            }
        },
    )


def _aad_error(status: int, code: str, aadsts: int, message: str) -> JSONResponse:
    """The AUTHORITY's envelope -- OAuth2, FLAT, with an AADSTS code.

    ┌─ THE TWO HOSTS DON'T SPEAK THE SAME ERROR DIALECT ─────────────────────┐
    │ Observed on 2026-09-03 against `login.microsoftonline.com`:            │
    │   {"error": "invalid_client", "error_description": "AADSTS7000215: …", │
    │    "error_codes": [7000215], "timestamp", "trace_id", "correlation_id"}│
    │                                                                        │
    │ Nothing like Graph's `{"error": {"code", "message"}}`. The mock         │
    │ returned Graph's envelope ON THE TOKEN PATH: a diagnostic written       │
    │ against it read `error.code`, and found `None` in production.          │
    │                                                                        │
    │ The AADSTS code is what makes an authentication failure diagnosable --  │
    │ 7000215 (wrong secret) and 700038 (unknown application) lead to two     │
    │ different actions. Losing it means no longer knowing which one.        │
    │                                                                        │
    │ The STATUS is not what distinguishes the two hosts: the authority       │
    │ returns 401 on a wrong secret and 400 on an unknown grant_type. It's    │
    │ the ENVELOPE. The statuses below are the ones that were observed.       │
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


def _bearer_valid(request: Request) -> bool:
    header = request.headers.get("Authorization", "")
    if not header.startswith("Bearer "):
        return False
    token = header.removeprefix("Bearer ")
    try:
        payload = json.loads(base64.urlsafe_b64decode(token + "=" * (-len(token) % 4)))
    except Exception:
        return False
    # Expiry is checked for real: that is what makes client-side renewal
    # testable, the main operational difference with BoondManager's static
    # credential.
    return bool(payload.get("aud") == CLIENT_ID and payload.get("exp", 0) > time.time())


def _denied(request: Request) -> JSONResponse | None:
    """The preamble common to every Graph call: token, then throttling."""
    if not _bearer_valid(request):
        return _graph_error(401, "InvalidAuthenticationToken", "Access token is empty or expired.")
    global _calls  # noqa: PLW0603 -- a mock counter, single-process
    _calls += 1
    if THROTTLE_EVERY and _calls % THROTTLE_EVERY == 0:
        response = _graph_error(
            429, "TooManyRequests", "Too many requests. Please retry after some time."
        )
        response.headers["Retry-After"] = RETRY_AFTER
        return response
    return None


def _project(
    objects: list[dict[str, Any]], select: str | None, universe: frozenset[str]
) -> list[dict[str, Any]]:
    """Applies `$select` -- and refuses a property the TYPE doesn't carry.

    Refusing is the point: Graph returns 400 on an unknown `$select`, and
    does so for the WHOLE COLLECTION. A mock that ignored the parameter would
    let a typo through to the first real run, where it would cost not one
    row but the entire source.

    ┌─ AGAINST THE TYPE, NOT AGAINST THE PAGE ───────────────────────────────┐
    │ First version: the universe of valid properties was inferred from the  │
    │ objects of the current page. Wrong, and caught on insights360's first  │
    │ end-to-end run -- `$select=id,userPrincipalName` on a group none of     │
    │ whose members is a person returned 400, while Graph accepts it: it     │
    │ validates against the collection's TYPE, independently of what the      │
    │ page contains.                                                         │
    │                                                                        │
    │ A mock stricter than the provider is a defect just as much as a more    │
    │ permissive one: it fails correct code, which pushes people to write     │
    │ wrong code to appease it.                                              │
    └──────────────────────────────────────────────────────────────────────────┘

    OData annotations (`@odata.type`) cannot be selected and are always
    returned: Graph treats them as metadata, not as properties. A valid
    property absent from the object is simply omitted -- Graph does not
    return a null key for a field a derived type doesn't carry.
    """
    if not select:
        return objects
    requested = [c.strip() for c in select.split(",") if c.strip()]
    if unknown := [c for c in requested if c not in universe]:
        raise _UnknownSelect(unknown)
    return [
        {c: v for c, v in o.items() if c in requested or c.startswith("@odata.type")}
        for o in objects
    ]


#: The properties each collection may carry -- the universe `$select` is
#: validated against. Union of the derived types for `/members`: the
#: collection is declared `Collection(graph.directoryObject)`, and Graph
#: accepts selecting a `user` property there even if the page contains none.
GROUP_PROPERTIES = frozenset(_rendered_group(next(iter(GROUPS))))
MEMBER_PROPERTIES = frozenset(
    set(_user_member("firstname.lastname@example")) | {"id", "displayName", "deletedDateTime"}
)


class _UnknownSelect(Exception):
    def __init__(self, fields: list[str]) -> None:
        self.fields = fields


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "entra-mock"}


@app.post("/{tenant}/oauth2/v2.0/token")
async def token(tenant: str, request: Request) -> JSONResponse:
    """Client credentials flow -- Graph's real authentication path.

    The four refusals reproduce responses OBSERVED on 2026-09-03: status,
    OAuth2 code and AADSTS code. None of it is invented.
    """
    form = await request.form()
    if tenant != TENANT_ID:
        # AADSTS900021, observed on a non-existent tenant GUID.
        return _aad_error(
            400, "invalid_request", 900021, f"Requested tenant identifier '{tenant}' is not valid."
        )
    if form.get("grant_type") is None:
        # AADSTS900144, observed on a JSON body -- which FastAPI can't read
        # as a form, exactly as the authority doesn't read it either.
        return _aad_error(
            400,
            "invalid_request",
            900144,
            "The request body must contain the following parameter: 'grant_type'.",
        )
    if form.get("grant_type") != "client_credentials":
        # AADSTS70003, observed.
        return _aad_error(
            400,
            "unsupported_grant_type",
            70003,
            f"The app requested an unsupported grant type '{form.get('grant_type')}'.",
        )
    if form.get("client_id") != CLIENT_ID:
        # AADSTS700038, observed on an application unknown to the tenant --
        # 400, NOT 401: it isn't the secret that's wrong, it's the
        # application that doesn't exist. The two failures call for two
        # different actions.
        return _aad_error(
            400,
            "unauthorized_client",
            700038,
            f"{form.get('client_id')} is not a valid client id for this tenant.",
        )
    if form.get("client_secret") != CLIENT_SECRET:
        # AADSTS7000215, observed -- and a 401, unlike the other three.
        return _aad_error(
            401,
            "invalid_client",
            7000215,
            "Invalid client secret provided. Ensure the secret being sent in the request "
            "is the client secret value, not the client secret ID.",
        )

    duration = int(os.environ.get("ENTRA_MOCK_TOKEN_TTL", "3600"))
    payload = json.dumps({"aud": CLIENT_ID, "exp": int(time.time()) + duration}).encode()
    issued_token = base64.urlsafe_b64encode(payload).rstrip(b"=").decode()
    return JSONResponse(
        {
            "token_type": "Bearer",
            "expires_in": duration,
            # Returned by the real authority alongside `expires_in`: a client
            # that reads it must not find a mock that ignores it.
            "ext_expires_in": duration,
            "access_token": issued_token,
        }
    )


@app.get("/v1.0/groups")
def groups(request: Request) -> JSONResponse:
    """The groups, PAGINATED -- like Graph, which serves 100 per page.

    A client following `@odata.nextLink` on `/members` but not here would see
    only the first `GROUPS_PAGE_SIZE` groups, without any error.
    """
    if (denied := _denied(request)) is not None:
        return denied

    offset = int(request.query_params.get("$skiptoken", "0"))
    ids = list(GROUPS)[offset : offset + GROUPS_PAGE_SIZE]
    try:
        values = _project(
            [_rendered_group(gid) for gid in ids],
            request.query_params.get("$select"),
            GROUP_PROPERTIES,
        )
    except _UnknownSelect as exc:
        return _graph_error(
            400,
            "Request_UnsupportedQuery",
            f"Could not find a property named '{exc.fields[0]}' on type 'microsoft.graph.group'.",
        )

    body: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
        "value": values,
    }
    next_offset = offset + GROUPS_PAGE_SIZE
    if next_offset < len(GROUPS):
        body["@odata.nextLink"] = _next_link(request, "/v1.0/groups", next_offset)
    return JSONResponse(body)


@app.get("/v1.0/groups/{group_id}/transitiveMembers")
def transitive_members(group_id: str, request: Request) -> JSONResponse:
    """EFFECTIVE membership, nested groups flattened -- like Graph.

    Same envelope, same opaque cursor, same `$select` as `/members`: only the
    collection's composition changes. That is what lets a consumer switch
    from one path to the other without rewriting its pagination.
    """
    return _serve_members(group_id, request, transitive=True)


@app.get("/v1.0/groups/{group_id}/members")
def members(group_id: str, request: Request) -> JSONResponse:
    """A group's members, paginated by opaque cursor -- like Graph.

    The cursor here is the starting index encoded in `@odata.nextLink`. A
    pipeline that didn't follow this link would see only the first PAGE_SIZE
    members, without any error.

    Members are NOT all people: `members` is declared
    `Collection(graph.directoryObject)` in the official CSDL, and returns a
    nested group as well as a service principal. Both lack a UPN, and a
    consumer must treat them differently.
    """
    return _serve_members(group_id, request, transitive=False)


def _serve_members(group_id: str, request: Request, *, transitive: bool) -> JSONResponse:
    """The body common to both member collections."""
    if (denied := _denied(request)) is not None:
        return denied
    if group_id not in GROUPS:
        return _graph_error(404, "Request_ResourceNotFound", f"group {group_id} does not exist")

    offset = int(request.query_params.get("$skiptoken", "0"))
    all_members = _transitive_members(group_id) if transitive else _rendered_members(group_id)
    path = "transitiveMembers" if transitive else "members"
    try:
        values = _project(
            all_members[offset : offset + PAGE_SIZE],
            request.query_params.get("$select"),
            MEMBER_PROPERTIES,
        )
    except _UnknownSelect as exc:
        return _graph_error(
            400,
            "Request_UnsupportedQuery",
            f"Could not find a property named '{exc.fields[0]}' on type 'microsoft.graph.user'.",
        )

    body: dict[str, Any] = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects",
        "value": values,
    }
    next_offset = offset + PAGE_SIZE
    if next_offset < len(all_members):
        body["@odata.nextLink"] = _next_link(
            request, f"/v1.0/groups/{group_id}/{path}", next_offset
        )
    return JSONResponse(body)


def _next_link(request: Request, path: str, cursor: int | str) -> str:
    """The next link, COPYING parameters over -- like Graph.

    Graph carries `$select` (and other options) over into `@odata.nextLink`.
    A mock that didn't would make the second page wider than the first, and
    would hide a client that passes its parameters back by hand -- which
    would then overwrite the cursor and get stuck on the first page.

    The cursor is a rank for groups, an opaque token for a folder
    (`drive.py`); `$top`, when given, follows the same path.
    """
    base = str(request.base_url).rstrip("/")
    params = [f"$skiptoken={cursor}"]
    for option in ("$select", "$top"):
        if value := request.query_params.get(option):
            params.append(f"{option}={value}")
    return f"{base}{path}?" + "&".join(params)


# ── The "files" surface, and the control plane ───────────────────────────────
#
# Imported at the END of the module, deliberately: `drive.py` reuses
# `_denied`, `_graph_error`, `_project` and `_next_link` above -- the same
# token, the same throttling, the same error envelope as the directory.
# Duplicating them would let the two surfaces drift apart at the first fix.
from .drive import router as drive_router  # noqa: E402

app.include_router(drive_router)

# The control plane isn't "mounted then forbidden": when disabled, the
# surface doesn't exist.
if ADMIN_ENABLED:
    from .admin import router as admin_router

    app.include_router(admin_router)
