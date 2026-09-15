# entra-mock

A mock of the **Microsoft Graph** endpoints used for **Entra group
membership** — the OAuth2 client-credentials flow, `/v1.0/groups` and
`/v1.0/groups/{id}/members` with opaque-cursor pagination.

Extracted from the `mock/` directory of
[insights360](https://github.com/LittleBigCode/insights360) so it follows the
same shape as the other mocks of the ecosystem
([boondmanager-mock](https://github.com/LittleBigCode/boondmanager-mock),
[linkedin-mock](https://github.com/LittleBigCode/linkedin-mock),
[ga-mock](https://github.com/LittleBigCode/ga-mock)): its own repository, a
published container image, a smoke-tested CI, and tests that live with the
code they protect.

## Why this mock is separate from the BoondManager one

Entra group membership is **not** a BoondManager resource, contrary to what
the insights360 spec suggests (§4.2 — see `docs/SPEC-DEVIATIONS.md` #4 in that
repo). Nothing is shared:

|  | BoondManager | Microsoft Graph |
|---|---|---|
| auth | static HS256 JWT in a custom header | OAuth2 client credentials → **expiring** Bearer |
| envelope | `data[]` + `meta.totals.rows` | `value[]` + `@odata.nextLink` |
| pagination | `page` / `maxResults` | **opaque cursor** in a URL |
| errors | `{"errors": [{status, code, detail}]}` | `{"error": {"code", "message"}}` |

Mixing them in one client produces a client that lies about what it speaks.

## Start in one command

```bash
docker run --rm -p 8011:8000 ghcr.io/littlebigcode/entra-mock:latest   # prebuilt
docker compose up --build                                              # local build
make bootstrap && make run                                             # from source
```

Then:

```bash
TENANT=00000000-0000-0000-0000-000000000000
curl http://localhost:8011/health

TOKEN=$(curl -s -X POST \
  -d "client_id=$TENANT" -d 'client_secret=change-me-entra' \
  -d 'grant_type=client_credentials' \
  "http://localhost:8011/$TENANT/oauth2/v2.0/token" \
  | python3 -c 'import json,sys; print(json.load(sys.stdin)["access_token"])')

curl -H "Authorization: Bearer $TOKEN" http://localhost:8011/v1.0/groups
```

(`make run` binds 8000; compose maps host port **8011** — the insights360
allocation: boond 8010, entra 8011, linkedin 8012, ga 8013.)

## Served surface

| Endpoint | Notes |
|---|---|
| `POST /{tenant}/oauth2/v2.0/token` | client-credentials flow; **form-encoded body** (JSON is refused, as in the real API); yields a Bearer with a real expiry. Refusals use the **flat OAuth2 envelope with an `AADSTS…` code** — *not* the Graph one |
| `GET /v1.0/groups` | the groups, **paginated by opaque cursor** like the members |
| `GET /v1.0/groups/{id}/members` | **direct** members, paginated by **opaque cursor** (`@odata.nextLink`); last page carries no link; not all members are people |
| `GET /v1.0/groups/{id}/transitiveMembers` | **effective** membership — nested groups flattened, cycles cut, nested group objects kept alongside their members |
| `GET /health` | unauthenticated probe |

Both collections honour `$select` (and report it on the next link); without it
they return a **wide default property set**, as Graph does. `429` can be
injected. Errors carry `innerError.request-id`.

Tokens are checked for real: audience **and** expiry. That is what makes the
client-side renewal path testable — the main operational difference with
BoondManager's static credential.

## The dataset, and the two edge cases that justify it

Four groups, one per authorization rule of insights360
(`grp-bi-rh`, `grp-bi-sales`, `grp-bi-direction` → rule 3, scope by group;
`grp-bi-comex` → rule 4, full visibility).

⚠️ **The UPNs are those of boondmanager-mock's `realiste` dataset, and that is
load-bearing.** The UPN is the *only* join key between the directory and the
HR system: a domain that drifts produces no error, it produces zero
memberships — hence zero visibility, hence tests that pass by vacuity. A
dataset change on the BoondManager side breaks this file, and that is
intended: better a red test than an authorization model whose edge cases have
silently ceased to exist.

Two of those UPNs exist **only** to exercise the join between the two sources:

- **`ext.consultant@boreal-conseil.example`** — a group member who is **absent from the HR
  system**. The pipeline must drop it at the join, never emit a row with an
  unknown key (which would degrade the inner-layer predicate into "no
  filter");
- **`kevin.silva@boreal-conseil.example`** — present in HR, member of **no group**. He must see
  only himself (rule 1).

`grp-bi-comex` deliberately overlaps `grp-bi-direction` (`arthur.ivanov`, the top
of the hierarchy, the only one without a manager): the Comex grants him *every*
collaborator — Nantes included — while the outer RLS perimeter of `bi_rh` and
`bi_sales` excludes Nantes. That couple is what makes the `inner ⊆ outer`
invariant do real work: it holds *because* RLS trims, and would fail loudly if
the inner query were ever run outside its role. Without it, the invariant would
be true by vacuity.

**Page size defaults to 1** (`ENTRA_MOCK_PAGE_SIZE`). That is not a
performance choice: with a page of one, *every* group of more than one member
is paginated, so the `@odata.nextLink` path is exercised by construction
rather than by luck. A pipeline that ignored the link would see only the first
member — silently, without any error.

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `ENTRA_TENANT_ID` | `00000000-0000-0000-0000-000000000000` | tenant accepted on the token path |
| `ENTRA_CLIENT_ID` | `00000000-0000-0000-0000-000000000000` | client id, and token audience |
| `ENTRA_CLIENT_SECRET` | `change-me-entra` | client secret |
| `ENTRA_MOCK_TOKEN_TTL` | `3600` | token lifetime in seconds — lower it to rehearse renewal |
| `ENTRA_MOCK_PAGE_SIZE` | `1` | members per page (see above) |
| `ENTRA_MOCK_GROUPS_PAGE_SIZE` | `2` | groups per page — same reasoning as above |
| `ENTRA_MOCK_THROTTLE_EVERY` | `0` (off) | return `429` on every Nth authenticated call |
| `ENTRA_MOCK_RETRY_AFTER` | `1` | the `Retry-After` served with those `429` |
| `ENTRA_MOCK_UPN_DOMAIN` | `boreal-conseil.example` | UPN domain — must match the BoondManager mock's dataset (see above) |
| `ENTRA_MOCK_HOST` / `_PORT` | `0.0.0.0` / `8000` | uvicorn bind |

## Development

```bash
make bootstrap   # uv sync
make test        # pytest — OAuth2 flow, pagination, dialect, dataset
make lint        # ruff + mypy --strict
make contract    # regenerate contracts/msgraph.openapi.yaml
```

The version is bumped in lockstep in `pyproject.toml`,
`src/entra_mock/app.py` (`FastAPI(version=…)`) and `docker-compose.yml`.

## Versions

| Version | Dataset |
|---|---|
| `0.3.0` | same UPNs, plus `grp-bi-imbrique` (a nested group + a service principal, no direct membership) and `/transitiveMembers`. Faithful to six dialect facts the mock used to get wrong — the current one |
| `0.2.0` | UPNs of boondmanager-mock's **`realiste`** dataset (`@boreal-conseil.example`) |
| `0.1.0` | the historical `@ent.fr` dataset, superseded; it joins with nothing since boondmanager-mock 0.3.0 |

### What 0.3.0 fixes, and how it was found

`insights360:scripts/relever_dialecte_entra.py` challenges this mock against
the **real** service — and needs **no credential at all** to do it. Microsoft
publishes its dialect: `GET graph.microsoft.com/v1.0/$metadata` is the official
CSDL (1.8 MB, 200 OK, anonymous), and both hosts show their error envelopes to
anyone they refuse. Run on 2026-09-03, it found five divergences, every one of
them in the same direction — **the mock was kinder than the provider**, so a
consumer could be green here and broken in production:

| # | The mock said | Microsoft actually says |
|---|---|---|
| 1 | token errors in the Graph envelope `{"error": {"code", "message"}}` | the **flat OAuth2** envelope with an `AADSTS…` code, `error_codes`, `trace_id` |
| 2 | errors without `innerError` | `error.innerError` carries `date`, `request-id`, `client-request-id` — the id Microsoft support asks for |
| 3 | all groups in one page | `/v1.0/groups` **paginates** (100 per page) |
| 4 | four properties per member | the **default property set** — 79 are declared on `user` in the CSDL |
| 5 | never throttles | `429` + `Retry-After`, and an extraction's call profile is what triggers it |
| 6 | no `/transitiveMembers` at all | nesting is the norm, not the exception — see below |

Divergence 6 is the one that bites hardest, and it is measured, not argued.
Probed against a **real corporate directory** (2026-09-03, figures rounded and
the tenant left unnamed):

| Group | `/members` | `/transitiveMembers` |
|---|---|---|
| a company-wide group | **1** person | **~33** (19 nested groups) |
| four regional leadership groups | 2 to 4 | 5 to 7 (one nested group each) |

A consumer building an ACL from `/members` strips **~97%** of the rightful
holders on that first group — with no error and no warning. Nesting was the
norm there, not the exception.

Not serving the endpoint did worse than hide the defect: it **locked consumers
into it**, since fixing their code would have broken their local stack. A mock
must never make the correct behaviour more expensive than the broken one.

The rule that comes out of it, and it holds for the four mocks of the
ecosystem: **a mock must be at least as strict as the provider.** A permissive
mock does not make CI greener, it makes it less informative.
