# entra-mock

A mock of the **Microsoft Graph** endpoints used for **Entra group
membership** — the OAuth2 client-credentials flow, `/v1.0/groups` and
`/v1.0/groups/{id}/members` with opaque-cursor pagination — and, since 0.5.0,
of the **SharePoint drive** surface the same Entra app uses to read the
Finance drop folder (monthly trial balances and FX rates, as Excel files).

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
docker run --rm -p 8011:8000 ghcr.io/diametralgroup/entra-mock:latest  # prebuilt
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
| `GET /v1.0/sites/{hostname}:/{server-relative-path}` | the site **by path** (e.g. `boreal-conseil.sharepoint.com:/sites/depot-finance`, trailing `:` optional, case-insensitive) → `{id: "<host>,<guid>,<guid>", webUrl, displayName, name}` |
| `GET /v1.0/sites/{site-id}` | the same site by its composite id |
| `GET /v1.0/sites/{site-id}/drive` · `/drives` | the default document library ("Documents", `b!…` id) · the list of libraries |
| `GET /v1.0/drives/{drive-id}` | the library by id |
| `GET /v1.0/drives/{drive-id}/root` · `/root/children` | the root folder · its children |
| `GET /v1.0/drives/{drive-id}/root:/{path}:/children` | a folder's children, **paginated** (`@odata.nextLink`, opaque `$skiptoken`), `$select` honoured, `$top` as a ceiling |
| `GET /v1.0/drives/{drive-id}/root:/{path}` | one item by path |
| `GET /v1.0/drives/{drive-id}/root:/{path}:/content` | **302** to a pre-authenticated download URL |
| `GET /v1.0/drives/{drive-id}/items/{item-id}` · `/children` · `/content` | the same three, by item id |
| `GET /sites/depot-finance/_layouts/15/download.aspx?UniqueId=…&tempauth=…` | the **pre-authenticated** download URL the 302 points to — the bytes; **401 if an `Authorization` header is present**, 401 on an unknown, tampered or expired `tempauth` |
| `GET /health` | unauthenticated probe |

Every collection honours `$select` (and reports it on the next link); without
it, it returns a **wide default property set**, as Graph does. `429` can be
injected — on the drive routes too: they go through the same preamble as
`/v1.0/groups` (bearer check, then throttling). Errors carry
`innerError.request-id`.

Mock-only surfaces, outside the published contract: `/__admin` (control plane,
see below) and `/__fixtures/depot_finance` (invalid file variants).

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

## The SharePoint drive — the Finance drop folder

The consumer (insights360, extraction source `depot_finance`) reads a
SharePoint folder where Finance drops a trial balance per entity and per
month, plus an FX-rate file. It reads it with **the same Entra app** as the
groups — same authority, same client-credentials token, same scope
`https://graph.microsoft.com/.default`. Hence the same mock: a second mock
would have its own token, and let through a client that asks for two.

### What the real Graph does, and the mock does too

| # | Behaviour | What a kinder mock would hide |
|---|---|---|
| 1 | An app on **`Sites.Selected`** with no grant on *that* site gets **403 `accessDenied`** (`{"error":{"code":"accessDenied","message":"Access denied","innerError":{…}}}`) — on the site and everything in it. `ENTRA_MOCK_SITE_GRANT=none`, or `POST /__admin/drive/grant {"grant":"none"}` at runtime | the first-deployment failure, which must be diagnosed as "ask for the grant", not "renew the token" (401) or "fix the path" (404) |
| 2 | `/children` **paginates** with an opaque `$skiptoken` (keyset on the last name, SharePoint-style). **Two items per page** by default; `$top` is a ceiling, never a promise | a client that ignores `@odata.nextLink` sees `balance_BOG_2026-01` and `-02`, and nothing else — silently |
| 3 | Without `$select`, every item carries **`createdBy` / `lastModifiedBy` with a person** (display name, email) | a client that forgets `$select` collects personal data in production. With `$select`, the mock returns **exactly** the selected properties — no implicit `id`: ask for it |
| 4 | Files carry **only `quickXorHash`** — SharePoint / OneDrive for Business never serve `sha1Hash` / `sha256Hash`. The hash is the real algorithm (checked against reference vectors) | a client comparing sha256 finds a missing key in production |
| 5 | `/content` answers **302** to an absolute, **pre-authenticated** URL (`tempauth`, valid `ENTRA_MOCK_TEMPAUTH_SECONDS`). That URL **refuses a request that carries an `Authorization` header** (401) | a client that forwards its Graph bearer to another host |
| 6 | `eTag` = `"{GUID},n"`, `cTag` = `"c:{GUID},n"`, `n` incremented when the content changes; ids have SharePoint's shapes (`<host>,<guid>,<guid>`, `b!…`, `01…`); paths are **case-insensitive** | a client that parses or stores them wrongly |

⚠️ **The 302 is same-origin here.** In production the redirect leaves
`graph.microsoft.com` for the tenant's SharePoint host, and httpx/requests
then **drop** the `Authorization` header by themselves (cross-origin
redirect). The mock redirects to itself, so a client that simply follows the
redirect keeps its bearer and gets 401 where production would have answered.
Deliberate: follow the 302 by hand, without the header — that works
everywhere. To reproduce the cross-origin behaviour instead, set
`ENTRA_MOCK_DOWNLOAD_BASE_URL` to a second name of the same container (e.g. a
compose network alias `http://sharepoint-mock:8000`).

Not implemented (ignored if sent): `$filter`, `$orderby`, `$expand` on
children. Downloads are not throttled.

### The dataset — built from code at startup

Site `boreal-conseil.sharepoint.com` / `/sites/depot-finance` ("Dépôt
Finance"), default library **Documents**, folder **`Insights360`**, whose
direct children are:

- `balance_<CODE>_<YYYY-MM>.xlsx` for `NTE`, `BOG`, `MTL` × `2026-01`…`2026-06`
  (18 files). One sheet `balance`, header **exactly**
  `entite | mois | compte | libelle_compte | debit | credit | devise`; `compte`
  and `mois` are text; `debit` / `credit` are numbers ≥ 0 rounded to the cent,
  the **movements** of the month (a balance-sheet account carries both sides).
  Every file is a **complete, balanced** trial balance (Σ debit = Σ credit to
  the cent — it is the aggregate of a double-entry journal), **leaf accounts
  only**, no account a strict prefix of another. Revenue on the credit side,
  costs on the debit side. Seeded, deterministic values: revenue ≈ 150k EUR
  (NTE), ≈ 400M COP (BOG), ≈ 120k CAD (MTL) per month;
- `taux_2026.xlsx`: one sheet `taux`, header
  `type_taux | periode | devise | devise_pour_1_eur` — `moyen` rows for
  `2026-01`…`2026-06` for `CAD` (1.45–1.52) and `COP` (4400–4700), `budget`
  rows for period `2026` (text). Units of currency for **1 EUR**; **no EUR
  row**;
- a sub-folder **`modeles`** holding `modele_balance.xlsx` — consumers must
  ignore sub-folders. It sorts *between* the balances and the rate file, so it
  shows up mid-stream, not first.

Intercompany flows are coherent: what BOG (`413595`) and MTL (`4090`) invoice
NTE is what NTE books in `604800`, and NTE's management fees (`706800`) are
BOG's `511095` and MTL's `6090` — converted at the month's average rate, so an
elimination matches to the FX rounding (≤ 0.01 EUR).

| NTE — Boréal Conseil Nantes — EUR (French PCG) | BOG — Boreal Conseil Bogota S.A.S — COP (Colombian PUC) | MTL — Boreal Conseil Montreal — CAD (QuickBooks-style) |
|---|---|---|
| `401000` Fournisseurs | `110505` Caja general | `1000` Chequing |
| `411000` Clients | `111005` Bancos - moneda nacional | `1200` Accounts Receivable |
| `421000` Personnel - rémunérations dues | `130505` Clientes nacionales | `2000` Accounts Payable |
| `431000` Sécurité sociale | `220505` Proveedores nacionales | `2100` Payroll Liabilities |
| `445660` TVA déductible sur autres biens et services | `237005` Aportes de nómina por pagar | `2200` GST/QST Payable |
| `445710` TVA collectée | `240805` Impuesto sobre las ventas por pagar (IVA) | `4000` Consulting Revenue |
| `512000` Banque | `250505` Salarios por pagar | `4090` Intercompany Revenue |
| `604800` Achats de prestations intragroupe | `413505` Ingresos por servicios de consultoría | `6000` Salaries and Wages |
| `613200` Locations immobilières | `413595` Ingresos por servicios - intercompañía | `6010` Payroll Taxes |
| `625100` Voyages et déplacements | `510506` Sueldos | `6090` Intercompany Management Fees |
| `627800` Frais et commissions bancaires | `510568` Aportes parafiscales y seguridad social | `6200` Rent |
| `641100` Salaires, appointements | `511095` Honorarios - intercompañía | `6300` Travel |
| `645100` Cotisations à l'URSSAF | `512010` Arrendamientos - construcciones y edificaciones | `7000` Bank Fees |
| `706100` Prestations de services | `515505` Gastos de viaje - alojamiento y manutención | |
| `706800` Prestations de services intragroupe | `530505` Gastos bancarios (financial, class 53) | |

The default dataset is **byte-stable**: two builds give the same bytes, hence
the same `quickXorHash` and `size`. openpyxl rewrites `dcterms:modified` with
the wall clock on every save, and `zipfile` dates every archive member — both
are pinned (see `dataset/depot_finance.py`). Timestamps are fixed too: each
balance is dropped on the 6th of the following month (`lastModifiedDateTime`
e.g. `2026-02-06T08:14:00Z` for NTE January); `taux_2026.xlsx` is at version
6 (`cTag` `…,6`).

### Invalid variants — `/__fixtures/depot_finance`

They are **not** in the default folder (they would turn the consumer's CI
red). They are served, **without authentication** (synthetic test data,
useful without the control plane too), at
`GET /__fixtures/depot_finance/{name}`; `GET /__fixtures/depot_finance` lists
them with their defect. A test downloads one and drops it with
`PUT /__admin/drive/files/Insights360/{name}`. **One defect per variant, and
only one** — a consumer test expecting "rejected as unbalanced" must not pass
because the file was *also* rejected for its header.

| Name | | Defect |
|---|---|---|
| `balance_BOG_2026-07_desequilibree.xlsx` | invalid | Σ debit − Σ credit = 1,000.00 COP (account `515505`) |
| `balance_MTL_2026-07_formule.xlsx` | invalid | the debit of `6300` is a formula (`=a+b`) with **no cached value** (`data_only=True` reads `None`) |
| `balance_NTE_2026-07_sous_total.xlsx` | invalid | `6411` **and** `641100` in the same file — a prefix pair; balanced otherwise |
| `balance_BOG_2026-07_mauvaise_devise.xlsx` | invalid | `USD` on every row — BOG keeps its books in COP |
| `balance_NTE_2026-07_entete.xlsx` | invalid | hand-typed header (`Entité`, `Libellé`, `Débit`…) instead of the exact one |
| `balance_XXX_2026-07.xlsx` | invalid | unknown entity code `XXX`, in the name and the `entite` column |
| `taux_2026_eur.xlsx` | invalid | an `EUR` row (`moyen 2026-01 EUR 1`) |
| `~$balance_NTE_2026-01.xlsx` | invalid | Excel owner/lock file: 165 bytes, not a workbook. Microsoft documents `~$` names as invalid in SharePoint/OneDrive; the control plane accepts it anyway, so the consumer's filter can be exercised |
| `notes.csv` | invalid | outside the naming convention |
| `balance_NTE_2026-07.xlsx` | **valid** | one more month — to accept (incremental drop) |
| `balance_NTE_2026-06_corrigee.xlsx` | **valid** | June re-issued with a +250.00 EUR rent correction — drop it **as** `balance_NTE_2026-06.xlsx` to exercise an overwrite |

The list lives in one place: `FIXTURES` in `src/entra_mock/dataset/depot_finance.py`.

### Control plane — `/__admin`

Same mechanics as the other mocks of the ecosystem: mounted **only** when
`ENTRA_MOCK_ADMIN_ENABLED=true` (off means *not mounted*), every call carries
`X-Mock-Admin-Token: <ENTRA_MOCK_ADMIN_TOKEN>` (403 otherwise).

| Endpoint | Effect |
|---|---|
| `PUT /__admin/drive/files/{path}` | raw body = file bytes, `{path}` from the drive root (`Insights360/…`); **201** on create, **200** on overwrite — same item `id`, `eTag`/`cTag` `n+1`, new `lastModifiedDateTime`; missing folders are created; returns the driveItem |
| `DELETE /__admin/drive/files/{path}` | 204 (a folder goes with its content), 404 if absent. Re-dropping a deleted path yields a **new** item id, as in SharePoint |
| `POST /__admin/drive/reset` | back to the default dataset, counters and clock included |
| `GET /__admin/drive/counters` | `{site, drive, drives, item, children, content, download}` — authenticated, non-throttled calls since the last reset: proves that an idempotent second run downloads nothing |
| `POST /__admin/drive/grant` | `{"grant": "none"}` closes the site (403 everywhere), `{"grant": "read"}` reopens it |
| `POST /__admin/reset` | everything: the drive, its counters, the grant, and the throttling cadence |

**The mock clock.** A drop through the control plane must move
`lastModifiedDateTime` forward — that is what an incremental consumer
compares — but the wall clock would make every run different and the
consumer's idempotence snapshot unstable. Writes therefore use a monotonic,
deterministic clock: `2026-07-15T08:00:00Z` (after the whole default dataset)
plus one minute per write since the last reset. The same sequence of writes
yields the same timestamps, run after run.

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
| `ENTRA_MOCK_SITE_GRANT` | `read` | the app's `Sites.Selected` role on the site; `none` → 403 `accessDenied` on every drive route |
| `ENTRA_MOCK_DRIVE_PAGE_SIZE` | `2` | items per `/children` page — small so that pagination is always exercised |
| `ENTRA_MOCK_TEMPAUTH_SECONDS` | `3600` | lifetime of a pre-authenticated download URL |
| `ENTRA_MOCK_DOWNLOAD_BASE_URL` | *(request base URL)* | origin of the download URLs — set a second hostname to make the 302 cross-origin, as in production |
| `ENTRA_MOCK_ADMIN_ENABLED` | `false` | mounts `/__admin` |
| `ENTRA_MOCK_ADMIN_TOKEN` | `mock-admin-token` | expected in `X-Mock-Admin-Token` |
| `ENTRA_MOCK_HOST` / `_PORT` | `0.0.0.0` / `8000` | uvicorn bind |

## Development

```bash
make bootstrap   # uv sync
make test        # pytest — OAuth2 flow, pagination, dialect, datasets, drive
make lint        # ruff + mypy --strict
make contract    # regenerate contracts/msgraph.openapi.yaml
```

The version is bumped in lockstep in `pyproject.toml`,
`src/entra_mock/app.py` (`FastAPI(version=…)`), `docker-compose.yml` and the
contract (`make contract`; `tests/test_drive.py` fails on a stale contract).

A release is a tag `vX.Y.Z` on `main`: CI tests, smoke-tests and pushes
`ghcr.io/diametralgroup/entra-mock:X.Y.Z` (every push to `main` also moves
`:latest`).

## Versions

| Version | Dataset |
|---|---|
| `0.5.0` | same directory, plus the **SharePoint drive** surface and the `depot-finance` site (18 trial balances, the FX-rate file, invalid variants), the `/__admin` control plane — the current one |
| `0.4.0` | `grp-comex` renamed `grp-bi-comex`, the tenant's real name (breaking for a consumer keyed on the old name) |
| `0.3.0` | same UPNs, plus `grp-bi-imbrique` (a nested group + a service principal, no direct membership) and `/transitiveMembers`. Faithful to six dialect facts the mock used to get wrong |
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
