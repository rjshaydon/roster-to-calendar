# At a glance implementation status

## Baseline

- Branch: `codex/at-a-glance-d1-optimization`
- Protected production base: `aa9eed8`
- Production deployment, remote D1/R2, external automations and provider services were not accessed or changed.
- The separate `codex/durable-doctor-identity-aliases` branch remains untouched.

## Phase 0: local safety harness and baseline

Status: gate passed locally.

Completed locally:

- Adapted the deterministic local lifecycle without importing Doctor ID feature code.
- Added fail-closed outbound-network protection for automation, FindMyShift, GitHub dispatch and email.
- Added the missing explicit schema migration for tables previously created by runtime repair; ordinary requests remain DDL-free.
- Verified reset, all migrations, idempotent synthetic seeding, Creator/user login, incorrect-password rejection and restart persistence using local Wrangler D1/R2 only.
- Closed the paused `/api/automation/pending` gap before any D1 access and made queue listing read-only.
- Revalidated the existing D1 emergency guards and facility-access regression tests.
- Built a focused deterministic fixture with 5 files, 600 doctors and 109,200 events.
- Recorded baseline query plans and separate future acceptance gates in `at-a-glance-phase-0-baseline.md`.
- Completed the 50-visible-page capacity worksheet using a 12-hour day and 60-second refresh interval.

Tests passed:

- `npm run local:check` on loopback port 8798
- `npm run test:local-lifecycle` on loopback port 8799
- `npm run test:local-isolation`
- `npm run test:d1-quota`
- `npm run test:facility-access`
- `npm run check`
- `npm run test:database-costs`
- Existing roster, Excel/FindMyShift, contact and queue correctness suites

Deferred to the separately approved rollout: verify external deployment and
binding inventory and compare these estimates with actual Cloudflare metrics.

## Phase 1: coverage and term staff materialisation

Status: local Phase 1 gate passed; live reads remain unchanged.

Completed locally:

- Added explicit compact storage for per-file coverage, per-file term-staff
  contributions and the 14-day term visibility boundary.
- Added explicit provider staff ID columns previously supplied only by runtime
  schema repair.
- Verified the migration from a completely fresh disposable local database.
- Added compact read models that inspect only active coverage rows and the
  selected ED/term staff contributions, never `roster_events`.
- The normal complete-import path now hashes canonical parsed content and
  returns before all writes for an identical import.
- A changed normal import diffs doctors, events and issues and updates daily
  presence only for changed/removed events.
- Activation, promotion, deletion, overlap trimming and daily-presence repair
  now refresh or remove the affected compact facts.
- Local integration tests cover zero-write identical imports, one-event
  sickness correction, overlapping file contributions, SMS continuity and
  the exact 14-day visibility boundary.
- Automated routine ingestion now submits one complete parsed change set and
  diffs it against the source's stable active file instead of writing a full
  inactive copy in chunks and then promoting/deleting whole-file rows.
- An unchanged automated import stops before supersession, membership,
  presence and snapshot work. Only the small sync-run/source bookkeeping
  records are updated by the automation endpoint.
- Existing-file revisions above the automatic 250-fact budget fail before any
  roster write. The hard ceiling is 500 facts; larger changes require a later,
  explicitly controlled ingestion path.
- A newly discovered roster is populated while inactive and made visible only
  after its bounded core insert succeeds.
- Added compact coverage and term-staff repository reads with no fallback to
  `roster_events`. They are not connected to live handlers yet.

Focused Phase 1 evidence:

- `test:facility-materialization`: exact repeat writes zero roster facts; one
  sickness correction changes one event and writes at most six rows; an
  over-budget revision writes zero rows; overlap, SMS continuity and 14-day
  visibility pass. The same test invokes the token-protected automation
  handler end to end: a repeat and a one-event correction both reuse the
  stable active file and create no inactive D1 copy.
- `test:database-costs`: compact Coverage returns one row and compact Staff 120
  contribution rows on the 109,200-event fixture, with zero `roster_events`
  access in either query plan.
- Fixture, queue-failure, D1-quota, local-isolation, VHH automation,
  facility-access and contact safeguards pass.
- Fresh local migration and safety check passes using local D1/R2 only.

Still required before any online rollout:

- Confirm the explicit provider-ID schema preflight for each eventual remote
  environment before approving migration execution.
- Implement later phases that switch live Staff/Coverage/On-shift handlers to
  the compact read/cache path. Until then, this branch must not be deployed as
  the production quota fix.

## Phase 2: access/session materialisation

Status: local Phase 2 gate passed behind a disabled-by-default reader flag.

Completed locally:

- Added one short-lived compact access decision per account. It records the
  authorised site scope, SMS/all-site status, preferred site, working-today
  state, roster date and term.
- The decision is keyed by a digest of the account role, explicit feature
  entitlement, non-clinical status and source-specific claims. Account or
  claim changes therefore cannot reuse an older decision.
- Decisions expire at the next fixed 15-minute boundary and are also rejected
  whenever the Australia/Melbourne roster date changes. Explicit feature
  revocation is checked before reading the access cache and is immediate.
- Cache misses derive access only from exact compact term-staff, continuing-SMS
  and daily-presence lookups. They never inspect `roster_events`.
- Missing compact evidence fails closed with `503 preparing` and `Retry-After`;
  that negative result is itself cached so multiple visible pages cannot keep
  probing the compact tables.
- Creator/owner and eligible non-clinical accounts retain all-site access.
  Creator-entered views continue to key access from the entered account, not
  the Creator.
- `FACILITY_ACCESS_MATERIALIZATION_ENABLED` guards the new handler path and is
  false when absent. No production or preview configuration has been changed.

Focused Phase 2 evidence:

- `test:facility-access` verifies term-specific trainee access, working-today
  preference, ambiguous access denial, permanent SMS continuity, immediate
  revocation, claim invalidation, roster-date invalidation, 15-minute expiry,
  positive and negative cache reuse, and entered-account handler wiring.
- A repeated decision performs one indexed primary-key read, zero writes and
  zero `roster_events` queries. A cold decision performs at most three exact
  compact lookups per claimed source identity plus one cache read and one
  idempotent single-row publication.
- With 50 visible pages refreshing every 60 seconds, the access layer's steady
  state is at most 50 single-row cache reads per minute. Fixed expiry values
  and a conditional upsert prevent duplicate refreshes from changing the row
  after the first equivalent publisher succeeds.

Still required before enabling the flag online:

- Apply and verify migration `0027_facility_access_materialisation.sql` in a
  separately approved rollout with all live readers and automations disabled.
- Populate/verify Phase 1 compact facts before enabling this reader; otherwise
  affected users correctly receive `preparing` rather than an event-table
  fallback.
- Phase 3 must route shared Staff and metadata reads through this access layer
  and their compact objects before any online test.

## Phase 3: shared Staff and metadata reads

Status: local Phase 3 gate passed behind disabled-by-default build and reader
flags.

Completed locally:

- Added per-file, per-term stream-catalogue contributions during ingestion.
  Catalogue updates are diffed alongside coverage and staff facts and never
  require a reader-time `GROUP BY roster_events` query.
- Added one shared manifest per ED and immutable, content-addressed Staff
  objects per ED/term in the local R2 substitute.
- Staff objects include compact membership, SMS continuity, coverage,
  designation and seniority-override data. Metadata manifests include compact
  coverage, visibility boundaries and stream signatures.
- Publication compares stable content revisions. An unchanged publication
  retains existing object keys and performs zero R2 writes.
- Staff and metadata handlers use the Phase 2 access decision before reading
  shared objects. Site-scoped users cannot request another ED or All EDs, and
  Creator-entered views retain the entered account's restrictions.
- Missing manifests or Staff objects return `503 preparing` with
  `Retry-After`; readers never build objects or fall back to legacy Staff or
  catalogue SQL.
- The 14-day pre-term boundary is enforced when Staff/catalogue objects are
  selected. The object may be prepared earlier but is not returned early.
- Roster ingestion, removal/reset and staff overlay changes schedule a shared
  Staff/metadata refresh only when the separate build flag is enabled.

Safety controls:

- `FACILITY_SHARED_METADATA_BUILD_ENABLED` controls publication.
- `FACILITY_SHARED_METADATA_ENABLED` controls reads and is effective only
  while `FACILITY_ACCESS_MATERIALIZATION_ENABLED` is also enabled.
- Both new settings are false when absent. Repository production and preview
  configuration remains unchanged.

Focused Phase 3 evidence:

- The materialisation integration test invokes the actual authenticated
  `/api/state` Metadata and Staff action chain with both flags enabled.
- Repeated shared reads perform zero D1 `roster_events` queries; after ordinary
  authentication and the one-row access decision, payload retrieval is R2
  only and is shared across accounts.
- A first publication creates immutable Staff data and its manifest. Repeating
  publication without a fact change performs zero R2 writes.
- A missing R2 object returns `preparing` from the handler and performs no
  legacy Staff query.
- Stream catalogue, Staff membership, overlap preservation, SMS continuity,
  access restrictions and the exact 14-day visibility boundary pass locally.
- Fresh migration and local safety checks pass using local D1/R2 only.

Still required before an online test:

- Migration `0028_facility_stream_catalog_materialisation.sql` and the Phase 1
  compact facts require a separately approved, measured, bounded backfill.
  Identical existing imports intentionally do not rewrite facts merely to fill
  a newly added projection.
- Phase 4 must replace the high-value On shift reader with shared ED/day
  snapshots and add the durable publication/recovery state machine. Until that
  gate passes, this branch is not a production quota fix and must not be
  deployed or have its new flags enabled.
