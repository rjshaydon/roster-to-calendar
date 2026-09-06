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

## Phase 4: shared On shift day snapshots

Status: local Phase 4 gate passed behind disabled-by-default build and reader
flags.

Completed locally:

- Changed ingestion reports the exact roster dates affected by event or staff
  changes. Unrelated ED/dates are not rebuilt.
- Added immutable, content-addressed ED/day objects and day pointers within the
  existing shared ED manifest. All authorised users reuse the same object.
- The shared day contains the base roster facts. Staff seniority overrides are
  read from the separately versioned shared Staff object and applied without
  rebuilding the day object.
- Added a durable per-ED publication record with a monotonic generation,
  operation/fencing token, bounded lease, base/candidate revisions and failure
  state.
- Publishers write immutable day and candidate-manifest objects before the
  fixed manifest pointer. The pointer uses the previous R2 ETag when available
  and is updated only after the D1 fencing token is rechecked.
- An unchanged date retains its existing object and an unchanged candidate
  performs zero R2 writes. A publication is capped at 120 affected dates.
- If object or pointer publication fails, the previous complete manifest stays
  readable. If the pointer succeeds but completion marking fails, retry
  recognises the committed operation and completes idempotently.
- An expired publisher cannot overwrite a newer owner. Roster deletion/reset
  can republish the bounded set of dates already present in the manifest.
- The authenticated On shift action reads shared R2 day/Staff objects and then
  applies the existing working-shift and Clinical Support rules. A miss returns
  `503 preparing`; it never queries or builds from `roster_events`.
- Direct day selection enforces the 14-day pre-term visibility boundary using
  the current Melbourne roster date, not merely the requested future date.

Safety controls:

- `FACILITY_SHARED_DAYS_BUILD_ENABLED` controls day publication.
- `FACILITY_SHARED_DAYS_ENABLED` controls the shared reader and is effective
  only when the Phase 2 access and Phase 3 metadata readers are also enabled.
- Both are false when absent. No deployed configuration has changed.

Focused Phase 4 evidence:

- The real authenticated `/api/state` On shift action returns the shared day
  while its traced D1 calls contain zero `roster_events` queries.
- A missing day returns `preparing` and performs no reader-time build or event
  fallback.
- Repeat publication reuses the immutable day and writes zero R2 objects.
- Injected failure before the manifest pointer retains the old day; retry
  publishes the corrected day.
- Two simulated Worker instances contend for one ED: the second is refused
  while the first lease is valid. After forced lease expiry, a new owner
  publishes and the stale worker is fenced out.
- Injected failure after the pointer leaves the new complete day readable and
  is recovered idempotently without another R2 write.
- The 15-day/14-day future-term boundary is tested against a known object key.
- Fresh migration `0029_facility_day_publications.sql` and local safety checks
  pass using local D1/R2 only.

Remaining before any online test:

- Phase 5 must remove repeated D1/contact-object discovery from the 60-second
  contact refresh loop. On shift roster reads are now shared, but its contact
  overlay remains on the existing path.
- Phase 6 must add safe browser persistence/revalidation and convert range
  views. Production migration/backfill, configuration and deployment remain
  separately prohibited until the rollout package is reviewed and approved.

## Phase 5: shared contact overlay

Status: local Phase 5 contact-overlay gate passed behind disabled-by-default
build and reader flags.

Completed locally:

- Contact ingestion publishes an immutable object keyed by source, operational
  date and content revision, plus a small conditionally updated source
  manifest. An identical extract performs no R2 writes.
- The On shift contact reader discovers the current or permitted night
  carryover extract entirely from predictable R2 keys. It performs no
  `contact_list_files` or contact-resolution D1 query.
- Manual allocation corrections remain separate from the base contact object.
  Their existing optimistic revision, same-shift and duplicate-allocation D1
  mutation checks are unchanged; a successful mutation republishes only the
  small resolution overlay.
- Browser refreshes send the revision already rendered. An unchanged response
  contains no contact payload and does not rerender the page.
- Polling now follows the agreed 60-second interval, stops while the document
  is hidden, runs immediately after it becomes visible, and stops when On
  shift is closed.
- Permission is still checked by the existing authenticated At a glance
  request and the Phase 2 access record, whose validity is capped at 15
  minutes. Contact objects are never exposed through an unauthorised public
  URL.

Safety controls:

- `FACILITY_SHARED_CONTACTS_BUILD_ENABLED` controls contact and correction
  publication.
- `FACILITY_SHARED_CONTACTS_ENABLED` controls shared contact reads and is
  effective only with Phase 2 access materialisation.
- Both are false when absent. Existing contact ingestion and reading remain
  unchanged until a separately approved rollout enables them.

Focused Phase 5 evidence:

- Sixty simulated unchanged one-minute refreshes use zero D1 reads, zero
  writes and three bounded R2 reads per refresh (manifest, date object and
  correction overlay).
- A repeated identical contact publication performs zero R2 writes.
- A correction changes only its independent overlay revision.
- MMC/MCH/DDH mappings, shift-change filtering, night carryover and the
  explicit exclusion of VHH contacts retain their existing tests.
- D1 quota guards, local network isolation, shared roster materialisation and
  facility-access tests remain green.

Remaining before any online test:

- The normal `/api/state` envelope still performs its existing small account
  authentication and materialised access checks. The contact refresh itself
  performs zero D1 contact discovery; removing all envelope reads would require
  a separately designed short-lived signed endpoint and is not being smuggled
  into this phase.
- Existing production contact extracts and corrections require a bounded
  publication step or a fresh normal ingestion before enabling the reader.
- Phase 6 browser persistence and range-view work, followed by a reviewed
  deployment/backfill package, remains outstanding. No production flag should
  be enabled yet.

## Phase 6: browser persistence and range views

Status: local Phase 6 gate passed. Phase 7 controlled rollout remains
separately prohibited pending review and approval.

Completed locally:

- Day publication also derives one immutable content-addressed month object
  for each affected ED/month. Updating one date rebuilds only its month; it
  does not scan D1 or rebuild unrelated months.
- By stream and Working together use the shared monthly objects when the Phase
  4 reader is enabled. Their request paths perform no `roster_events` or
  `roster_daily_presence` queries and retain the existing one-year UI limit.
- A one-year request is bounded to at most 13 monthly objects per requested ED,
  plus the small ED manifests and applicable Staff objects. Empty covered days
  remain represented without requiring event rows.
- Range and On shift responses carry content revisions. The browser sends its
  saved revision and unchanged responses omit the roster payload.
- Added schema-versioned IndexedDB storage for roster-only Metadata, Staff, On
  shift, By stream and Working together snapshots. Contact details are never
  persisted there.
- Cache keys include the exact viewed account, authorised access scope, action
  and canonical query. A record cannot be reused for another account, ED scope
  or query.
- Persisted data is rendered only while a current server-issued access expiry
  remains valid. Access records retain the established maximum 15-minute
  window. An expired or absent grant fails closed.
- A valid saved roster can remain visible with a clear last-saved warning when
  revalidation fails. Successful revalidation replaces the warning and stored
  revision.

Focused Phase 6 evidence:

- Actual authenticated By stream and Working together handlers read shared
  month snapshots with zero roster-history D1 queries.
- Repeated matching revisions return `unchanged` without retransmitting roster
  events; On shift can still refresh its non-persisted contact overlay.
- Browser-cache tests reject another account, another ED scope and expired
  authorisation, and fail closed when IndexedDB is unavailable.
- Existing incremental ingestion, publication fencing/recovery, contact,
  access, fixture, quota and local-isolation checks remain green.

Remaining:

- Phase 7 is operational rollout, not another automatic coding phase. Before
  any push or online test, prepare and review exact migrations, flags, bounded
  initial publication steps, expected D1/R2 costs and rollback checks.
- Existing online data is intentionally not backfilled by this branch. No
  production reads, writes, migration, deployment, automation or flag change
  has occurred.
