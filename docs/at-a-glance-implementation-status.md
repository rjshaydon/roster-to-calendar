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
