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

Status: in progress; live reads remain unchanged.

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
- Added compact coverage and term-staff repository reads with no fallback to
  `roster_events`. They are not connected to live handlers yet.

Still required for the Phase 1 gate:

- Replace chunked automated D1 staging with a bounded staged change set. The
  current automation still inserts a complete inactive D1 copy and promotion
  then deletes/moves whole-file rows; this remains a write-quota release blocker.
- Extend the focused integration test through that automated chunk/finalise
  route and prove unchanged and small-correction write counts.
- Confirm the explicit provider-ID schema preflight for each eventual remote
  environment before approving migration execution.
