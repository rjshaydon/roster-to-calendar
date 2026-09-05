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
