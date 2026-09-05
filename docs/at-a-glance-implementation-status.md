# At a glance implementation status

## Baseline

- Branch: `codex/at-a-glance-d1-optimization`
- Protected production base: `aa9eed8`
- Production deployment, remote D1/R2, external automations and provider services were not accessed or changed.
- The separate `codex/durable-doctor-identity-aliases` branch remains untouched.

## Phase 0: local safety harness and baseline

Status: in progress.

Completed locally:

- Adapted the deterministic local lifecycle without importing Doctor ID feature code.
- Added fail-closed outbound-network protection for automation, FindMyShift, GitHub dispatch and email.
- Added the missing explicit schema migration for tables previously created by runtime repair; ordinary requests remain DDL-free.
- Verified reset, all migrations, idempotent synthetic seeding, Creator/user login, incorrect-password rejection and restart persistence using local Wrangler D1/R2 only.
- Closed the paused `/api/automation/pending` gap before any D1 access and made queue listing read-only.
- Revalidated the existing D1 emergency guards and facility-access regression tests.

Tests passed:

- `npm run local:check` on loopback port 8798
- `npm run test:local-lifecycle` on loopback port 8799
- `npm run test:local-isolation`
- `npm run test:d1-quota`
- `npm run test:facility-access`
- `npm run check`

Still required for the Phase 0 gate:

- Build the large deterministic synthetic fixture.
- Capture query plans and reproducible read/write estimates for Staff, coverage, On shift, access and contacts.
- Add broad-query regression gates and a whole-account capacity worksheet.
- Record the external deployment/binding inventory at rollout time without enabling or mutating it.
