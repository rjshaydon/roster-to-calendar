# At a glance immediate D1 safety plan

## Purpose and authority boundary

This plan is the emergency prerequisite for resuming the Phase 7 production
rollout. Writing it does not authorise a deployment or any production D1
operation. Implement it only after separate approval.

The immediate objective is to make the known historical At a glance queries
unreachable before the next UTC quota day begins. At a glance may be
temporarily unavailable. Login, calendars, account management and other roster
features must not be intentionally disabled.

Do not apply a migration, inspect application data, bootstrap retained files,
publish shared objects, open At a glance, restore ingestion or enable an
automation as part of this plan.

## Confirmed incident state — 6 September 2026

- GitHub `main`, the rollout branch and the active Pages Production deployment
  identify `0fab128`.
- Production D1 is `roster-converter-calendar`, UUID
  `237d0d52-3a7c-4e02-8648-9f4dedbc1cb0`.
- Migration `0024_contact_allocation_resolutions.sql` is applied. Migrations
  `0025` through `0030` remain pending.
- Immediately after `0024`, Wrangler's rolling 24-hour snapshot reported
  143,110 rows read and 1,465 rows written.
- After the Creator opened At a glance for MMC, MCH, DDH and VHH and visited
  On shift and ED Staff, the rolling snapshot reported 5,260,178 rows read and
  1,558 rows written: approximately 5.12 million additional reads.
- Cloudflare's one-hour query insights attributed 3,440,682 reads to 11
  executions of the legacy ED Staff membership query. Each execution examined
  an average of 312,789 rows. Legacy active-file, coverage and staff-event
  queries accounted for further material reads.
- The migration was not the source of the spike. The deployed routing state
  was: shared rollout inactive, shared emergency pause active, and legacy reads
  explicitly permitted. The requests therefore used historical D1 SQL.

The rolling 24-hour figure is incident evidence, not the authoritative UTC
billing-day counter. It will continue to include pre-reset activity after
midnight UTC. Morning go/no-go decisions must use a metric bounded from
00:00 UTC for the new quota day.

## Required configuration change

Change exactly this safety setting for both the Production and Preview Pages
environments:

```text
FACILITY_LEGACY_READS_PAUSED=true
```

Keep all other Phase 7 settings inert:

- shared rollout inactive;
- shared emergency pause active;
- all shared build and reader switches false;
- all source/reader allowlists and cohorts empty;
- roster and contact writes false;
- queue, maintenance and watchdog false.

The Pages project manages these text variables through `wrangler.toml`. The
effective Production value changes only when the reviewed configuration commit
is deployed. Do not rely on a dashboard-only value or an absent variable.

## Serial implementation sequence

### I1. Freeze the unsafe path

Until this plan passes, do not open At a glance in Production or Preview. Do
not ask another user to test it. Record the current commit, active deployment
and effective variables without querying application tables.

Gate: the exact release identity is known and no production mutation has
occurred.

### I2. Prepare the configuration-only patch

On `codex/at-a-glance-d1-optimization`, change only the two explicit
`FACILITY_LEGACY_READS_PAUSED` values in `wrangler.toml` from `false` to `true`.
Do not change application code, migrations, bindings, compatibility dates,
automation controls or allowlists in this emergency commit.

Gate: the tracked diff contains only the two intended value changes.

### I3. Run focused local verification

Use the existing local harness. At minimum run:

- `npm run test:facility-rollout`
- `npm run test:d1-quota`
- `npm run test:local-isolation`
- `npm run check`
- `git diff --check`

Verify specifically that an inactive shared rollout plus paused legacy reads
returns the blocked route before Staff, metadata, range, On shift or contact
handlers can invoke their historical repository functions. Confirm that the
configuration diff does not alter login or unrelated API routing.

Gate: all focused checks pass locally and no remote resource was accessed.

### I4. Review, commit and deploy

Record the exact commit and obtain approval for the Production mutation. Push
the reviewed branch, fast-forward `main` to the exact commit and allow the
connected Pages project to deploy it. Do not run a migration or application
request during this step.

Gate: GitHub `main`, the rollout branch and Pages Production identify the same
reviewed commit.

### I5. Verify the deployed safety state without D1

Read back the effective Pages Production variables and bindings. Confirm:

- `FACILITY_LEGACY_READS_PAUSED=true`;
- `FACILITY_SHARED_ROLLOUT_ACTIVE=false`;
- `FACILITY_SHARED_EMERGENCY_PAUSED=true`;
- every builder, reader, ingestion, queue, maintenance and watchdog control
  remains inert; and
- the Production D1 and R2 bindings are unchanged.

Do not verify this by opening At a glance. Do not perform a login smoke test
while the D1 daily quota may be exhausted.

Gate: the deployed configuration is fail-closed and no D1 operation was used
to establish that fact.

### I6. Pause until the new UTC quota day

Make no further Production D1 calls tonight. In particular, do not apply
`0025`, inspect application tables, bootstrap a file, publish a term or restore
contacts/rosters.

After 00:00 UTC, obtain the authoritative new-day read and write usage from a
UTC-bounded Cloudflare metric. Do not use the rolling 24-hour total as the
go/no-go figure. Stop if the new day shows unexplained activity or less than
50% headroom.

Gate: the new UTC day is healthy before the main rollout resumes at isolated
migration `0025`.

## Rollback and failure handling

Do not roll back to a configuration that permits legacy reads. If the Pages
deployment fails, leave the prior deployment serving and correct the
configuration forward. If an unrelated regression is discovered, prepare a
forward patch that preserves `FACILITY_LEGACY_READS_PAUSED=true`.

An unavailable At a glance page is the intended safe failure mode until shared
objects exist. Quota exhaustion is not an acceptable availability fallback.

## Completion evidence

The implementation report must include:

- before/after commit and deployment IDs;
- the exact tracked diff;
- focused local test results;
- effective deployed safety variables and bindings;
- confirmation that no D1 migration, bootstrap, publication or application
  request occurred; and
- the morning UTC-day quota reading used to authorise or block `0025`.

