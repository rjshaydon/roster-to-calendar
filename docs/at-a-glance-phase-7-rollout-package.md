# At a glance Phase 7 rollout package

> **Rollout suspended 7 September 2026.** Read and complete
> [`d1-account-quota-safe-rollout-plan.md`](./d1-account-quota-safe-rollout-plan.md)
> before any further Production or Preview D1 request. The account-wide plan
> supersedes the quota evidence, 50% headroom thresholds, bootstrap controls
> and canary order below. Per-database Metrics are diagnostic only and cannot
> authorise work.

This is an operational worksheet, not permission to deploy. Each remote step
requires separate approval and must be performed serially.

## Release identity

- Branch: `codex/at-a-glance-d1-optimization`
- Protected production base: `aa9eed8`
- Reviewed code commit: `20c4824`
- Deployed release candidate: `0fab128`
- GitHub `main` and the rollout branch identify `0fab128`.
- Active Pages Production deployment: `42af062c-cbe8-4eb1-8912-d0ad7abef9df`.
- Doctor identity work remains outside this rollout.
- The 2026-09-06 Phase 7B preflight accessed production metadata and D1 only
  through read-only commands. It made no production changes and did not
  deploy, migrate, bootstrap, publish or change configuration.

### Current checkpoint — 7 September 2026

- Current fail-closed Production commit is `a293251`; explicit configuration
  deployment is `5ac63a19`. All optional readers, builders, automation and
  maintenance are paused.
- Migrations `0026` through `0031` have been applied individually and no
  migration remains pending.
- A protected bootstrap inspection was rejected at 04:55 UTC because the
  account had exhausted its daily row-read allowance. No bootstrap executed.
- Do not resume at the old canary sequence. Resume only at Phase A of the
  account-wide quota safety plan linked above.

- At the earlier pre-suspension checkpoint, GitHub `main`, the rollout branch
  and active Pages Production deployment identified `95bc9ad`
  (`d692bdc9-95a7-4d82-a6d1-47fb9a13851a`).
- At that earlier checkpoint, migrations `0024` and `0025` were applied and
  `0026` through `0030` remained pending.
- Legacy At a glance reads are paused. The Creator `calendarStoreStatus` route
  is also guarded before D1 while roster writes are paused.
- Before roster writes are restored, implement the separate
  [calendar store status remediation plan](./calendar-store-status-d1-remediation-plan.md).
  Its proposed `0031` migration follows, and does not modify, the existing
  `0026`–`0030` sequence.
- The locally reviewed `0031_roster_file_status_summaries.sql` SHA-256 is
  `9beb3aa67bf7445a0950d2390818de750c4acf1f69a8f2eb31bb62e810866311`.
  Recompute it from the committed release before any remote migration.

This checkpoint supersedes older release and ledger statements below where
they describe the then-current state. Their measurements remain historical
rollout evidence.

Before deployment, record the exact release and production commits, migration
ledger, effective settings, D1 daily usage and UTC reset time. Stop if any state
differs from the reviewed record.

The original branch push caused an automatic Preview deployment at `20c4824`.
Gate 4 later fast-forwarded GitHub `main` and deployed Production at `0fab128`.
Do not exercise Preview endpoints during a production preflight, and do not
treat the existence of Preview as authority to merge or deploy Production.

## Mandatory incident correction — 6 September 2026

Read [the immediate D1 safety plan](./at-a-glance-immediate-d1-safety-plan.md)
before any further rollout action. It supersedes every statement in this
worksheet that permits temporary Production fallback to historical At a glance
SQL.

After migration `0024`, the Creator opened four hospitals' At a glance, On
shift and ED Staff views. Wrangler's rolling 24-hour total rose from 143,110 to
5,260,178 rows read. One-hour D1 insights attributed 3,440,682 reads to 11
executions of the legacy ED Staff membership query, averaging 312,789 examined
rows per execution. This proves that even a small compatibility canary is not
safe.

The cause was the deployed combination of an inactive shared rollout and
`FACILITY_LEGACY_READS_PAUSED=false`. The shared emergency pause protected the
new builders/readers; it did not stop the legacy fallback. Therefore:

- deploy `FACILITY_LEGACY_READS_PAUSED=true` before `0025` or any further D1
  work;
- keep legacy reads paused for the remainder of the rollout and permanently
  thereafter;
- never test an ED until its shared objects exist and its exact shared route is
  enabled; and
- return `preparing` or unavailable when a shared object or permission is
  missing. Never use historical SQL as a compatibility or rollback route.

Migration `0024_contact_allocation_resolutions.sql` is applied and recorded as
ledger row 24. Migrations `0025` through `0030` remain pending. Do not repeat
`0024`.

## Verified Phase 7B preflight baseline

Read-only observations recorded on 2026-09-06:

- GitHub `main` and the latest Pages Production deployment both identify
  `aa9eed8`. The remote rollout branch identifies `20c4824`.
- Production D1 is `roster-converter-calendar`, UUID
  `237d0d52-3a7c-4e02-8648-9f4dedbc1cb0`, in region `OC`, with a reported size
  of 73,850,880 bytes.
- Wrangler's rolling 24-hour snapshot reported 142,417 rows read, 1,520 rows
  written, 9,498 read queries and 608 write queries. This is evidence of ample
  headroom at that moment, not the authoritative UTC billing-day counter.
- The Creator subsequently verified the production database's Cloudflare
  Metrics page directly. It showed approximately 150,000 rows read over the
  last 24 hours; the separate Preview database showed zero rows read. Because
  the elapsed portion of the current UTC billing day is wholly contained in
  that trailing 24-hour interval, current-day reads cannot exceed 150,000
  (3% of the 5,000,000-row daily limit). The contemporaneous Wrangler snapshot
  reported 1,445 rows written over 24 hours (1.445% of the 100,000-row daily
  limit). Gate 2's requirement for at least 50% headroom therefore passes at
  this checkpoint. Refresh both figures immediately before any mutation.
- One bounded schema inventory consumed 186 rows read and zero rows written.
- The largest observed read class was the existing doctor/date
  `roster_events` lookup: 119,378 total rows read across 416 calls. This is a
  current legacy workload, not evidence that the new shared reader is active.
- Contact ingestion was still active under Production code and accounted for
  approximately 293 `contact_list_files` inserts and 303 retention deletes in
  the observed window. The inert release will pause this workload.
- At this initial preflight checkpoint, the migration ledger reported `0024`
  through `0030` as unapplied. The current state is recorded in the mandatory
  incident correction above: `0024` applied, `0025` through `0030` pending.
- The tables described by `0024` and `0025` already exist, consistent with the
  former runtime-schema path. The subsequent Gate 1 inspection confirmed that
  their columns, constraints and named indexes match the migration files.
- The new materialisation tables from `0026` through `0029` were not present in
  the bounded schema inventory.
- Existing tables touched by `0030` are small at this checkpoint:
  `facility_staff_designations` has 0 rows,
  `facility_staff_seniority_overrides` has 41 rows and `roster_files` has 12
  rows. Recheck immediately before migration; do not rely on these counts later.
- Gate 1's exact-name schema query and PRAGMA checks read 963 rows and wrote
  zero. All six named indexes from `0024`/`0025` already exist. The indexed
  table counts were: contact resolutions 22, resolution history 30,
  subscription tokens 42, parser rules 146, console messages 50 and roster
  dispatches 529. None reached its 10,001-row sentinel.
- The read-only Pages configuration download showed that Production explicitly
  defines only `ROSTER_AUTOMATION_WRITES_ENABLED=false` from the inert-settings
  table. Every other listed rollout, emergency, allowlist, queue, maintenance
  and contact setting is absent. The current code defaults the new activation
  switches and allowlists closed, but absence is not accepted as the recorded
  Production configuration for deployment.

No subsequent mutating step may use these figures without refreshing the
deployment identity, UTC-day quota position, migration ledger and relevant
table counts. If the authoritative UTC-day quota position cannot be obtained,
the mutating step remains blocked.

## Inert configuration

| Purpose | Setting | Inert value |
| --- | --- | --- |
| Activate cohort routing | `FACILITY_SHARED_ROLLOUT_ACTIVE` | `false` |
| Stop shared canary/builds | `FACILITY_SHARED_EMERGENCY_PAUSED` | `true` |
| Stop historical reads | `FACILITY_LEGACY_READS_PAUSED` | `true` permanently |
| ED build allowlist | `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST` | empty |
| ED reader allowlist | `FACILITY_SHARED_READER_SOURCE_ALLOWLIST` | empty |
| Reader cohort | `FACILITY_SHARED_READER_COHORT` | empty |
| Access decision writes | `FACILITY_ACCESS_MATERIALIZATION_ENABLED` | `false` |
| Metadata builds | `FACILITY_SHARED_METADATA_BUILD_ENABLED` | `false` |
| Day builds | `FACILITY_SHARED_DAYS_BUILD_ENABLED` | `false` |
| Contact-overlay builds | `FACILITY_SHARED_CONTACTS_BUILD_ENABLED` | `false` |
| Shared metadata reads | `FACILITY_SHARED_METADATA_ENABLED` | `false` |
| Shared day reads | `FACILITY_SHARED_DAYS_ENABLED` | `false` |
| Shared contact reads | `FACILITY_SHARED_CONTACTS_ENABLED` | `false` |
| Roster writes | `ROSTER_AUTOMATION_WRITES_ENABLED` | `false` |
| Manual roster writes | `MANUAL_ROSTER_WRITES_ENABLED` | `false` |
| Compact roster status | `ROSTER_STATUS_SUMMARY_ENABLED` | `true` in Production; this bounded reader is independent of the facility canary |
| Roster source allowlist | `ROSTER_AUTOMATION_SOURCE_ALLOWLIST` | empty |
| Queue | `ROSTER_AUTOMATION_QUEUE_ENABLED` | `false` |
| Advanced maintenance | `ROSTER_ADVANCED_MAINTENANCE_ENABLED` | `false` |
| Contact ingestion | `CONTACT_AUTOMATION_WRITES_ENABLED` | `false` |
| Contact source allowlist | `CONTACT_AUTOMATION_SOURCE_ALLOWLIST` | empty |
| External watchdog | `ROSTER_AUTOMATION_ENABLED` | `false` |

Keep all access, metadata, day and contact shared build/reader settings false.
Missing allowlists fail closed. Keep legacy reads paused even during the
Creator canary. Contact ingestion is not enabled by the roster automation
switch.

The table above is the required deployed state, not an assumption about the
current Pages environment. Before merging, capture the effective Production
variables and explicitly set every inert value. Gate 3 confirmed that Pages
text variables for this project are managed through `wrangler.toml` and cannot
be edited separately in the dashboard. The release configuration must
therefore declare every value for both Production and Preview; re-read the
effective Production values after deployment.

Pausing contact ingestion is an intentional safety effect of the inert deploy.
Existing contact objects may remain readable only while their existing
authorisation/freshness rules permit; expired contact information must not be
extended. Keep this pause short, tell the Creator when it begins, and restore
one contact source only through the separate guarded restoration gate below.

## Migrations

Migration `0024` was applied in isolation on 2026-09-06 and is recorded in the
Production ledger. The pending ledger now lists `0025`, `0026`, `0027`, `0028`,
`0029` and `0030`. Do not reapply `0024`, and do not describe `0025` as ordinary
empty-table creation: its tables already exist outside the ledger.

Before applying anything:

1. Inspect `PRAGMA table_info`, `PRAGMA index_list` and the relevant index SQL
   for every `0024`/`0025` table, using exact-name schema queries only.
2. Compare the result field-for-field with the migration files and record row
   counts for every table on which either migration may create an index.
3. If the schema differs, stop. `CREATE TABLE IF NOT EXISTS` will not repair an
   existing table. Prepare and locally test a separate explicit reconciliation
   migration instead of editing an old migration or relying on runtime DDL.
4. If the schema matches, treat `0024` and `0025` as ledger reconciliation that
   may still create missing indexes. State the maximum expected index work from
   the refreshed counts before approval.

Gate 1 completed on 2026-09-06 with an exact schema match and every expected
index already present. `0024` then completed with no change in reported
database size or rolling usage at displayed precision. The expected
application-table and index work for `0025` remains zero; it should only need
its D1 migration-ledger update and Cloudflare's normal migration bookkeeping.
Recheck the schema and ledger immediately before applying it.

Only after the immediate safety plan passes and the new UTC quota day is proven
healthy may the remaining migrations be applied individually in this order:
`0025`, `0026`, `0027`, `0028`, `0029`, then `0030`. Re-read the ledger,
database size and authoritative UTC-day quota after each migration. Never batch
migrations, and never retry an uncertain result without first checking whether
it committed.

Wrangler's `d1 migrations apply` command has no single-migration selector and
would apply every pending file visible in its configured migrations directory.
Do not run it against this repository's normal migration directory while more
than one migration is pending. For each migration, prepare an isolated temporary
Wrangler configuration and migration directory containing exactly that one
unchanged reviewed SQL file. Verify its checksum, run `migrations list` and
proceed only if exactly one expected filename is shown. Apply it, re-read the
production ledger and quota, then discard the temporary files. This mechanism
must be rehearsed against disposable local D1 before its first production use.

The mechanism was rehearsed locally on 2026-09-06 using an isolated copy of
`0024` whose SHA-256 matched the repository file
(`8132442dff613a272f18d2d9a2cbbc540d9c03c1eaecec5f297c17b61c056265`).
Wrangler listed exactly that migration, applied it over an already-existing
matching schema, recorded exactly one `d1_migrations` row and then reported no
pending migrations. The disposable local database and configuration were
removed afterward. This rehearsal is evidence for the mechanism only; it does
not authorise or predict a production application result.

Migrations `0026`–`0029` create new empty materialisation structures and do not
backfill data. Migration `0030` creates two indexes over those new empty tables
and three over existing tables; the preflight counts for the existing tables
were 0, 41 and 12 rows. Recheck those counts before applying it. No migration
step authorises bootstrap, publication or a reader switch.

## Serial existing-data bootstrap

Use token-protected `POST /api/automation/facility-bootstrap`. Every request
names exactly one ED and one active roster file.

Inspection omits `execute`, reads only file metadata and its compact marker,
and returns a `planRevision`. Execution requires identical inputs,
`"execute": true` and that revision.

The executor performs one indexed `file_id` event read capped at 25,001 rows,
without a database sort, and applies separate sentinel limits to doctors and
existing per-file facts. It permits at most 750 proposed compact mutation statements in one D1 batch
(conservatively estimated as at most 2,250 rows written including indexes). An extra event row, an
excess write count or a stale plan performs zero compact writes. There is no
all-files mode, pagination, background retry or automatic continuation.

Process one file, disable advanced maintenance, inspect actual usage and check
the resulting facts before considering another file. Other active-file
contributions must remain intact.

## One-term publication

Use `POST /api/automation/facility-materialize` only after compact facts exist.
Every request names one ED and one actual medical-term start and is capped at
120 dates.

Dry run returns exact dates, months, a `planRevision` and distinct ceilings for
D1 read statements, estimated rows examined, returned rows,
publication-state writes, R2 reads and R2 writes. Rows examined remain labelled
as local estimates, not Cloudflare measurements.

Execution requires identical inputs plus the returned revision. Changed Staff,
SMS, designation, override, visibility, compact roster or manifest inputs make
it stale before writes. Inputs are checked again before the fixed manifest
pointer changes. Staff publication touches only the requested term and
preserves other terms. Its R2 ceiling includes the
Staff object, Staff/metadata pointer, day/month objects, candidate manifest and
fixed manifest.

Reject unexpected dates, more than 120 dates, any broad source/history scan or
a plan without comfortable quota headroom.

## Canary sequence

Each numbered mutation is a separate approval gate. Finish and record one gate
before requesting the next.

1. Reconcile the `0024`/`0025` schema and indexes using read-only inspection.
   Completed on 2026-09-06 with an exact match.
2. Refresh the Git/Pages release identity, authoritative UTC-day D1 usage,
   migration ledger and effective Production settings. Completed before the
   initial deployment with more than 50% headroom.
3. Declare and locally verify every inert configuration value above for both
   Production and Preview. The initial configuration incorrectly permitted
   legacy reads; the immediate D1 safety plan must correct that value to true.
4. Deploy the exact reviewed release commit to Production. Initial deployment
   `0fab128` completed, but its legacy-read value is unsafe. Complete the
   separately approved configuration-only safety deployment before continuing.
5. Migration `0024` completed in isolation. After the new UTC quota day begins,
   apply `0025` through `0030` individually under the migration rules above,
   with all new activity paused and a quota/ledger check after every migration.
6. Run a bounded smoke check proving ordinary login remains functional and all
   shared builders/readers remain inactive. Do not browse At a glance through
   the legacy path merely as a smoke test.
7. Allow one ED for building. Enable only advanced maintenance, inspect one
   retained file, then separately approve its bootstrap. Keep general roster
   writes disabled and disable maintenance immediately afterward.
8. Repeat only for files needed by the same ED/term, checking usage each time.
9. Dry-run one ED/term publication. Separately approve and execute only that
   plan, then disable maintenance and inspect usage.
10. Allow that ED for readers, select cohort `creator`, activate shared rollout
   and enable access/metadata/day/contact readers in dependency order. Test
   only the Creator's own view.
11. Confirm no other account's login or access calculation changed. Creator
   impersonation must not enter the canary.
12. Observe actual D1/R2 metrics through an agreed quiet interval. Stop on any
   unexplained growth.
13. Bootstrap, publish and validate further EDs serially. Never allow an ED to
   read before its objects exist.
14. Before cohort `all`, re-prove that the already permanent
    `FACILITY_LEGACY_READS_PAUSED=true` setting and the code-level fail-closed
    default prevent shared, missing and disallowed routes from reaching
    historical SQL. Do not change the setting during or after this proof.
15. Restore one contact source with its exact source allowlist. Observe one
    changed extract and repeated unchanged refreshes; unchanged refreshes must
    write zero D1 rows and zero R2 objects. Pause again on any unexplained
    insertion or retention loop before adding another source.
16. Restore roster sources separately. Restore queue/watchdog last.

During the short Creator canary, accounts outside the cohort receive
`preparing` or unavailable. They must not retain a legacy compatibility route.
A canary request also never falls back: missing or disallowed data returns
`preparing` or unavailable.

## Capacity and acceptance

The planning envelope is 50 visible On shift pages for 12 hours at a 60-second
interval: 36,000 requests/day. Unchanged contact refresh must perform zero D1
contact reads/writes and only its documented bounded R2 reads.

Local evidence uses the 109,200-event fixture and is not a Cloudflare bill.
Production acceptance requires:

- no `roster_events`/`roster_daily_presence` scan on shared or blocked reads;
- no Production route to legacy At a glance SQL, including during canary,
  missing-object, error and rollback states;
- unchanged ingestion rewrites no roster or cache facts;
- paused contact ingestion touches neither D1 nor R2;
- bootstrap/publication remain within their reviewed ceilings;
- actual usage matches request volume and estimates; and
- at least 50% of each daily quota remains after the test window.

Investigate and pause optional work at 50% usage. Stop it by 70%, or immediately
on an unexplained spike.

## Rollback

Rollback must not restore expensive SQL:

1. Confirm `FACILITY_LEGACY_READS_PAUSED=true`; it must already be permanent.
2. Set `FACILITY_SHARED_EMERGENCY_PAUSED=true` before disabling faulty shared
   readers. At a glance may temporarily be unavailable.
3. Disable shared readers/builders and empty their allowlists/cohort.
4. Disable roster/contact writes, queue, maintenance and watchdog.
5. Confirm other app functions remain available and inspect usage.
6. Retain compact rows and immutable objects for diagnosis. Do not repair,
   replay, clean up or republish during the incident.

## Evidence ledger

For every checkpoint record UTC time, commit, settings, ED/file/term, plan
output, requests, D1 rows read/written before and after, R2 operations,
correctness and rollback result. Cloudflare numbers and local estimates must
remain clearly distinguished.
