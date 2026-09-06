# At a glance Phase 7 rollout package

This is an operational worksheet, not permission to deploy. Each remote step
requires separate approval and must be performed serially.

## Release identity

- Branch: `codex/at-a-glance-d1-optimization`
- Protected production base: `aa9eed8`
- Release candidate: the reviewed remediation commit containing this worksheet
- Doctor identity work remains outside this rollout.
- Production services have not been accessed while preparing this package.

Before deployment, record the exact release and production commits, migration
ledger, effective settings, D1 daily usage and UTC reset time. Stop if any state
differs from the reviewed record.

## Inert configuration

| Purpose | Setting | Inert value |
| --- | --- | --- |
| Activate cohort routing | `FACILITY_SHARED_ROLLOUT_ACTIVE` | `false` |
| Stop shared canary/builds | `FACILITY_SHARED_EMERGENCY_PAUSED` | `true` |
| Stop historical reads | `FACILITY_LEGACY_READS_PAUSED` | `false` initially |
| ED build allowlist | `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST` | empty |
| ED reader allowlist | `FACILITY_SHARED_READER_SOURCE_ALLOWLIST` | empty |
| Reader cohort | `FACILITY_SHARED_READER_COHORT` | empty |
| Roster writes | `ROSTER_AUTOMATION_WRITES_ENABLED` | `false` |
| Roster source allowlist | `ROSTER_AUTOMATION_SOURCE_ALLOWLIST` | empty |
| Queue | `ROSTER_AUTOMATION_QUEUE_ENABLED` | `false` |
| Advanced maintenance | `ROSTER_ADVANCED_MAINTENANCE_ENABLED` | `false` |
| Contact ingestion | `CONTACT_AUTOMATION_WRITES_ENABLED` | `false` |
| Contact source allowlist | `CONTACT_AUTOMATION_SOURCE_ALLOWLIST` | empty |
| External watchdog | `ROSTER_AUTOMATION_ENABLED` | `false` |

Keep all access, metadata, day and contact shared build/reader settings false.
Missing allowlists fail closed. Contact ingestion is not enabled by the roster
automation switch.

## Migrations

Inspect the production ledger and quota headroom first. Apply only missing
migrations, one at a time: `0025`, `0026`, `0027`, `0028`, `0029`, then
`0030`.
Migrations `0025`–`0029` create empty structures and do not backfill data.
Migration `0030` adds bounded-maintenance indexes; four should cover empty new
tables and one covers the small retained-file table. Record usage after each
and never retry an uncertain result without rechecking the ledger.

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

1. Deploy inertly. Do not bootstrap or publish.
2. Apply reviewed empty migrations individually with new activity paused.
3. Allow one ED for building. Enable only advanced maintenance, inspect one
   retained file, then separately approve its bootstrap. Keep general roster
   writes disabled and disable maintenance immediately afterward.
4. Repeat only for files needed by the same ED/term, checking usage each time.
5. Dry-run one ED/term publication. Separately approve and execute only that
   plan, then disable maintenance and inspect usage.
6. Allow that ED for readers, select cohort `creator`, activate shared rollout
   and enable access/metadata/day/contact readers in dependency order. Test
   only the Creator's own view.
7. Confirm no other account's login or access calculation changed. Creator
   impersonation must not enter the canary.
8. Observe actual D1/R2 metrics through an agreed quiet interval. Stop on any
   unexplained growth.
9. Bootstrap, publish and validate further EDs serially. Never allow an ED to
   read before its objects exist.
10. Before cohort `all`, set `FACILITY_LEGACY_READS_PAUSED=true` and prove
    shared, missing and disallowed routes cannot reach historical SQL.
11. Restore roster and contact sources separately. Restore queue/watchdog last.

During the short Creator canary, accounts outside the cohort may retain the
explicit legacy compatibility route. A canary request never falls back:
missing or disallowed data returns `preparing` or unavailable. Remove legacy
compatibility before broad rollout.

## Capacity and acceptance

The planning envelope is 50 visible On shift pages for 12 hours at a 60-second
interval: 36,000 requests/day. Unchanged contact refresh must perform zero D1
contact reads/writes and only its documented bounded R2 reads.

Local evidence uses the 109,200-event fixture and is not a Cloudflare bill.
Production acceptance requires:

- no `roster_events`/`roster_daily_presence` scan on shared or blocked reads;
- unchanged ingestion rewrites no roster or cache facts;
- paused contact ingestion touches neither D1 nor R2;
- bootstrap/publication remain within their reviewed ceilings;
- actual usage matches request volume and estimates; and
- at least 50% of each daily quota remains after the test window.

Investigate and pause optional work at 50% usage. Stop it by 70%, or immediately
on an unexplained spike.

## Rollback

Rollback must not restore expensive SQL:

1. Set `FACILITY_SHARED_EMERGENCY_PAUSED=true`.
2. Set `FACILITY_LEGACY_READS_PAUSED=true` before disabling faulty shared
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
