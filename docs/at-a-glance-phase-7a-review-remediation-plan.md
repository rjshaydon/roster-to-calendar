# Phase 7A second-review remediation plan

## Purpose and current state

This plan addresses the safety defects found during the read-only review of
commit `587bbca` on `codex/at-a-glance-d1-optimization`.

Commit `587bbca` must not be pushed or used for an online test in its current
form. Production access, deployment, remote migrations, remote backfills,
automation changes and feature-flag changes remain prohibited. The protected
production base remains `aa9eed8`.

The remediation is deliberately narrow. It does not resume the Doctor identity
work and does not redesign the At a glance interface.

## Review findings to resolve

1. The bootstrap's `LIMIT 25001` bounds returned rows, but does not currently
   prove a bound on rows examined. `ORDER BY id` requires a temporary sort when
   using the existing `file_id` index, so SQLite may examine every matching
   event before applying the limit. Other file-scoped bootstrap reads also lack
   explicit row ceilings, and two compact-contribution lookups lack a leading
   `file_id` index.
2. The publication `planRevision` omits mutable Staff, designation, seniority,
   SMS-continuity and daily-roster inputs. An approved dry run can therefore
   execute against different data without being rejected as stale.
3. The nominally one-term metadata query reads and groups catalogue rows for
   every historical term belonging to an ED, then filters the result in
   JavaScript.
4. The dry-run response reports statement counts but not defensible upper
   bounds for D1 rows examined. Existing tests count calls and returned rows,
   but do not validate the actual bootstrap plan or stale-plan behaviour after
   real input changes.
5. The rollout worksheet unnecessarily enables general roster writes while
   performing a compact-fact bootstrap.

## Required outcome

Phase 7A may return for review only when:

- every bootstrap read has a documented hard row ceiling;
- `EXPLAIN QUERY PLAN` proves the event query uses its exact `file_id` index
  without a table scan or temporary sort;
- over-budget bootstrap paths perform zero writes;
- one-term planning reads only the selected ED and term, apart from the small
  set of active-file coverage revisions required to detect daily changes;
- every input capable of changing published Staff or day objects participates
  in the plan revision;
- execution cannot publish a manifest pointer when inputs changed after
  planning or while immutable objects were being assembled;
- the dry run distinguishes statements, rows returned and estimated rows
  examined for each query class; and
- the rollout procedure grants only the minimum switch required for each
  operation.

## Workstream 1: make bootstrap reads genuinely bounded

### Event query

Keep the existing single-file operation and 25,000-event ceiling, but change
the database query so the limit is applied while walking the `file_id` index:

- query `WHERE file_id = ? LIMIT 25001` without database ordering;
- reject 25,001 rows as over budget before preparing any mutation;
- when 25,000 or fewer rows are returned, sort the bounded result by event ID
  in memory before calculating digests or facts; and
- retain the exact file identity and active/source validation.

The order of an under-limit complete result does not affect which events are
processed. Sorting in memory preserves deterministic digests without forcing
D1 to inspect and sort all matching history.

### Supporting queries

Add explicit sentinel ceilings to every collection loaded by bootstrap. The
initial proposed ceilings are:

| Input | Maximum accepted | Query maximum |
| --- | ---: | ---: |
| roster events for one file | 25,000 | 25,001 |
| doctors for one file | 512 | 513 |
| existing Staff contributions for one file | 750 | 751 |
| existing catalogue contributions for one file | 750 | 751 |
| existing coverage record | 1 | 1 |

Validate these constants against the representative Excel and FindMyShift
fixtures before adopting them. A fixture exceeding a proposed ceiling should
cause the ceiling to be reviewed explicitly; the endpoint must not silently
raise it or paginate.

Add migration `0030` containing the missing bounded-maintenance lookup indexes:

- `facility_term_staff_contributions(file_id, term_start, doctor_key)`;
- `facility_stream_catalog_contributions(file_id, term_start, catalog_key)`;
- `facility_staff_designations(source_type, active, term_start, doctor_key)`;
  and
- `facility_staff_seniority_overrides(source_type, active, term_start,
  doctor_key)`; and
- `roster_files(source_type, active, id)`.

The existing roster-event and roster-doctor indexes should be reused. Do not
duplicate them. Update the local schema helper only to maintain fresh-schema
parity with the migration; retain the permanent prohibition on runtime schema
creation in production.

Perform all bounded reads before constructing or executing mutations. Any
sentinel row returns an `over-budget` result naming the exceeded collection and
performs zero writes. Keep the 750-statement write ceiling so the mutation fits
within one D1 batch.

### Bootstrap cost contract

The inspection response must publish a fixed upper-bound worksheet for one
execution, including:

- exact metadata/marker statements;
- at most 25,001 event rows examined;
- the doctor and existing-contribution sentinel ceilings;
- the maximum total estimated rows examined across all bootstrap reads;
- at most 750 mutation statements and the conservative indexed-write estimate;
- zero R2 operations; and
- no automatic continuation.

These are local structural bounds, not Cloudflare billing measurements.

## Workstream 2: create an immutable publication input snapshot

Replace the current partial `planRevision` input with one canonical,
term-scoped publication plan assembled by a shared planner used by both dry run
and execution.

For one ED and one medical term, the canonical plan must contain or digest:

- active compact coverage rows and their `contentRevision`, `staffDigest` and
  `dailyDigest` values for files intersecting the term;
- the exact term visibility row;
- catalogue facts for that term only;
- aggregated term Staff contributions;
- applicable continuing-SMS membership facts;
- applicable Staff designations;
- applicable seniority overrides;
- exact covered dates and affected months;
- the current fixed-manifest revision and ETag; and
- the schema/parser revision used to interpret these facts.

Stable ordering must be applied before hashing. Timestamps that change without
changing content must be excluded.

The compact file digests deliberately participate in the revision even if this
occasionally invalidates a plan because another date in the same retained file
changed. Safe over-invalidation is preferable to publishing data that was not
approved.

### Execution contract

Execution must:

1. rebuild the canonical input plan from the same term-scoped queries;
2. compare its revision with the supplied dry-run revision before any D1 or R2
   write;
3. build Staff and day/month objects from that validated plan and the exact
   planned date set;
4. write new content-addressed objects only as unpublished candidates;
5. recompute the canonical input revision immediately before changing the
   fixed manifest pointer; and
6. abandon the candidate and return `stalePlan` if the revision or base
   manifest ETag changed.

Orphaned immutable candidate objects are acceptable after a stale concurrent
build; publishing a mixed or unapproved manifest is not. Their deletion must
not occur automatically in this phase.

The Staff publication helper should accept the already validated term snapshot
instead of independently re-querying broader live state. Day publication may
perform its planned indexed ED/date reads, but every returned day must have a
hard per-day row ceiling and the final revision check must detect intervening
roster changes.

## Workstream 3: make every publication query term-scoped and bounded

Extend the compact metadata query with an exact `termStart` option used by the
planner and executor:

- select one visibility row with `source_type = ? AND term_start = ?`;
- select catalogue contributions with both `source_type = ?` and
  `term_start = ?` in SQL;
- retain coverage only for active files intersecting the requested term;
- keep Staff, designations and overrides constrained to the same ED/term; and
- never fetch all terms and filter them in JavaScript on this path.

Apply explicit `LIMIT + 1` safety ceilings to term catalogue, aggregated Staff,
SMS continuity, designations, overrides, active contributing files and each
ED/date day query. Exceeding any limit invalidates the plan or execution before
the manifest pointer changes. There must be no automatic pagination or broader
fallback.

General metadata publication used elsewhere may retain its existing API if it
is still required, but the controlled one-term endpoint must not call it.

## Workstream 4: replace call counts with an honest cost worksheet

The planner and executor must share one constants object and one operation
budget. The dry-run response should report, separately:

- D1 read statements by query class;
- maximum estimated rows examined by query class and in total;
- maximum rows returned by query class;
- D1 mutation statements and conservative rows written including indexes;
- R2 GETs for the fixed manifest and existing day objects;
- R2 PUTs for Staff, changed days, changed months, candidate manifest and fixed
  manifest;
- maximum dates and affected months; and
- zero permitted broad roster-history scans.

The estimate should use hard constants and the exact planned date/month counts,
not the observed size of an accidentally small test result. Label all row
figures as local upper-bound estimates until the separately approved online
rollout records actual Cloudflare metrics.

Do not state `broadRosterScans: 0` merely because no query contains the table
name. Require the query-plan assertions described below.

## Workstream 5: focused tests that close the review gaps

### Bootstrap query-plan tests

On the 109,200-event synthetic database, run `EXPLAIN QUERY PLAN` for the exact
SQL used by bootstrap and assert:

- the event lookup uses `idx_roster_events_file` or an explicitly approved
  equivalent;
- no `SCAN roster_events` appears;
- no `USE TEMP B-TREE FOR ORDER BY` appears;
- both compact-contribution lookups use their new leading-`file_id` indexes;
  and
- the sentinel response reports zero writes.

Test every supporting ceiling independently. Use a database adapter that
records all statements and rows returned, and assert the recorded totals do not
exceed the advertised worksheet.

### Publication stale-plan tests

Create a valid dry run, then separately change each of the following before
execution:

- a roster event without changing its shift catalogue signature;
- Staff membership or grade;
- SMS continuity;
- a Staff designation;
- a seniority override;
- term visibility;
- compact coverage; and
- the fixed manifest.

Every case must return `stalePlan` before D1/R2 writes. Add one concurrency test
that changes an input after candidate objects are created but before the fixed
manifest pointer; it must leave the previous fixed manifest intact.

### Term and cost tests

Seed multiple historical and future terms, then prove that planning one term:

- issues SQL containing the requested `term_start` predicate;
- does not return or aggregate other-term catalogue/Staff rows;
- remains within the advertised row and operation ceilings;
- preserves other manifest term entries byte-for-byte; and
- rejects 121 dates and every other over-budget collection before publication.

Retain the existing routing, access, contact, quota, snapshot, fixture,
FindMyShift, queue, local-isolation and syntax regression suites. No remote or
exhaustive browser testing is required for this remediation.

## Workstream 6: minimise rollout authority

Correct the operational worksheet so compact bootstrap requires only:

- the automation token;
- the one-ED materialisation build allowlist; and
- the advanced-maintenance switch.

Keep general roster automation writes, contact ingestion, queue processing and
the watchdog disabled during bootstrap and initial publication. If the
implementation unexpectedly requires the general roster-write switch, treat
that as a defect rather than enabling it.

Add migration `0030` to the serial migration list and record its expected cost
as five index creations before any data bootstrap. The four compact and
maintenance indexes should still be empty at this point; the roster-file index
is bounded by the small retained-file table. It remains a
separately approved remote step.

## Implementation order

1. Add fresh-schema parity and migration `0030` for the two exact file indexes.
2. Replace the bootstrap event ordering with bounded indexed retrieval and
   deterministic in-memory ordering.
3. Add ceilings to all other bootstrap reads and centralise its cost constants.
4. Introduce the canonical one-ED/one-term publication input planner.
5. Make term-scoped metadata, Staff and day builders consume that plan.
6. Add pre-write and pre-pointer stale checks.
7. Replace the cost response with statement, returned-row and examined-row
   upper bounds derived from shared constants.
8. Add the missing query-plan, input-mutation and concurrency tests.
9. Correct the rollout worksheet and implementation status.
10. Run the focused local regression set and perform another independent
    read-only safety review.

## Acceptance gates before push consideration

All of these gates are mandatory:

- no high-severity finding remains in the follow-up review;
- the actual bootstrap SQL has approved query plans and hard ceilings for every
  loaded collection;
- every over-budget bootstrap test records zero writes;
- every changed publication input makes the previous plan stale;
- a change during publication cannot replace the fixed manifest;
- one-term planning proves zero other-term catalogue or Staff reads;
- actual mocked D1/R2 operations remain within every advertised ceiling;
- the worksheet no longer enables unrelated write paths;
- `git diff --check`, syntax checks and the focused regression set pass; and
- unrelated local files remain untouched.

Passing these gates authorises only consideration of pushing the branch. It
does not authorise deployment, production inspection, migration, bootstrap,
publication, feature activation or automation restoration.
