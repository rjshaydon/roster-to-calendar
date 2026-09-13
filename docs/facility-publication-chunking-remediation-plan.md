# Facility publication chunking remediation plan

Prepared 11 September 2026 after the first MMC shared-cache publication could
not complete as one 91-day Pages Function request.

This plan authorises design and local implementation only. It does not
authorise another Production publication, reader activation, Power Automate,
roster/contact ingestion, remote migration or legacy At a glance access.

## Settled evidence

The three active MMC roster files are compact-ready:

- Term 1: 148 doctors and 4,295 events;
- Term 2: 147 doctors and 4,264 events; and
- Term 3: 152 doctors and 4,173 events.

The current-term publication dry run planned 91 dates and reported no broad
roster scans. Its maximum was 145,835 examined D1 rows, 111 D1 statements,
eight indexed D1 writes, 126 R2 gets and 98 R2 puts.

Two execution attempts failed safely:

1. The first was stopped by the generic 64-statement middleware ceiling.
2. Commit `3e913a7` added the reviewed route-specific ceiling. The second
   progressed further but returned a non-JSON HTTP 500 after about two minutes.

Settled Analytics Engine evidence attributes 6,556 reads/two writes and 10,418
reads/three writes to those two isolated buckets: 16,974 reads and five writes
in total. Later buckets show no continuation or write loop. Daily account usage
at the 18:33 AEST check was 66,837 reads and 4,210 writes.

The failure is therefore request size/runtime duration, not D1 quota
exhaustion. The existing publisher performs one bounded ED/date lookup for each
of 91 dates and then reads/writes the corresponding R2 day and month objects in
one request. It must become resumable.

## Required outcome

Publish one hospital/term through small, explicit batches while preserving
these invariants:

- readers see either the previous complete generation or the new complete
  generation, never a partial build;
- no request scans roster history or processes more than its declared batch;
- retries are idempotent;
- input changes invalidate the operation before the public manifest changes;
- no request schedules an automatic continuation;
- all execution remains hospital-, term-, operation- and revision-specific;
- a failed or abandoned operation can be resumed or discarded safely; and
- the permanent legacy-read pause remains true.

## Design

### 1. Immutable operation plan

The dry run continues to compute the complete fixed plan. It must return:

- `operationRevision` derived from hospital, term, input revision, old manifest
  ETag/revision and the complete ordered date list;
- total dates and deterministic batches;
- a default batch size of seven dates and a hard maximum of seven;
- expected batch count (13 for 91 dates);
- per-batch and whole-operation cost ceilings; and
- the finalisation cost as a separate figure.

The client/workflow must supply the exact operation revision to every later
request. Never accept an arbitrary date list that was absent from the plan.

### 2. Staging manifest

Create an immutable R2 staging object for the operation. It contains the
prepared staff/metadata, base public revision, complete planned dates, completed
batch digests and candidate day pointers. It is not the public manifest.

D1 keeps one small operation record per hospital with status `planned`,
`building`, `ready`, `complete`, `failed` or `superseded`; operation revision;
input revision; next incomplete batch; and lease timestamps. Prefer extending
the existing `facility_day_publications` row unless local implementation proves
that a separate table materially simplifies correctness. Any migration remains
local until separately approved.

### 3. Explicit batch action

Add a token-protected batch mode to the publication endpoint:

- exact hospital, term, operation revision and batch index are required;
- at most seven predetermined dates are processed;
- each date uses the existing indexed hospital/date lookup and 513-row sentinel;
- day objects are immutable and content-addressed;
- the staging manifest is updated only after every day in that batch succeeds;
- repeating a completed batch performs zero D1 day queries and zero R2 puts;
- an out-of-order, stale, oversized or foreign batch is rejected before writes;
- one request cannot process the next batch automatically.

Use a dedicated middleware ceiling derived from the locally measured seven-day
path. Do not reuse or enlarge the 112-statement whole-publication allowance.

### 4. Separate month assembly

Do not rebuild all four month objects during every date batch. After all date
batches complete, assemble month objects in a separate explicit phase, one
month per request. Month assembly reads only immutable day objects referenced by
the staging manifest. Repeating an unchanged month performs no put.

### 5. Atomic finalisation

Finalisation is a separate request. It must:

1. prove all planned date batches and month objects are complete;
2. recompute and compare the current input revision and base manifest ETag;
3. reject stale input before touching the public pointer;
4. write the immutable candidate manifest;
5. atomically replace the public manifest using the expected R2 ETag; and
6. mark the D1 operation complete.

If the pointer write succeeds but the final D1 status update fails, the next
request must recognise the matching `publicationOperationId` and repair only
the status row. Existing readers must continue to use the old public manifest
until step 5 succeeds.

### 6. Workflow orchestration

Replace the current single execution option with manually dispatched workflow
modes: `plan`, `build-batch`, `build-month` and `finalize`. Inputs use choices or
validated integers and exact revisions. The workflow prints concise JSON
evidence but never secrets or roster/contact content.

For the first Production canary, dispatch each step separately and reconcile
settled Analytics Engine usage between steps. Automated iteration may be added
only after the manual sequence has passed and must still stop on the first
failure; it must never run on a schedule.

## Local implementation sequence

1. Refactor planning into a deterministic operation plan without changing the
   current read queries.
2. Implement staging and one seven-day batch, with explicit recovery states.
3. Implement one-month assembly and atomic finalisation.
4. Update middleware with distinct plan, batch, month and finalise ceilings
   based on measured statements, not a broad shared limit.
5. Update the manual GitHub workflow to expose the four modes.
6. Keep all Production variables closed in committed configuration.
7. Update the restoration register and quota-safe rollout ledger with the
   final tests, checksums and rollback controls.

## Focused tests

Use existing synthetic and representative fixtures. Prove:

- a 91-day term becomes 13 deterministic batches of at most seven dates;
- every date query is constrained by hospital and exact date and uses the
  intended index;
- each request remains below its declared D1 statement/row and R2 limits;
- no batch reads `roster_events` outside its seven exact dates;
- retrying a completed batch or month writes nothing;
- a failure on every possible batch boundary leaves the public manifest intact;
- missing, duplicate, skipped, foreign and out-of-order batches are rejected;
- changed coverage, staff, grade, designation, visibility or daily digest
  invalidates finalisation;
- concurrent operations cannot publish over one another;
- pointer-success/status-failure repair is idempotent;
- readers cannot access staging keys; and
- emergency pause and malformed/missing configuration stop before D1 or R2.

Do not expand parsing regression tests unless the refactor changes parsing.

## Production gates after implementation

1. Start on a fresh, reconciled account-wide budget sample with Power Automate
   and all optional readers/writers off.
2. Deploy only MMC planning, run once, close and reconcile.
3. Execute one seven-day batch, close and reconcile.
4. If clean, repeat batches serially; do not combine unobserved early batches.
5. Build one month at a time, reconciling the first before continuing.
6. Finalise once, close all builder controls and wait for settled telemetry.
7. Verify the public manifest through R2/control-plane evidence before issuing
   at most one tiny application read.
8. Only then enable the Creator-only MMC reader with contacts off.
9. Observe the Creator canary before expanding At a glance readers, publishing
   contacts or enabling any Power Automate flow that builds At a glance data.

Core roster Power Automate restoration is independent and now precedes this
facility canary under
[`core-calendar-sync-restoration-plan.md`](./core-calendar-sync-restoration-plan.md).
It may update personal calendars while every facility builder and reader stays
closed. Contact flows remain later and separate; the historical global watchdog
is not a prerequisite. The user will be told explicitly when an individual
flow may be enabled.

## Stop conditions

Close all optional controls immediately on any quota error, non-JSON response,
unexpected query fingerprint, batch outside its plan, ordinary request above
10,000 examined rows, missing telemetry, pointer ambiguity, unexplained write,
or evidence of automatic continuation. Do not raise a ceiling or retry until
the failed request has settled and its state is understood.

## Implementation status — 13 September 2026

Implemented locally without a schema migration. The production endpoint now
accepts only `plan`, `build-batch`, `build-month` and `finalize`; the former
whole-term execution input is no longer a callable path. The manual GitHub
workflow exposes exactly those four modes. Local scale and failure-recovery
tests pass, while committed Production and Preview flags remain closed.
