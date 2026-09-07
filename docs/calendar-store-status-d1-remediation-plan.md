# `calendarStoreStatus` D1 remediation plan

## Implementation status — 7 September 2026

Implemented locally on `codex/at-a-glance-d1-optimization`; not deployed and
not applied to any remote database. The dedicated reader remains disabled in
Production and Preview configuration.

The implementation adds migration `0031`, incremental summary maintenance,
the bounded compact reader, one-file bootstrap integration, a separate
Creator-only doctor-discovery action, in-flight browser request coalescing and
active-job-only polling with hidden-tab suspension and backoff.

Focused local evidence:

- the 109,200-event scale fixture returns five status rows through the compact
  active-summary index with no `roster_events` access or temporary sort;
- an identical roster import writes zero rows;
- a one-event correction writes at most nine rows, including the `building`
  and final `ready` crash-safety transitions;
- fresh local migrations, isolated login/restart lifecycle, failure recovery,
  fixture behaviour and quota guards pass; and
- local Cloudflare bindings remain unused.

A subsequent pre-commit review found blockers in historical sync/dispatch
metadata lookups, active retained-file semantics, transactional summary
publication, browser polling and revision handling. Those findings have now
been remediated and verified locally. The corrective work and acceptance
evidence are recorded in
[`calendar-store-status-precommit-remediation-plan.md`](./calendar-store-status-precommit-remediation-plan.md).

Production still requires the rollout gates below. In particular, migrations
`0026` through `0030` precede `0031`, existing summaries are populated only by
the one-file bootstrap, and `ROSTER_STATUS_SUMMARY_ENABLED` stays `false` until
an explicitly approved canary.

## Purpose and authority boundary

This plan removes repository-wide D1 scans from the Creator's roster-storage
status. Writing it does not authorise code changes, a deployment, a migration,
a bootstrap or any other Production operation.

The current Production guard must remain active until this plan is implemented,
tested and separately rolled out. While `ROSTER_AUTOMATION_WRITES_ENABLED` is
`false`, `calendarStoreStatus` returns unavailable before reading roster
repository data.

## Confirmed Production evidence — 7 September 2026

- Production code and configuration identify `95bc9ad`; active Pages deployment
  is `d692bdc9-95a7-4d82-a6d1-47fb9a13851a`.
- Migration `0025_runtime_schema_parity.sql` is applied. Migrations `0026`
  through `0030` remain pending.
- A one-hour Insights window inspected at 00:34 UTC—covering the first 34
  minutes of the new quota day and 26 minutes before reset—showed this query ten
  times around an ordinary Creator app load:

  ```sql
  SELECT file_id, COUNT(*) AS count
  FROM roster_events
  WHERE file_id IN (...)
  GROUP BY file_id
  ```

- Those ten calls examined 364,620 rows, averaging 36,462 rows per call. The
  window proves the query shape and its cost, but does not assign every call to
  the new UTC day.
- The same status operation also loaded all retained `raw_roster_files`; 21
  observed calls examined another 19,026 rows.
- After the pre-D1 Production guard was deployed, the event-count query remained
  at ten executions in the following observation. No further broad status count
  was observed.

This is separate from the legacy At a glance incident. The query is used by
Creator roster-management status and is not fixed by the shared At a glance
reader.

## Current implementation problem

`calendarStoreStatus` currently performs overlapping work on every request:

1. `queryRosterFiles(..., { includeInactive: true })` counts doctors and events
   with correlated subqueries and loads every doctor row for every file.
2. `queryRosterFileRanges` derives coverage from stored roster data.
3. `queryRawRosterFiles` loads the complete retained-file history.
4. `countDerivedEventsByFile` counts the active files' event rows again.
5. `countDerivedDoctorsByFile` counts the active files' doctor rows again.
6. A non-lightweight request may perform further per-file/per-doctor counts and
   event loads.

The browser can call this operation from initial Creator reconciliation,
management surfaces, manual refresh, import completion and a five-second sync
poller. Multiple callers are not coalesced. A small response therefore does not
represent a small database read.

## Required outcome

The normal status path must read a bounded set of compact rows whose cost grows
with the number of current roster files, not with roster-event history or
retained-file history.

It must not query, count, aggregate or join `roster_events`,
`roster_file_doctors`, `roster_daily_presence` or the complete
`raw_roster_files` table. A missing summary returns `unknown` or unavailable;
it never falls back to a historical count.

## Data model

Add a new migration after the already-reviewed sequence. Do not edit migrations
`0026` through `0030`. The proposed migration is
`0031_roster_file_status_summaries.sql`.

Create one compact row per relevant roster file:

```text
roster_file_status_summaries
  file_id                    primary key
  source_type
  source_id
  name
  active                     0 or 1
  derived_state              retained | building | ready | error | removed
  expected_doctor_count
  indexed_doctor_count
  event_count
  raw_source_available       0 or 1
  size
  last_modified
  uploaded_at
  content_revision
  status_revision
  updated_at
```

Use non-negative integer constraints for counts and Boolean fields. Add one
index supporting the only collection lookup:

```text
(active, source_type, updated_at, file_id)
```

The primary key supports exact expected-file lookups. Coverage dates remain in
`roster_file_coverage` from migration `0026`; join it by `file_id` only when the
UI needs coverage. Do not duplicate or recalculate coverage from events.

`roster_sync_runs.doctor_count` and `event_count` may help validate an exact
completed automated import, but they are not the source of truth for manual
files and must not be used without an exact file/revision match.

## Incremental write rules

Maintain the summary in the same controlled change set as its roster file:

- When retained input is accepted, create or update one `retained` summary row
  from metadata already in memory. Do not enumerate older retained files.
- When parsing begins, record `building` and the authoritative expected doctor
  count only if that state actually changed.
- When a complete import commits, set the indexed-doctor and event counts from
  the final parsed change set already in memory. Do not issue a verifying
  `COUNT(*)` query.
- For a correction, use the final authoritative parsed set or the known
  inserted/deleted deltas within the same transaction. Update only that file's
  summary.
- An identical import must not rewrite the summary or change its revision.
- Activation, supersession, overlap replacement and removal update only the
  affected summary rows. Preserve unrelated file contributions.
- Removing only the retained object clears `raw_source_available`; removing the
  derived file records `removed` or deletes the summary only when no current
  reference needs it.
- A failed or interrupted build must never retain a false `ready` state.
  Publication of `ready` is the final transactional step after the roster
  facts have committed.

All import and maintenance entry points must use one summary writer. No route
may maintain a private approximation.

## Existing-file population

Do not run a database-wide summary backfill.

Populate existing rows through the bounded, one-file-at-a-time facility
bootstrap:

1. Name exactly one active file and ED.
2. Read its retained R2 input and compact D1 markers under the bootstrap's
   existing limits.
3. Obtain counts from the parsed retained input already needed for that file's
   compact facts.
4. Write one status summary alongside the file's compact coverage/staff facts.
5. Verify the revision and stop before another file.

If an exact completed `roster_sync_runs` record and content revision already
prove the counts, the bootstrap may reuse them without reading event rows. If
retained input and exact revision evidence are both unavailable, record the
file as `unknown`. Do not count its D1 events merely to improve a diagnostic
label.

Do not create summaries for all 906 historical retained files. Include only:

- current active derived files;
- each source's explicit active file;
- an exact bounded list of files currently expected by the Creator workspace;
  and
- a file explicitly selected for inspection or repair.

## Read API redesign

Replace the normal `calendarStoreStatus` repository assembly with a compact
query that:

- selects active summary rows plus explicitly supplied expected file IDs;
- enforces a hard maximum of 100 rows;
- optionally joins `roster_file_coverage` by primary key;
- loads source state and at most the latest bounded sync-run/dispatch metadata;
  and
- returns `unknown` for absent or non-ready summaries.

Do not call `queryRosterFiles`, `queryRosterFileRanges`,
`queryRawRosterFiles`, `countDerivedEventsByFile` or
`countDerivedDoctorsByFile` from this route.

Detailed per-doctor/file diagnostics must become a separate explicit Creator
action. It must name one file and, where applicable, one doctor, use exact
indexed predicates and have its own small returned-row ceiling. It must not be
invoked by application startup, modal opening or polling.

## Browser request control

The browser must treat status as management data, not a general application
heartbeat:

- Coalesce simultaneous callers behind one in-flight promise.
- Load once when the Creator genuinely needs roster-management state.
- Manual refresh starts one request and disables its button while pending.
- Poll only while at least one roster job is genuinely `pending`, `parsing` or
  `saving`.
- Never poll a hidden tab and stop immediately when all jobs settle.
- Start at five seconds only for an active job, then back off to 10 and 20
  seconds if its revision is unchanged.
- Do not retry a successful unchanged response.
- Bound errors to one retry with backoff; never multiply retries across callers.
- Send the last `status_revision`; an unchanged response may omit the rows.

Ordinary login, account switching, calendar display and At a glance must not
depend on this status call.

## Independent safety control

Do not couple status safety to the broad roster-write switch permanently.
Introduce a dedicated default-off setting such as:

```text
ROSTER_STATUS_SUMMARY_ENABLED=false
```

While false or missing, the route keeps the current pre-D1 unavailable
response. When true, it permits only the compact summary reader. There is no
Production setting that re-enables the old scanning implementation.

Roster writes must not be restored until this dedicated control exists. Turning
`ROSTER_AUTOMATION_WRITES_ENABLED` on must not expose legacy status SQL.

## Implementation sequence

1. Preserve the deployed pre-D1 status guard and all current At a glance and
   automation pauses.
2. Add migration `0031` and fresh-schema/local migration tests.
3. Implement the single incremental summary writer and connect every retained,
   import, correction, activation, supersession and removal path.
4. Add the bounded one-file bootstrap integration without changing the
   approved Production bootstrap limits.
5. Implement the compact reader and dedicated default-off control.
6. Replace startup/management callers and add in-flight request coalescing,
   visibility handling and active-job-only backoff.
7. Separate detailed diagnostics from the normal status response.
8. Run the correctness and performance gates below against local D1/R2 only.
9. Review the full diff and update the Phase 7 rollout package with the exact
   migration checksum, maximum bootstrap cost and activation settings.
10. During a separately approved Production rollout, apply `0031` in isolation
    after `0030`, bootstrap one active file, enable the compact status reader
    for the Creator, and observe actual metrics before processing another file.
11. Remove the old scanning function only after parity is proven. Until then it
    remains unreachable in Production and excluded from rollback.

This sequence is a prerequisite to restoring roster automation. It complements
the main At a glance rollout rather than replacing its migration, bootstrap or
reader gates.

## Correctness tests

Cover:

- retained-only, building, ready, error, inactive and removed files;
- complete Excel and FindMyShift imports;
- identical imports, one-shift corrections, swaps, sickness and removals;
- overlapping files and current/next-term coexistence;
- activation, supersession and deletion without changing unrelated summaries;
- a crash before and after the final `ready` transition;
- stale revisions and concurrent duplicate imports;
- exact status parity for file counts and availability where authoritative
  evidence exists;
- `unknown` rather than a fallback scan where it does not; and
- Creator-only access and no dependency from ordinary login or At a glance.

## Performance and safety gates

Use the existing 109,200-event synthetic database and representative roster
fixtures. Keep correctness and scale measurements separate.

The gate passes only when:

- normal status query plans contain no access to `roster_events`,
  `roster_file_doctors`, `roster_daily_presence` or unbounded
  `raw_roster_files`;
- status cost is bounded by at most 100 compact rows, independent of event
  history size;
- a repeated unchanged status request writes zero D1 rows;
- an unchanged import writes zero status rows;
- a one-file correction changes at most that file's summary row in addition to
  its already-budgeted roster facts;
- hidden-tab and settled-job tests make zero status requests;
- concurrent callers produce one request;
- a missing summary produces `unknown` or unavailable and zero historical
  queries;
- disabled and malformed configuration stop before repository reads; and
- no test, benchmark or fixture accesses Cloudflare.

Record rows examined separately from rows returned. `LIMIT` is not accepted as
proof of bounded reads without an indexed query plan.

## Production rollout gates

Every Production mutation remains a separate approval:

1. Apply pending migrations `0026` through `0030` under their existing rollout
   rules; none activates this feature.
2. Apply `0031` alone from an isolated checksum-verified directory.
3. Inspect one exact active-file bootstrap plan and its maximum D1/R2 cost.
4. Execute only that file's bootstrap and verify its summary.
5. Deploy/confirm `ROSTER_STATUS_SUMMARY_ENABLED=false` before code activation.
6. Enable it for the Creator only, request status once and inspect query
   fingerprints and UTC-day usage.
7. Repeat one unchanged request and prove no writes and a small indexed read.
8. Observe a quiet interval before bootstrapping another file.
9. Restore roster writes only after every active required file has a trustworthy
   summary and the old status scan is unreachable.

Stop immediately on a `roster_events`/`roster_file_doctors` status query,
unexpected writes, more than 100 summary rows, an unexplained usage increase or
less than 50% UTC-day quota headroom.

## Rollback

- Set `ROSTER_STATUS_SUMMARY_ENABLED=false`.
- Keep the current pre-D1 unavailable response.
- Keep roster writes and maintenance paused if summary correctness is in doubt.
- Retain summary rows for diagnosis; do not delete or rebuild them during an
  incident.
- Do not restore the old count queries as a compatibility path.
- At a glance legacy reads remain independently paused.

The safe rollback is reduced Creator diagnostic availability, not a historical
D1 scan.
