# `calendarStoreStatus` pre-commit remediation plan

## Status and authority boundary

This plan addresses the blocking findings from the local review on 7 September
2026. Its remediation has now been implemented and verified locally. This
document does not authorise a deployment, remote migration, Production query,
bootstrap or configuration change.

Work remains on `codex/at-a-glance-d1-optimization`. All Production safeguards
stay fail-closed, including `ROSTER_STATUS_SUMMARY_ENABLED=false`, paused roster
writes and paused legacy At a glance reads. Migration `0031` is still
uncommitted, so its schema and indexes may be corrected in place without adding
a follow-up migration.

## Required outcome

Before commit, the complete Creator status operation must have a cost bounded
by current configured sources and at most 100 relevant file summaries. No part
of the operation may scan or sort growing event, retained-file, sync-run or
dispatch history.

The compact summary must also remain correct through retained uploads,
incremental imports, corrections, supersession, interruption and deletion.
Browser callers must share one request budget.

## Remediation sequence

### 1. Bound every query used by the status response

Keep the indexed summary lookup, but repair its accompanying metadata reads:

- Add an index for exact latest-run lookup:
  `(source_id, started_at DESC, id DESC)`.
- Load no more than the latest relevant run for each source in the bounded
  source set. Do not globally sort the complete `roster_sync_runs` table.
- Add an index for the latest dispatch lookup:
  `(requested_at DESC, id DESC)`.
- Bound roster sources to a small explicit maximum, initially 16, and add an
  index supporting `(label, id)` if that ordering remains part of the API.
- Keep expected file IDs deduplicated and capped at 100 before preparing SQL.
- Preserve primary-key joins to `roster_file_coverage` only.

The normal status operation must contain no access to `roster_events`,
`roster_file_doctors`, `roster_daily_presence`, or a collection read of
`raw_roster_files`.

### 2. Define active-summary semantics precisely

`active` means “currently relevant derived roster,” not “a retained source
object exists.” `raw_source_available` records retained-source availability
independently.

- A newly retained object starts as `retained`, `active=0`, unless the same
  controlled transaction is deliberately creating the current derived file.
- Updating retained metadata on an existing summary preserves its current
  derived state, active flag, counts and content revision.
- Superseding or deleting derived data always makes the old summary inactive.
  Its state becomes `retained` when raw input remains and `removed` otherwise.
- An inactive retained summary remains discoverable only through an exact
  expected-file or explicitly selected-file lookup.
- Activation must deactivate only the known superseded summaries and must not
  alter unrelated source or term contributions.

Add a test with more than 100 historical retained files proving that only
current active and explicitly expected files appear and that current files
cannot be crowded out.

### 3. Replace revisions instead of appending to them

Use one summary mutation helper for every lifecycle path. Each genuine state
change receives a fixed-size revision token or digest; SQL must never append
text to the previous revision.

The helper must distinguish “preserve this field” from an explicit false or
zero. In particular, preserving `raw_source_available=1` must produce a
revision representing `1`, not a revision calculated from an omitted value.

Requirements:

- revisions have a fixed documented maximum length;
- identical input leaves the revision and `updated_at` unchanged;
- revision input represents the complete resulting summary row; and
- no import, activation, removal or materialisation path maintains a private
  approximation outside the shared writer.

### 4. Make publication crash-safe

Use `building` as the only externally visible state while facts or dependencies
are changing. A previous `ready` summary must not remain ready during a
replacement.

- Start/full replacement: mark the affected file `building` in the same batch
  that begins changing its roster facts.
- Chunk append: insert the chunk and apply its known event-count delta to the
  summary in the same transactional batch.
- Corrections: use the authoritative parsed totals or known mutation deltas;
  update only that file's summary in the same controlled change set.
- Finalisation: build daily presence and facility materialisations while the
  file is not published as ready. Publish the roster-file activation and final
  `ready` summary together only after all required dependencies succeed.
- Failure before final publication must leave `building` or explicitly set
  `error`; it must never expose stale `ready` counts or revisions.
- Supersession and deletion must atomically update the affected file records
  and summary lifecycle state wherever D1 batching permits it.

Add deterministic fault injection after fact writes, after a chunk, after
materialisation and immediately before final publication. Each failure must
prove that no false `ready` state is returned.

### 5. Remove ingestion-time verification counts

Eliminate the before/after `COUNT(*)` queries in overlap trimming. Load the
single compact summary by primary key, use the delete statement's affected-row
metadata, and derive the new total from that known delta. If the summary is
absent or cannot authoritatively support the calculation, leave it unknown or
building rather than scanning event history.

No correction path may count events or doctors merely to improve a status
label.

### 6. Correct one-file bootstrap integration

The bootstrap remains explicit, authenticated, one-file-at-a-time and subject
to the existing hard limits. It must never start another file automatically.

- Establish retained-source availability with one exact file/object lookup.
- If the approved facility bootstrap is already reading one exact file's
  bounded rows to create compact facility facts, its known result totals may be
  reused for the status summary; do not add `COUNT(*)` or another event read.
- Otherwise use parsed retained input or an exact completed sync-run whose file
  and content revision match.
- If none of those authoritative sources is available, publish `unknown` or
  leave the summary absent. Never scan D1 history for a diagnostic label.
- Write the summary only after the compact facts succeed, with accurate raw
  availability and a fixed-size revision.

Update the dry-run estimate so it separately states rows examined, rows
returned, summary writes and retained-object operations.

### 7. Enforce one browser request budget

Route every browser status request through the existing shared in-flight
promise. Remove direct status polling from persistence confirmation and
post-save refresh paths.

- One logical refresh permits one initial request and at most one retry.
- Persistence confirmation uses the shared 5/10/20-second active-job schedule;
  it must not run a second 2.5-second polling loop.
- Hidden tabs issue no requests and cancel scheduled polling immediately.
- Settled, failed or timed-out jobs stop polling.
- Manual refresh joins an in-flight request rather than starting another.
- A successful unchanged response is not retried.
- Expected file IDs needed by an active save are merged into the shared request
  without exceeding 100.

## Focused verification

### Query-plan and scale gate

Extend the deterministic local cost fixture with growing history independent
of roster events:

- 109,200 roster events;
- at least 10,000 sync runs;
- at least 10,000 dispatch records;
- more than 100 inactive retained summaries; and
- a small current active/expected set.

Record `EXPLAIN QUERY PLAN` and distinguish rows examined from rows returned.
The complete status path passes only when:

- every growing-history lookup uses an appropriate index;
- no plan reports a full history scan or temporary ordering B-tree;
- no forbidden roster or raw-history table is accessed;
- at most 100 summary rows, 16 source rows, one run per source and one dispatch
  row are returned; and
- cost remains effectively constant when historical rows are multiplied.

### Correctness gate

Use existing representative Excel and FindMyShift fixtures to cover:

- retained-only, building, ready, error, inactive, removed and unknown states;
- identical import with zero summary writes;
- one-shift correction, swap, sickness and removal;
- overlap trimming and current/next-term coexistence;
- active-file supersession without unrelated-summary changes;
- raw-source deletion independently of derived status;
- fixed-length revisions through many chunks and lifecycle transitions;
- accurate bootstrap raw availability; and
- all crash points listed above.

### Browser gate

Use deterministic mocked timers and requests to prove:

- concurrent callers make one request;
- an active job follows 5/10/20-second backoff;
- one failed request causes no more than one retry;
- persistence confirmation creates no parallel poller;
- hidden and settled states make zero requests; and
- manual refresh joins an existing request.

### Regression gate

Run only the focused local suites required by the touched paths:

```text
npm run check
npm run test:d1-quota
npm run test:database-costs
npm run test:facility-materialization
npm run test:fixtures
npm run test:local-isolation
```

Run the fresh local migration check if `0031` changes. No test may access
Cloudflare or reuse Production bindings.

## Pre-commit review gate

After remediation, perform another read-only diff review. A commit is permitted
only when all of the following are true:

- all blocking findings above are closed with test evidence;
- both the summary query and its sync/dispatch/source metadata queries have
  indexed, history-independent plans;
- no false-ready failure case remains;
- no status caller bypasses request coalescing or retry limits;
- `git diff --check` passes;
- the rollout documentation contains the final `0031` checksum and revised
  maximum cost;
- Production/Preview flags remain false; and
- only intended project files are staged. Existing unrelated untracked files
  must not be included.

The commit itself remains local until separately reviewed. Committing does not
authorise pushing, deployment, remote migration, bootstrap or feature
activation.

## Completion evidence — 7 September 2026

- Status summaries, expected files, bounded sources, latest per-source runs and
  latest dispatches all use their intended indexes.
- The scale fixture contains 109,200 events, 10,000 sync runs, 10,000
  dispatches and 150 inactive retained summaries.
- Retained history remains inactive unless requested by exact file ID.
- Chunk retries preserve the same count and revision; successful finalisation
  publishes `ready` only after dependency materialisation.
- Injected chunk failure rolls back its event facts, and injected
  materialisation failure leaves `building` rather than false `ready`.
- Overlap trimming uses mutation deltas instead of event verification counts.
- The browser has one coalesced status request path with no independent
  persistence polling request.
- The focused test suite, fresh local migration/restart lifecycle, local safety
  check and diff whitespace check pass without Cloudflare access.
