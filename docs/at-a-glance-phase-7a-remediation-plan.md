# Phase 7A remediation plan

> **Superseded rollout assumption — 6 September 2026:** Production evidence
> proved that even a short legacy compatibility canary is unsafe. Wherever this
> completed remediation plan permits temporary legacy fallback, the
> [immediate D1 safety plan](./at-a-glance-immediate-d1-safety-plan.md) and
> [Phase 7 rollout package](./at-a-glance-phase-7-rollout-package.md) now take
> precedence. Production legacy At a glance reads must remain paused throughout
> migration, canary, general rollout and rollback.

## Purpose and current state

This plan fixes the safety gaps found in the review of commit `26ae9da` before
any Phase 7B or online rollout begins.

The work remains on `codex/at-a-glance-d1-optimization`. Production access,
deployment, remote migrations, remote backfills and automation restoration are
still prohibited. The protected production base remains `aa9eed8`, and the
Doctor identity branch remains out of scope.

Commit `26ae9da` is inert while its new settings are disabled, but it is not an
approved rollout candidate. Its rollout worksheet must be corrected after the
remediation is implemented and measured locally.

## Non-negotiable outcomes

The remediation is complete only when:

1. a Creator-only canary cannot change another user's access calculation;
2. an enabled shared route never falls back to a legacy event-history query;
3. existing roster data can be prepared in explicit, inspectable, bounded
   batches without a source-wide or account-wide event scan;
4. a dry run describes the same immutable inputs that execution will use and
   states honest upper bounds for all D1 queries/writes and R2 operations;
5. contact ingestion can be stopped before D1 or R2 access independently of
   roster automation; and
6. every disabled, over-budget, stale-plan and missing-cache path stops before
   prohibited database work.

## Workstream 1: one routing decision for access and data

Replace the current collection of partly independent Boolean checks with one
request-scoped routing decision. It should return an explicit mode rather than
a simple allowed/denied value:

- `shared`: this actor, viewed account and every requested ED are inside the
  active cohort and source allowlist;
- `legacy`: an explicitly isolated local comparison permits the old route;
  this mode is never valid in Production or Preview; or
- `blocked`: the request is part of an active shared rollout but cannot safely
  use the shared object, or legacy reads have been paused.

The routing decision must be calculated before access resolution and reused by
Metadata, Staff, On shift, contact refresh, By stream and Working together.
Do not recalculate slightly different eligibility rules inside each action.

### Access behaviour

- Use materialised access only in `shared` mode during the Creator canary.
- A Creator's own view may enter the Creator cohort. A Creator-entered view of
  another account must use that subject's permitted route and must never gain
  Creator-canary treatment.
- Accounts outside the canary receive `preparing` or unavailable; they never
  retain a Production legacy compatibility route.
- Enabling materialised access for a Creator must therefore have no effect on
  login, account preparation or At a glance access for other users.
- A `shared` request with missing access facts returns `preparing`; it never
  derives access from `roster_events`.

### No-fallback behaviour

- Once a request is routed to `shared`, a missing manifest, object, ED, term or
  contact overlay returns `preparing` or a clear unavailable response.
- A disallowed ED in a shared canary is blocked before any legacy Staff,
  coverage, catalogue or range query.
- “All EDs” is shared only when every requested ED is prepared and allowlisted;
  otherwise it is blocked for a shared-canary actor.
- No exception handler, cache miss or quota error may retry through the legacy
  query.

### Operational controls

Keep separate controls with unambiguous effects:

- stop shared builders and repair jobs;
- stop legacy expensive At a glance reads before D1;
- disable shared readers if the new code is faulty; and
- retain valid, already-published shared reads during a build pause.

The emergency procedure must be capable of returning At a glance as
temporarily unavailable. It must not require restoring the known expensive
queries merely to produce a response.

Legacy compatibility is prohibited throughout the Production canary and every
rollback. If retained temporarily for local parity tests, it requires an
explicit isolated-local setting that cannot be enabled by Production defaults.

## Workstream 2: bounded compact-fact bootstrap

Add a separate bootstrap operation before snapshot publication. Do not make
the publication endpoint silently discover or repair all existing data.

### Inspection

The inspection request must specify exactly one ED and one active roster file.
It may read only small `roster_files` metadata and existing compact-state rows.
It returns:

- file identity, ED and active state;
- whether compact facts already exist;
- the declared maximum event rows and compact writes for the next batch;
- the requested operation and term scope; and
- a stable plan revision derived from the file identity, stored file revision,
  parser version and requested scope.

It must not enumerate `roster_events` merely to estimate their total size.

### Execution

Execution must require the inspection revision and the same explicit file/ED
inputs. It performs one indexed `file_id` event read with a hard `limit + 1`.
If the extra row is present, stop with `over-budget` and perform zero writes.
Never continue with pagination automatically.

After reading the bounded file, derive coverage, term staff and stream facts in
memory. Calculate the exact proposed compact-row mutations before writing. If
that count exceeds its separate write ceiling, return `over-budget` with zero
writes. Commit one file's compact mutations atomically and preserve every
other file's contributions.

The initial ceiling should be set from the local 109,200-event fixture and
current representative roster fixtures, then written as a constant and a
documented rollout input. It must not be inferred from the remaining free-tier
quota at runtime.

Process files serially with an approval and usage check between files. Do not
offer “all files,” “all EDs,” automatic continuation or background retry.

### Bootstrap correctness

- Repeating a completed file bootstrap performs zero writes.
- Overlapping active files retain independent contributions.
- Current and next terms remain separate.
- Trainee membership exists only for terms represented by that hospital's
  roster.
- Continuing SMS rules and manual removals remain unchanged.
- A failed or over-budget batch leaves existing compact facts untouched.

## Workstream 3: honest one-term publication planning

Change initial Staff/metadata publication to accept an exact term and update
only that term's manifest entry. It must not loop over all historical terms.
Existing entries for other terms are preserved byte-for-byte unless their own
separate publication is approved.

Split publication into a plan and execute contract:

1. `plan` reads only compact state and existing object pointers for one
   ED/term;
2. it returns the exact dates, affected months, current input revisions and a
   `planRevision`;
3. `execute` requires that revision and rejects stale input before any write;
4. execution uses exactly the planned ED, term, dates and months; and
5. another roster or override change invalidates the plan and requires a new
   dry run.

The dry run must report separate upper bounds for:

- D1 statements and estimated rows examined for compact metadata, Staff,
  designation, override and indexed ED/date queries;
- D1 publication-state writes, including recovery/error writes;
- R2 reads for manifests, day objects used to assemble months and existing
  revision checks;
- R2 writes for the one Staff object, Staff/metadata pointer, changed day
  objects, changed month objects, candidate manifest and fixed manifest; and
- Worker requests and maximum dates.

Where local code cannot prove Cloudflare billable rows, label the value as a
local upper-bound estimate. Do not label returned rows as rows examined.

The estimate and executor must share the same constants and planning object;
tests should compare actual mocked operation counts against the reported
ceilings. A response must never claim `broadRosterScans: 0` unless the test
captures every SQL statement used by both bootstrap and publication and proves
that only the separately approved indexed file/date reads occur.

## Workstream 4: contact-ingestion containment

Add contact-ingestion controls independent of roster automation:

- a global contact-ingestion write switch;
- an exact contact-source allowlist; and
- a payload-size and retained-object ceiling.

Apply these controls after token validation but before schema checks, D1 reads,
R2 reads/writes or pruning. Missing settings must fail closed. An unchanged
extract must not write merely to update bookkeeping or prune unless an
explicit, bounded maintenance operation was separately requested.

Limit contact discovery to the small retained window using the existing
source/date index. Cleanup must have its own maximum deletions and must never
scan or delete unbounded history. Manual single-contact corrections may remain
a separate bounded user action, but their optional shared-object publication
must continue to respect the ED build allowlist and build pause.

Before an inert deployment, verify whether the external contact flows are
currently active. Record them as paused or explicitly allowed source by source;
do not assume the roster automation switch controls them.

## Implementation order

Complete the work in this order because later evidence depends on the earlier
safety boundary:

1. Implement the unified routing decision and cohort-scoped access.
2. Make shared misses fail closed and add the explicit legacy-read stop.
3. Add contact-ingestion pre-D1 controls.
4. Add one-file compact-fact inspection/execution with hard read and write
   ceilings.
5. Restrict Staff/metadata publication to one term and implement the immutable
   plan-revision contract.
6. Replace the provisional cost report with shared planner/executor budgets.
7. Update the rollout worksheet and implementation status from measured local
   results.

Do not combine bootstrap and publication into one request. The operator must be
able to inspect compact facts after bootstrap and before approving any R2
publication.

## Focused test plan

Testing should remain local and proportionate, using the existing in-memory
and synthetic infrastructure.

### Routing and access matrix

Test Creator self, Creator impersonation, regular user and non-clinical user
against allowed ED, disallowed ED and All EDs. For every case assert:

- selected route (`shared`, `legacy` or `blocked`);
- materialised versus legacy access calls;
- exact SQL fingerprints; and
- zero legacy roster-history calls for every `shared` or `blocked` request.

Also assert that enabling the Creator canary does not change a regular user's
login/account response or access decision.

### Bootstrap tests

Use one representative Excel roster, one FindMyShift roster and the existing
large synthetic database. Cover:

- successful single-file preparation;
- `limit + 1` read rejection with zero writes;
- compact-write ceiling rejection with zero writes;
- idempotent repeat with zero writes;
- overlapping files and unrelated contributions;
- SMS continuity, trainee term boundaries and the 14-day visibility rule; and
- interruption before commit.

Capture `EXPLAIN QUERY PLAN` for the exact file-bound query and reject a broad
source/history scan.

### Publication tests

Exercise the real protected endpoint through successful dry run and execution,
not only its disabled paths. Assert:

- only the requested term is queried and changed;
- stale `planRevision` fails before writes;
- actual D1/R2 operations never exceed every reported ceiling;
- 120 dates succeeds within its declared maximum and 121 is rejected before
  publication;
- repeat execution is idempotent; and
- failure/recovery preserves the previous complete manifest.

### Contact tests

Assert disabled, missing-allowlist and disallowed-source requests make zero D1
and R2 calls. For an allowed source, test unchanged, changed, expired and
retention-limit cases with exact read/write ceilings.

### Regression set

Run only the existing focused suites covering materialisation, access,
snapshots, contacts, fixtures, VHH/queue behaviour, D1 quota guards, database
costs, local isolation and syntax. Do not expand to remote or exhaustive UI
testing during this remediation.

## Acceptance gates before Phase 7B

Phase 7B remains blocked until all of the following are recorded:

- the full routing/access matrix passes;
- a Creator canary has no observable effect on any non-canary account;
- all shared and blocked paths prove zero legacy Staff, coverage, catalogue or
  range queries;
- a fresh-schema database can bootstrap existing representative files in
  bounded, serial batches;
- bootstrap and publication over-budget cases perform zero writes;
- successful publication stays within its reported D1/R2 ceilings and changes
  only one ED/term;
- contact ingestion is zero-work when paused;
- the capacity worksheet includes bootstrap, publication, authentication,
  blocked/unavailable routes, contacts and rollback headroom;
- the rollout worksheet no longer recommends a step that globally changes
  access during a Creator-only canary; and
- an independent focused review finds no high-severity rollout issue.

Passing these local gates authorises only consideration of a push. Deployment,
migrations, bootstrap, publication and enabling readers each still require the
separate approvals and measured checkpoints in Phase 7.

## Deliverables

- cohort-aware routing and access implementation;
- bounded compact-fact bootstrap endpoint and shared planning types;
- one-term publication planner/executor;
- independent contact-ingestion guards;
- focused operation-count and query-plan tests;
- corrected `at-a-glance-phase-7-rollout-package.md`;
- updated implementation status with known limitations stated accurately; and
- one reviewed local remediation commit, with unrelated files untouched.
