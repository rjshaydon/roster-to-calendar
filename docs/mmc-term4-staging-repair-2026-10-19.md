# MMC Term 4 staging repair — 19 October 2026

## Diagnosis

The five-minute Power Automate reconciliation is enabled and its runs succeed.
The app has retained `AdultTerm4.2026.xlsx`, latest observed provider version
19.0. Discovery and file transfer are working. No next-term MMC file is active.

Provider version 7 saved all 68 event batches (4,078 events) and 35 presence
batches, but deferred before preparation/activation. Versions 8 and 11 started
new inactive copies, stopping at 11 and 15 event batches respectively. Latest
versions remain queued. The processor was then repeatedly parsing the current
term before discovering that the write grant could not admit another step.
GitHub's successful workflow conclusion represents a handled deferral, not a
successfully published roster.

Read-only analytics at 04:49 UTC: 429,940 account reads / 71,453 writes, almost
entirely production. Event and presence insertion dominate measured writes.
Subsequent production grant inspection: 458,787 observed reads / 71,614 writes,
2,258,022 usable reads but only 39 usable writes after allocations/forecast.
No unresolved incomplete maintenance receipt was found. This is not the earlier
one-off preview forecast issue; real partial import writes have been consumed.
The internal maintenance ceiling remains 4,000,000 reads / 80,000 writes.

## Repair

- Additive migration 0053 adds inactive-file lease/progress records and bounded
  source/term/file indexes. It does not rewrite or activate clinical facts.
- Only `monash-adults` initially enables staged reuse. An exact provider file
  and exact shift-start range can reuse a fully staged inactive event baseline.
  Read at most eight indexed candidates in that exact source/range; never use
  an active roster or a different provider file as the baseline.
- Compare on GitHub, reading small indexed pages, and apply only changed facts
  and missing presence/compact rows. Large true changes retain the existing
  guarded full-staging fallback and all daily reservations.
- A durable lease prevents an old worker from modifying the borrowed file.
  A newer provider revision can take over even a partially applied rebase.
  Existing provider-order checks reject stale activation and completion.
- Save hash-pinned completion batches privately in the existing R2 source bucket
  before applying facts. Lost replies and restarts resume the original plan;
  they do not rediff a partially modified baseline into a conflicting plan.
  Checkpoint objects remain private retained recovery evidence, under the same
  storage scope as raw rosters; no new public routes or services are introduced.
- Each 16-entry batch and its cursor commit atomically while the file remains
  inactive. Exact final counts, compact metadata and current-source ordering
  precede the existing guarded activation/publication pipeline.
- An unpublished term receives queue priority over current-term corrections.
  Dispatch/pending/runner preflight require usable headroom (2,048 writes /
  8,192 reads), not merely a positive remainder. No cap, forecast, actual usage
  or unknown reservation is relaxed.

## Verification

`npm run test:roster-staging-reuse` exercises the actual derived HTTP routes:
4,320 synthetic events, 4,318 unchanged rows retained, partial presence, changed/
removed/added shifts, private R2 checkpoints, lost reply after a committed batch,
transaction rollback, a later autosave taking the lease, old-worker rejection,
exact compact-builder parity, normal finalization, current-term preservation,
queue priority and an unusable nonzero grant.

Existing source-isolation, wire-protocol (real MMC/MCH fixture workbooks and
100,001 historical deliveries), patch, restoration-mutation, delivery-reliability,
queue-failure, database-cost, account-budget/summary, cold-start and syntax checks
pass. Tests are local/synthetic; no additional disposable cloud databases were
created. No browser/calendar reader behaviour or password/session feature changes.

Automatic approval review rejected an optional local export of identifiable
production roster rows. No such export was performed; this release uses
synthetic correctness tests and aggregate production inspection instead.

## Release and remaining gate

Migration 0053 is applied to the production database. Deploy the tested code and
verify normal five-minute metadata reconciliation stops launching unaffordable
jobs. Retained version 19 and the inactive baseline must remain available.

**Real MMC Term 4 publication is not yet verified.** Its remaining work must be
admitted by the existing grant. With only tens of usable writes, do not manually
dispatch a doomed run, clear reservations or increase ceilings. Existing automatic
reconciliation can resume when usable allowance returns. Then verify a successful
latest-term activation, completed shared/doctor publications and visible calendar
coverage. Merely deploying this repair or seeing a green workflow does not close
that live gate. Do not start other restoration batches before resolving it.

The separate manual/automatic authority refinement remains planned, not implemented
by this repair.
