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

An aggregate-only follow-up counted 8,290 planned presence rows and 4,309 already
prepared for version 7. About 3,981 rows remain before newer-revision differences,
compact facts and publication. Their conservative indexed-write cost alone is
19,905, so the observed remaining grant cannot complete this work today.

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

Migration 0053 is applied to the production database. Source commits `00657fe6`
and `431340ff` are pushed. Production deployment `85c392f1` is successful;
the root and session-options endpoints return HTTP 200, with Requests and
passkeys unchanged. Retained version 19 and the inactive baseline remain
available. A post-deployment primary read confirms no active MMC Term 4 file
and only 42 usable writes, with no overrun/stop flag. Check normal five-minute
metadata reconciliation stops launching unaffordable jobs.

Post-release observation: the normal 16:11 Melbourne reconciliation succeeded
in two seconds. Subsequent primary reads showed no new dispatch after 16:06,
the latest MMC Term 4 input still queued, and the same 42-write remainder. This
confirms safe retained/deferred behaviour for that cycle, not live import or
publication acceptance. Ordinary five-minute checking remains On.

The first deployment hit Workers Free's 64-variable limit. Remove the redundant
production `FACILITY_BOOTSTRAP_EXECUTION_ENABLED=false` entry: its tested default
is false, so recovery execution stays closed. This leaves room for the scoped
staging-reuse control without upgrading the plan or changing permissions.

**Real MMC Term 4 publication is not yet verified.** Its remaining work must be
admitted by the existing grant. With only tens of usable writes, do not manually
dispatch a doomed run, clear reservations or increase ceilings. Existing automatic
reconciliation can resume when usable allowance returns. Then verify a successful
latest-term activation, completed shared/doctor publications and visible calendar
coverage. Merely deploying this repair or seeing a green workflow does not close
that live gate. Do not start other restoration batches before resolving it.

The separate manual/automatic authority refinement remains planned, not implemented
by this repair.


## Write-efficiency follow-up, 19 October

The original 80,000 ceiling was conservative rather than derived from measured
uncertainty. Admission now starts from the Free allowance of 100,000 indexed
writes and subtracts actual account usage, unsettled reservations, the ordinary
rest-of-day forecast, and a disclosed uncertainty allowance. That allowance is
one extra ordinary-traffic window (25-minute maximum analytics lag + 10-minute
grant lifetime) at the measured ordinary rate, plus 256 writes for bounded
admission/recovery bookkeeping. The existing 4M read ceiling, atomic shared
reservations, stale-analytics rejection, unknown-work retention and cost-overrun
stop remain. This is not permission to spend the entire account allowance on imports.

For the reviewed `ROSTER_STAGING_REUSE_SOURCES` scope (currently MMC adults), the
server negotiates a hash-pinned `presenceMode: cached`. New full imports advance
the duplicate-presence cursor in one guarded transaction; incremental/reused
staging omits daily-presence expansion and inspection. Existing pinned legacy
jobs retain their original mode. Clinical event rows, compact term/grade facts,
source/version ordering, invisible staging and atomic promotion remain intact.
Existing daily-presence rows are retained for legacy recovery, not bulk deleted.

The remaining live daily-presence reader in facility access now probes only
active source files with applicable coverage and the exact file/doctor index.
Its cached decision revision changes, and repeated access still costs one
indexed cache read with no writes. Inactive staged rows cannot grant access.
Comparison, On shift and subscription publication continue using their normal
published caches. No frontend loading or contact refresh changes were introduced.

Local verification includes both legacy and cached 4,320-event autosave/rebase
journeys, rollback/lost response/supersession, current-term preservation, cached
full imports with overnight/leave events, server rejection of unapproved mode,
MMC/MCH real-fixture incremental clinical-fact parity, access scope/index/cache
checks and shared-grant concurrency/recovery tests. Production acceptance follows
only after a real Term 4 activation and publication; this section alone is not a
claim that the missing roster has been restored.


### Production acceptance checkpoint after follow-up

App/runner commit `24638644`, production deployment `848dfafb`.
Actual provider revision 19.0 succeeded at 2026-10-19T05:47:21.643Z:
132 doctors, 4,075 events, coverage 2026-11-02 through the overnight
boundary 2027-02-01. Reused 4,078 staged events with 33 changed facts,
18 cached-mode batches, prepared and activated. Existing current term retained.

At the last verified publication checkpoint all 13 Term 4 day batches and the
first month were complete; remaining monthly/personal source publication was
pending, without a recorded error. Normal continuation workflow
39205895062 was running after initial workflow 39205844065 reached its bounded
pass limit. Do not equate that workflow's generic deferred message with a D1
cost-overrun: stop_reason remained empty and admitted writes remained positive
(last inspected 5,478). Latest settled analytics still lag this activity.

Next acceptance: check facility_refresh_jobs for MMC Term 4 completion, normal
processor result and doctor-source publication, then inspect MMC 2 Nov in the
owner's authenticated Safari. The Mac locked during the visual check; unlock
requested. Do not report full end-to-end publication as verified yet. Scope of
no-presence mode remains the reviewed MMC adult source; other sources retain
legacy presence behavior pending their rollout. Tests passed before deployment.
