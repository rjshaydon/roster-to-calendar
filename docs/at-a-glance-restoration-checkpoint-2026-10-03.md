# At a glance restoration checkpoint — 3 October 2026

The clinician-access implementation is saved in commit `3a3a1b8a` on
`codex/clinician-access-history`. This checkpoint adds the coordinated
Working together and ED Staff restoration work. Production migration 0038
is applied and verified. Application changes have not been deployed or merged.

## Implemented behaviour

- Combined views return enabled, authorised hospitals even when another
  hospital's reader is unavailable. They identify unavailable hospitals and
  dates instead of implying complete results. Explicit hospital requests
  continue to respect rollout gates and clinician permissions.
- SMS and CMO retain equivalent all-site permissions. Other clinicians use
  verified current-term sites, including multiple rotations or locums.
  Historical Working together searches use memberships for the searched period
  and restrict colleague names and shifts to the authorised sites and dates.
- Published range, daily and staff readers distinguish absent manifests,
  staff objects, month objects and coverage gaps from genuinely empty rosters.
  Available results survive partial publication failures.
- Browser snapshots include reader configuration in their scope and revisions.
  Disabled readers clear displayed cached roster details. Temporary publication
  failures can retain previously authorised cached data with the existing error
  handling. Partial results cannot reuse a complete-results response silently.
- Incomplete monthly publication cannot replace a user's personal calendar
  snapshot. Explicitly empty published coverage still supports roster removal.

## Verification

Passed focused suites: clinician access, facility access, facility rollout,
facility snapshots, facility contact access, cached roster insights, restoration
mutations, D1 account budget and client request budget. Syntax and whitespace
checks passed. Clinician-access tests cover CMO parity, locums, historical
rotations, direct API denial, impersonation, corrections, missing publication,
reader changes, emergency pause and bounded indexed reads.

Three broader failures were independently reproduced on the original main
commit: fixtures (automated correction identity assertion), facility maintenance
(launch flag expectation), and facility materialization (import SQL missing
`name` column). The updated materialization suite reaches that same existing
SQL failure; assertions before it pass. These are not claimed as passing suites.

## Publication audit and remaining work

Read-only production/R2 audit found current-term publication at MMC, DDH and
MCH, and current/next-term publication at VHH. Earlier active membership and
file coverage exist for MMC, DDH and MCH, but the published manifests largely
lack those earlier terms. Casey has neither a manifest nor retained term
coverage; enabling its reader alone would not restore usable roster results.
Private audit output contains no data committed to Git.

1. Verify the queued DDH pilot for 4 May–2 August 2026 after maintenance runs:
   job complete, old and current terms retained, all expected months/day
   pointers readable, and historical staff/shift searches available.
2. Recheck account budget before publishing other retained historical terms.
   Use bounded per-term jobs and verify each publication. Do not reset or
   globally rebuild the repository.
3. Deploy the completed application change through the normal release process,
   then smoke-test SMS, CMO, single-site, multi-site, historical-only and entered
   accounts. Confirm Casey notices and direct-site denial, and normal calendar
   and contact access. Deployment remains pending.
4. Restore Casey only after its roster input and publication are available and
   validated. Advanced administration and durable Doctor Names merging remain
   separate work.

The DDH pilot is pending, not verified complete. Automatic approval review
rejected manually dispatching the existing general maintenance workflow:
it can process other pending production jobs beyond the DDH pilot. No manual
dispatch occurred. The normal scheduled workflow may process the queued job.
Explicit approval is needed before retrying that broader manual dispatch.

## Recovery

Migration 0038 is additive (one index and ten permission-cache invalidation
triggers) and preserves staff contribution rows. A fresh Time Travel recovery
point was captured before application. Prefer reverting application behaviour
over restoring the entire database, which would also affect intervening writes.
See `clinician-access-release-preflight-2026-10-03.md` for measured quota and
migration evidence. No other migration was applied.


## Approved maintenance outcome — 3 October, 2:13 pm AEST

The user explicitly approved the general maintenance dispatch after its broader
scope was explained. Run [37095222824](https://github.com/rjshaydon/roster-to-calendar/actions/runs/37095222824)
completed successfully on `main`, with current-view seeding disabled. It also
processed a current DDH FindMyShift roster update (149 doctors, 3,649 events).
No application deployment or merge was performed.

The historical DDH job is now **complete**, with 13 daily batches and four
monthly builds. Its manifest was published at 2:09 pm AEST and retains both
4 May and 3 August term entries. Read-only verification at 2:10 pm AEST checked:

- All 91 historical dates have day pointers, and all 91 previously published
  current-term dates still have pointers. Two current day revisions changed
  during the ordinary roster update; exact old revision equality was therefore
  not required.
- Seven monthly objects (May through November), both term staff objects, and
  historical boundary/current sample day objects are readable: 13 R2 objects.
- The published-reader functions return 128 historical staff entries and 3,597
  historical events with no missing coverage. Current-term range reads return
  3,649 events with no missing coverage.

The same-day maintenance ledger contains 28 finished receipts, all with complete
metadata, reporting 37,600 reads and 997 writes. No reservations are unfinished,
and the account maintenance stop reason is empty. These are maintenance-meter
figures, not the entire account's daily usage.

The latest settled account-wide sample returns GO with 40,351 reads and 1,343
writes, but its interval ends at 1:56 pm AEST, before this workflow started.
It therefore does not yet measure the run's entire account-wide cost. An initial
comparison against the seven-minute-old start sample returned STOP solely for
`sample-gap-too-short`; comparison with the original valid baseline returned GO.
No quota exhaustion was observed.

The earlier pending/approval-blocked status above is superseded by this outcome.
Remaining work is publication of other retained historical terms after budget
admission, application release and account smoke tests, and restoration of Casey
when its roster input is available. The pilot's roster publication alone does
not deploy the new historical trainee access rules.

Private evidence:

- `/private/tmp/clinician-ddh-maintenance-run.log`
- `/private/tmp/clinician-ddh-history-publication/summary.json`
- `/private/tmp/clinician-ddh-history-completion-budget.json`
- `/private/tmp/clinician-d1-maintenance-latest.json`
