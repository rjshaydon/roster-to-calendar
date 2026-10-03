# At a glance restoration checkpoint — 3 October 2026

**Restoration completed and deployed on 3 October 2026.** Migration 0038,
clinician access, Working together, ED Staff and retained historical publications
are complete for MMC, DDH, MCH and VHH. Casey Terms 3 and 4 are explicitly
deferred by the user until its authoritative input and syncing are available.
The earlier pending statements below describe historical checkpoints and are
superseded by this completion record.

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

## Final release and verification — 16:20 AEST

Production code is merge commit `0246edf26831275314bdbc446bb3458f6ca6abda`,
successful deployment `a8b89ad4-0f71-41b0-a262-955bd66b75b2`.
PRs [9](https://github.com/rjshaydon/roster-to-calendar/pull/9),
[10](https://github.com/rjshaydon/roster-to-calendar/pull/10) and
[11](https://github.com/rjshaydon/roster-to-calendar/pull/11) are merged.
The final fixes bound timezone work and preserve coverage separately for each
term, including older publications without the new manifest field. An older
empty term cannot clear a current personal calendar; inclusive final dates are
covered correctly.

Both additional maintenance runs succeeded (37100438538, 37101292192).
All seven targeted historical jobs are complete. All 155 maintenance receipts
have complete cost metadata and finished: 121,664 measured reads and 5,247
writes, with no budget stop. Account-wide analytics at 06:14 UTC reported
195,601 reads and 3,328 writes and admission GO; analytics lag the synchronous
receipt ledger. These are different overlapping measurements, not additive
whole-account totals. Five-hour Codex usage was 35% used at the final release.

74 production publication objects were verified through the actual readers.
MMC, DDH and MCH cover 2 February–1 November. VHH retained Term 2 covers
27 July–2 August; current coverage includes 3–23 August and 21 September–
1 November. Earlier Term 2 dates and 24 August–20 September are reported as
unavailable. Its retained next term remains hidden until 19 October. Existing
current/future daily pointers were preserved. No missing Casey data was invented.

Live acceptance used the existing authenticated creator session: Working
together rendered full current and Term 2 history across enabled sites;
ED Staff All displayed all four hospitals with a Casey notice; DDH On shift,
contacts and By stream Fast Track displayed actual current assignments.
The existing personal calendar remained available. CMO parity, multiple current
sites/locums, historical trainee scope, direct API denial, impersonation,
corrections and cache invalidation were verified by behavioural tests using
isolated accounts; production trainee/CMO credentials were not required or changed.

Clinician access, facility access/rollout/contact access, personal snapshots,
cached insights and restoration mutation checks pass. The Worker compiles.
Three pre-existing broader-suite failures were reproduced on the base revision:
fixture identity assertion, maintenance automatic-launch expectation, and
materialization import fixture missing column `name`. They remain documented
limitations, not passing checks. Live bounded publication completed successfully.

Advanced global administration stays disabled and durable doctor-name merging
remains a separate project. Casey input/sync is the explicitly deferred follow-up.
Private evidence is in `/private/tmp/clinician-restored-history-verification/summary.json`,
`/private/tmp/clinician-restoration-final-ledger.json`, and
`/private/tmp/clinician-restoration-final-budget.json`; no credentials or raw roster
objects have been committed.
