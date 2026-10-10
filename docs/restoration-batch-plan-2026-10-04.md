# Remaining restoration and development batches — 4 October 2026

Status: Batch 1 roster traffic cutover accepted; Batch 2 production rollout is
merged and enabled. Automatic registration resumed after exact queue
reconciliation on 10 October; historical catch-up completion and the first real
weekly audit remain acceptance checks. See the 10 October checkpoint in
`restoration-batch2-acceptance-2026-10-05.md` for verified results and limitations.
The 11 October MMC incident exposes an operational reliability gap: the pipeline
is enabled, but reservation exhaustion can stall changed rosters. Complete the
reservation-recovery batch below before proceeding to Batch 6.
The user's agreed priority is **1 → 2 → 6 → 7 → 3 → 4 → 5 → 8**,
retaining the original batch numbers for continuity. This plan supersedes earlier
fifteen-minute freshness targets and batch ordering.

Baseline: production `ba5ee7ad` restored VHH history from 4 May, DDH five-minute
metadata polling and the four-site cached views. Existing account-wide admission,
source isolation, bounded imports and publication safeguards remain in force.

## Common delivery method

Implement each complete feature set, run one coordinated correctness/cost test
pass, deploy the reviewed batch, and perform representative live acceptance.
Avoid serial cosmetic releases or repeated full-suite runs without a new failure.
Retain independent feature/source stop controls inside each batch. A failing batch
can be disabled/reverted and its components isolated without stopping unrelated
working sites. Rollback cannot recover already consumed D1 allowance: retain
fail-closed reservations and stop before daily limits are reached.

Do not automatically restore broad SQL readers, global snapshot warm-up, runtime
DDL or Creator startup fan-out. Provide the requested outcomes through bounded
paths. Preserve clinical inputs, personal edits, existing subscription URLs and
the rollback release. Keep raw rosters and credentials outside Git.

## Batch 1 — roster freshness, next-term delivery and historical grades

### Metadata before retrieval

Every FindMyShift and SharePoint roster check must compare lightweight provider
metadata with the last successfully activated version of that exact site/file/
term window before downloading workbook content or running extraction/parsing.
Use Modified/LastModified plus a stable file identity and provider version/ETag
where available. Successful unchanged checks perform no download, parse, import,
publication, workflow dispatch or D1 status write. Avoid per-tick analytics calls.

DDH already applies LastModified and exact successful version/window checks.
SharePoint server ingestion deduplicates received versions, but that is after
the Flow supplies the file. Inspect each live Flow and put the metadata decision
before Get file content/Run script; server deduplication alone is insufficient.
Keep SharePoint change triggers as the fast path and add metadata reconciliation
for missed triggers. A file save with identical roster facts may still require
one extraction to discover the semantic no-op, but must not rewrite those facts.

Necessary exceptions are explicit: an initial import; a newly eligible term/file
not imported before; a user-authorised parser repair; and resuming failed work.
Resume queued/failed work using retained input where possible rather than fetching
the provider again. Preserve the previous valid active roster until completion.
A failed candidate must not become the successful-version watermark. Re-check
version stability around extraction and reject stale/concurrent deliveries.

### Timing and cost

Target retrievable, valid provider changes reaching the app, affected At a glance
views and subscription-feed responses within **five minutes** under normal
conditions. Start with a two-minute metadata interval, change triggers, prompt
bounded processing and a visible-browser revision check at most once per minute.
A five-minute check alone leaves no processing time. Measure the complete path,
especially GitHub dispatch/startup latency; adapt the runner if it cannot meet
the target with existing inexpensive infrastructure. Do not declare success from
metadata checks alone or silently introduce paid infrastructure.

Coalesce rapid saves and duplicate tabs. Hidden tabs stop polling, then check on
foreground return. An unchanged browser revision does not fetch/rebuild calendars
or re-run account discovery. Existing contact extraction remains at its agreed
five-minute frequency. Review HTTP feed caching so it does not add another five
minutes of server-controlled staleness. Phone calendar clients control their
own fetch schedule; the deadline applies to server/feed and visible-app freshness.

DDH's last measured unchanged tick used nine D1 reads and zero writes: at two
minutes this is about 6,480 reads/day before other sites. Measure each site's
actual costs, ordinary traffic and changed-import batches; include Power Automate
licence/action allowances, provider throttling, R2, Worker and GitHub limits.
Keep FindMyShift calls serial per API key, obey 429/backoff and never defeat an
admission stop to meet the timing target. Expose deferred/invalid/outage status.

### Next-term eligibility

Use the first day of the calendar month preceding the term's start month, in
Australia/Melbourne: 2 November 2026 → 1 October 2026; February → 1 January.
Continue checking the current term while checking eligible next-term input.
Identify each input/window independently so an unchanged team version does not
prevent the first fetch of a newly eligible term. Unpublished terms remain waiting
without repeated downloads at the same provider version.

Proposed calendar availability follows the same eligibility date once an
authoritative roster exists, replacing the fourteen-day visibility hold. Keep
this assumption explicit in release acceptance; importing and visibility must
not accidentally use different dates. Verify source selection, cached manifests,
identity suggestions, permission boundaries, personal calendars and feeds together.

### Historical grades

Grade belongs to the roster/site/term, not the person's permanent identity.
Historical views use the supplied grade for that roster; current directory grade
comes from the current term. No promotion, merge, alias correction or current
account update may rewrite past grades. Retain provenance and permit a reviewed
term-specific correction. Never infer missing historical grade from today's
provider staff list; mark uncertainty and retain authoritative historical inputs.

Correct both identified paths: manual grade overrides carrying forward despite
later roster evidence, and multi-term shared history combining overrides without
term-specific keys. Inspect parser, membership, correction publication, grouping
and historical permission readers together. Include a Junior Registrar → Senior
Registrar → SMS fixture across sites and terms, with an older manual correction,
merged identities, missing grade evidence and boundaries. Keep authorisation
policy separate from displaying a historical grade.

Acceptance covers unchanged ticks, rapid saves, missed triggers, term eligibility,
same timestamp/different provider version, concurrent arrivals, crash/resume,
budget deferral, open/hidden tabs, subscriptions and historical grade accuracy.
Use real needed provider changes for live checks, not synthetic clinical edits.
Export recoverable Flow definitions and inventory obsolete flows with this batch;
delete only after verifying replacements and the existing cleanup authorisation.

## Batch 2 — durable identities and account linking

Reconcile the parked alias branch and identity/merge plans with production.
Choose one stable ID design before migration; keep grade history separate.
Deliver preferred names, aliases, bounded duplicate suggestions, merge preview,
conflict checks, audit/history, exact reversal and preserved account/profile/feed
relationships as one feature set. Add bounded automatic linking/repair only for
unambiguous, permitted matches; ambiguous identity changes require review.
Prove concurrency, rollback, historical grade preservation, real subscription
continuity and scale costs. Candidate discovery must not scan every person pair.

5 October: implemented and accepted in an isolated synthetic cloud Preview.
Production activation awaits the identity plan’s separate approval. See
[Batch 2 acceptance](restoration-batch2-acceptance-2026-10-05.md) for measured
costs, rollback and remaining follow-up work.

## Batch 6 — disputes, messaging and secure account sessions

Build on stable identities: dispute submission, Creator queue, claim transfer,
temporary alias/rejection and in-app notifications. Implement secure session/
refresh-cookie authentication and account lifecycle integration together, with
upgrade/rollback compatibility for existing accounts and external subscriptions.
Add password recovery or additional email notifications only where requirements
and provider costs are agreed. Test complete user/Creator journeys and concurrency.

## Batch 7 — Director and reporting workflows

Existing non-clinical accounts, Director entitlement and At a glance access stay
available. First settle hospital/program scope, reporting/export/sharing and any
approval/delegation/editing requirements. Implement the agreed workflow as a batch
using the stable identities and session/access model from preceding batches.
Requirements are an external dependency, not permission to invent capabilities.

## Batch 3 — Casey and richer contacts

Obtain authoritative Casey input and join the same bounded delivery/freshness
contract. Restore its roster-backed views without inventing missing data. Add
clearly labelled on-call/non-rostered entries, switchboard instructions and
remaining contact matching acceptance. Preserve existing contact frequency/costs.

## Batch 4 — recovery and repository administration

Provide exact site/file/term dry runs, budgeted reconciliation, repair/rebuild,
facility bootstrap, version conflict review, checkpoints and recoverable outcomes.
Keep advanced capabilities scoped and normally closed; do not enable global
automatic rebuilding. Validate interrupted execution and unrelated-term preservation.

## Batch 5 — resilience and diagnostics

Deliver coalesced bounded background retries/backoff, recoverable publication
failures, persistent sampled/batched UI console history and clear freshness/error
status. Preserve current retry protections during earlier batches; earlier batches
must still handle their own errors safely. Test no retry storms or per-render writes.

## Batch 8 — optional work, polish and consolidation

Prioritise remaining proposals, mobile improvements, retention policy, test debt,
architecture tidying and documentation consolidation. Reassess deferred physical-
presence, any-pair matching and overview editing against real requirements.
DDH source subscription input is removed from the backlog: administrative
FindMyShift API access retrieves the team-wide shifts. Apple/Google subscription
outputs remain live and are unrelated to source ingestion.

## Effort and next step

Recommend **Medium** for Batch 1 and Batch 2 implementation: they change source
watermarks, concurrent delivery, calendar revisions, historical grades and identity
relationships. Low is suitable for routine cleanup and measured acceptance once
the design and regression checks pass. Higher effort is not currently required.
Pause after this plan/document cleanup; the next resumed implementation starts
with Batch 1 rather than another planning cycle.


## Batch 1 implementation checkpoint — 4 October

- Incorporated production `d318bf7b` and preserved the other chat's clinician
  visibility, startup and navigation changes; their regression checks pass.
- Added Melbourne preceding-month eligibility in provider windows and shared
  readers, plus exact-term grade corrections and removal of undated historical
  membership fallback. Historical DDH repairs must use retained term-specific
  evidence rather than today's live staff list.
- Added edge/R2-only opaque change fingerprints, coalesced visible-tab checks,
  safe calendar/view refresh, shorter feed caching and two-minute DDH polling.
  Unchanged browser checks contain zero D1 operations. Imports retain account
  admission and source isolation; processing is serial within each site.
- Added authenticated SharePoint pre-download checks, including unchanged,
  pending, invalid-candidate and budget-stop behavior. Added a private exported
  Flow package preparation tool; the MCH backup/update package is ready outside
  Git. Existing contact schedules are unchanged.
- Production admission passed at approximately 94,000 account-wide reads and
  570 writes. The compact visibility migration cost 12 reads and 12 writes.
- **Outstanding:** Mac unlocked for MMC/MCH/VHH export/import and live Flow
  validation; verify the actual Power Automate allowance before enabling
  two-minute metadata reconciliation, verify the VHH library/path from its
  export, then measure real changed-roster delivery across all sites. Five-minute
  all-site delivery has not yet been declared complete.
- Microsoft documents 6,000 daily requests for seeded/trial licences and 40,000
  for Power Automate Premium. Reconciliation must fit the verified existing
  licence alongside contact flows; do not silently add paid capacity.
  Sources: https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/faqs
  and https://learn.microsoft.com/en-us/power-platform/admin/api-request-limits-allocations.


Production acceptance: PR #22 merged as `f4883633`, deployed in
`0c505357-22f7-42cf-b5f8-4e27d21e7bdb`. All four public fingerprints and the
unchanged MMC/MCH/VHH preflight passed. Watchdog version
`c2447802-0253-4b89-9418-7ab4519bdd06` reports two-minute polling, configured and
unpaused. DDH current term remained unchanged; the newly eligible next term
imported 78 doctors/1,654 events, with shared publication complete at 08:47:50 UTC,
4 minutes 36 seconds after submission. This measures submission-to-publication,
not worst-case provider-change-to-visible-browser delivery. The latter and other
sites still need acceptance after their Flow changes.

The ten-minute post-deployment account-wide comparison passed GO at 96,010 rows
read/570 written (analytics excludes its latest settlement window); current
maintenance admission also passed. No extraordinary credit spending was used.
A final failed-save regression extends edit protection beyond the pending save
interval until a successful save. Actual browser apply tests cover in-flight
edits, account switches, stale payloads, date filters and scroll.

## Batch 1 checkpoint — 5 October

- Exported and retained recoverable MMC/MCH/VHH definitions privately. MMC had
  downloaded immediately; MCH and VHH had five-minute delays. Live MCH history
  showed saves taking 5–24 minutes because queued saves waited behind the delay.
- Successfully imported updates to the existing MMC and MCH Flows. They reject
  superseded trigger versions, check the successful import watermark before
  downloading, and use a 30-second stability interval only for changed files.
  ETags are rechecked after settling and after downloading. Existing connections,
  ingestion credentials and serial processing are preserved. Contact Flows stay
  at five minutes. VHH's real library ID was verified from its export.
- Added bounded library reconciliation: two SharePoint metadata requests, one
  authenticated batched preflight, and a loop containing only changed eligible
  files. Missing next-term files wait independently; incomplete inventories and
  provider outages defer only their own library. Unchanged metadata performs no
  content retrieval, parse, import, dispatch or D1 write. Tests cover same version
  across different term files, wrong folders, missing files and source outages.
- Prepared its private Flow package with a future start time; it cannot run
  automatically until the schedule is deliberately enabled. Two-minute idle
  reconciliation is approximately 3,600 Power Automate actions/day, in addition
  to contacts and changed imports. The UI showed Free and Per App Baseline Access
  with premium capability, but did not establish the owner's actual request
  allowance. Confirm allowance before enabling this schedule.
- D1 comparison passed GO: 128,759 rows read / 663 written; projected daily
  usage about 197,000 reads / 1,706 writes. No database migration is required.
- **Live work pending:** the Mac locked during VHH import mapping, before the
  update was applied. Unlock, finish VHH, verify all three enabled Flow versions
  and unchanged runs, import/validate reconciliation and enable a schedule that
  fits the existing allowance, then measure provider-to-visible-app delivery.
  Batch 1 is not declared complete until those acceptance checks pass.
- Batch 2 recommendation remains Medium for transactional identity merges,
  reversal and subscription/account mapping; Low is suitable for cleanup and
  routine acceptance after implementation passes.

### Live acceptance checkpoint — 5 October afternoon

All three existing roster Flows (MMC, MCH, VHH) are imported and enabled.
The new reconciliation Flow is `f46e3e52-ff98-4db9-9067-51a5de909236`.
Its fresh manual run succeeded in eight seconds: both metadata inventories
succeeded and only two changed files downloaded. MMC imported 153 doctors and
4,150 events in about 40 seconds. VHH exposed a remaining global dispatch lock:
MMC's active lease prevented VHH dispatch. Dispatch leases are now bounded to
the exact source ID, retaining one active processor per site. A newly observed
provider version with identical workbook content now records one successful
receipt, preserving old import evidence and preventing subsequent downloads;
already verified versions remain zero-write. Regression and account budget
checks pass.

Automatic reconciliation remains held at a future 2030 start time pending
confirmation of the existing Power Automate allowance; no paid capacity has
been added. All-site five-minute delivery remains an acceptance requirement,
not a claimed result. Batch 2 waits for Batch 1 acceptance.

## Revised traffic policy — user clarification, 5 October

The user clarified that five minutes is the maximum roster retrieval frequency,
not a provider-change-to-browser deadline. Supersedes the earlier two-minute
polling proposal. Use one scheduled five-minute SharePoint metadata reconciliation
Flow and retire the overlapping save-trigger roster Flows after acceptance.
Check DDH metadata every five minutes. Current and eligible next-term windows
remain independent. Only a newly observed provider version needs retrieval;
autosaves are coalesced into the latest snapshot, not replayed. The preflight
also defers a newly changed exact file within five minutes of its last ingestion,
without D1 status writes. A later scheduled tick handles that pending change.

Five-minute reconciliation has five fixed actions/run, or 1,440/day. Existing
contacts have 3 MMC/MCH actions, 4 DDH actions and 3 VHH actions/run, or 2,880/day.
Fixed total: 4,320/day before changed downloads, retries and other Power Platform
use. No shared Process capacity is assigned or purchased. Historical roster
Flow usage on 4 October was MMC 238 / MCH 208 / VHH 64 actions; those older layouts
are not a forecast ceiling. Contact extract-before-transmit optimisation, including
DDH's unnecessary unchanged workbook retrieval, remains a separate follow-up.
Identical content under a genuinely new provider version may need one download to
establish equivalence; the successful version receipt prevents repeated retrieval.

D1 comparison returned GO again after the import burst settled: 313,762 reads /
2,936 writes, projected daily reads approximately 589,457. Admission controls,
account budgets, source isolation and bounded publication remain enabled.

### Five-minute roster cutover completed — 5 October evening

Production code is merged at `488734a9` (PR 26), Pages deployment
`bb030e61-56b1-458a-83c5-4c24a5bc9dd5`. The DDH watchdog reports healthy,
enabled and `intervalMinutes: 5`; its deployed cron is `*/5 * * * *`.
All existing production admission and publication safeguards remain enabled.

The reconciliation Flow's recurrence is now five minutes with the future start
removed. Automatic runs at 19:01, 19:06, 19:11, 19:16, 19:21, 19:26, 19:31 and
19:36 succeeded. The 19:36 run (`08584104174729700971009164296CU12`) completed
both SharePoint metadata queries and version checking; its changed-file loop
skipped retrieval. Earlier fresh reconciliation runs verified actual changed
MMC/VHH import and publication. No synthetic clinical edits were made.

The overlapping save-trigger Flows are verified **Off**, retained for rollback:
- MMC `1dd007f9-868d-4fce-a0fd-ed03d413a4ac`
- MCH `0d191bae-d7c5-47f1-a1c9-e754cb288790`
- VHH `2984416c-7c89-4400-a9e5-6ff1e990f191`

Contact Flows were left running on their existing five-minute schedules.
Their extraction/transmission optimisation remains follow-up work, especially
DDH's unchanged contact workbook download. The roster cutover does not claim
that contact traffic has been eliminated.

Settled analytics check `/private/tmp/oct5-batch1-ninth.json` returned GO:
388,149 reads / 3,507 writes; projected daily usage about 1.17 million reads /
9,509 writes. The previous sample was held only because its comparison gap was
too short. Code regression checks passed before deployment. This completes the
roster traffic cutover under the revised policy; Batch 2 remains separate.

## Priority reliability batch — reservation recovery, 11 October

Status: implementation and coordinated regression tests complete; production
rollout and real MMC acceptance are in progress. This is one coordinated delivery batch spanning Batch 2 maintenance and Batch 5 resilience.
It takes priority over new features. Medium effort is recommended.

### Evidence and scope

MMC version 451 was received at 07:06 AEDT on 11 October but remained queued.
Repeated GitHub jobs reported success after logging “Account budget deferred;
queued work retained.” At inspection, actual account usage was approximately
451,000 reads / 4,600 writes. Twenty-seven unfinished maintenance receipts held
76,176 reserved writes, exhausting the maintenance write allowance. This is not
evidence of a five-million-read limit breach.

Diagnose the requests that lost settlement, including their routes, timing,
CPU/resource failures, durable progress and reservation sizes. Identity work is
a suspect given the reservation pattern, not a confirmed attribution for every
unfinished receipt. Keep diagnostics and credentials private. Preserve roster
inputs, historical grades, identities, human decisions and subscription URLs.

### Implementation and controlled recovery

1. **Trace and contain the faulty maintenance path.** Correlate unfinished
   receipts with available invocation telemetry and durable jobs/checkpoints.
   Pause only the implicated optional identity backfill/audit if it continues
   stranding allowances; retain routine roster checks and cached app readers.
   Obtain fresh account-wide admission evidence before recovery writes.
2. **Repair reservation lifecycle and request sizing.** Reserve tested read/write
   ceilings appropriate to each bounded operation, including indexes and
   bookkeeping. Reduce/split the failing request's work where measurements show
   CPU or statement pressure. Make settlement idempotent and durable progress
   recoverable. Add bounded reconciliation outside the originating request so a
   killed request cannot depend solely on its own `finally` block for recovery.
3. **Recover old reservations conservatively.** Define the evidence required to
   reconcile completed, partially completed and abandoned work. Account for
   settled measured usage without double counting or refunding work twice.
   Expiry alone, a successful GitHub job, or an unknown response does not prove
   that no D1 work occurred. Retain uncertain reservations unless a reviewed
   conservative accounting bound proves safe headroom. Recovery must not raise
   limits, bypass admission, erase receipts or fabricate actual usage.
4. **Protect changed-roster delivery.** Give routine imports/publication priority
   over historical identity registration and weekly suggestions. Budget optional
   work separately within the shared account envelope; it must pause/resume at
   durable checkpoints before consuming roster headroom. Preserve the five-minute
   metadata cadence and unchanged-version download suppression.
5. **Expose real outcomes and avoid repeated no-op jobs.** Distinguish completed,
   deferred, failed and awaiting-publication states through the processor and
   dispatch lifecycle. Check admission before launching expensive GitHub work;
   coalesce retries with bounded backoff and resume when admission returns.
   Report materially delayed roster syncing in the existing exception/status
   UI, without normal-operation banners or per-render D1 writes.
6. **Recover and verify the current MMC update.** Revalidate which provider
   version is now current, then resume the retained latest import through the
   ordinary bounded path. Do not re-download unchanged files or replay superseded
   versions. Check Josh Feek's actual 11 October source entry, parsed result,
   published On shift and personal calendar output. Retain the previous valid
   publication until the replacement validates and promotes atomically.

### Coordinated tests and live acceptance

- Cover failure before reservation, after reservation, during partial writes,
  before/after durable progress, during settlement, and after settlement when the
  response is lost. Simulate resource termination as well as caught exceptions.
- Prove concurrent reconciliation/settlement cannot double refund; partial or
  unknown work remains conservatively accounted for; old UTC-day receipts cannot
  refund today's grant. Test midnight and analytics delay/outage boundaries.
- Exercise identity backlog alongside a changed roster: optional work defers,
  the roster proceeds when safely admitted, and exhausted account capacity still
  stops all maintenance. Verify actual operation costs against reservations.
- Test deferred workflow outcomes, admission-before-dispatch, retry coalescing,
  delayed-sync warnings and recovery without repeated interface/database writes.
- Run the affected maintenance/account-budget, import/queue, identity and cached
  publication suites as one pass. Repeat only for failures or subsequent changes.
- Deploy the tested batch and observe a real changed-roster import/publication,
  automatic identity progress and request settlement. Compare existing ordinary
  and Creator subscription URLs/event UIDs for unintended changes. Obtain settled
  account-wide analytics after the operation, not just a pre-operation sample.

Acceptance requires no unexplained abandoned-grant accumulation or retry storm,
measured requests within their ceilings, and admission below the existing
4,000,000-read / 80,000-write maintenance thresholds while retaining Cloudflare's
5,000,000-read / 100,000-write daily limits. No quota increase is part of this work.
The real weekly identity audit remains a separate observed acceptance check; an
audit outside its intended window must not be forced merely to complete testing.

### Rollback and completion

Retain the preceding deployment and targeted optional-maintenance controls.
If measured costs, reconciliation or concurrency violate the safety envelope,
stop the affected optional work and revert the batch while retaining receipts,
source files, checkpoints and the last valid publication. Never restore stale
reservation counters over newer allocations or undo already consumed usage.

Completion means the current roster is correctly published, the demonstrated
failure mode recovers safely, routine imports retain priority, deferred outcomes
are honest, and live settled usage passes the unchanged safety gate. Record any
remaining external/provider or scheduled-audit verification explicitly.

Implementation checkpoint, 11 October: automatic identity work now reads the
indexed historical names cache in one-name checkpoints, with a 96-statement
ceiling and smaller reservations. Three unfinished requests pause optional
maintenance; 100,000 reads and 20,000 writes are protected for routine imports
inside the existing account grant. The five-minute roster metadata cadence is
unchanged. Deferred GitHub processors expose a deferred result and do not
repeatedly launch while the account grant is exhausted. A small R2 status object
shows an exception after 15 minutes of deferred delivery; ordinary use is quiet.

Migration 0041 adds receipt purpose, execution deadline, recovery cutoff and
reconciliation markers. New SQL calls stop after two minutes or UTC midnight,
whichever is earlier. Recovery waits for the full admitted statement bound,
settlement calls and margin to drain, then for validated account analytics to
cover that cutoff. This uses the documented [D1 maximum 30-second query duration](https://developers.cloudflare.com/d1/platform/limits/).
Legacy receipts without an enforced deadline stay reserved through their UTC
day. Unknown metadata is never rewritten as measured zero usage. Recovery and
settlement use atomic conditional updates to prevent duplicate refunds.

The pre-release settled analytics check returned GO at 10:13 AEDT: 472,689
reads and 4,761 writes. Migration 0041 was the sole pending migration and was
applied after saving the current grant and receipts privately. Recovery,
identity operations, cached registration, queue outcomes, delayed warnings,
source isolation, import idempotency, On shift access, contact health, login
containment and request attribution checks pass. Live publication and subsequent
settled measurements still need to be recorded before declaring completion.

Publication continuation is also checked through the existing five-minute
metadata path. A successful file import with a pending indexed shared-view job
now resumes that job without fetching or parsing the unchanged workbook. The
processor and GitHub queue step accept publication-only work, retain delayed
warnings until both import and publication finish, and expose deferred results
when another publication checkpoint remains. Tests exercise this exact path.

The expanded publication tests also pass after updating their old 14-day
visibility fixtures to the already-restored preceding-month policy and adding
the retained raw-file metadata required by the existing delivery ordering guard.
Cross-midnight tests now prove that an in-flight bounded receipt from yesterday
reduces today's available headroom until a settled cutoff covers it. Recovery
only refunds that receipt's original UTC-day ledger, never today's allocations.

Optional maintenance reserves 32,768 reads to cover the fixed receipt-inspection
ceilings as well as the cached-name completion check; measured unused capacity
is refunded. Its write estimate remains 512 per registration/audit checkpoint,
with the existing 4,096 estimate retained for identity publication.
