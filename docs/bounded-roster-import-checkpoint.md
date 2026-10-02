# Bounded roster restoration batch — 1 October 2026

This checkpoint replaces the pre-reset work list. Production overwrite protection
was released as 098bc2a2; MMC/MCH/DDH contacts were verified before this batch.
User authorised a combined restoration batch and safe automatic continuation.

## Implemented and locally verified

- Deterministic manifests and hash-pinned, ordered chunks for MMC, MCH, DDH and
  VHH; 1,250 weighted facts per chunk, 512 KiB payload, 512 doctors, 25,000
  events, 5,000 issues, and 180-day maximum workbook span (120 days per event).
- Inactive staging, transaction receipts, durable event/presence cursors,
  initialization lease, count validation and fenced activation of disjoint terms.
  Current terms remain intact; incomplete terms remain hidden.
- Exact retained-term corrections keep the existing diff path. Large overlapping
  replacements remain guarded. Completion retries repair bookkeeping after an
  activation without deleting live data.
- Shared UTC-day reservations: 10,000 indexed writes and 500,000 reads for
  automated import/publication work across all sites. This leaves headroom under
  the restoration stop thresholds for ordinary traffic. It does not replace
  account-wide analytics admission. Failed work consumes its reservation;
  progress survives and resumes after the allowance resets.
- Corrections reserve old/new daily presence, metadata and SMS continuity costs.
  Local migration-index checks prevent write estimates ignoring added indexes.
- Source-scoped overview queue commits with roster activation. Seven-date
  publication steps retain the previous R2 manifest until successful completion.
  Partial monthly updates preserve untouched dates. Every step reserves an upper
  read bound; unused reads are refunded only with complete D1 billing metadata.
- Immediate publication after queue imports plus a six-hour continuation workflow,
  using the existing global roster concurrency lock and rotating source order.
  DDH checks LastModified before downloading; retained MMC/MCH/VHH jobs resume
  without fetching unrelated sources. Paused/exhausted work skips workbook parsing.
- Optional one-time current-term view seeding uses at most 32 compact active-file
  records per source. It does not rebuild historical terms.
- Cached automatic On shift launch restored for eligible on-duty users. Legacy
  broad builders, history queries, identity fan-out and global warm-up stay closed.
  Preview remains closed. Existing contact flows are unchanged.

## Evidence and rollout

Local checks pass planner/full-term execution (18,000 shifts per source), HTTP
middleware/protocol interruption and retries, receipt rollback, budget deferral,
inactive visibility, current-term preservation, atomic activation, completion
repair, partial-month publication and failed manifest promotion. Contact-cache
checks exercise 36,000 unchanged refreshes with zero D1 use; ingress idempotency,
source isolation, request attribution, client request budgets, rollout guards,
queue-failure safeguards and indexed database-cost fixtures pass. Pages Functions
build succeeds. Synthetic fixtures were used only in local ephemeral databases.

Account-wide admission before release: GO, 27,489 reads / 46 writes. The pending
0032 indexes were already present in Production. Bounded inspection read 694 sync
runs and 34 claims; migrations 0032–0034 applied successfully without backfilling
roster events. New tables are empty on creation.

Safari readback: Sync Monash roster files, Sync Monash Paediatrics roster files,
VHH Active Medical Roster to Production, Sync MMC clinician contacts and Sync DDH
clinician contacts are enabled. Legacy MMC allocations/bootstrap and VHH Preview
flows are disabled.

First Production deployment: 49834951-d550-423a-88f2-09cafed3acbb, commit
874c2f63, canonical deploy stage successful. Bounded live compact checks found
three retained files each for MMC/MCH/DDH and two for VHH. Two older DDH files
lacked coverage/status preparation; VHH's active workbook spans September–January.

Follow-up from real readback: ownership now uses compact stream-catalogue shift
START-date bounds, so overnight attendance cannot claim the next term and disjoint
VHH windows can coexist. No history scan or date guessing. The first maintenance
pass prepares missing exact-source compact records one file at a time (25,000
events, 512 doctors, 750 mutations maximum), reserves indexed costs before work,
and refunds measured unused costs only with complete billing metadata. It does
not rewrite shifts. Coverage preparation is a separate explicit bounded mode;
legacy bootstrap remains closed. Workbook span allows 180 days; chunks/day limits
are unchanged. Future-date publication waits safely for prepared term membership.

Final follow-up deployment/readback and maintenance outcome are recorded in the
final report and subsequent checkpoint. Do not infer real next-term completion from local tests.

Rollback: disable ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED,
FACILITY_AUTOMATIC_PUBLICATION_ENABLED and automatic launch, and cancel scheduled
maintenance. Keep progress tables and valid published objects for recovery;
never delete current terms or re-enable historical whole-workbook contact flows.

Remaining service work: compact colleague tools, bounded manual roster mutations,
Creator switching/directory enrichment, doctor discovery, cross-device settings,
and VHH contact support. All-site next-term completion still depends on provider
publication and may span UTC budget windows.


Admission refinement: the checker's twofold optional-cost reserve rejects proposing
the entire daily allowance as one pass. Every scheduled/manual maintenance pass
now carries a stable ledger baseline and SQL-enforced additional cap of 5,000
indexed writes / 250,000 reads. Daily global caps remain 10,000 / 500,000.
Preparation and publication refunds require complete billing metadata. First-pass
admission with these bounds is GO (31,460 reads / 73 writes before the pass).
Initial current-view seed intent is persisted for all four sources before costly
work, so deferral does not lose the request. Later scheduled passes finish it.


First live pass 36813726504 completed successfully. VHH's current view published;
MMC reached batch 5 before the 5,000-write pass reservation cap. MCH/DDH initial
seed intents remain durable. Reserved costs were 94,826 reads / 4,948 writes;
settled account analytics were still only 43,331 reads / 140 writes.

Reservation refinement from this result: compact correction mutations are planned
against proposed file facts and bounded existing contributions BEFORE active facts
change, including orphaned/stale compact rows. Coverage preparation reserves the
actual planned compact statement count, with all indexes included, before executing
it. Publication refunds unused writes only with complete billing metadata and keeps
16 control-write units plus ledger overhead. Hard pass/daily ceilings are unchanged.
Local tests verify dry planning is read-only, refused reservation writes nothing,
stale-row repair is counted and replay/publication failures preserve active data.
Second-pass admission is GO including unsettled first-pass reservations in the
twofold estimate (297,413 reads / 6,474 writes proposed including carryover).


## Verified wrap-up after two live passes

Production is active at deployment d105d480-91cb-4df7-9d79-7f47a5e0f7a3,
commit 8cf7019237b9f5b47ab09afa4aeffbdacfabafcc. Both maintenance runs
36813726504 and 36815329456 completed successfully.

Compact queue readback: MMC current term complete (13 day batches / four
months); VHH current and next term complete. VHH's February spillover is waiting
for prepared term membership. MCH and DDH current-term seed intents remain
queued. Shared ledger holds 149,850 reads / 9,525 indexed writes reserved for
1 October; no further database work was undertaken at wrap-up. These are
conservative reservations, not account billing totals. Automatic six-hour
continuation preserves progress and resumes when its allowance admits work.
DDH's two old missing compact records remain to be prepared before new-term
activation; current contact automation is unchanged.

Usage checkpoint: 93% weekly and 57% five-hour allowance remained; credit
balance unchanged from the reset checkpoint. No additional credit spend.

Do not describe all four sites as fully refreshed: MMC/VHH publication is
verified, MCH/DDH publication is pending. Provider-triggered future deliveries
and large overlapping replacements still require their own completion evidence.

## Account-aware budget batch prepared — 1 October, awaiting secret approval

Implemented replacement admission at 4,000,000 account reads / 80,000 account
writes, accounting for settled analytics, unresolved legacy reservations,
unsettled measured request receipts and projected account traffic. Grants expire
in ten minutes; workflow drivers refresh after five minutes. Atomic SQL fences
concurrent allocation and stale snapshots. Request cost overruns stop the UTC
budget; analytics failures close admission until a valid refresh. Lost D1
responses retain full reservations rather than being measured as zero.

Local account policy/SQL race/HTTP middleware tests, full-term materialization,
request attribution, source isolation, queue failure, quota guards, indexed
cost fixtures and the Pages build passed. Live pre-release admission is GO:
109,444 settled reads / 1,524 writes; CLI projections 1,974,590 reads /
40,569 writes before the proposed bounded pass. Report is in
`/private/tmp/account-budget-release.json` and is time-specific evidence.

Production migration 0035 was created and recorded successfully; the initial
bookkeeping import used five reads and ten writes. No roster history or live
roster data was changed. Code is NOT yet deployed. Automatic approval review
rejected uploading the existing analytics credential as a persistent Production
Pages secret, requiring explicit user authorization. An approval question is
pending. Do not push/deploy this batch with the Production budget flag enabled
until `ROSTER_ACCOUNT_ANALYTICS_TOKEN` is installed; afterwards verify canonical
deployment, live budget admission, queued MCH/DDH publication and all-site state.
The previously deployed small caps remain active meanwhile.

## Approved budget deployment and cached colleague batch — 1 October

User explicitly approved the protected analytics secret. Installed in Production
and verified present in the canonical deployment. Hosted Pages Wrangler 3 rejected
JSON import attributes; changed the inventory to a plain JS module with a drift
test. Production deployment afaf4a2e-913a-4068-bcf4-a19fc4890634 / 50e40491
includes this compatibility fix and atomic measured-unused-grant refunds.

First admitted pass exposed a bookkeeping envelope mismatch: one read-only
reservation request measured 29 write units including settlement against 24
reserved. Corrected reservation overhead to 48 while retaining 24 conservative
settlement units; no provider quota approached. Cleared only that investigated
pause with a SQL fence excluding larger unexplained overruns. Receipt replay
cannot refund twice, incomplete metadata retains the full reservation, and raw
bounded settlement works even after a route exhausts its statement allowance.
Run 36819943554 resumes all four sources under the new guard. Initial live account
admission: 112,002 reads / 1,539 writes. Subsequent progress had 340,417 reads /
10,198 writes allocated (not provider totals), with no stop reason; MMC/MCH/VHH
current views complete and DDH publication advancing.

Additional batch: replace queryRosterInsights/queryRosterOverlapDoctors event
history joins with published R2 range reads and in-memory same-site/date overlap.
Preserve entitlement, cohort/source controls, term visibility, leave/unknown
filtering and clinical-support options. Requests allow 180-day ranges and at most
64 identities per filter. Missing publication has no D1 fallback. Five-minute
client cache expiry and new cache keys prevent indefinite/legacy stale results;
background insight warm-up remains disabled. Local helper, full SQLite/R2 HTTP,
entitlement/missing-data, client-request and Creator-containment tests pass.

Traffic projection now excludes only confirmed settled measured maintenance
from ordinary daily traffic; no estimated settlement units are subtracted.
Otherwise a legitimate one-off import was incorrectly forecast as repeating all
day. Unknown/legacy costs remain conservative. Local forecast tests pass.
Canonical deployment and visible colleague-tool verification still required for
this additional batch; do not infer completion from the commit alone.

## Successful all-site readback and published-object verification

Maintenance run 36819943554 completed successfully at 15:39 AEST. Latest source
records are successful: MMC 153 doctors / 4,168 events; MCH 80 / 2,372; DDH
149 / 3,650; VHH 47 / 958. The run also processed the queued MCH correction
(2,376 parsed events); later latest-file counts can differ, so do not conflate
those figures. Current-term publication jobs are complete for all four sites;
all errors empty. VHH's next term is published and remains hidden until
19 October. MCH/DDH November spillover and VHH February requests remain pending
term membership; they are not evidence of an imported next-term workbook.
DDH's February, May and August compact coverage/status rows are all ready.

Downloaded the four exact R2 manifests. MMC/MCH/DDH each have all 91 current-term
day pointers and four monthly pointers; VHH has current/next term pointers.
The actual cached colleague reader returned today's published working rows for
all four sites: MMC 46, MCH 27, DDH 43, VHH 15, using 12 R2 object reads and zero
D1 operations. No clinician details were printed or committed. The reader now
uses daily attendance for one-day queries and the range plus first-day facts for
longer queries, preserving overnight overlap without a historical lookback.
Request-scoped object caching pins one manifest version per site. Tests cover
attendance starting before the selected date.

Canonical colleague deployment ca62f18b / 7d45e7c1-5ed0-4755-aeca-5f3d3d89bcc7
was successful. The daily-reader optimisation is verified locally against actual
published objects and requires its final deployment. Safari loaded the logged-in
calendar, but the Mac then locked before the live colleague click. User unlock
question is pending; visual UI confirmation remains outstanding.

Settled account sample (15-minute delayed) returned GO: 160,452 reads / 2,645
writes through 15:31 AEST. Import/publication grant at final maintenance readback:
225,968 reads / 10,534 writes allocated, with no stop reason; these are not provider
billing totals. Existing contact flow flags remain unchanged. Full restoration
still includes VHH contact support, large overlapping replacements, bounded manual
mutations, Creator directory/discovery and dedicated settings persistence.

Final canonical deployment verified: b7380c4e-7506-46b5-9a0c-e16879e0ea1f,
commit 5481698fd185da7f4f9d7b20513c08df67d968aa, deploy stage success,
account-budget flag true and approved analytics secret present. The daily/range
colleague-reader optimisation is now deployed. Latest Codex usage check: 24%
five-hour and 88% weekly allowance remaining; credit balance unchanged.
Mac unlock remains required for visual click verification and the next Microsoft
source-inspection work. Existing contact flows remain enabled and unchanged.

### 1 October — bounded Creator tools batch

Restored explicit Production Creator directory and doctor picker; Preview and automatic startup hydration remain closed. Directory pages use indexed email cursors (100 profiles, at most 1,000 claims); credentials and full sessions are excluded. The browser consumes at most ten pages. Visible staff from at most two terms per site supplies the picker through R2, with no history fallback. Switching account resolution uses exact site/doctor indexed lookups and fails on ambiguous ownership. Unclaimed doctor calendars now read published monthly artifacts, preserving saved overrides/custom events and locations without D1 roster discovery, snapshot writes or background builds. Missing source publications fail closed.

Build, Creator containment/client budgets, source isolation, cached insights and full materialization HTTP/SQLite/R2 tests pass. New tests cover directory pagination/index selection, credential exclusion, authorization, published doctor calendars and indexed account resolution. Both colleague tools were visually verified in live Safari: DDH shift colleagues and future overlapping shifts.

Settled account analytics at 05:58 UTC: 220,772 reads / 10,739 writes. The standalone CLI checker reported STOP using its own forecast; it does not reconcile settled maintenance with ordinary traffic in the same way as the deployed account guard. Live grant has no stop reason (225,968 reserved reads / 10,534 writes against 3,223,512 / 47,614). This batch adds read-only Creator views, not an import run. Dedicated settings API, manual overlapping replacement, automatic identity discovery and VHH contacts remain separate work. Profile calendars expose currently published visible roster terms; historical terms absent from publication are not rebuilt on demand.

Deployment confirmation: commit `98f32019599edf36b7abf7cf3eabaafa1bf2babb` became canonical Production deployment `50d47437-3fdb-4e37-ba52-ce174d8836ac` successfully at 06:14:58 UTC. Account-budget flag and protected analytics secret remained present. Live Safari directory displayed 45 accounts, picker populated, Bob Seith calendar rendered 34 current-term events (MCH shifts and long-service leave), and Back to creator restored Richard's calendar. No real account settings or roster files were changed by this UI check. Five-hour allowance readback before final checks: 10% remaining; credit balance unchanged at 182.1872700000.

## 2 October combined settings and manual-roster batch

- Migrations 0036/0037 add settings revisions and promotion fences; no history backfill.
- Manual and automatic corrections share the pure batch planner, durable staging and account admission. A promotion compares the pinned active file identities/revisions inside its atomic transaction. A failed/partial replacement preserves the old active files; disjoint future terms remain active. Retired roster data remains recoverable.
- Removal deactivates up to four files, queues bounded view publication and preserves events/presence. Existing automatic source flows remain enabled and can subsequently deliver a replacement.
- Settings saves touch exact account/profile rows plus at most 40 changed custom events and five site locations. Independent field changes merge; conflicting edits/replays are tested. Settings never trigger roster import, history rebuild or identity discovery.
- Current personal shifts follow committed R2 publications, preserving cached historical shifts outside visible terms. Missing/incomplete publication retains the working cache until finalization.
- Local SQLite/HTTP tests cover 1,600-event replacement, interruption, transactional rollback, replay, concurrent promotion races, partial-window rejection, future-term preservation, soft removal, settings conflicts, ownership and feature gates. Materialization, account-budget, client-request, Creator containment and source-isolation regressions passed; Worker compiled.
- Account-wide settled check at 01:22 AEST: 304,036 reads / 14,094 writes for the UTC day; 80,427 reads / 3,350 writes since yesterday's release. Five-hour Codex usage 19% used; additional credit balance unchanged.

### Release verification and VHH discovery

Canonical Production commit `d6ad1ddf4252bc055868d972decfcc0b529f5ca8`, deployment `5ee7ec92-624b-4ae8-8278-1d0607c79338`, succeeded 16:26:24 UTC on 1 October (02:26 AEST on 2 October). Migrations applied to the verified Production database UUID. Safari showed Richard's 45-event current-term calendar and existing settings. Saving unchanged settings left the account revision and timestamp unchanged (expected no-op). HTTP fixtures additionally verify personal account calendars use published shifts and active roster references without querying roster history.

The final empty-active-set publication test exposed an older compact filtering edge case: an empty active-file set previously admitted inactive contributions. Corrected staff filtering for explicitly supplied empty sets and catalog filtering for all bounded calls; empty-roster R2 publication passes. No live roster was removed or replaced for these tests.

Read-only VHH review: existing flow `e8b16ab2-978a-44b1-b331-070f17504b14` remains Off, unchanged. Its 10 September 18:44 run successfully executed the Office Script; HTTP failed `NotFound`. The bounded result has schemaVersion 1, sourceId `vhh-shift-phone-allocations`, a CIC entry, and nine doctor handset rows (role/name/phone). It contains no source date and most rows have no period; consultant roles carry start times. Workbook and script identifiers agree with the retained flow. Before enabling Production, establish whether it represents current handset holders or whole-day allocations, add trustworthy source freshness/version metadata, and validate matching against VHH On shift. User clarification is pending. No VHH contact data has been sent to Production.

## 2 October — VHH current handset restoration

Morning account analytics remains low (8,468 reads / 688 writes through 10:58
AEST; final 11:00 session minutes awaited telemetry settlement). Application
changes add VHH to the existing doctors-only JSON ingress and R2 contact reader;
no roster reconstruction, migrations or new D1 reader queries are introduced.

VHH uses a single changing list, not AM/PM/Night contact blocks. Only named,
validated clinical rows from Zebra Allocations are eligible. Their civil sheet
date must be current and their uniquely matched VHH roster event must have
explicit hours and be active now. Previous-date Night events can match after
midnight until their actual end; blank, conflicting, ambiguous, off-shift and
unknown-hour entries remain unassigned. Unmatched VHH entries do not expose a
manual override that could bypass this check. Open views recompute VHH matches
at every bounded contact refresh even if the JSON revision is unchanged.

The existing Office Script reads C1:E80. Its original extraction was preserved
and only sourceDate: cell(2, 3) was added to the return envelope. The application
normalizes the formatted date. A replacement draft was discarded after editor
accessibility output caused automatic save review to reject it; the minimal
one-field edit was verified and saved successfully. No worksheet content was
edited.

Automatic architecture: replace the disabled VHH workbook-modified trigger
with a five-minute recurrence, preserving the exact workbook/script and JSON
HTTP action. This caps extraction/submission at 288 runs per day and avoids
Excel reopening its own SharePoint change trigger. An unchanged submission uses
one indexed deduplication lookup and writes zero D1 rows/R2 objects. Reader
refreshes use only bounded R2 objects. HTTP retries must remain disabled. The
existing flow remains disabled until Production and a fresh-source canary pass.

Validation: VHH timed and ambiguous-match fixtures, cross-midnight/end boundary,
UTC timestamps, stale/future sheet dates, duplicate phones, unknown hours,
SQLite-backed VHH HTTP publication and zero-write replay, existing contact
access/ingress safeguards, facility materialization, client request budgets and
Worker compilation passed. Live flow/deployment verification pending.

### Live checkpoint at 11:25 AEST

Commit `5722ecdc` pushed through the normal Git-triggered deployment. The public
Production app.js SHA-256 matches the local release, including VHH polling and
expiry checks. Canonical deployment/control-plane variables still need readback:
Wrangler's existing OAuth login expired and renewal reached the normal 26-scope
consent screen, then the Mac locked before authorization could finish.

Settled account telemetry through 11:06 AEST covers the user's session ending
11:00: UTC quota-day totals 8,912 reads / 695 writes; maximum five-minute bucket
3,224 reads / 666 writes; largest sampled average query 468 reads. No read quota
pressure appeared. These are account totals, not attribution solely to the user.

The existing VHH script's minimal date addition is saved. The Power Automate
flow remains Off, with an unsaved draft removing the SharePoint modification
trigger and the Add a trigger search set to Recurrence. Excel and HTTP actions
are retained. The rename attempt did not apply: the displayed flow title is
still VHH Shift Phone Allocations to Preview. No fresh live submission or visual
VHH contact check has run yet. Unlock Safari, complete recurrence settings
(five minutes, one concurrent run), production JSON URL and disabled retries;
confirm canonical flags; save/enable and verify fresh-source then unchanged runs.
Additional contradictory-surname and implausibly-long-shift fixtures pass locally
and await the final verification/checkpoint commit with the completed flow work.

### VHH automatic restoration completed, 2 October 11:44 AEST

Canonical Production deployment `67597134-138d-442c-91a2-55d5d1627bce`
confirmed successful for commit `5722ecdc30b60bb48dff4947643f9a33d6d67c6b`.
Both contact allowlists include VHH. Existing Wrangler authorization refreshed.

Existing flow `e8b16ab2-978a-44b1-b331-070f17504b14` renamed **Sync VHH
clinician contacts**, saved as Scheduled and verified On. Recurrence interval
5 minutes, concurrency enabled at 1, HTTP retry policy None. Workbook and
Office Script references preserved; only C1:E80 is extracted. HTTP sends the
script result directly to the Production contact-list-extract endpoint.
The old Preview credential produced two initial 401 failures; replaced only
that flow header with the existing working MMC Production contact credential.
No new credentials or broader endpoint authentication were introduced.

Fresh run at 11:36 succeeded and published today's three named worksheet
entries. Unattended recurrence run at 11:41 succeeded with HTTP 200,
`status: unchanged`, sourceDate `2026-10-02`, contactCount 3 and the same fileId.
The R2 contact manifest remained byte-for-byte identical after that run,
including receivedAt/publishedAt. The tested unchanged ingestion path performs
one indexed deduplication lookup and zero D1/R2 writes.

Bounded live verification read five R2 objects and no D1 data. Today's published
roster contained 15 timed rows; only Asare Amoafo (08:00–17:30) matched an active
handset allocation at 11:39. Daryl and Nidhal remained unmatched. Production
At a glance / VHH visibly confirmed updated 11:36, one matched, Asare 12018,
and two entries needing review, without handset assignment to inactive names.
All three names are synchronized; eligibility to display a handset is narrower
than presence on the mutable contact sheet.

Settled telemetry through 11:17 AEST: quota-day 9,901 reads / 700 writes,
maximum five-minute 3,224 reads / 666 writes; largest sampled average query
468 reads. This confirms the user's morning session remained well below the
5,000,000-row read quota, but telemetry lags and does not yet cover the 11:36
first VHH publication. Live unchanged response and stable R2 revision verified
its replay behavior independently. Extra surname and >24-hour safety fixtures
passed. Five-hour usage last checked at 34% used (66% remaining), credits
182.1872700000 unchanged.

Final settled check through 11:29 AEST: 10,592 reads / 700 writes today,
unchanged maximum five-minute and largest-query figures. Final five-hour
allowance 53% remaining; credit balance still unchanged. UI flow restoration
and live verification are complete; no further source edits are required.

### All-site five-minute contact schedules, 2 October 12:14 AEST

User requested contact source syncing no more frequently than every five
minutes. Existing MMC/MCH shared contact flow `64d9fad7-3462-4b6a-bb6e-95ce1951f4fb`
and DDH flow `67ee065a-c230-4d1d-9a20-d9425f9cacdf` were converted from
SharePoint change triggers to Recurrence interval 5 / frequency Minute.
Existing source destinations and credentials preserved; obsolete event-version
conditions and trigger-version headers removed. No additional contact flows
were created. Both flows remain On and Scheduled. VHH retains its verified
five-minute schedule. Cached app contact refresh remains an R2-only read;
this change governs source extraction/submission.

MMC keeps the exact Office Script/workbook reference. Removed old delay and
trigger-only file metadata dependency. The live Office Script differs from the
repository copy: it returns sourceDate and contacts, but omits sourceId. A
controlled run initially returned HTTP 400 Invalid doctor contact extract;
the request envelope now uses:
`setProperty(outputs('Run_script_from_SharePoint_library')?['body/result'], 'sourceId', 'mmc-shift-allocations')`.
An unattended scheduled run at 12:13 succeeded after this correction. No broad
workbook download was introduced for MMC/MCH.

DDH retains its proven small-workbook base64 transport and bounded clinician
parser. Get file metadata is bound to its verified existing identifier
`Shared%2bDocuments%252fGeneral%252fDaily%2bContact%2bSheet.xlsx`;
Get file content still uses the metadata Id; providerModifiedAt remains the
metadata LastModified, preserving contact operational-date semantics. The old
two-minute debounce delay and edit-version gate were removed. Automatic runs
at 12:00, 12:05 and 12:10 all succeeded.

Trigger concurrency initially produced Power Automate throttling/skipped-run
warnings. Automatic approval review rejected restarting MMC while that warning
remained. The configuration was corrected before retry: concurrency control
Off, MMC script timeout PT2M and HTTP timeout PT1M; DDH metadata/content/HTTP
each timeout PT1M. Retries None for every processing action. Total configured
processing timeout stays below the five-minute interval; the current warning
cleared, MMC activation was approved, and its scheduled delivery succeeded.
No unresolved approval blocker remains.

Contact sync safeguard fixtures passed. Latest settled account telemetry
through 11:46 AEST: 11,753 reads / 705 writes today, maximum five-minute bucket
3,224 reads / 666 writes. These lagged totals cover the VHH first publication
and replay; they do not yet cover the new MMC/DDH schedules. All contact paths
retain semantic deduplication and unchanged submissions perform zero D1/R2
writes. Usage last checked at 90% used / 10% remaining; credits unchanged at
182.1872700000. Stop additional restoration batches after final verification
and checkpoint to conserve the remaining allowance.


### 2 October — bounded identity and directory enrichment batch

Prepared together: automatic name suggestions, explicit account linking/removal,
Creator claim correction and visible-term directory grades. Discovery reads the
same bounded R2 term staff publications as the Creator picker. There is no
canonical-directory/history fallback, automatic account repair, identity seeding,
login claim acquisition or global canonical rebuild when discovery is enabled.
Users confirm suggestions through the existing Confirm action. Exact site/key
ownership and the prior account claim set are asserted inside the atomic D1
mutation batch. Concurrent ownership attempts fail with HTTP 409; replay writes
nothing. Claim changes preserve settings, custom events, locations and unrelated
profile fields and schedule no snapshot work.

Account preparation also replaces full profile listing and historical file
membership discovery with indexed bounded lookups: 16 claims, 32 matching
profiles, and at most 32 active files per claimed source. Existing historical
claims survive absent publications; missing/oversized identity publications
produce an unavailable state without breaking existing links. Grades come from
current-term published membership/overrides, preferring the current term over an
already-visible next term. Directory pages retain their 100-account/1,000-claim
limits and add no per-account D1 enrichment query. No migration or backfill.
Production enables IDENTITY_DISCOVERY_ENABLED; Preview remains disabled. Global
account snapshot building and Creator startup hydration remain disabled.

Local evidence: test-bounded-identity exercises 100,001 historical events and
100,001 historical file memberships, plus 1,000 unrelated profiles/accounts;
read-only discovery, ambiguous suggestions, missing data, current/future grades,
indexed query plans, authenticated gates, duplicate ownership, transactional
races, failed replacement rollback, replay and removal pass. Existing restoration
mutations, Creator startup, client request budgets, account budget, colleague
insights, facility materialization, source isolation and rollout suites pass.
Worker compilation passes. The legacy monolithic test-fixtures suite stops at
its pre-existing outdated automated-correction source assertion (line 922) in
unchanged derived-import code; this suite is not reported as passing.

Live pre-release R2 verification: eight object reads, zero D1, no missing sites;
MMC 152 doctors/152 graded, MCH 79/79, DDH 169/162 and VHH 50/50. Settled account
analytics through 04:26 UTC: 43,553 reads / 16,601 writes; maximum five-minute
reads 12,938 and writes 15,867. Import statements are present in the write burst.
The exact current-day admission record read cost one D1 row, zero writes; its
prior grant had expired, so additional maintenance needs a fresh valid grant.
Release/deployment and authenticated live UI verification remain pending.


#### Delivery failure discovered and corrected in the same release batch

Exact indexed readback (16 D1 rows read, zero writes) found recent MMC/MCH runs
still queued, while VHH completed a real SharePoint-triggered bounded import at
04:13 UTC (47 doctors, 958 events). GitHub logs for the latest MMC/MCH jobs and
the six-hour continuation identify `Staged batch differs from the pinned plan`.
Both checked-in real Monash workbook fixtures reproduce this: parsed events have
`sources: undefined`, which the old canonical digest included and JSON transport
omitted. The manifest itself survived transport; every first fact batch failed.

Canonical hashing now follows JSON omission/null semantics. Only an initialized,
empty, inactive automatic job with identical source/content, doctors, ranges and
batch counts may replace its old batch hashes. CAS assertions require no facts,
receipts, preparation or activation and preserve the original promotion fence.
Transactions also assert their pinned revision/cursor before staging facts or
presence, so a concurrent replan cannot write stale batches. Manual requests
cannot perform this recovery. No migration or historical reprocessing.

Same-term automatic corrections now first use the existing 1,250-fact bounded
diff path instead of copying the entire term. If the actual diff exceeds that
ceiling, the server retains the queued job and returns the precise budget code;
the processor then uses full inactive staging under account admission. New terms
and partially staged jobs continue through bounded staging. This avoids spending
the daily write allowance on complete term copies for small changes.

`test-roster-wire-protocol` verifies real MMC (4,295 events) and MCH (2,363 events)
through JSON serialization, rejection before writes, safe legacy empty-job
recovery, refusal of changed identities/partial jobs, transaction races, replay,
full staging/activation, same-term small diffs and safe large-diff fallback.
Production deployment and exact queued-source retries remain pending.
Control-plane inventory is complete: 103 Production deployments, zero Preview;
cleanup will retain the accepted current release and at most one verified rollback.


### Combined identity/import release and delivery verification — 2 October

Production `239d215a1a0b8fba70d7071e8f87aaa610b15e6a` is canonical in
`fe6de1d2-c0b6-4a8b-85de-65664bac46db`, deploy success at 05:19 UTC.
Readback verifies identity discovery true, session settings/manual bounded
imports/replacements true, account budget true with the analytics secret present,
legacy manual writes false and global snapshot builder false. Both restoration
commits were pushed together through the existing Git deployment pipeline.

Controlled retries succeeded on the deployed code:
- MMC workflow `36968437530`: 153 doctors / 4,167 shifts imported at 05:20 UTC;
  current-term R2 manifest published at 05:21:27. Same-term correction retained
  active file `automation:monash-adults:47c0951dd2582465d59b19a9`.
- MCH workflow `36968682143`: 82 doctors / 2,367 shifts imported at 05:23 UTC;
  current-term R2 manifest published at 05:25:32. Active file retained as
  `automation:monash-paeds:16c4bd4d0ac403ef6f371e2e`.
- VHH real SharePoint workflow `36963520835` succeeded before this release;
  current and next-term objects remain published, next visible from 19 October.
- All four exact source rows remain enabled and have empty last_error. DDH's
  last successful import is 1 October 13:10 UTC; one non-seeding all-site
  maintenance verification `36969058680` is running to check continuation and
  its routine LastModified detection.

Post-delivery directory check used eight exact R2 objects, zero D1: no missing
sources/preparing state; published MMC 152 graded, MCH 81 graded, DDH 169 with
162 graded, VHH 50 graded. Published term staff counts need not equal all parsed
workbook names. MMC/MCH/DDH currently publish 3 August–1 November; no real next
term has yet been verified for those sources. This is a provider-delivery
acceptance dependency, not a reason to run synthetic Production imports.

Added SQLite-backed subscription acceptance to restoration-mutations: an active
future-term correction appears immediately in the ICS feed without snapshot
warm-up; prior/inactive files are excluded; subscription reads write zero and
query plans use indexes. Local test passes. Real subscribed-client refresh still
needs user acceptance; client caching is separate from server feed freshness.

Account-wide checker at 05:23 UTC returned GO, settled totals 48,055 reads /
16,631 writes. Analytics lags recent imports. The exact live reservation row
read at that time has no stop reason, allocated reads 181,164 and writes 18,665
against maxima 3,837,273 and 79,653. Readback of four source rows and this row
cost nine D1 reads, zero writes. No quota thresholds or per-chunk limits raised.

Retained deployment inventory: 103 prior Production, zero Preview; new release
makes 104. Prepared cleanup deletes 102, preserves this release and verified
rollback `2140de5d-9811-4a46-8e17-dc63d4d5b97e`. Automatic approval review rejected
execution because the exact mass-deletion scope needs explicit user approval;
no deployments deleted. Approval requested. Safari authenticated verification
remains blocked by the locked Mac. Medium implementation is complete; user has
been advised that remaining verification/cleanup is suitable for Low.


#### Provider revision ordering defect found during final continuation

All-site maintenance `36969058680` completed successfully and imported DDH's
changed source (149 doctors / 3,650 shifts). However, exact MMC latest-four-run
readback found version 398 imported at 05:20, then queued version 397 at 05:28
and 396 at 05:35. Counts agree but content hashes differ. Queue selection only
excluded newer queued/processing entries, not newer successful ones; successful
completion had remapped file_id but retained immutable source_file_id correctly.
This is a confirmed same-term stale-overwrite path, distinct from term retention.

The follow-up batch excludes newer successful deliveries during polling and
source dispatch. Every automatic mutation checks at most 65 recent source runs
and 32 active source files, comparing immutable raw-file provider timestamps for
the same filename. Late older arrivals are skipped; other named term workbooks
remain independent. Unprovable old work fails closed without a history scan.
Exact callbacks mark obsolete work superseded under budget admission. Atomic
assertions repeat the ordering check before small correction writes and staged
activation, rolling back if a successor arrived after preflight. Local SQLite
coverage includes 100,001 historical runs, late arrival, different term names,
no obsolete dispatch, unchanged active facts, callback skip and transactional
races. Wire, source-isolation, queue-failure and mutation/subscription tests pass.

Deploy this correction, then requeue only the retained real MMC version 398
using guarded exact-row recovery and verify its import/publication and provider
last_modified. No synthetic Production workbook or broader roster repair. D1
sample at 05:35 UTC was GO, 74,792 reads / 17,026 writes through 05:20 UTC;
additional credits remain unchanged, five-hour allowance approximately 49%.


### Ordering protection deployed and MMC latest revision recovered

Production `6513315a1591cf5e9ce73be46de4f74afc68c525` is canonical in
`bd4924c6-05e6-4299-bc40-929ec8535954`, deploy success 05:52 UTC. All restored
feature controls and account admission remain enabled; global builders and legacy
mutations remain closed. Fresh guard admission has no stop reason. The standalone
short-interval checker at 05:51 projected the recent maintenance burst as ordinary
traffic and returned STOP despite measured totals 161,154 reads / 18,866 writes.
The deployed admission reconciles those maintenance receipts; at recovery it
allowed allocations 268,935 reads / 20,776 writes against its admitted maxima.
No costly recovery bypassed the account guard.

Exact guarded recovery requeued only the real retained version-398 run with its
hash/timestamp and no newer delivery. Control update cost 743 reads / four
indexed writes. Workflow `36971041074` succeeded: 153 doctors / 4,167 shifts,
completed 05:55:02 UTC. Final PK readbacks cost two reads, zero writes: active
MMC file last_modified is 1790900725000, exactly version 398; run status success.
All-site latest-two-run and active-file verification cost 34 reads, zero writes;
MCH, DDH and VHH active timestamps match their latest checked provider revisions.
Obsolete queued versions are excluded from queue polling and dispatch. No
historical/whole-term reimport was used for this recovery.

Updated control-plane inventory: 105 Production deployments, zero Preview.
Concrete revised cleanup list contains 103 deletions and retains this current
release plus verified rollback `fe6de1d2-c0b6-4a8b-85de-65664bac46db`.
No deletion occurred. The earlier 102-deletion request was rejected by automatic
approval review; explicit approval is needed for the updated scope. The Mac
remains locked. Authenticated identity acceptance and cleanup are pending;
Medium implementation is complete again and remaining work is suitable for Low.
Latest five-hour allowance 42% remaining, additional credits unchanged at
182.1872700000. Untracked user work has not been modified.


Final post-burst assessment at 06:01:50 UTC returned GO with no reasons:
162,287 reads / 18,866 writes through 05:46:50 UTC; projections 276,340 reads /
18,866 writes. The settled window does not yet include the 05:55 recovery;
its live reconciled admission and indexed completion proof are recorded above.
One exact R2 read verifies MMC's recovery publication at 05:55:27 UTC, 91 current
term dates. No further import/maintenance job was dispatched. Final five-hour
allowance 40% remaining, credits 182.1872700000 unchanged. Remaining work needs
Mac unlock, effort setting change to Low, user functional acceptance and explicit
approval for the prepared 103-deployment cleanup. Final records are committed
locally without creating another documentation-only Production deployment.


## Approved deployment cleanup and live acceptance — 2 October, Low effort

Deleted all 103 approved superseded deployments through the Pages control plane,
with zero failures. Independent full inventory confirms exactly two Production
deployments and zero Preview: canonical `bd4924c6-05e6-4299-bc40-929ec8535954`
(commit `6513315a`) and verified rollback `fe6de1d2-c0b6-4a8b-85de-65664bac46db`
(commit `239d215a`). Cleanup did not call application/D1 endpoints.

Unlocked Safari acceptance loaded the 45-account Creator directory. Selecting
SMS filtered it to 14 accounts, confirming published seniority lookup in the
live UI. Owner account locations remained present; the account screen displayed
no explicit linked roster names, while the selected doctor profile retained its
45-event calendar. This does not verify an intended ordinary-account linking
mutation; that remains an acceptance item. No claims, permissions, credentials
or preferences were changed. Safari was returned to the personal calendar.

At 06:09 UTC analytics showed 163,498 reads and 18,870 writes, settled through
05:54 UTC. The checker stopped solely because the prior sample was less than
ten minutes earlier; no expensive fingerprints were detected. This sample does
not yet cover the 05:55 recovery. Final properly spaced sample follows below.

Documentation is retained locally to avoid creating another deployment solely
for operational records. Provider future-term delivery, subscribed-client refresh
and remaining genuine user journeys remain external/live acceptance work.

Final sample at 06:12:22 UTC: GO, no reasons, settled through 05:57:22 UTC
(includes the MMC recovery). Account usage 191,377 reads / 18,967 writes;
recent-rate projections 3,138,581 reads / 29,200 writes remain below admission
thresholds. The extrapolation includes maintenance, not solely ordinary UI
traffic. Latest UI checks remain within the analytics lag. Five-hour allowance
30% remaining; credits unchanged at 182.1872700000. Stop here to conserve usage.
