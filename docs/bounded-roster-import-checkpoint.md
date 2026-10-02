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
