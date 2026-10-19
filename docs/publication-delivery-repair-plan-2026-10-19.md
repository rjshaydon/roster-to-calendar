# Roster publication and delivery repair — resume on Medium

Status: implementation tested; production acceptance pending below.
The original planning turn changed no application code, deployment, flow setting or retry. Resume only when the user returns after
the usage reset. Do not create an automation or wake-up for this pause.

## Required outcome

An imported roster must reach On shift, personal calendars and subscription
artifacts without another import consuming its delivery allowance or displacing
its publication operation. Resume saved work rather than re-importing unchanged
rosters. Keep Cloudflare Free, the account-wide read/write safeguards, fast cached
UI reads, unchanged historical grades and existing source/version ordering.

## Verified starting point

- App/runner code: `24638644`; production deployment: `848dfafb`.
  Documentation checkpoint: `b6431066`.
- MMC Term 4 provider revision 19.0 imported successfully at
  2026-10-19T05:47:21.643Z: 132 doctors, 4,075 events. Reused the inactive
  4,078-event baseline with 33 changed facts and 18 small completion batches.
- Latest read-only check during this planning turn: MMC Term 4
  `facility_refresh_jobs` is **complete**, with 13 daily batches and 3 monthly
  batches; `facility_day_publications` is complete and owned by that Term 4
  operation. This changed during the automatic processing, not through a new
  action in this planning turn. User-visible personal delivery remains unverified.
- A 2027-02-01 publication job remains pending with zero batches. Investigate
  whether this is an overnight boundary rather than a genuine next-term roster.
- Last inspected account snapshot: 79,118 settled writes at cutoff 05:56:52Z,
  4,804 admitted writes remaining, no cost-overrun stop. These are historical
  observations, not permission to reuse that grant after resuming.
- Manual retry `39207875357` failed with HTTP 409 from `facility-refresh`.
  Its response reason was lost by the runner. Later automatic run
  `39208006998` was progressing when inspected. Do not claim the exact 409 reason
  was proved; reproduce it with diagnostic preservation below.
- No personal roster-row export is authorized. A previous local JSON export
  was rejected. Use existing normal authenticated interfaces and aggregate or
  in-place verification; do not work around that rejection.

## Confirmed implementation defects

1. `scripts/process-roster-queue.mjs` imports queued files before doing publication.
   After its 24-step publication limit, the next invocation can import another
   term first. That happened: the continuation imported AdultTerm3.2026.xlsx
   while Term 4 delivery was unfinished.
2. Publication and imports share a grant with no protected publication completion
   allowance. Dispatch/pending also use the import-oriented 2,048-write minimum,
   even though a publication request's base reservation is 64 + 48 accounting
   writes. The latter is a per-request base, not a proven total delivery cost.
3. Publication plans include the entire source manifest's revision and ETag.
   Another term's publication therefore changes their identity. The publication
   ownership row is also keyed by source, allowing another term to displace an
   unfinished operation. Current stale-plan handling resets all cursors, while
   the runner throws on HTTP 409 and discards `reason`.
4. The early already-published path in `runFacilityPublicationStep` marks the
   operation complete without enforcing `requireDoctorCalendars`. The deeper
   finalize path does enforce that gate. A lost response between the source
   manifest write and doctor-source promotion can consequently report completion
   without proving personal-calendar delivery.
5. The runner/lifecycle uses a generic account-allowance deferral message for
   distinct causes, including its step cap and remaining publication work.

## Implementation order — one coordinated repair batch

### 1. Verify actual delivery once before mutating anything

Refresh aggregate production status and account admission. Check whether the
latest automatic process has finished personal publication. Through the normal
authorized interfaces, prove the MMC facility manifest and doctor-source pointer
agree, required month completion indexes exist, and the owner's personal calendar
contains MMC Term 4 shifts matching the retained roster. Check On shift on 2 Nov
and the subscribed calendar output. Use authenticated Safari for the visible
check; the Mac may need unlocking. If delivery already works, record that and do
not rebuild it. Continue fixing recurrence defects regardless.

### 2. Finish publication before starting another import

Extract a reusable bounded publication driver. At startup, drain/resume pending
publication before downloading or parsing a queued workbook. After an import,
finish its publication before admitting another import. If the run yields,
continuation must resume publication rather than import first. Preserve the
existing per-source GitHub concurrency and durable cursors; do not raise CPU
limits, start parallel workers for the same source, or disable flows.

Add a publication-only processor mode for recovery and acceptance. It must not
download a workbook or execute import phases. It remains authenticated and
subject to normal measured D1 admission.

### 3. Protect delivery allowance across sources

Give unfinished publication priority in the shared account reservation system,
not only within one site's runner. Before admitting bulk import writes, protect
the bounded outstanding delivery requirement across sites. Derive the requirement
from the remaining plan steps and reviewed per-phase costs, including accounting,
finalization and checkpoint writes; do not introduce another arbitrary 20,000
buffer. R2 artifact work must remain outside the D1 write estimate unless it
actually writes D1.

Use durable, uniquely keyed per-operation holds and atomic SQL admission. Bulk
work cannot consume held publication capacity. Convert a hold to a metered step
reservation atomically, without double charging; release unused capacity only
after verified settlement/completion. Preserve holds through grant refresh and
midnight, and handle unknown/lost responses conservatively. A superseded job may
release its hold only after ownership and termination are proved. Publication
dispatch must use its own next-step affordability test, not the 2,048-write
import-launch floor. Insufficient capacity retains the job and does no work.

### 4. Make publication resume and conflicts safe

Preserve source-wide ownership of an unfinished publication across invocations.
A different term waits rather than taking that ownership. Preserve `reason`,
phase, source, term and operation in structured server/runner diagnostics, so an
expected stale plan can be distinguished from a malformed request or real error.

Separate immutable term input identity from the mutable shared-manifest base.
For an unrelated term's manifest change, reuse verified daily results and rebuild
only the candidate merge and any affected shared-month objects from cached day
pointers. For example, Term 3 and Term 4 share November: preserve the newest
1 November data while merging Term 4's days. Do not simply remove the ETag check
or merge old monthly objects blindly. Retain compare-and-swap promotion and
check relevant clinical inputs before activation. Actual changed inputs invalidate
the affected work; stale workers cannot publish over a newer roster.

Handle a recoverable 409 with one bounded reconciliation/rebase, then resume its
cursor. Do not blindly retry it, discard all staged work, or mark it successful.

### 5. Make completion mean delivery

Unify all finalize/retry paths. Even when the facility manifest already points
at the operation, require complete monthly indexes and the matching prepared and
promoted doctor-source generation before marking the job complete. If the source
manifest was written but doctor promotion was interrupted, repair that promotion
from its pinned artifact; do not rebuild the roster or silently bypass the gate.
Keep the last consistent published generation readable during recovery.

Resolve the 1 February boundary job explicitly. If it exists solely because a
31 January night shift ends the next day, preserve carry-in in views and personal
events without declaring a full new-term roster. Complete/cancel only a proved
unnecessary job; a real future term must remain eligible. Clear delayed warnings
only when the applicable delivery is proved complete.

Distinguish budget deferral, run-step/time yield, stale-plan recovery and actual
failure in existing admin diagnostics and workflow results. Add no routine UI
messages or normal-operation clutter.

## Required regression tests before deployment

Use synthetic SQLite/R2 fixtures and actual HTTP handlers plus the runner driver,
not only source-text assertions. Extend the existing facility materialization,
facility maintenance, doctor artifact, maintenance-account-budget and delivery
reliability suites.

- Reproduce the incident: Term 4 pauses after daily batches/one month, Term 3
  update arrives, allowance is low, continuation finishes Term 4 delivery first.
- Simultaneous site processors: import reservations cannot spend delivery holds;
  concurrent hold conversion, grant refresh, midnight, unknown receipts, duplicate
  callbacks and releases cannot create credit or reserve twice.
- Step-cap/restart, lost HTTP response, process failure and replay at each cursor
  leave resumable work and do not rewrite unchanged roster facts or artifacts.
- Cross-term manifest change/ownership conflict preserves both terms and their
  shared month, without rerunning unaffected day queries. A genuinely changed
  term invalidates the right results. Stale CAS/worker cannot overwrite newer work.
- Crash after facility manifest promotion but before doctor promotion: the next
  call must not return complete until personal generation/indexes match.
- Overnight 31 Jan/1 Feb and term visibility boundaries; no phantom full next term
  and no lost overnight shift. Historical grades and source isolation unchanged.
- Unchanged UI/calendar reads keep their existing cache path and add zero writes.
  Record total metered D1 costs of the whole delivery journey, not just import.

Run the targeted suites once, then relevant source isolation, cached calendar,
subscription and startup checks. No paid services, large live test imports or
production roster exports. Broaden tests only if a failure/change warrants it.

## Deploy and acceptance in one continuous session

Push compatible runner/server changes together, deploy, and reconcile existing
pending publication without re-uploading or re-importing unchanged work. Use
publication-only mode, a fresh measured grant, and existing credentials. Do not
relax the account cap or bypass a stop. Check aggregate costs after completion.

Acceptance requires the actual MMC Term 4 to appear in the owner's authenticated
calendar and subscription output, plus On shift for 2 November, with matching
facility/doctor generations. Also prove a harmless synthetic restart/current-term
arrival cannot repeat this starvation. No live roster edit merely to create a
test. Report remaining uncertainty explicitly if the user-visible check is blocked.

No other restoration batch starts until this acceptance is satisfied. Medium is
appropriate for implementing this plan. Batch routine edits/tests together and
stop only for a real safety blocker, required user test or the user's usage limit.


## Implementation — 19 October, after usage reset

The user confirmed MMC Term 4 was visible by 18:10. Before this deployment,
Safari's authenticated personal calendar visibly contained MMC shifts on 4, 5,
6, 16, 18 and 20 November and later dates. Production tracking marked Term 4
complete; a carry-in boundary job at 2027-02-01 remained pending.

Changes implement publication-first continuation, an explicit publication-only
runner mode, source-owner priority across term jobs, protected delivery capacity,
cache-preserving base reconciliation, mandatory personal-generation promotion on
all completion paths, and distinct run-limit/conflict/budget diagnostics.

`0054_publication_delivery_holds.sql` is the only new migration. Holds are unique
by automatic source/term, owned by the run/request revision, and are not actual
usage. Staging protects the bounded write tail and the next read step; once its
publication plan is pinned, the hold tracks the complete remaining step envelope.
A step retains its capacity commitment until its durable cursor/remaining-tail
checkpoint commits. In-flight receipt and capacity overlap conservatively until
that checkpoint; unknown/lost responses cannot erase future capacity. Settlement
refunds actual unused reservations once, never the hold as if it were usage.
Publication can advance against its own admitted step even if aggregate future
commitments exceed the current grant; bulk/optional work cannot spend those
commitments. The account caps, forecast, stale-grant rules and unknown receipts
remain intact. Staging holds transfer to publication at activation; ended owners
no longer count, and releases cannot remove a successor's commitment.

Reconciliation gives the candidate a new manifest-base namespace, copying only
verified immutable day-batch results. Affected month objects are assembled again
from cached day pointers, preserving updated days from other terms. No clinical
event re-read is needed for an unrelated manifest-base change. Clinical changes
still invalidate the appropriate plan. Conditional promotion and source-owner
fences remain required.

A verified carry-in-only boundary completes without deleting or changing its
night event, fabricating next-term coverage, or rebuilding published terms.
No frontend assets or normal contact/calendar refresh behavior changed. The
existing access-policy test fixture was updated to provide its already-required
file/doctor index and distinguish its bounded access-evidence probe from a
forbidden cached-view roster scan; this does not change production access code.

Local regression coverage passed actual HTTP handlers and the real monthly and
personal artifact builders, interrupted checkpoint and post-manifest crashes,
source priority, cross-term November merge, preserved grades, held capacity,
exactly-once refunds, unknown metadata and UTC carryover, run limits and raw-file
exclusion. Existing real MMC/MCH transport, staging reuse, mutation, source
isolation, facility, subscription, doctor-artifact, cached-insight, cold-start and
syntax checks also passed. Local SQLite metadata is explicitly not claimed to be
native Cloudflare billing or CPU measurement.

Production deployment and aggregate read-only acceptance will be recorded after
completion. No source workbook edits or production roster exports are needed.


## Production acceptance — 19 October, 22:00 AEDT

Applied migration 0054 and deployed code through `3064b59e` to production
(`https://2e047630.roster-to-calendar.pages.dev`, custom domain unchanged).
Publication-only workflow run `39241418324` completed successfully, using the
normal existing credentials and no roster download, parse or import. Its
read-only delivery result was:

```json
{"sourceType":"mmc","term":"2026-11-02","ready":true,"matchingParent":true,"termPresent":true,"monthCount":12}
```

All five tracked MMC term jobs are complete with no last error, including the
2027-02-01 carry-in-only boundary. The final aggregate D1 check read seven rows
and wrote zero. The budget snapshot showed settled account writes 81,096,
usable writes 13,740, and no stop reason; these are the safeguard accounting
snapshot, not a guarantee against future account-wide activity or analytics lag.

After deployment, refreshing the authenticated Safari personal calendar again
visibly showed MMC shifts throughout November and December alongside DDH shifts.
The Mac then became unavailable (`noWindowsAvailable`) before the On shift
2 November visual check could be completed. A live private subscription response
was not downloaded: subscription artifacts/invalidation were covered by the
local regression suite, while production personal-generation readiness was
verified above. These two visual/output checks remain explicit acceptance
limitations; they do not require reimporting the roster.

All implementation changes are committed/pushed. No frontend assets changed,
no Cloudflare plan upgrade or cap relaxation occurred, and no other restoration
batch was started. Remaining five-hour Codex allowance at the final check: 65%.

## Follow-up: MMC Orientation abbreviation

User requested MMC stream `O` be expanded to Orientation. Aggregate indexed
production inspection confirmed `0900-1300 O` for eight Intern events on
19 October and one on 2 November. Implemented source-specific parser recognition
and normalization of already-cached calendar/On shift events, preserving hours,
raw roster text and historical grade without database writes or reimport.
The cached stream catalogue display expands the label while retaining its key.
Deployed `b0bfe68f` to `https://877984b2.roster-to-calendar.pages.dev`.

Focused `node scripts/test-mmc-orientation.mjs`, subscription-artifact,
doctor-calendar-artifact, syntax and diff checks passed. The broader legacy
fixture suite stops at an unrelated existing login-markup assertion at line 888;
the On shift suite passes its server checks then stops at an existing startup
coalescing assertion at line 177. Neither failure was changed as part of this
label correction. No claim that these entire suites pass is made.

Final Orientation release: `451e9483`, deployment
`https://e240b49b.roster-to-calendar.pages.dev`. Safari confirmed a stale module
export caused the intermediate reload failure; both entrypoint and parser import
versions are now updated. Final reload completed normally. Live MMC On shift
showed Orientation with eight 09:00–13:00 Intern entries on 19 October. Selecting
2 November displayed the full Term 4 roster and Orientation containing Hieu HA
09:00–13:00 alongside the separate 11:00–13:00 orientation allocation. This also
completes the previously outstanding 2 November On shift visual check. Existing
private subscription-output live verification remains unperformed; artifact tests
pass. The orientation fix required no roster reimport or D1 row mutations.
