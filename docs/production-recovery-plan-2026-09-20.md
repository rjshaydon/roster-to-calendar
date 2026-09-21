# Production recovery and restoration plan — 20 September 2026

## Purpose and authority

This is the current operational sequence for restoring the roster application.
It supersedes stale rollout checkpoints in the earlier plans, but retains their
permanent safety requirements, source-isolation design and feature-restoration
register. It does not authorise a deployment, D1 migration, backfill, Power
Automate change or Production test merely by existing.

Priorities remain:

1. preserve reliable login, calendars and subscription feeds;
2. prove and restore doctor switching;
3. confirm MMC and MCH automatic syncing, then restore VHH and DDH;
4. verify users receive current calendars and feeds; and
5. resume the separate At a glance rollout only after roster syncing is stable.

## Reconciled starting state

- Production `main` is at `6534991`.
- Creator login succeeded at approximately 14:50 AEST on 16 September after
  the login-specific statement ceiling was returned to the normal guarded
  request ceiling.
- A Creator switch to Toby VANHEST remains unverified. The failed attempts
  showed that Toby uses the unclaimed doctor-profile path. The last deployed
  repair replaced the history-sized file/doctor predicate with an indexed
  active-file doctor-key query, but that repair was not tested after login was
  restored.
- Monash Adults and Monash Paediatrics were successfully ingested on 15
  September and their isolated routine flows were enabled at that checkpoint.
  Their present flow state must be read back before relying on it.
- VHH routine syncing was restored on 20 September after a bounded canary and
  unchanged replay. DDH remains paused because its isolated provider check
  returned HTTP 503.
- At a glance, contact automation, Doctor Names/identity discovery, manual
  roster mutation and unrelated maintenance remain paused unless the feature
  register explicitly records otherwise.
- The read-only account-wide sample generated at 12:46 AEST on 20 September,
  settled through 12:31 AEST, recorded 5,478 reads, zero writes and no expensive
  fingerprint. It is a valid first sample only; the checker correctly returned
  `STOP` because a second sample is required.
- Today's dominant fingerprint is the indexed active-roster doctor-key event
  query used by personal calendar/feed delivery. It read 4,989 rows across 15
  executions, rather than scanning global roster history.

## Rules for every stage

- Never use remaining daily quota as permission for an unbounded query.
- Use account analytics and request attribution; do not query Production D1 to
  measure D1.
- Change one independently reversible subsystem at a time.
- Before any controlled Production action, record the deployment, flags, flow
  state, expected requests and explicit stop threshold.
- After it, wait for settled attribution and reconcile rows examined, rows
  returned, writes, action and deployment before continuing.
- Do not ask the operator to repeat a failing action until the failure has been
  reproduced locally with a production-shaped history fixture.
- A UI error, unexplained write, missing attribution, expensive fingerprint or
  safety-stop response closes the gate. Do not compensate by merely raising a
  ceiling.

## Gate 0 — read-only reconciliation

No application code or Production state changes.

1. Take the required second account-budget sample at least ten minutes after
   the valid first sample.
2. Read back the active Production deployment and effective feature flags.
3. Read back Power Automate states for the isolated Adults and Paediatrics
   flows and all still-paused flows.
4. Read back recent GitHub processor runs and confirm no unexpected scheduled
   or concurrent processing.
5. Reconcile settled request attribution for the final 16 September login and
   profile-switch attempts.

Gate: two samples reconcile to `GO`; zero unexplained writes; no expensive or
unknown fingerprint; one intended Production deployment; the enabled/paused
source inventory is explicit.

## Gate 1 — make failures diagnosable locally

This is a small observability correction, not a feature change.

1. Preserve a machine-readable distinction between:
   - the application's per-request statement guard;
   - a native D1 SQL/statement failure; and
   - the account-wide daily quota response.
2. Record a privacy-safe operation phase for doctor-profile loading (registry,
   revision, active events, issues, snapshot write) without recording names,
   credentials or SQL parameters.
3. Ensure API handlers do not swallow the guard's typed error and return the
   misleading raw text `D1 statement budget exceeded.`
4. Test that a stopped request records its actual phase and performs no retry,
   background build or navigation save.

Gate: a deliberately stopped local request identifies one cause and phase;
ordinary login remains the fast two-query shape; no safety ceiling is raised.

## Gate 2 — prove doctor switching locally

1. Build a deterministic Toby-shaped fixture: one MCH doctor associated with
   many historical files, several active files, aliases and a realistic number
   of events.
2. Prove that the switch queries active events/issues by indexed doctor keys,
   not by an expanding file/doctor `OR` predicate.
3. Assert exact ceilings for statements, rows returned and writes. A read-only
   switch must write nothing when a current snapshot exists. A missing snapshot
   may write only its compact registry row and R2 artifact once.
4. Prove no automatic save of the calendar being left, no hidden `waitUntil`
   build and at most one bounded foreground attempt.
5. Prove UI transactionality: success changes both calendar and label; failure
   changes neither.

Gate: the production-shaped fixture passes repeatedly with cost independent of
historical file count, and all focused login, quota and attribution tests pass.

## Gate 3 — deploy closed and observe login

1. Deploy the Gate 1–2 repair without changing roster flows or feature flags.
2. Confirm the intended commit is the sole active Production deployment.
3. Observe ordinary feed traffic and login only. Do not test doctor switching
   in the same interval.
4. Reconcile a settled post-deployment sample.

Gate: login succeeds; no unintended writes; no safety error; usage matches the
bounded feed/login baseline.

## Gate 4 — one doctor-switch canary

1. From a fresh browser window, perform exactly one Creator switch to Toby
   VANHEST and no other optional action.
2. Record the time and visible result, then close the window.
3. Reconcile the settled request record and account-wide delta before any
   second switch.

Gate: Toby's calendar and picker label both change; no error; no navigation
save; attributed cost stays within the locally proven bound. Otherwise roll
back or close profile switching and return to Gate 2 with the recorded phase.

## Gate 5 — confirm restored Monash syncing

1. Read back the Adults and Paediatrics flow definitions and enabled state.
2. Verify their latest provider versions and successful GitHub processor runs.
3. Use a real provider revision when available, or one reviewed one-shot exact
   source submission—not a fabricated workbook edit.
4. Verify one affected personal calendar and its subscription feed, then replay
   the unchanged revision and prove zero fact/object rewrites.

Gate: MMC and MCH changed revisions update calendars; unchanged revisions are
idempotent; their flows remain independently pausable.

## Gate 6 — restore remaining roster sources

Restore one implementation class at a time:

1. DDH FindMyShift ingestion;
2. VHH Office Script/SharePoint extraction.

For each source: inspect configuration, run the local representative fixture,
open only that exact source, observe one changed and one unchanged run, verify
an affected calendar/feed, reconcile settled usage, then continue. A failure
pauses only that source.

Gate: all four roster sources sync automatically and independently; unchanged
runs perform no derived rewrites; user calendars and feeds reflect current
rosters.

## Gate 7 — restore remaining desirable features

After at least one stable observation period with all roster sources working:

1. restore Creator manual roster operations;
2. resume the chunked, cached At a glance publication plan;
3. restore cached At a glance readers one ED/date surface at a time;
4. restore contacts, insights, identity review and other paused features from
   the feature-restoration register.

Never restore the legacy live At a glance history scans.

## Immediate next action

Gates 0–6 are complete for routine roster synchronisation. MMC, MCH, VHH and
DDH are enabled behind exact source isolation and the permanent incremental
fact ceiling. The 22 September DDH import completed successfully with 149
doctors and 3,651 events; representative Dennis CHUNG and Steve GUASTALEGNAME
calendar checks passed. The next operational priority is Gate 7: preserve fast
core calendar loading, then begin the one-MMC-term cached At a glance rollout
defined in the 22 September readiness packet in
`core-calendar-sync-restoration-plan.md`. Do not restore contacts or any legacy
At a glance reader during the first canary.

## Completed evidence — 20 September 2026

### Gate 0

- The delayed account-wide sample returned `GO`: 6,004 settled reads, zero
  writes, no expensive fingerprint and a projected daily total of about 28,500
  reads. The delta from the first sample was 526 reads over about 29 minutes.
- Production was serving commit `6534991`.
- `Sync Monash roster files` and `Sync Monash Paediatrics roster files` were
  enabled. Every VHH, DDH, contact and bootstrap flow visible in Power Automate
  was disabled.
- Recent GitHub roster processors were only `monash-adults` and
  `monash-paeds`, and completed successfully.
- Settled 16 September attribution proved that login was two statements and
  three rows read. The failed Toby switch stopped in `loadDoctorProfile` after
  14 statements; it was the application's request ceiling, not the account
  daily quota. The state handler had converted the typed guard into the
  misleading raw 400 response.

### Gates 1 and 2

- Guard stops, native D1 failures and account daily quota errors now remain
  machine-distinguishable. Profile requests record a privacy-safe operation
  phase, and guard errors are no longer swallowed by the state handler.
- A guard-stopped profile request is not retried by the browser.
- Doctor-profile event and issue reads enumerate active files for the profile's
  source through `idx_roster_files_source_active`, then probe by exact
  file/doctor index. They do not expand historical file/doctor `OR` clauses.
- The deterministic cost fixture contains 150 inactive MCH Toby histories and
  one active record. The query returns the one active record; its plan uses the
  active-source index and exact file/doctor index, not the doctor-history index.
- The existing uncached profile fixture remains within 20 statements. A ready
  cached profile switch performs no `INSERT`, `UPDATE` or `DELETE`.
- Focused quota, attribution, client retry, Creator login containment, source
  isolation, database-cost and full representative fixture suites pass. No
  statement ceiling was raised.

### Gate 3 first attempt and remediation

- The first login-only canary on `99d150d3` failed at 15:36 AEST. Attribution
  recorded `login`, status 503, `d1-statement-budget`, one statement, zero rows
  and zero writes. This ruled out expensive authentication and account quota.
- The request meter had replaced `ROSTER_DB` on the supplied environment
  object. A reused isolate could therefore wrap an already-metered binding;
  feed requests accumulated against an older inner meter and a later login was
  rejected on its first statement.
- A first correction using a cloned environment restored login but Pages did
  not pass that replacement environment to the downstream handler, so its
  attribution recorded zero metered statements. That version is not the final
  safety implementation.
- The final correction unwraps any stale meter, installs a fresh meter only for
  the duration of the downstream request, and restores the raw binding in the
  middleware `finally` block. A regression test runs two 40-statement requests
  against the same supplied environment: both retain independent 64-statement
  budgets and the original binding is restored after each request.
- Gate 3 must restart after the corrected deployment. The failed login is not
  evidence against the bounded doctor-profile query, which was never reached.

### Gate 3 restart and Gate 4 canary

- Production commit `390700ff` restored the raw D1 binding after every request
  and retained request-level attribution through the downstream handler.
- The 15:47 AEST login-only canary succeeded. Login used 2 statements, read 3
  rows and wrote none; calendar loading used 15 statements, read 11 rows and
  wrote none. The independently attributed requests confirm that the meter no
  longer accumulates work between requests.
- The first Toby VANHEST canary succeeded. Account resolution used 2
  statements and read 162 rows. The uncached doctor-profile request used 25
  statements, read 6,518 rows and wrote 5 rows while creating its compact
  snapshot state. No guard or quota error occurred.
- The repeat Toby canary used the cached path: 14 statements, 1,501 rows read
  and zero writes. Account resolution and the separate access check each used
  2 statements and read 162 rows. The calendar and picker changed together and
  no error was shown.
- This is a 77% reduction in profile rows read after the one-time snapshot
  build and proves that repeat switching does not rewrite the snapshot. The
  remaining 1,501-row cached read is bounded but remains a later optimisation
  target; it is not a reason to delay restoring roster syncing.

### Gate 5 controlled Monash verification

- Both `Sync Monash roster files` and `Sync Monash Paediatrics roster files`
  remained enabled. Other roster/contact/bootstrap flows remained disabled.
- Paediatrics processed `Paeds - Term 3 2026.xlsx` successfully on 20 September:
  80 doctors and 2,375 events.
- SharePoint showed `AdultTerm3.2026.xlsx` modified on 19 September, but the
  Adults processor had not run since 17 September. Power Automate history
  showed that both modification-triggered runs fetched the workbook and then
  failed at the HTTP submission with status 422.
- Attribution for those failed submissions showed one D1 statement, zero rows
  read and zero writes on the older deployment. This matches the defective
  shared request meter fixed by `390700ff`; it was not a workbook/parser error.
- A single resubmission on the corrected deployment queued and processed
  `AdultTerm3.2026.xlsx` successfully: 152 doctors and 4,173 events. Ingest used
  13 statements, read 3 rows and wrote 20; the derived update used 40
  statements, read 54,706 rows and wrote 133. No guard or quota error occurred.
- Replaying that exact provider revision returned unchanged with 2 statements,
  zero rows read, zero writes and no GitHub processor dispatch. This proves the
  unchanged path does not rewrite roster facts or derived data.
- At 16:15 AEST the known MMC roster change was correct in both the application
  calendar and the subscribed Apple Calendar feed. This completes Gate 5.

### Gate 6 DDH opening check

- Production commit `6b695d4c` added only `dandenong-findmyshift` to the
  automated roster source allowlist. The separate watchdog Worker was not
  deployed, so this did not create background polling.
- The first Creator-controlled refresh at 17:55 AEST failed before ingestion.
  Attribution recorded 5 statements across the outer refresh and provider
  check, 4 rows read and 2 compact source-status writes. No GitHub processor
  was dispatched and the active DDH roster was not replaced.
- The provider-check handler previously discarded its safe failure diagnostic,
  leaving no historical distinction between rate limiting, an HTTP/provider
  rejection and an unexpected response. The next deployment preserves only
  the safe code, HTTP status, content type and response size; it does not expose
  credentials, provider response bodies, roster content or clinician details.
- The diagnostic retry at 18:02 AEST returned `http-error, HTTP 503`.
  Attribution recorded four statements, one row read and two compact status
  writes for the provider check. No ingestion or GitHub processor was started
  and the retained DDH roster remained active. This is an upstream provider
  availability failure, not a D1 quota or statement-budget failure. Remove DDH
  from the active canary allowlist and retry it on a later provider window.

### Gate 6 VHH preflight

- Local VHH extraction, source-isolation and roster-ingress idempotency suites
  pass. An unchanged VHH provider revision or identical content writes zero D1
  rows and zero R2 objects.
- `Sync VHH Active Medical Roster to Production` is the disabled instant
  canary. It reads exactly
  `/Shared Documents/Medical/Rosters/Active Medical Roster.xlsx`, uses the
  SharePoint ETag as `providerVersion`, and submits the raw workbook with exact
  source ID `vhh-active-medical-roster` to the bounded ingress endpoint.
- `VHH Active Medical Roster to Production` is the disabled automated flow.
  Its trigger currently watches the entire `/Shared Documents/Medical/Rosters`
  folder and its retained run history contains multiple long-running failures.
  It is not authorised for routine enablement until its trigger is narrowed to
  the exact workbook and its retry behaviour is reviewed.
- For the controlled VHH canary, retain MMC and MCH, replace DDH with exact VHH
  in the server allowlist, keep both VHH flows off, then run the instant flow
  exactly once. Reconcile ingestion and processor attribution before any
  calendar test or unchanged replay.

### Gate 6 VHH completion

- Production deployment `413534f9-cd56-4de7-b75e-bedeb8d2a9b5` permits only
  `monash-adults`, `monash-paeds` and `vhh-active-medical-roster`; DDH remains
  excluded.
- The one-shot VHH canary completed successfully: 46 doctors and 971 events
  were parsed, and the derived phase used 41,057 D1 row reads and 2,128 writes
  within the request-local ceiling. Claire CHARTERIS's VHH calendar then loaded
  without an error.
- Replaying the identical workbook used two statements, zero D1 row reads and
  zero writes, and dispatched no processor run.
- The automated flow now has concurrency one, an exact filename trigger for
  `Active Medical Roster.xlsx`, and no HTTP retries. Flow checker reported zero
  errors and zero warnings before it was saved and enabled. The instant canary
  flow remains off.

### Gate 6 completion update — 22 September 2026

- Guarded MMC, MCH and VHH routine flows are active and source-isolated.
- DDH FindMyShift ingestion completed successfully at 01:41 AEST: 149 doctors
  and 3,651 events, with no duplicate or abandoned queue sibling.
- MCH's apparent permanently queued state was traced to one manual refresh
  queueing two retained files at the same timestamp while the one-item worker
  processed only one. `90bec097` now queues only the canonical active file and
  prefers a completed sibling when displaying legacy tied rows.
- The Creator doctor directory is retained locally while browsing claimed and
  unclaimed calendars (`79c5c7ec`), adding no D1 request.
- Dennis CHUNG's payroll-transfer shifts and Steve GUASTALEGNAME's 1 November
  Rover shift were confirmed in Production.
- Cached repeat switching was fast, but one cold claimed-account switch took
  approximately 30 seconds and one Creator login approximately 10 seconds.
  These are not accepted as resolved merely because later cached attempts were
  faster. The protected performance gate in the core restoration plan governs
  the follow-up.
