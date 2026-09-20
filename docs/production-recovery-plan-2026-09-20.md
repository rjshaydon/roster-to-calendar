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
- VHH and DDH routine syncing are not yet restored.
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

1. VHH Office Script/SharePoint extraction;
2. DDH FindMyShift ingestion.

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

Gate 0, Gate 1 and Gate 2 are complete locally. The next action is Gate 3: a
closed deployment followed by a login-only observation. Do not test doctor
switching until that interval has settled and reconciled.

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
