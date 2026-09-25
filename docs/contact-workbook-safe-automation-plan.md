# Contact workbook safe-automation plan

## Objective

Restore automatic DDH and MMC/MCH live contact publication without allowing an
Excel extraction action to modify its own SharePoint trigger source, without
reintroducing broad D1 work, and without retaining complete contact workbooks.

This plan selects one implementation. Do not trial the rejected alternatives in
Production unless this plan is formally revised.

## Current state and confirmed cause

- The `Sync DDH clinician contacts` and `Sync MMC clinician contacts` Power
  Automate flows are Off.
- Existing published contact overlays remain available through the already
  restored R2-backed DDH, MMC and MCH readers until their normal expiry.
- Production contact automation writes must remain closed while this work is
  prepared.
- Both flows currently use SharePoint file-modified triggers followed by Excel
  Online `Run script` actions.
- The DDH canary on 25 September 2026 proved that the settling/revision guard
  discards stale autosave triggers correctly. It also proved that the accepted
  `Run script` execution causes a later SharePoint modification and therefore a
  new flow run. Runs then repeated roughly once per minute until the flow was
  turned Off.
- Consequently, debounce alone cannot make the present architecture safe. The
  action that opens the trigger workbook in Excel must be removed from the
  automatic path.

## Chosen architecture

Use the SharePoint connector to read the stable workbook bytes, then perform the
already-supported bounded contact extraction in the application request:

```text
SharePoint source file changed
  -> two-minute settling delay
  -> reread exact source metadata
  -> stop if the trigger revision is stale
  -> SharePoint Get file content
  -> POST the workbook bytes plus source metadata
  -> bounded in-memory contact extraction
  -> semantic allocation hash
  -> zero-write unchanged response, or publish one small contact overlay
```

The automatic flow must not use Excel Online, Office Scripts, a temporary copy,
scheduled polling, or a self-editor filter. The existing Office Scripts may be
retained only as manual diagnostic tools.

## Safety invariants

Every implementation step must preserve all of the following:

1. DDH and MMC flows stay Off until their individual canary gates.
2. Contact workbook ingress is disabled by default and guarded by the existing
   exact source allowlist.
3. Authentication, source validation and declared-size validation happen before
   D1 access.
4. Only these exact source identities are accepted:
   - `ddh-daily-contact-sheet` for the approved DDH workbook; and
   - `mmc-shift-allocations` for the approved MMC workbook.
5. Each source has an explicit maximum body size derived from the real workbook
   size with limited headroom. Requests without a valid content length must
   still be stopped by a streaming/read ceiling.
6. Parsing is bounded to the existing approved worksheet/range:
   - DDH: `ED Clinicians!A1:N120`;
   - MMC/MCH: `SHIFT ALLOCATIONS!A1:I42`.
7. Raw workbook bytes are never written to D1 or R2 and are discarded after the
   request.
8. The semantic hash excludes provider timestamps and SharePoint versions.
9. An unchanged clinical extract writes zero D1 rows and zero R2 objects.
10. A changed extract updates only its small contact source object and dependent
    contact overlays. It must not rebuild roster, membership, coverage, daily
    presence, or base At a glance artifacts.
11. Existing request-local D1 statement ceilings and privacy-safe attribution
    remain active.
12. Power Automate HTTP retries remain disabled and flow concurrency is one.

## Phase A — local endpoint and parser work

1. Introduce a new contact-workbook endpoint rather than reopening the legacy
   broad binary route. Give it an explicit name and contract that identifies it
   as transient parsing, not workbook storage.
2. Authenticate with the existing automation token mechanism.
3. Carry source ID, filename, provider modification time and provider
   version/ETag in fixed request headers. Do not embed a second workbook copy in
   JSON or Base64.
4. Validate the source, exact filename, content type and size before parsing.
5. Pass the bytes to the existing bounded functions in
   `functions/_lib/contact-list-workbook.js`.
6. Feed the resulting normalized doctors/clinicians-only extract into the same
   semantic deduplication and small-overlay publication path used by
   `/api/automation/contact-list-extract`.
7. Refactor shared ingestion only as much as necessary so JSON and workbook
   entry points cannot drift in validation, hashing, retention or publication.
8. Do not add a new D1 table, backfill, migration, cron, queue, scheduled Worker
   or background processor.

### Phase A tests

Use synthetic/fixture workbooks locally. Prove:

- correct DDH extraction, including continuation roles and night phone layout;
- correct MMC and MCH extraction from the shared allocation workbook;
- exact source/filename isolation;
- missing/incorrect authentication is rejected before D1;
- wrong content type, malformed ZIP, missing sheet and oversized body are
  rejected safely;
- parser reads cannot escape the approved row/column bounds;
- an unchanged semantic replay writes zero D1 rows and zero R2 objects;
- a metadata-only replay writes zero D1 rows and zero R2 objects;
- one changed contact produces one bounded metadata update and only the expected
  contact overlay invalidation/publication;
- roster and facility base-artifact paths are not invoked; and
- the request-local statement budget stops unexpected expansion.

Run only the focused contact, quota and source-isolation suites plus syntax
checks. Do not expand to unrelated full-suite work unless a shared module was
changed.

## Phase B — prepare Power Automate while still Off

Apply the same pattern to each saved flow:

1. Keep the exact SharePoint site, library, folder and filename restriction.
2. Keep the two-minute delay and latest-revision comparison.
3. Keep the stale branch terminating successfully before file retrieval.
4. Remove the Excel Online `Run script` action from the automatic path.
5. Add SharePoint `Get file content` using the exact file identifier.
6. POST the binary body to the new stable Production endpoint with the fixed
   source and metadata headers.
7. Set concurrency to one and HTTP retry policy to none.
8. Confirm Flow checker reports zero errors and zero warnings.
9. Save both flows and confirm both remain Off.

Record the final flow IDs, exact filenames, endpoint, source IDs, body-size
limits and rollback action in the feature-restoration register. Never record a
token value.

## Phase C — disabled Production deployment

1. Commit the code, focused tests and documentation together.
2. Deploy with contact-workbook automation globally disabled and its source
   allowlist empty.
3. Confirm a request would be rejected before D1 while disabled.
4. Confirm existing contact readers and roster/calendar syncing are unchanged.
5. Reconcile the deployment and account-wide quota checker before admitting a
   source.

## Phase D — serial DDH canary

1. Open only the DDH contact-workbook source and its existing DDH contact build
   target.
2. Turn on only `Sync DDH clinician contacts`.
3. User action required: make one harmless DDH contact change and undo it, then
   report the completion time.
4. Wait for the two-minute settling window plus normal connector latency.
5. Verify in Power Automate:
   - stale save events stopped before `Get file content`;
   - exactly one stable run retrieved the file and called HTTP;
   - no later flow run was caused by the retrieval; and
   - the HTTP action succeeded once.
6. Verify in application telemetry:
   - the exact request has complete attribution;
   - statement/row counts remain within the contact ceiling;
   - an undo restoring the published clinical state produces zero D1 writes and
     zero R2 writes; and
   - no roster/facility-base builder ran.
7. Observe for at least two further settling intervals. Any unexplained trigger,
   repeat HTTP request, incomplete attribution or unexpected write is a failed
   gate: turn the flow Off and close the DDH source immediately.
8. If all gates pass, leave DDH enabled and document routine restoration.

## Phase E — MMC/MCH restoration

Proceed only after DDH passes Phase D.

1. Open only `mmc-shift-allocations` and its existing MMC/MCH contact build
   targets.
2. Turn on only `Sync MMC clinician contacts`.
3. Retrieve the current stable workbook through the new path once. A workbook
   edit is not required merely to prove the non-writing retrieval architecture;
   use a controlled test/resubmission method that cannot change the source.
4. Confirm one successful HTTP request, no follow-on SharePoint trigger, bounded
   D1 attribution, and correct current-date MMC and MCH extracts.
5. User action required: verify one MMC and one MCH On shift view, including the
   current AM/PM period as appropriate, then leave one view visible beyond the
   60-second refresh.
6. Confirm the visible refresh remains R2/session-only with zero D1 statements.
7. If all gates pass, leave MMC contact automation enabled and document routine
   restoration.

## Rollback

Rollback is source-specific and must not remove existing published data:

1. Turn Off the affected Power Automate flow.
2. Remove only that source ID from the contact-workbook automation allowlist.
3. If necessary, disable global contact-workbook ingress while leaving contact
   readers and retained R2 overlays intact.
4. Do not delete contact corrections, source objects, roster data or cached base
   facility artifacts.
5. Record the observed request/run IDs and the failed gate before making another
   attempt.

## Completion criteria

Contact automation is restored only when:

- neither automatic flow contains an Excel `Run script` action;
- genuine source edits publish automatically after settling;
- file retrieval never modifies or retriggers the source workbook;
- stale/autosave triggers stop before workbook transfer;
- unchanged clinical content performs zero D1/R2 writes;
- changed content updates the correct hospital and period;
- DDH and MMC/MCH readers continue their zero-D1 visible refresh behaviour;
- all request attribution is complete and account-wide quota status remains
  safe; and
- the feature-restoration register contains the final enabled state and exact
  rollback instructions.

## Operator checkpoints

Low-effort implementation should continue autonomously except at these points:

1. Stop for the user's DDH edit-and-undo action in Phase D.
2. Stop for the user's MMC/MCH visual verification in Phase E.
3. Stop immediately on any failed safety gate, describe the evidence, and do not
   try an alternative architecture without revising this plan first.

