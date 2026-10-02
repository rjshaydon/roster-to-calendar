## Current batch — 2 October bounded identity and delivery restoration

Production `6513315a`, deployment `bd4924c6-05e6-4299-bc40-929ec8535954`,
restores bounded automatic name suggestions, explicit confirmed linking and
Creator directory grades using visible R2 term staff data. It also fixes the
JSON transport hash mismatch blocking real MMC/MCH imports and restores the
bounded diff path for small same-term automatic corrections. Global repair,
history discovery, Creator startup fan-out and snapshot builders remain closed.

MMC/MCH deliveries and all-site continuation completed. A stale-provider ordering
defect found during verification is fixed; MMC version 398 was recovered at
05:55 UTC. Read-only provider comparisons confirm the latest checked revisions
for the other three sites.

Remaining verification: authenticated
identity UI, calendar/subscription freshness and consolidated user acceptance;
check settled account-wide usage and remove superseded deployments. Real
MMC/MCH/DDH future terms remain dependent on provider publication. VHH completed
a real automatic import at 04:13 UTC. VHH contacts and five-minute contact
schedules were restored earlier today. Ordinary manual mutations, session
settings, colleague tools and Creator switching are already deployed.
Implementation and evidence: `bounded-roster-import-checkpoint.md`. The separate
parked Doctor Names/merge project is not part of baseline functional restoration.
Historical checkpoints below describe their date, not current feature state.

## Restoration batch — 2 October 2026

Step 1: settled account-wide analytics through 01:07 AEST recorded 304,036 reads and 14,094 writes for the UTC quota day; 80,427 reads and 3,350 writes since the previous release. Telemetry reconciled across all three databases. Normal-use traffic and scheduled work stayed comfortably below admission thresholds.

Step 2 deployed as `d6ad1ddf` and locally/live verified: dedicated field-level session settings saves with optimistic conflict detection; bounded manual imports and automatic large corrections; inactive staging, pinned active-set promotion, preserved future terms and recoverable removal. Current personal calendars use published R2 shifts with cached historical shifts retained outside visible published terms. Publication deferral preserves committed imports and continues through scheduled maintenance. Full-range replacements are required: a partial overlapping upload is rejected rather than discarding retained dates.

Advanced whole-database rebuilding remains disabled. VHH contacts, bounded automatic identity discovery and real MMC/MCH/DDH next-term delivery verification remain next. Production release evidence is recorded in `bounded-roster-import-checkpoint.md`. VHH flow review confirms successful JSON extraction and a missing Preview HTTP destination; the extract lacks a date and most shift periods. Current-holder versus whole-day sheet semantics require clarification before live matching.

## Verified restoration checkpoint — 1 October 2026, 16:15 AEST

Current Production includes bounded automatic roster imports and continuation, all-site current-term publication, VHH next-term publication subject to visibility, automatic On shift, MMC/MCH/DDH contacts, explicit cached colleague tools, and the bounded Creator directory/picker with published doctor-profile calendars. Creator startup fan-out and automatic identity discovery remain disabled. Commit `98f32019` is deployed; Safari verified the directory and doctor switching, plus both colleague tools.

Remaining service work: VHH contact automation; manual roster mutations and larger overlapping replacements; bounded automatic identity discovery/claim suggestions; dedicated cross-device settings persistence. End-to-end real MMC/MCH/DDH next-term delivery remains to be demonstrated. Doctor-profile views show visible published terms and do not rebuild unpublished historical calendars. See `bounded-roster-import-checkpoint.md` for deployment and test evidence. Historical ordered steps below are superseded by this checkpoint and the account-aware batching policy.

# Full service restoration plan

## Quota and batching adjustment — 1 October 2026

This adjustment supersedes the temporary restoration ceilings and serial
micro-canary ordering below. It does not claim the quota implementation has
already changed. Current deployment evidence is in
`bounded-roster-import-checkpoint.md`.

Cloudflare Free-plan account quotas are 5,000,000 rows read and 100,000 rows
written per UTC day. The importer ledger's 500,000-read / 10,000-write ceilings
are temporary application limits, not provider quotas. Reserved costs are not
measured billing totals. Index maintenance must be included in write estimates.

Target admission policy: stop admitting additional restoration work before
projected account usage exceeds 4,000,000 reads or 80,000 writes per UTC day.
These are maximum safety thresholds, not spending targets; reduce available
maintenance capacity further if normal-traffic projections require more
headroom. Account usage includes all databases/callers. Count settled analytics,
unsettled measured work and outstanding reservations without double-counting.
Allow for analytics delay and concurrent activity. Missing or inconsistent
telemetry stops additional costly maintenance, while cached readers continue.
Retain bounded per-operation limits, atomic reservations, durable continuation,
idempotency, indexed access and stale-publication protection. Replace the old
daily and per-pass caps and the fixed pending cutoff together; do not simply
remove the write guard or increase one constant.

Execution order, using combined functional batches:

1. Implement and test the reconciled account-aware budget policy. Check live
   account usage, deploy once, then finish queued MCH/DDH publication and DDH
   compact preparation. Verify all-site current views and next-term readiness;
   distinguish provider deliveries not yet received from failed imports.
2. Restore VHH contact extraction/publication using bounded source-specific
   extraction, together with remaining routine roster correction/overlap support.
   Preserve active terms and test interruption, replay and large replacements.
3. Restore colleague tools, Creator switching/directory and doctor discovery
   as a batch using compact indexed lookups/cached projections. No roster-history
   scans, identity fan-out or global warm-up.
4. Restore bounded manual roster mutations and cross-device settings, then run
   one combined user-visible regression check and account-usage review.

Keep a successful batch enabled. If a batch fails, disable or revert its new
feature switches while preserving valid data; isolate the failing portion using
request-cost evidence and local tests. Avoid repeated deployments or waits for
each small change. A batch is complete only when its visible behaviour and
measured costs are verified. Full restoration remains the objective.

Quota reference: https://developers.cloudflare.com/d1/platform/pricing/

Status: DDH, MMC and MCH current contact publication and automatic flow
execution confirmed on 1 October. User confirms contacts visible in all three
sites. MMC uses the recovered Microsoft-side Office Script / small-JSON path;
DDH retains its working workbook extraction path. Other restoration batches
remain open.

### MMC/MCH contact restoration — confirmed 1 October

**1 October confirmation checkpoint (11:47–11:50 AEST):**
- User logged into the app and confirmed visible contacts for MMC, MCH and DDH.
- Both `Sync MMC clinician contacts` and `Sync DDH clinician contacts` read
  back **On**. No flow edits, submissions, deployments or workbook edits made
  during this confirmation.
- MMC automatic run `08584107890701849482674731631CU22`, started 11:23,
  succeeded: settling delay, metadata, condition, Office Script (26.3 seconds)
  and HTTP (0.3 seconds) all executed. This is not a skipped-extraction replay.
  Latest ten runs include settled accepted runs at 11:13 and 11:23 and shorter
  superseded runs; no ongoing minute-by-minute execution after 11:23 was shown
  at 11:47. This observation does not prove every future edit cannot retrigger.
- Specific MMC/MCH R2 manifest has operational date **2026-10-01**, received
  11:16:13 and published 11:16:14 AEST, provider modified 11:13:17. Extract
  contains both Adult and Paediatric sections. Existing contact reader run
  locally against these exact published objects returns `available` for MMC
  and MCH, with 17 and 9 populated AM contacts respectively. Contact details
  were not printed or added to Git.
- DDH run `08584107968727202331917818539CU06`, started 09:13, executed
  file retrieval and HTTP successfully. R2 date **2026-10-01**, received
  09:15:33 and published 09:15:35 AEST. Same reader verification returns
  `available`, 14 populated AM contacts. DDH remains unchanged.
- Fresh account analytics sample: 12,033 reads / 27 writes. The checker returns
  STOP only for the new UTC day's two-hour passive baseline and missing second
  sample; this is not a quota exhaustion or admission for another rollout.
  Prior temporary reports were no longer present. Do not start a new rollout
  on the basis of this single sample.
- Functional automatic contacts are now confirmed for all three hospitals.
  Exact per-request D1 telemetry, a controlled unchanged replay and a new
  60-second visible-page observation were not repeated during this narrow
  confirmation. Existing semantic no-op and zero-D1 refresh safeguards remain.
  Bootstrap MMC remains the incompatible disabled historical flow; do not run
  it or restore whole-workbook transfer.

**1 October trigger diagnostic checkpoint (00:20–00:32 AEST):**
- User performed the requested edit. SharePoint library confirms the exact
  `Contact lists/shift allocation/SHIFT ALLOCATIONS.xlsx` was modified by the
  user at 00:20 AEST. Do not ask for another edit merely because no run appears.
- MMC flow remains On. Trigger readback: correct SharePoint site/Documents
  library, exact filename condition, Split on enabled, concurrency Off, polling
  definition one minute. No failed trigger checks; last recorded no-data check
  at 30 September 23:52. No fresh run appeared during this diagnostic window.
- Armed manual live-trigger test; it remained waiting. Restarted only the MMC
  flow Off/On at approximately 00:28 AEST, preserving its definition and guards.
  Readback confirmed On. This is a trigger recovery attempt, not success proof.
- Budget checker at 00:30:20 returned GO: 84,490 reads / 229 writes.
  Report `/private/tmp/mmc-oct1-budget.json`. Specific R2 manifest read still
  shows only 25 September; no current extract has been proven published.
- Inspected retained `Bootstrap MMC shift allocations`
  (`09e5c1c1-4697-4b62-a536-c2531a8fcba0`). Its current chain reads the entire
  workbook and sends base64 to deployment `4a12dbe2`'s `contact-list-binary`
  endpoint. It is not the recovered JSON implementation. Left Off and unchanged;
  do not run it as a shortcut or upload the full workbook.
- Mac locked at approximately 00:32 AEST; UI access blocked, unlock requested.
  Next: after unlock, refresh MMC runs and skipped/failed trigger checks after
  the restart; distinguish actual polling delay from an unhealthy trigger before
  further edits. Verify current publication and recursion only after a fresh
  accepted run. DDH and other flows remain unchanged.

**30 September enablement checkpoint (23:49–23:58 AEST):**
- Reused the existing valid budget sample; fresh account checker returned
  **GO** at 23:49:44 AEST, 78,035 reads / 215 writes, no stop reasons.
  Report: `/private/tmp/mmc-restoration-budget.json`.
- Read back the saved MMC flow: Office Script extraction followed by JSON POST
  to `/api/automation/contact-list-extract`, with script sourceDate/contacts
  and provider metadata. No whole-workbook transfer or replacement flow.
  Flow checker: zero errors and zero warnings. No definition changes made.
- Enabled `Sync MMC clinician contacts` at approximately 23:51 AEST;
  Power Automate readback confirms **On**. DDH and other flows unchanged.
- Focused contact-sync and contact-allocation checks pass. Specific R2 manifest
  read still contains only 25 September. Run history has no fresh run yet;
  yesterday's stale successful replay remains the latest entry.
- Requested one harmless workbook edit/undo and close from the user to generate
  a fresh revision. Awaiting that action; current-date publication, both hospital
  sections, unchanged replay and absence of self-trigger recursion remain
  **unverified**. Do not label this checkpoint end-to-end restoration.
- Next: inspect the fresh run, verify Run script AND HTTP actually succeeded,
  read the specific R2 manifest/object without printing contacts, observe for
  follow-on runs, and perform the combined MMC/MCH visible check. If recursion
  occurs, turn this flow Off immediately; retain the proven JSON extraction path.

**30 September read-only checkpoint (23:31–23:35 AEST):**
- Account analytics: 76,015 reads / 215 writes, no flagged expensive
  fingerprints. Single sample `/private/tmp/contacts-sep30-budget.json`;
  admission is STOP solely because a second sample is required. No enablement
  or production data query performed in this check.
- MMC flow confirmed **Off**. Latest replay at 29 September 15:57 reports
  success but its condition cancelled and both Run script and HTTP were
  **Skipped** (run `08584109454089369892445051106CU03`). Do not treat that
  replay or the user's visible contacts as proof of new JSON publication.
- MMC/MCH R2 manifest still contains only 25 September, published that day.
  DDH manifest contains 30 September, published 30 September 23:14 AEST;
  leave working DDH unchanged.
- Next narrow operation: fresh budget admission, enable the existing MMC
  flow and use a fresh real workbook modification (not replay of a stale
  trigger); verify Run script AND HTTP ran, both Adult/Paediatric records
  reached the current-date manifest, and no self-trigger loop. Then perform
  one combined MMC/MCH visual check. No replacement helper or workbook
  download is warranted by these findings.
- VHH is not a switch-only restoration: the current Production contact source
  normalizer supports MMC/MCH and DDH only. Its retained Preview flow must not
  be enabled as though it were a validated Production contact path.

- Edited existing `Sync MMC clinician contacts`
  (`64d9fad7-3462-4b6a-bb6e-95ce1951f4fb`), not a duplicate flow.
- Reused the script and workbook identifiers from its successful 25 September
  13:48 run. Removed the superseded whole-workbook `Get file content` action.
- HTTP now posts source ID, script date, doctor contacts and provider metadata
  to the stable Production `/api/automation/contact-list-extract` endpoint.
  Existing authentication and protected HTTP inputs/outputs retained; retries,
  HTTP chunking and asynchronous polling disabled.
- Existing two-minute settling delay and stale-trigger check retained.
  Flow saved successfully while Off. DDH and other flows unchanged.
- Local `test:contact-sync` and `test:contact-allocations` passed.
- Admission samples at 15:06 and 15:12 AEST were low (111,512/505 and
  112,038/510 reads/writes), but the second sample was too early for the
  ten-minute admission interval. No test was admitted on that basis.
- Next: admitted single replay, verify current-date MMC/MCH R2 publication,
  then verify automatic trigger behaviour before leaving the flow enabled.
  A successful save alone is not restoration evidence.
- 15:16:55 AEST admission returned GO: 112,159 reads / 510 writes, no
  flagged expensive fingerprints. The Mac then locked before the replay
  confirmation could be submitted. Flow remains saved and Off; test and
  automatic restoration are not yet verified. Unlock is required to continue.

## Authority and objective

This is the authoritative ordering for restoring the Production service after
the September 2026 D1 incident. It supersedes the serial rollout ordering in
the older quota, calendar-sync, At a glance and contact plans. Their permanent
safety constraints, parser requirements and recorded evidence remain valid.

The objective is to restore the user-visible service that existed on
31 August 2026 as quickly as possible without restoring the mechanisms that
scanned roster history, performed runtime schema work, retried blindly or
wrote unchanged state. Work proceeds in functional batches. A safe batch stays
enabled. A batch is split only if measured evidence identifies it as unsafe.

## Reconciled Production position — 27 September 2026

The repository and restoration register do not currently describe the same
state. The tracked Production configuration at `3161b23` has:

- automatic roster ingress and queue processing enabled for Monash Adults,
  Monash Paediatrics, Dandenong FindMyShift and the VHH active roster;
- cached At a glance readers enabled for MMC, DDH, MCH and VHH;
- cached contact readers enabled for DDH, MMC and MCH;
- contact-workbook ingress enabled for the exact DDH and MMC sources;
- automatic At a glance launch, ordinary manual roster mutations, colleague
  insights, identity discovery, broad Creator hydration, global snapshot
  warm-up and advanced maintenance disabled; and
- legacy historical At a glance reads permanently disabled.

The 27 September user test found no DDH contact details. Therefore a successful
Power Automate run is not accepted as proof that DDH contact publication works.
DDH and MMC/MCH contact automation are **not restored** until the correct
current-date overlay is present and visible in the application. The current
state of every external flow must be read back before changing it; Git cannot
prove whether a saved flow is On or Off.

Cached At a glance objects are also not yet kept current automatically after a
roster change because the metadata/day builders remain disabled. Existing
views can work while silently becoming stale. This is a restoration blocker.

## Features still requiring restoration

### Required for the core service

1. End-to-end automatic roster freshness for all four sources, including a
   changed source reaching personal calendars and subscribed calendar feeds.
2. Incremental publication of affected At a glance metadata/day objects after
   a roster change, without a whole-hospital or whole-term rebuild.
3. Automatic DDH and MMC/MCH contact publication and visible current-period
   contacts. VHH contact support remains a separate product task because the
   Production parser/source was never completed.
4. Automatic At a glance launch for entitled users and its existing
   visibility-aware 60-second contact refresh.
5. Explicit colleague tools: “Who else is working with me?” and “When am I
   working with…?”, served from compact daily facts rather than event joins.
6. Ordinary Creator manual imports, replacement and removal with incremental,
   source-scoped mutations.
7. Fast Creator switching and directory enrichment without roster-history
   discovery.
8. Bounded automatic doctor discovery/claim suggestions.
9. Cross-device calendar/session-setting persistence through a dedicated
   single-record API, not the former broad save path.
10. Creator contact corrections if the end-to-end contact review shows that
    they are not already operating on the compact overlays.

### Operational capabilities, restored only if needed

- A source-specific schedule for DDH checking if no existing external trigger
  provides routine automation.
- Advanced repair/bootstrap as short-lived operator actions with exact scope.
- A watchdog only if a demonstrated operational gap remains; never the old
  global poller.

### Work that follows service restoration

- Durable Doctor Names and identity aliases.
- VHH contact-list discovery and implementation.
- Other planned product additions that were not working Production features on
  31 August.

## Mechanisms that must not return

- legacy At a glance queries over roster-event history;
- runtime schema inspection or `CREATE ... IF NOT EXISTS` on user requests;
- automatic colleague-insight warm-up;
- global snapshot warm-up after ordinary saves;
- one D1 write per displayed status message;
- repeated blind login/calendar retries;
- duplicate same-open metadata requests; or
- automatic Excel actions that edit their own SharePoint trigger workbook.

Restoring the user outcome does not mean restoring its unsafe implementation.

## Restoration method

### Independent releases

Group independently ready capabilities into one reviewed release and one
Production deployment. A capability requiring substantial implementation must
not delay another that is ready. Contacts, roster readiness and At a glance
publication below are independently releasable; their numbering is priority,
not a requirement to wait for every preceding item. During
recovery, use one deployment mechanism; do not create both a GitHub deployment
and a manual deployment for the same commit. Keep the active deployment and at
most one verified predecessor until the batch passes, then delete older
deployments through the control plane. Git history is the long-term rollback
record.

All external callers must use the stable Production domain, never a hashed
deployment URL. Inventory Power Automate, GitHub Actions, schedules and deploy
hooks once at the start; subsequently inspect only changes and unresolved
items. Deployment cleanup is release hygiene, not a fresh prolonged gate for
each feature. The incident record does not prove that deployment count caused
the read bursts; it separately records an approximately 998,000-row colleague
query. Preserve that distinction.

### Evidence and stop rules

Do not wait for long passive intervals after every small action. Use live
request attribution and the existing request-local statement ceiling while a
batch is exercised, followed by one settled account check for the batch.

Before enabling a capability, prove query cost locally with representative
scale, effective indexes, bounded input and query plans. A returned-row LIMIT
does not bound examined rows. The current middleware stops excess statements
but observes row counts only after execution; it cannot prevent one expensive
statement. Telemetry thresholds below are detection/containment, not a hard
pre-execution row budget. Do not enable a known unbounded path to discover its
cost in Production.

Retain the account-wide quota checker's admission rules and reserves: four
million reads and 80,000 writes reserved for ordinary availability; optional
work limited to one million reads/20,000 writes including 100% contingency;
start below 500,000 reads/10,000 writes and stop optional work at 750,000
reads/15,000 writes. Use the existing reconciled two-sample baseline and freshness
rules; reuse valid evidence rather than restart a passive observation period
for each feature. Missing/stale telemetry blocks new Production experiments,
but independent local implementation continues. Changing budget thresholds
requires an explicit documented revision, never silent relaxation.

Rollback the batch immediately, without D1, if any of these occurs:

- an unexpected query fingerprint or caller appears;
- a request reaches the statement ceiling;
- a request examines more than 10,000 D1 rows unless that exact bounded
  maintenance request was declared in advance;
- a five-minute bucket exceeds the existing checker's allowed baseline/action
  envelope, or 100,000 total reads (not merely incremental reads);
- unchanged input writes D1 or R2, dispatches processing, or republishes a
  shared object; or
- disabling the batch does not stop it before D1.

If a batch fails, close its flags/flows, retain the last valid data, divide
only that batch by capability or source and rerun the smaller halves. Continue
with independent safe batches rather than blocking the whole restoration.

### User checkpoints

Stop for the user only when credentials/access are required, a clinical
interpretation is needed, a visible result cannot be verified otherwise, or a
rollback/production decision is required. Do not stop merely to announce that
an internal test, deployment or analytics wait succeeded.

### Execution discipline and checkpoint

Maintain one short current checkpoint in this document: observed deployment
and flags; exact Flow IDs/states and observation times; completed acceptance
evidence; unresolved failures; next executable action. Link detailed evidence
instead of rereading the entire incident history. Distinguish repository
configuration, deployed configuration, and observed behaviour.

For each failure, inspect the saved Flow definition and actual action outputs,
HTTP result, normalized date/count/hash, published object and reader result.
Identify the first divergence, reproduce it locally where possible, then fix
that stage. Prefer structured definitions/run outputs over repeated designer
navigation when supported. Do not record tokens or contact details in logs.
After a failed fix, use new diagnostic evidence before trying another change.

Run focused meaningful checks for changed paths once; rerun only after relevant
changes or new failures. Reuse passing evidence for unchanged code. Perform
independent local work during telemetry settlement. Avoid repeated unchanged
polling. Combine necessary user visual checks into one concise list.

Use Medium effort for unresolved diagnosis and coupled ingestion/cache changes;
Low is suitable for a bounded correction with a reproduced cause and clear
acceptance checks. Request High only for a specific unresolved architectural or
data-integrity decision. Keep reports concise at every effort level.

## Batch 0 — truthful inventory and deployment hygiene

This batch is control-plane/read-only except for deleting superseded
deployments after their targets are verified.

1. Read back the active Production commit and effective flags.
2. Export or inspect the exact saved definitions and On/Off state of every
   roster/contact flow. Record trigger, filename filter, endpoint and source ID.
3. Inventory Pages Production and Preview deployments, Workers, schedules,
   GitHub workflows and deploy hooks.
4. Confirm every external application endpoint is the stable Production
   domain and no flow references a hashed/deleted deployment.
5. Select one deployment pipeline for the remainder of recovery.
6. Keep the active deployment and one verified predecessor; delete older
   callable deployments. This is hygiene and caller containment, not a claim
   that deployment count itself consumes D1.
7. Correct the feature register to the read-back state.

Gate: relevant callers, active configuration and rollback switches are known.
No application data request is needed. Reuse verified inventory and inspect
deltas. Do not let unrelated housekeeping delay a proven correction.

## Batch 1A — restore current contacts

### MMC/MCH implementation plan — 28 September revision (not authorised for execution)

**Superseding evidence and direction, 28 September evening:** Recover the
successful existing Microsoft-side JSON extraction, not the external-processing
proposal below. User rejects full-workbook transfer/processing elsewhere.
Do not execute the external design checkpoint or require that consent.

Exact historical evidence recovered from Power Automate:

- Flow `Sync MMC clinician contacts`, ID
  `64d9fad7-3462-4b6a-bb6e-95ce1951f4fb`, run
  `08584112987802616168552190919CU17`, 25 September 13:48 AEST (Test succeeded).
- Actual historical chain: SharePoint file-created/modified trigger →
  Run script from SharePoint library → HTTP. Script succeeded in 36.3 seconds;
  HTTP succeeded in 1.3 seconds, starting 13:49:01 AEST. This directly confirms
  Microsoft-side extraction before app ingestion, not a whole-workbook upload.
- Script action workbook is `Shared Documents/Contact lists/shift allocation/SHIFT ALLOCATIONS.xlsx`.
  Script source `me`, drive `b!yrpB7FRy2UGRxEjgaA5VRuUSetwWz9BOkoVIeGisKCBOCWqvE4jbRIG-pt-05Z_d`,
  script ID `013Y6OLOVIBMFTJSDL2ZHK5JXH56GAQ62Z`. Result date is the raw label
  `FRIDAY 25TH SEPTEMBER 2026`, so do not substitute the repository script
  without comparing the saved script/result contract. HTTP inputs/outputs are
  protected and remain protected.
- User-supplied retained JSON has 92 Adult/Paediatric entries for 25 September,
  with normalised date/contact keys. HTTP timing matches the 13:49 R2 receipt.
  Normalisation happens in app code, so the saved object is not assumed to be
  byte-identical to the submitted script result. No private contacts copied here.
- Register records replacement of the script actions on 26 September after a
  DDH self-trigger loop; it does not establish an MMC loop. Current definitions
  of the other two MMC flows are not evidence of this run's implementation.

Next bounded work: recover this exact saved script/HTTP contract, retain the
deployed small-JSON ingress and semantic no-op guards, and prepare restoration
of this flow instead of a new parser/job. Compare workbook identity and size
with the historical trigger before claiming the documented connector limit
precludes the demonstrated method. Its published limit remains a support risk,
not evidence that this successful run did not happen. Verify one controlled
extraction and unchanged repeat under existing quota admission, and observe
whether MMC actually self-triggers. Automatic recursion must be bounded before
leaving it unattended; if observed, resolve only that trigger behaviour rather
than replace the extraction/data path. Do not change working DDH or VHH.

This section supersedes the MMC whole-workbook implementation in
`contact-workbook-safe-automation-plan.md`, not DDH's working implementation.
Do not change flows, deploy, or execute this revision until the user resumes.

**Established constraints:** MMC workbook is 33,919,328 bytes and grows with
unrelated lost-property content. Our Worker ingress rejects above 5 MiB.
Microsoft documents a 25 MB maximum Excel Online (Business) workbook:
https://learn.microsoft.com/en-us/connectors/excelonlinebusiness/ . Therefore
restoring Run script against this workbook is not a supported long-term fix.
Do not trial it in Production or silently raise either limit. Small JSON output
does not remove the connector's input-workbook limit. Prior successful runs do
not prove support at its current size.

#### Medium design checkpoint — external extraction recommended, consent pending

Repository inspection confirms the existing GitHub-hosted Node 22 roster job
downloads source bytes from an authenticated Cloudflare route and processes them
outside the Worker. `extractMmcDoctorContactsFromWorkbook` already implements
the required extraction independently of Office Scripts. Reuse that function
and the existing JSON publisher; do not send contacts through roster ingestion.
The old `contact-list-binary` endpoint is deliberately 410 and must stay closed.

Recommend a source-specific contact job following that existing processing
pattern, rather than requiring roster writers to maintain another workbook.
This is NOT merely re-enabling an old flow. It needs a bounded transport/job
adapter and resource proof. The recommended data route is:

SharePoint read-only file retrieval → authenticated streaming staging into
private R2 → existing GitHub runner platform with a contact-specific job →
existing bounded MMC parser → small JSON through existing contact ingress →
existing R2 overlays. Never log/store the complete workbook in GitHub artifacts,
repository, Actions caches or D1. Do not hand a public/signed workbook URL to
arbitrary clients. Only the exact source and revision may be fetched by the job.

**Approval required before implementation or live transfer:** the complete
workbook, including unrelated lost-property records, would temporarily be held
in the account's private Cloudflare R2 and processed on a GitHub-hosted runner.
Existing roster authorisation is not treated as approval for those additional
contents. Confirm the user is authorised to permit this. If not, stop here;
contacts must be separated inside the approved Microsoft environment with the
workbook owner's assistance. No new paid service or tenant permission may be
introduced without explicit approval.

After consent, perform these bounded design/proof steps before Low handoff:

- Confirm available free GitHub Actions capacity and R2 capacity; do not enable
  paid overage. Contact jobs must not block the existing roster concurrency group.
- Specify a finite upload cap (initial proposal 64 MiB, not automatic growth),
  streaming raw transport rather than JSON/base64 buffering, exact source/file
  checks, short job timeout, no blind retries, and explicit stale-revision rejection.
- Staging remains private; delete on completion/failure, with a verified expiry
  backstop no longer than 24 hours. No claim of expiry merely from an object
  timestamp; confirm a real lifecycle mechanism or stop. Recovery cannot require
  retaining unrelated workbook contents indefinitely.
- Coalesce trigger revisions and prevent duplicate dispatch before expensive
  extraction using reviewed atomic state; do not invent another D1 polling queue.
  Write down the chosen lock/state mechanism and failure recovery before coding.
- Parser currently reads the ZIP directory and all shared strings into memory;
  reading only A1:I42 does NOT bound decompression. Extend that existing parser
  with explicit entry/bounds validation, decompressed-byte ceilings and measured
  peak memory. Reuse it; no parallel parser. Validate semantic parity with existing
  fixtures and synthetic large irrelevant content; keep actual private data out.
- Contact submission remains compact and idempotent. Trace measured R2, Actions
  and D1 costs separately, including duplicate/failed/stale requests. Stage and
  run a single exact source only; DDH and VHH unchanged.

Medium remains appropriate for the atomic state, security and resource proof.
Low can handle prescribed fixture additions, documentation and mechanical
configuration once these gates are settled. Do not label this integration fully
Low-ready or guarantee a usage allowance. Next action is the data-transfer
approval question, not live execution or additional broad investigation.

1. **Resolve one architecture decision, on Medium.** Preferred only if the
   workbook owner can maintain it: a contacts-only operational workbook with
   the existing layout/date and bounded size, retaining lost-property records
   separately. This requires explicit owner agreement and must not create a
   manually maintained duplicate or silently change roster-writer workflows.
   Otherwise design a bounded external extractor using the existing roster
   processing infrastructure, not a new parser/service by default. Confirm
   permissions, existing free capacity, credential handling, and approval for
   transferring the full workbook (including unrelated contents) before choosing
   storage/processing outside its current location. Do not assume GitHub/R2
   authorisation for these additional data just because roster jobs use them.
   Present that concrete choice once; no series of speculative flow trials.
2. **Reuse contracts and parsing.** Preserve `mmc-shift-allocations`, existing
   MMC/MCH range/role/date semantics, existing JSON ingress and R2 publication.
   Reuse `mmc-contact-allocations-office-script.ts` for the supported small-file
   path; reuse the existing workbook parser for an approved external processor.
   No new D1 tables, full-roster queries, migrations or duplicate contact helpers.
3. **Prevent recursion before enabling.** External byte retrieval must never
   open/edit the source through Excel. For an Office Script small-file path,
   use a separately agreed bounded schedule rather than a modification trigger
   on the script-opened workbook; do not rely on debounce or Modified By alone.
   Select cadence and its request budget explicitly, and preserve concurrent
   human edits. Do not create the schedule during planning. Existing blanket
   schedule/script prohibitions may be replaced only by this reviewed design.
4. **Bound costs and ordering.** Exact source filter; coalesce autosaves;
   reject stale revisions before publication; finite execution time/input size,
   bounded decompression/memory for external extraction; no blind retries.
   JSON <=512 KiB; no private content in logs, Git or artifacts. Source-specific
   rollback. Unchanged clinical data writes zero D1/R2; duplicate triggers must
   not repeatedly launch expensive extraction. Record measured read/write cost,
   including indexes, separately from SQL statement counts.
5. **Focused local proof.** Reuse contact fixtures to cover Adult/Paeds AM/PM/
   Night, nursing exclusions, blank names, date rollover, unchanged replay,
   corrected allocation and stale/out-of-order input. For external extraction,
   synthetic large irrelevant content must not change output and must stay
   within measured resource bounds. Do not copy lost-property data into tests.
6. **One release and one user check.** Fresh existing account admission;
   snapshot/export the affected flow definition privately; disable only the
   failing MMC whole-workbook sender before activating its replacement. Keep
   DDH working and original legacy/Preview flows Off. One current-source run,
   verify JSON/date/count, request costs, published MMC/MCH objects and readers,
   then one combined visual test. Verify next automatic update and unchanged
   replay without repeated user edit cycles. One settled usage reconciliation.
7. **Stop/rollback.** Unexpected source modification, recursive runs, wrong
   date/source, resource overrun or existing quota gate failure closes only the
   new MMC sender; retain valid cached contacts. Do not reopen legacy ingestion.

Effort handoff: Medium for step 1 and any unresolved concurrency/resource design.
Low is appropriate for prescribed configuration, narrow code edits and focused
test execution only after those decisions and acceptance bounds are written.
Pause and request Medium if new architecture questions appear; do not improvise
on Low. No claim that an effort level guarantees completion within allowance.

1. Trace DDH's latest relevant run from trigger to visible current-period
   contacts using the diagnostic sequence above. Absence of contacts invalidates
   the end-to-end completion claim; it does not by itself identify the failed
   stage or prove that no extract was stored.
2. Verify the known MMC/MCH configuration mismatch against deployed settings:
   ingress admits `mmc-shift-allocations`, but committed
   `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST=ddh` makes
   `contactPublicationEnabled()` return false for MMC. Correct contact
   publication scope without opening metadata/day builders.
3. Check transport decoding, operational date/expiry, populated contact count,
   semantic replay and missing-overlay recovery locally. An unchanged complete
   publication writes zero D1/R2; recovery of a missing derived overlay is a
   separately identified bounded repair, not an ordinary unchanged refresh.
4. Exercise DDH and MMC/MCH as one ready release with source-specific rollback.
   Use current source retrieval where possible; do not require repeated user
   edit-and-undo cycles when existing run evidence or a controlled replay is
   sufficient.

Gate: correct current-period contacts visibly reach DDH, MMC and MCH; unchanged
complete publication is a no-op; retrieval causes no recursive source edits;
visible refresh is zero-D1. A green Flow run alone cannot pass this gate.

## Batch 1B — roster automation and upcoming-term readiness

Proceed independently of contact failures.

1. Verify each source's actual recurring trigger, exact filename/folder routing,
   endpoint and most recent accepted provider revision. DDH scheduling belongs
   here: if missing, implement one source-specific bounded coalesced schedule.
   An enabled endpoint or successful manual refresh is not automatic syncing.
2. Test each ingestion class locally with a representative full new term,
   correction, unchanged replay, filename/version replacement and interrupted
   retry. Include supersession, unrelated contributions, current-term retention,
   14-day pre-term visibility, and personal/feed cache revision propagation.
3. Trace the configured 1,250 incremental-fact ceiling through staging and
   promotion. Prove a realistically sized new term completes within bounded
   resumable work; do not simply increase/remove the guard. If unsupported,
   implement bounded chunking with atomic activation and retry deduplication.
4. Verify Production source freshness and existing changed-run evidence; use
   one budgeted current-revision replay only where evidence is missing. Never
   send synthetic shifts to Production. Do not wait for a roster writer's next
   edit to progress: local changed/new-term evidence and current Production
   delivery evidence establish readiness. Record a later genuine edit as
   additional confirmation, without claiming it has already happened.

Gate: all four sources have working recurring triggers and current delivery
evidence; representative new-term ingestion, feeds and browser refresh pass
locally. Clearly label any live new-term observation still pending.

## Batch 1C — keep At a glance current

Release separately if this needs more work than 1A/1B.

1. Trace successful ingestion through compact facts to metadata/day publication;
   prove which disabled controls prevent freshness before changing them.
2. Publish only affected facility/date/term dependencies with bounded work and
   coalesced revisions. Test source corrections, deletion, a new term and
   duplicate delivery. Preserve the previous valid revision on failure and
   expose a clear stale/updating status with a bounded retry path.
3. Retain protected facility access, SMS continuity, non-SMS term membership,
   locum access and the 14-day pre-term rule. Never revive broad event scans.

Gate: changed roster facts update the affected shared views, unchanged imports
perform no publication work, and failures retain clearly dated valid data.

Expected normal costs remain those already demonstrated: unchanged ingestion
does not write; an authenticated cached At a glance request uses a few indexed
rows; visible contact refresh is R2/session-only with zero D1; contact
publication changes only a small contact source/overlay.

First operational milestone: current contacts visible, roster triggers verified
and new-term ingestion proven locally. Each completed release stays enabled
while remaining work continues.

## Batch 2 — user-facing feature restoration

Group ready items behind independent flags. If colleague insights require a
rewrite, release the other ready items first; never flip the old insight path
on for a trial. Items already live require focused regression evidence only:

1. automatic At a glance launch for entitled users;
2. visible-tab-only 60-second contact refresh and immediate refresh on reopen;
3. explicit colleague insight actions rebuilt on compact daily presence;
4. Creator contact corrections on the small contact overlay; and
5. the current multi-site access policy: SMS/CMO see all enabled sites; other
   grades see current-term membership sites plus genuine locum sites only.

The four cached At a glance tabs and all four facility selectors are already
enabled and remain part of this batch's regression check. No feature may fall
back to legacy event-history SQL when an artifact is absent.

Gate: one Creator and representative all-site and site-scoped accounts complete
the user journeys; hidden tabs issue no polls; repeated unchanged use stays
inside the declared request costs.

## Batch 3 — Creator and account workflows

Group independently ready items into a release with separate rollback flags.
Do not let identity/discovery implementation block a finished session-setting
or manual-import correction:

1. ordinary manual roster upload/replacement/removal, excluding bulk repair;
2. compact Creator directory enrichment;
3. bounded doctor discovery/claim suggestions generated when identities
   change, not on every login;
4. a dedicated one-record API for cross-device session settings; and
5. explicit bounded account repair, never an ordinary-login side effect.

Gate: unchanged saves/imports write zero; manual roster changes affect only the
declared source scope; ordinary and Creator login remain cache-first; switching
does not query the directory again.

## Batch 4 — operational completion

Contact-flow cleanup (explicit user request, 28 September): after the app is
fully functional, inventory existing flows and their callers; retain verified
working flows, identify duplicates/obsolete replacements, export recoverable
definitions and obtain deletion approval before removing unwanted flows.
Include `Sync MMC shift allocations`, `Bootstrap MMC shift allocations`,
`Sync MMC clinician contacts`, `Sync DDH clinician contacts` and
`VHH Shift Phone Allocations to Preview`. Do not delete during diagnosis.

1. Confirm the source schedules established in Batch 1B remain healthy.
2. Keep bootstrap and advanced repair normally Off; document the exact
   time-limited runbook rather than enabling them permanently.
3. Restore a watchdog only if measured operations show it is necessary after
   the source triggers are working.
4. Run the final end-to-end matrix and reconcile one settled account-wide
   sample.
5. Update every register entry to Restored, Replaced permanently, or Deliberate
   future work. Remove obsolete ordering statements from older plans.

Gate: the Production experience matches 31 August user functionality, except
where an unsafe mechanism has been replaced by an equivalent bounded one.

## Completion and next work

Service restoration is complete only when automatic roster and supported
contact updates are proven end-to-end, personal calendars and subscriptions
update, At a glance remains current, all intended user and Creator actions are
available, and every unchanged path is a no-op at the data layer.

After that point, rebase and review the durable Doctor Names branch against the
stable Production baseline. It must not be mixed into the restoration batches.

## Current checkpoint

### 28 September morning release admission

Original-flow review after renewed Safari access (read-only; no saves):

- `Sync MMC shift allocations` (`114b1e0b-071e-4f44-9bba-5d56fa9d5d25`)
  is Off, modified 10 September. Current chain is SharePoint modified trigger
  → Get file content → HTTP. Its JSON contains the complete workbook as
  `contentBase64`, not a contact extract. Destination:
  `https://4a12dbe2.roster-to-calendar.pages.dev/api/automation/contact-list`.
  Do not re-enable unchanged. Earlier flow versions were not established.
- `VHH Shift Phone Allocations to Preview`
  (`e8b16ab2-978a-44b1-b331-070f17504b14`) is Off. Chain is SharePoint modified
  trigger → Run script from SharePoint library → HTTP. Workbook:
  `/CONTACTS/Shift phone allocations.xlsx`; script:
  `/Documents/Office Scripts/VHH Extract Shift Phone Allocations.osts`.
  HTTP sends `outputs('Run_script_from_SharePoint_library')?['body/result']` to
  `https://vhh.roster-to-calendar.pages.dev/api/automation/contact-list-extract`.
  This is the existing small-JSON extraction pattern. Script internals,
  historical success and Production compatibility remain unverified.
- Correction: VHH has an external extraction script; absence of a repository
  parser does not establish absence of a flow solution. Reuse the existing MMC
  bounded script and JSON endpoint. Resolve Excel-trigger recursion before
  enabling the automatic path; do not create another parser or blindly enlarge
  the workbook upload ceiling. Both reviewed flows remain Off and unchanged.

- User identifies original small-JSON flows as the intended reusable solution.
  Shared-flow listing confirms `Sync MMC shift allocations` and
  `VHH Shift Phone Allocations to Preview` still exist and are disabled.
  Their internal actions/history are NOT yet inspected: Power Automate SSO
  repeatedly requires sign-in when opening details.
- Existing `scripts/mmc-contact-allocations-office-script.ts` reads only
  `SHIFT ALLOCATIONS!A1:I42` and returns Adult/Paediatric doctor JSON. Existing
  `/api/automation/contact-list-extract` still ingests that format and uses
  semantic deduplication/publication. Reuse these rather than add another parser.
- Register's 26 September entry says the Office Script actions were removed
  from the newer clinician-contact flows, not the old flows deleted. The older
  workbook plan cites DDH self-retriggering as the reason; that is not proof
  that the original MMC flow had the same behaviour. Review original MMC
  trigger, script, payload, destination and run frequency, then choose a bounded
  non-recursive JSON path. Do not blindly raise the workbook size ceiling or
  re-enable an unreviewed old flow. Revise the older blanket Office Script
  prohibition if the reviewed reuse strategy supplies equivalent safety.

- 12:22 user edited/reverted both contact workbooks. MMC first runs were
  superseded during the settling delay. Latest run
  `08584110445941620923997290915CU25` retrieved the workbook, then HTTP failed
  at 12:26:55 AEST with 413 `Contact workbook is too large.` SharePoint metadata
  reports **33,919,328 bytes**, versus the endpoint's 5 MiB cap. This is a
  confirmed additional blocker, not a D1 quota failure. Do not raise the cap
  blindly: the full workbook plus decoding/unzip memory needs an alternative
  bounded extraction path or measured resource proof. HTTP inputs stay protected.
  The MMC manifest still has only 25 September. DDH's current manifest remains
  the already-verified 28 September publication from 08:42; an unchanged
  edit/revert need not update its publication timestamp.

- 12:18 AEST second sample GO: 57,840 reads, 302 writes, no expensive
  fingerprints, settled through 12:03. Enabled `Sync MMC clinician contacts`
  (`64d9fad7-3462-4b6a-bb6e-95ce1951f4fb`); UI status shows On. This is the
  shared MMC/MCH source. R2 manifest still contains only 25 September, so a
  fresh workbook-triggered publication is required before today's visual test.
  Do not report end-to-end contact restoration merely from enabling the flow.

- Later checkpoint: Git push succeeded. Canonical Production deployment
  `3aa83a00-b4bb-4c2f-a28c-16533bffd4fc` serves `6737969b`, deployed successfully
  at 11:50 AEST; control-plane readback confirms `ddh,mmc,mch` publication.
  MMC flow has not yet been enabled.
- GitHub MCH queue run `36353791313` at 08:01 parsed 80 doctors and 2,377
  events, coinciding with the write burst. This supports roster-ingestion
  attribution but does not prove why every row changed or reconcile all writes.
- Post-reset sample at 11:49 AEST: 43,079 reads, 214 writes. STOP now only for
  two-hour baseline and missing second sample. Reuse
  `/private/tmp/contact-sep28-baseline.json` for a second sample at/after noon
  AEST (10-minute minimum separation); do not restart the baseline.

- 09:06 AEST user confirmed DDH contact display matches the source sheet.
- Fresh 09:07 sample (settled through 08:52) is STOP: 226,369 reads and
  26,238 writes. The 08:00–08:05 AEST bucket contains 25,823 writes. Sampled
  fingerprints attribute 13,845 writes to `roster_daily_presence` inserts;
  contact-list inserts account for 84 across the UTC day. Fingerprint totals
  undercount account writes by 11,928, so the full burst/source is not yet
  attributed. No expensive-read fingerprint was flagged. Do not blame the
  user's DDH visual check or claim full attribution from this evidence.
- MMC flow must remain Off until admission passes. The publication fix remains
  committed locally; GitHub connectivity blocked the first push and the retry
  is being checked. No manual alternate deployment should duplicate it.

- Account analytics at 08:14 AEST, settled through 07:59 AEST: 162,444 reads,
  225 writes for the UTC quota day; projected reads approximately 181,000;
  no expensive fingerprints. Checker GO against the previous valid sample,
  using the documented `free-plan-dashboard-omits-d1` billing mode.
- Production readback remains `baec779b-db7a-43a3-b5bc-53a1ca8cd17e` with
  publication allowlist `ddh`. The workbook-ingress regression and diff check
  passed again. User authorised commit/push; release changes only the contact
  publication allowlist plus tests/documentation. No roster/day builders opened.
- After deployment verification: restore/check the MMC contact flow and obtain
  one current workbook submission. Then one combined DDH/MMC/MCH On shift
  check, against contact names actually present in today's source sheets;
  verify AM/PM as applicable. A missing date's source is not a reader failure.
- Do not advance to the next restoration batch before reporting this result.
- Release committed locally as `6737969b`. Push failed because github.com:443
  was unreachable; an independent short IPv4 check also timed out. Production
  has NOT changed. MMC flow remains Off (Flow checker: zero errors, only the
  Off warning). Resume by pushing this existing commit, not making another
  implementation or deploying a duplicate manually. Verify deployment before
  enabling MMC or requesting its end-to-end test.

- Active Production readback: `baec779b-db7a-43a3-b5bc-53a1ca8cd17e`, commit
  `3161b23283f003b42216c6370022a7d67d127a26`. Git automatic Production deployment
  is enabled; avoid a second manual deployment of the same change.
- Power Automate: DDH contact Flow `67ee065a-c230-4d1d-9a20-d9425f9cacdf` On;
  MMC contact Flow `64d9fad7-3462-4b6a-bb6e-95ce1951f4fb` Off. Monash Adults,
  Monash Paediatrics and automated VHH roster flows show Enabled. DDH scheduling
  remains to be verified under 1B.
- DDH run `08584110963160113529091988067CU09`: HTTP succeeded at 22:04:50 AEST.
  HTTP inputs/outputs are protected; protections retained. Retained current-day
  R2 revisions start at 14:26:06 AEST, after the user's 13:59 check. This
  supports absence of a current publication at that earlier test; no earlier
  manifest history was available to prove the exact reader response. Current
  manifest points to the 18:53 revision with 11 populated AM contacts and 14 PM
  contacts. The local reader loads it as available with 25 populated phone
  contacts. Browser matching/display still needs one combined visual check.
- Analytics credential and Wrangler login work. At 12:26 UTC the settled
  sample through 12:11 UTC showed 59,373 reads, 44 writes, maximum five-minute
  reads 3,615 and no expensive fingerprint. Valid sample; checker STOP only
  for missing second sample. This is not deployment admission.
- MMC/MCH defect reproduced using tracked Production flags: ingestion succeeds
  but publication excludes MMC/MCH. Local correction expands the publication
  list to `ddh,mmc,mch`; metadata/day builders, advanced maintenance and Preview
  remain closed. Not deployed.
- Regression exercises workbook ingress through published MMC/MCH/DDH readers,
  existing-source overlay publication and zero-write subsequent replay. It
  failed before correction and passed afterward. Focused ingress, rollout,
  contact-sync and workbook suites passed.
- No Production D1 query, flow edit, deployment, push or deployment deletion was
  performed. Diagnostics used analytics/control-plane reads and specific R2
  objects. Contact details were not printed or added to Git.
- Next: fresh budget admission, one release of the tested configuration,
  controlled current MMC extract publication and routine Flow restoration,
  followed by a combined DDH/MMC/MCH visual check. Reuse passing evidence;
  do not repeat transport experiments or restart the full inventory.

### 1 October — roster overwrite safety release

- Released commit `098bc2a2` through the existing Git-triggered Production
  deployment `20a047ff-55a1-46b6-8dce-f85cb3e2f700`; Cloudflare lists it Active
  on main. The canonical derived endpoint returns the expected unauthenticated
  401 without any sync or D1 work. No synthetic Production roster was submitted.
- Complete ingestion selects the matching retained term using compact indexed
  coverage (32 active-file ceiling), gives disjoint new terms independent files,
  and preserves the latest-term pointer on older-term corrections. Unprepared,
  ambiguous or partially overlapping term ranges fail safely.
- Explicit 1,250-fact guard also covers initial imports. Large full-term
  ingestion remains blocked pending bounded staging/chunking; Batch 1B is open.
- Duplicate completion handles remapped file ids; failed callbacks cannot delete
  completed runs. Cleanup preserves active files following a bookkeeping failure.
- Local endpoint/materialisation regressions cover term retention, 14-day
  visibility, corrections, duplicate zero-write callbacks, oversized imports,
  overlap rejection, indexed lookup and interrupted-bookkeeping retry. Existing
  ingress-idempotency, source-isolation and queue-failure checks also pass.
- D1 release admission at 02:13 UTC: GO, 19,414 rows read / 36 written account-wide.
  Contact flows, contact code and deployed flags remain unchanged.
- Final Codex usage check: 3% weekly and 84% five-hour remaining; one free reset
  available and credit balance unchanged at 182.18727.


### 1 October — combined batches 1B / 1C and automatic launch

User asked to enact large safe batches and avoid micro-change approval cycles.
Bounded next-term importing, durable continuation, DDH LastModified checks,
affected-date overview publication and cached automatic launch are implemented
and validated together. Existing contacts remain unchanged. The daily maintenance
allowance includes indexed writes and read reservations; publication refunds reads
only when complete billing metadata supports it. Empty progress migrations are
applied. External automatic roster flows for MMC/MCH/VHH are enabled.

The [current checkpoint](./bounded-roster-import-checkpoint.md) contains release
limits, test evidence, remaining work and rollback. Next-term completion must be
read back after actual provider publication; this batch does not claim otherwise.
