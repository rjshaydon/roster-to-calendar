# Terra execution plan: VHH roster and contact-list JSON integration

## Current implementation scope

This document originally described a discovery-only exercise. That restriction
is now superseded for the **VHH branch and Cloudflare Preview environment**.
Implement the VHH JSON pipeline below; do not deploy, configure, or write any
VHH data to Production.

The SharePoint workbooks are read-only source material. Do not edit, save,
rename, copy, move, upload, or replace either workbook. Power Automate may use
the SharePoint modified-file trigger and Excel **Run script from SharePoint
library** only. Both Office Scripts are read-only and must have no write calls.

### Required live path

```text
VHH SharePoint workbook → Power Automate → VHH Office Script → Preview HTTPS ingress
  → Preview R2 retained JSON → Preview D1 queued run → GitHub Actions parser → Preview D1 calendar
```

Use the repository artifacts already created for this path:

- `scripts/vhh-roster-office-script.ts` extracts every visible `Shift Label`
  block in `Active Medical Roster`, stops before `JMS Teaching Timetable`, and
  retains only dates, shift labels, and exact roster-cell text.
- `scripts/vhh-contact-allocations-office-script.ts` extracts only CIC from
  `Zebra Allocations!D4:E4` and the rows strictly between `DOCTORS` and
  `CLERICAL` from columns C:E.
- `functions/api/automation/vhh-roster-extract.js` is the JSON-only roster
  ingress. It stores no workbook bytes.
- `/api/automation/contact-list-extract` stores the VHH CIC/doctor directory
  JSON alongside MMC/DDH extracts, without yet representing it as an On shift
  allocation.

### Power Automate configuration

Create two VHH-only flows in the Monash tenant after the Preview deployment is
available. Start each manually, verify its response, then enable its file
modified trigger. Set trigger concurrency to one and use the SharePoint
filename and ETag/version, not the flow-run time, as the change identity.

1. **VHH Active Medical Roster → Preview**
   - Trigger: SharePoint **When a file is created or modified (properties
     only)** for the VHH Cardiac Emergency / Medical channel library/folder.
   - Guard: filename exactly `Active Medical Roster.xlsx`.
   - Action: **Run script from SharePoint library** against the triggering
     workbook, using `vhh-roster-office-script.ts`.
   - HTTP POST: `https://<VHH Preview Pages URL>/api/automation/vhh-roster-extract`.
   - Header: `Authorization: Bearer <Preview ROSTER_AUTOMATION_TOKEN>` and
     `Content-Type: application/json`.
   - Body: the script `blocks`, plus `sourceId` `vhh-active-medical-roster`,
     SharePoint `Modified`, and SharePoint ETag/version. Do not send file
     content, Base64, or an Excel attachment.

2. **VHH Shift Phone Allocations → Preview**
   - Trigger: SharePoint **When a file is created or modified (properties
     only)** for the VHH Cardiac Emergency / General channel library/folder.
   - Guard: filename exactly `Shift Phone Allocations.xlsx`.
   - Action: **Run script from SharePoint library** against the triggering
     workbook, using `vhh-contact-allocations-office-script.ts`.
   - HTTP POST: `https://<VHH Preview Pages URL>/api/automation/contact-list-extract`.
   - Header: as above.
   - Body: script result `cic` and `doctors`, `sourceId`
     `vhh-shift-phone-allocations`, SharePoint `Modified` converted to a
     Melbourne `YYYY-MM-DD` `sourceDate`, and SharePoint ETag/version.

### Roster interpretation rules implemented now

- `Last, First` is normalised to `First Last` before the person key is built.
- Every name in an extracted roster cell becomes a staff record; raw cell text
  remains on the event for audit.
- Explicit time ranges such as `0800-1530` or `08:00-15:30` become timed
  calendar events. A multi-range cell creates one event for each stated range.
- A cell without stated hours remains an all-day calendar entry; it is excluded
  from VHH **On shift** until the default timings are confirmed. No times are
  invented from a shift label.
- `JMS Teaching Timetable` is never extracted or converted to an event. Roster
  rows named `AM JMS`, `PM JMS`, or `ON JMS` remain roster shifts.

### Validation and release gates

1. Use the supplied local workbooks only as ground truth; never commit their
   contents or received VHH JSON.
2. Confirm the Preview roster response queues, dispatches the VHH branch
   workflow, and reports a successful D1 derived-save run.
3. Compare Preview doctor count, date range, and representative events against
   the supplied roster. Verify name order and all explicit time ranges.
4. Confirm the VHH contact extract stored in Preview R2/D1 contains only CIC
   and the requested doctor rows. Do not match those contacts to staff or
   expose them in On shift yet.
5. Stop before any Production deployment, Production token, Production R2/D1
   write, or Production Power Automate endpoint is used.

The historical discovery notes below explain why this architecture was chosen.
Where they conflict with the current implementation scope, the current scope
above takes precedence.

## Objective

Retrieve structure-preserving JSON from the VHH roster and VHH contact list using the existing SharePoint and Power Automate security boundary. After both JSON captures have been verified, park the contact-list work and examine the roster JSON beside the supplied four-sheet Excel workbook.

This phase is discovery only. It must not add VHH to the live calendar, parse VHH shifts into events, expose VHH contacts in On shift, deploy anything, or write production D1/R2 data.

## Confirmed existing patterns

### MMC and MCH roster retrieval

The current Monash roster flow is documented in [`roster-automation.md`](./roster-automation.md):

1. Power Automate uses a SharePoint file-created-or-modified trigger.
2. It filters to approved roster filenames.
3. It gets the complete file content.
4. It sends a JSON envelope to `/api/automation/ingest`. The envelope contains `sourceId`, filename, content type, Base64 workbook content, SharePoint modification time, and the SharePoint ETag/version.
5. The ingress validates the bearer token and source, rejects files over 15 MB, deduplicates first by filename plus provider version and then by content hash, retains the original workbook in R2, and creates a queued D1 sync run.
6. A GitHub Actions worker downloads the retained workbook, parses it outside the Cloudflare request CPU limit, posts derived doctors/events in batches, and activates the staged roster only after the final save succeeds.

The configured automated roster sources are currently only:

- `monash-adults` → `mmc`
- `monash-paeds` → `mch`
- `dandenong-findmyshift` → `ddh`

VHH is not an automation source and is not a recognized parser source. `parseUploadForm` currently creates only `mmc`, `ddh`, `casey`, and `mch` source collections. Sending a VHH workbook to the current roster ingress would therefore be rejected or fail type detection. Do not use `/api/automation/ingest` for this discovery phase.

Relevant files:

- [`docs/roster-automation.md`](./roster-automation.md)
- [`functions/api/automation/ingest.js`](../functions/api/automation/ingest.js)
- [`functions/_lib/automation-import.js`](../functions/_lib/automation-import.js)
- [`scripts/process-roster-queue.mjs`](../scripts/process-roster-queue.mjs)
- [`.github/workflows/monash-roster-sync.yml`](../.github/workflows/monash-roster-sync.yml)
- [`public/static/roster.js`](../public/static/roster.js)

### MMC and DDH contact-list retrieval

The live contact feeds no longer upload complete workbooks. Power Automate runs a doctors/clinicians-only Office Script inside Microsoft 365 and sends the script result as small JSON to `/api/automation/contact-list-extract`.

The MMC script:

- Reads `SHIFT ALLOCATIONS!A1:I42`.
- Derives the roster date from `D2`.
- Reads Adult and Paediatric sections in AM, PM, and Night column groups.
- Excludes nursing roles.
- Returns `sourceId`, `sourceDate`, and a `contacts` array.
- Relies on Power Automate to add SharePoint modification/version metadata to the HTTP body.

The DDH script:

- Reads `ED Clinicians!A1:N120`.
- Treats `A:D`, `F:I`, and `K:N` as AM, PM, and Night blocks.
- Accepts the Melbourne source date and SharePoint metadata as script parameters.
- Handles continuation rows and the DDH-specific precedence of phone data across name, EMR, and number cells.
- Returns `sourceId`, `sourceDate`, provider metadata, and a `contacts` array.

The receiving endpoint accepts only the known MMC and DDH source/area combinations, up to 240 contact rows and a 512 KB request. It normalizes the operational date, validates shifts and roles, content-hash deduplicates the JSON, stores it in R2, and records metadata in D1. The former complete-workbook contact endpoints intentionally return `410 Gone`.

Relevant files:

- [`scripts/mmc-contact-allocations-office-script.ts`](../scripts/mmc-contact-allocations-office-script.ts)
- [`scripts/ddh-contact-clinicians-office-script.ts`](../scripts/ddh-contact-clinicians-office-script.ts)
- [`docs/mmc-contact-allocation-automation.md`](./mmc-contact-allocation-automation.md)
- [`functions/api/automation/contact-list-extract.js`](../functions/api/automation/contact-list-extract.js)
- [`public/static/contact-allocations.js`](../public/static/contact-allocations.js)
- [`functions/_lib/contact-list-workbook.js`](../functions/_lib/contact-list-workbook.js)
- [`scripts/test-contact-list-workbook.mjs`](../scripts/test-contact-list-workbook.mjs)
- [`scripts/test-contact-sync.mjs`](../scripts/test-contact-sync.mjs)

The repository documents the MMC Power Automate request in detail but does not contain an equivalent DDH flow guide. Before copying either pattern, Terra must inspect the actual current DDH and MMC flow definitions or user-supplied screenshots/export and record any configuration that exists only in Power Automate.

## Discovery design decision

Use separate, manually triggered discovery flows for the VHH roster and VHH contact list. Run Office Scripts against SharePoint workbooks and retain the returned JSON inside a restricted Microsoft 365 location or export it manually for local examination.

Do not initially post either result to the application's existing automation endpoints:

- The roster endpoint expects a supported workbook that can immediately enter the parsing queue.
- The contact endpoint rejects unknown source IDs and areas and is designed for live, normalized contact allocations rather than raw structural discovery.
- A temporary public or weakly protected capture endpoint would unnecessarily widen the exposure of roster and contact information.

Office Scripts run by Power Automate cannot make external HTTP calls themselves; the flow must handle any later HTTP step. Microsoft currently limits Office Scripts requests/responses to 5 MB and synchronous Power Automate runs to 120 seconds. Design the roster capture as a manifest call followed by one sheet per call, with row-chunk fallback, rather than returning all four sheets in one result.

## Required inputs before Terra runs retrieval

Obtain these from the user or the VHH SharePoint owner:

- The current VHH roster workbook in Excel format, with its four worksheets unchanged.
- The exact SharePoint site, library, folder, and approved roster filename or filename pattern.
- The exact VHH contact-list workbook and relevant SharePoint path.
- Confirmation of which workbook is authoritative if copies exist.
- Read/run permission for the Power Automate connection account and permission to run Office Scripts against both workbooks.
- Current MMC and DDH flow screenshots or exported definitions, especially trigger conditions, concurrency settings, retry settings, metadata expressions, tokens/connections, and failure notifications.
- A restricted SharePoint folder in which temporary discovery JSON may be stored, or an agreed manual method for downloading a flow run's `result` output.

Do not place identifiable roster or contact JSON in Git. Use a local ignored/private working location and create sanitized fixtures later only if the user separately authorizes parser implementation.

## Phase 1 — Confirm the existing automation baseline

Terra must first produce a short baseline note that traces the real flows, not only the repository's intended design.

1. Compare the live MMC/MCH roster flow with `docs/roster-automation.md`.
2. Compare the live MMC and DDH contact flows with their Office Scripts.
3. Record, without copying secret values:
   - Flow owner and connection type.
   - Trigger and filename filters.
   - SharePoint actions used.
   - How `Modified`, ETag/version, filename, and file identifier are obtained.
   - Whether concurrency is restricted to one run.
   - Retry and failure-notification behavior.
   - Where the bearer token is held, if an HTTP step exists.
4. Resolve any difference between documentation and the live flows before creating VHH flows.

Checkpoint: report the baseline and any missing access. Do not change production flows during this phase.

## Phase 2 — Build the VHH roster discovery script and flow

This is a structure capture, not a roster parser.

### 2A. Manifest mode

The first script call should return workbook-level metadata only:

- Discovery schema version.
- Source identifier such as `vhh-roster-discovery`.
- Workbook/provider metadata passed in by Power Automate: filename, SharePoint modification time, and ETag/version.
- Worksheet names in workbook order.
- Worksheet visibility.
- Values-only used-range address and dimensions for each sheet.
- Full used-range address and dimensions where different, so excessive formatting can be identified without dumping it.
- Table and named-range names/addresses if present.

Acceptance: the manifest identifies the expected four worksheets and provides stable dimensions for planning the per-sheet calls.

### 2B. Per-sheet mode

Run the script once for each worksheet name or index. Each result must preserve position and distinguish displayed content from underlying Excel content. Include:

- Sheet name and workbook provider metadata.
- Exact returned range address, starting row/column, row count, and column count.
- A rectangular `texts` matrix from Excel's displayed text.
- Raw `values`, because dates and times may be numeric serials.
- Formulas, so calculated dates or copied timetable values are not mistaken for literals.
- Number formats for interpreting date/time serials.
- Merged-area addresses.
- Hidden-row and hidden-column information where it affects interpretation.

Keep blank cells in their original matrix positions. Do not filter empty rows, collapse columns, normalize names, infer shifts, or separate fortnight/timetable blocks in the extraction script. Those are analysis decisions and must remain auditable against the workbook.

Use displayed `texts` as the primary human comparison layer. Use raw values, formulas, and number formats only to explain ambiguous dates/times and calculated cells.

### 2C. Size and timeout fallback

Treat approximately 4 MB as the operating ceiling so the flow stays below Microsoft's 5 MB response limit.

If a sheet result approaches that ceiling or times out:

1. Return the sheet in deterministic row chunks with explicit `startRow`, `endRow`, and unchanged absolute coordinates.
2. Repeat provider metadata and the used-range identity in every chunk.
3. Reject overlapping, missing, or differently versioned chunks during reassembly.
4. Do not truncate arrays silently.

### 2D. Power Automate roster discovery flow

Create a separate manually triggered VHH roster discovery flow using **Run script from SharePoint library**. Select the workbook through the SharePoint file picker and pass the provider metadata and requested mode/sheet/chunk to the script.

For the first run:

1. Run manifest mode.
2. Run one call per returned worksheet.
3. Save each `result` as a separate JSON artifact in the approved restricted location, using a non-overwriting name that includes source, provider version or modification identity, sheet index/name, and chunk number if applicable.
4. Record the flow run ID and retrieval time in a small manifest, but keep the SharePoint ETag/version as the source-change identity.
5. Fail the run if any expected sheet/chunk is missing; do not label a partial capture complete.

Do not yet enable a file-modified trigger and do not call `/api/automation/ingest`.

## Phase 3 — Build the VHH contact-list discovery script and flow

Follow the contact pattern: perform extraction inside Excel Online and return JSON, without uploading the complete workbook.

Because the VHH contact-list layout is not yet documented, begin with a restricted structural inspection of its relevant worksheet(s). Then define a VHH-specific discovery result that retains enough source context to verify every extracted field. At minimum capture:

- Discovery schema version and a provisional source ID such as `vhh-contact-list-discovery`.
- Source date if the workbook contains one, plus the exact source cell/address used.
- SharePoint modification time and ETag/version supplied by the flow.
- Sheet and source-cell/range coordinates.
- Role or allocation label, person name, phone/extension text, shift/period label, and populated status where those concepts are actually present.
- Raw displayed row values for any record whose interpretation is uncertain.

Do not assume the MMC AM/PM/Night blocks, DDH operational-day rollover, known areas, role exclusions, phone precedence, or 240-row live endpoint contract apply to VHH. Record the workbook evidence first.

Create a separate manually triggered Power Automate flow and save its `result` JSON to the same class of restricted discovery location. Verify that the artifact is valid JSON, carries the correct provider version, and can be tied back to the source workbook.

Checkpoint and stop: report the contact JSON filename, provider identity, worksheet/range coverage, record count, and any obvious omissions. Do not normalize it into the live contact schema, send it to `/api/automation/contact-list-extract`, match contacts to rostered clinicians, or alter the On shift UI. Contact-list work is parked here.

## Phase 4 — Retrieve and validate the discovery artifacts

For both sources:

1. Confirm every JSON file parses successfully.
2. Confirm the source ID, workbook name, provider modification time, and ETag/version are present and consistent.
3. Confirm the roster manifest lists four sheets and that every sheet/chunk belongs to the same workbook version.
4. Confirm each returned matrix has the declared dimensions and retains blank positional cells.
5. Spot-check at least ten distinctive cells per roster sheet against the supplied Excel workbook, including dates, names, shift codes, headings, formulas, and timetable entries.
6. Confirm merged ranges and date/time formatting explain the visible workbook layout where applicable.
7. Record artifact sizes and whether chunking was required.
8. Store a checksum for each local artifact in the private analysis notes so accidental recapture or modification is detectable.

If the workbook changes during capture, discard the mixed-version set and rerun all roster discovery calls against one ETag/version.

## Phase 5 — Examine the VHH roster structure

Analyze the supplied workbook and its JSON side by side. Do not begin with an assumed term model.

### 5A. Map each worksheet

For each of the four worksheets, record:

- Sheet name, visibility, used range, and dimensions.
- The exact row/column boundaries of each fortnightly roster block.
- The exact start and end dates found in each block.
- Whether there are exactly two fourteen-day periods per sheet, and how they are stacked.
- The exact boundaries of each JMS Teaching Timetable and how it is associated with the preceding or following fortnight.
- Header rows, date rows, day-of-week rows, staff-name columns, role/seniority groupings, shift cells, leave cells, comments, totals, legends, and spacer rows.
- Merged cells, hidden rows/columns, formulas, and formatting that carry structural meaning.

The statement that there are two fortnights per worksheet and a JMS timetable beneath each fortnight is a hypothesis to verify, not a parser rule to encode yet.

### 5B. Explain JSON representation

For every structural element, document how it appears in:

- `texts`.
- Raw `values`.
- Formulas.
- Number formats.
- Merged-area metadata.

Pay particular attention to Excel date serials, time-only numbers, formulas that copy dates/headings, merged headings whose text appears only in the top-left cell, and blank cells used to imply continuation.

### 5C. Compare all sheets and periods

Determine which anchors are stable and which vary:

- Row offsets of the first and second fortnight.
- Number and order of staff rows.
- Role/seniority section names.
- Date-column count and ordering.
- Position and dimensions of JMS timetables.
- Shift vocabulary and explicit time formats.
- Leave, education, non-clinical, and free-text annotations.
- Additional notes outside the principal roster blocks.

Do not infer that a layout seen once is universal. Classify each proposed anchor as invariant across all eight supplied fortnights, common but variable, or observed only once.

### 5D. Produce a parser-readiness report

Write `docs/vhh-roster-structure-report.md` containing:

- A worksheet-by-worksheet structural map.
- An annotated example of one roster block and one JMS timetable, using sanitized values where needed.
- A catalogue of date, staff, role, shift, leave, teaching, and annotation fields.
- Stable anchors and known variations.
- Ambiguities requiring the roster writer or user to explain.
- A recommendation for how future VHH parsing should separate clinical roster data from JMS timetable data.
- Suggested sanitized regression fixtures and acceptance cases for a later parser phase.
- A clear list of information that must not become calendar events.

Stop after this report and request user review. Do not add a VHH source type, parser, live automation source, contact source, database rows, or UI behavior in this phase.

## Expected Terra deliverables

1. Existing-flow baseline note, including differences between repository documentation and live Power Automate flows.
2. VHH roster discovery Office Script.
3. VHH contact-list discovery Office Script.
4. Manual Power Automate setup notes for both discovery flows, with secret values omitted.
5. Private roster manifest and per-sheet/chunk JSON captures tied to one provider version.
6. Private VHH contact JSON capture and receipt summary, then no further contact processing.
7. `docs/vhh-roster-structure-report.md`.
8. Final handoff stating files changed, tests/checks run, artifact sizes, data-handling precautions, unresolved questions, and confirmation that no production deployment or data change occurred.

## Acceptance criteria for this phase

- JSON has been retrieved from both VHH workbooks through the approved Microsoft 365/Power Automate connection.
- The VHH contact artifact is validated and then deliberately parked.
- All four roster sheets are captured completely and tied to one SharePoint version.
- The JSON can be reconciled positionally with the supplied Excel workbook.
- All fortnight and JMS timetable boundaries are documented from evidence.
- Variations across the eight apparent fortnights are identified.
- No identifiable JSON is committed to Git.
- No VHH data reaches the live roster/contact endpoints.
- No parser, calendar event, contact allocation, deployment, or production data change is made.

## External platform references

- [Run Office Scripts with Power Automate](https://learn.microsoft.com/en-us/office/dev/scripts/develop/power-automate-integration)
- [Pass data to and from scripts in Power Automate](https://learn.microsoft.com/en-au/office/dev/scripts/develop/power-automate-parameters-returns)
- [Office Scripts platform limits](https://learn.microsoft.com/en-us/office/dev/scripts/testing/platform-limits)
- [ExcelScript Range API](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts)
