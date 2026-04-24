# Roster to Calendar Tool PRD

## Goal

Build a multi-account roster conversion tool that ingests supported hospital roster formats for a selected doctor, preserves each user's workspace between sessions, and exports a single `.ics` calendar file for Apple Calendar.

## Accounts And Persistence

### Account Model

- Every user must create or log into an account using:
  - email address
  - password
- `rhaydon@gmail.com` is the Creator / Admin account.
- All other accounts are standard user accounts.

### Account Behaviour

- Logging into a new email address creates a new empty account.
- Logging into an existing account requires the correct password.
- Logging out must not destroy the account workspace.
- Logging back in must restore the same workspace for that account.
- Switching accounts must load only that account's workspace and must not show another user's imports, edits, or custom events.

### Persisted Workspace

Each account workspace persists:

- imported roster references
- doctor selection
- settings
- conflict selections
- imported-event overrides
- custom events
- preview/export working state

### Storage Model

- Uploaded source files are deduplicated by file identity/content fingerprint and stored once.
- Account workspaces store references to shared source files rather than duplicating the same file for each account.
- If multiple users import the same MMC workbook, the physical source file should exist only once in storage, with multiple account references pointing to it.
- DDH uploads are expected to be per-doctor / my-shifts-only exports to reduce storage.

### Cloud Persistence

- Cross-device persistence requires server-side account storage.
- Account metadata and workspace state should be stored server-side.
- Deduplicated shared source files should be stored in a shared object/blob store, with account workspace records referencing them.
- If cloud bindings are unavailable, the app may fall back to browser-local persistence, but this is not sufficient for the full product goal.

### Retention Limits

- Creator/Admin account: unlimited retained roster history.
- Standard user accounts: maximum active retained roster window of 6 months.
- When a standard user upload pushes the account beyond 6 months from the latest date in the current upload, the app must ask whether to remove events older than that threshold.
- The newer data is the default to keep.

## Supported Sources

1. MMC workbook export
2. DDH FindMyShift Excel export

## Core Workflow

1. Create an account or log in
2. Upload one or more supported roster files
3. Detect source type automatically
4. Detect names present in the relevant sources
5. Auto-select the doctor when there is only one valid doctor
6. Show a dropdown when multiple valid doctors are available
7. Restore previous account workspace state where available
8. Preview and edit the normalized roster
9. Export a single `.ics` file

## Name Matching

- Extract candidate names from both sources
- Match names case-insensitively
- Normalize whitespace and punctuation for matching
- Preserve a readable display version for the UI

## Event Output

- One `.ics` file containing all normalized events across the uploaded date range
- Timezone: `Australia/Melbourne`

### Title Rules

- MMC events: `MMC: <normalized shift>`
- DDH events: `DDH: <normalized shift>`

Examples:

- `0800-1730 AAC` -> `MMC: Amber AM`
- `0800-1730 AAR` -> `MMC: Amber AM Float`
- `1430-0000 PCC` -> `MMC: Clinic PM`
- `1430-0000 PSSR` -> `MMC: SSU PM Float`
- `Clinical Support` -> `DDH: CS`
- `Orange PM (on-call)` -> `DDH: Orange PM`
- `onsite CS` -> `DDH: CS onsite`

## Location Rules

### MMC On-Site

- Label: `MMC Car Park`
- Address: `Tarella Road, Clayton VIC 3168, Australia`

### DDH On-Site

- Label: `DDH Car Park`
- Address: `135 David St, Dandenong VIC 3175, Australia`

### Off-Site

- No location

## Time Rules

If a source contains an explicit time, that time overrides canonical defaults.

### Canonical MMC Defaults

- AM regular: `08:00-17:30`
- AM SSU: `07:30-17:30`
- PM any: `14:30-00:00` next day

### Canonical DDH Defaults

- AM regular: `08:00-18:00`
- AM SSU: `07:30-17:30`
- AVAO AM: `07:30-17:30`
- PM regular: `15:00-00:00` next day
- AVAO PM: `14:30-00:00` next day

### All-Day Special Cases

- MMC `CS`: all-day, off-site unless explicit time is provided
- MMC `CSO`: all-day, MMC location unless explicit time is provided
- DDH real shift labels without a time row: all-day
- `PHNW`: all-day single-day event
- `Annual leave`: seven-day all-day event, Monday through Sunday
- `Conference leave`: seven-day all-day event, Monday through Sunday

## MMC Parsing Rules

- Ignore cross-hospital placeholders such as `Dandenong`, `Dandenong AL`, `Dandenong CL`
- Team-code map:
  - `G` -> `Green`
  - `A` -> `Amber`
  - `R` -> `Resus`
  - `C` -> `Clinic`
- Shift prefix:
  - `A` -> `AM`
  - `P` -> `PM`
- Role suffix:
  - `C` -> no suffix in title
  - `R` -> append `Float`
- SSU codes:
  - `ASSC`, `ASSR`, `PSSC`, `PSSR`
  - Normalize team label to `SSU`

## DDH Parsing Rules

Normalize labels as:

- `Clinical Support` -> `CS`
- `SSU SMS` -> `SSU`
- `Orange PM (on-call)` -> `Orange PM`
- `AVAO PM` -> `AVAO PM`
- `AVAO AM` -> `AVAO AM`
- `PM FAST IC` -> `FAST PM`
- `Orange AM IC` -> `Orange AM`
- `Silver AM IC` -> `Silver AM`
- `onsite CS` -> `CS onsite`
- `PHNW clinical` -> `PHNW`
- `AM` / `PM` alone -> ignore

## Ignore Rules

Ignore:

- sick leave labels
- meetings
- exams
- teaching notes
- cross-hospital placeholders
- DDH `AM` / `PM` reference markers

Include:

- clinical shifts
- `PHNW`
- `Annual leave`
- `Conference leave`

## Overnight Rule

Any shift ending at `00:00` ends at midnight on the following calendar date.

## UX Requirements

- Require login before workspace use
- If one common doctor exists, auto-select and hide the dropdown
- If multiple common doctors exist, show a dropdown
- If no common doctors exist, show a validation error
- Show a preview count before export
- Keep event descriptions minimal
- Creator account can view a list of other user accounts
- Standard accounts see only their own account details
- Files list should show only the current account's imported roster references
- Reloading the page after login must restore the correct account workspace

## Storage Requirements

- Deduplicate identical uploaded files across all accounts
- Separate file/blob storage from account workspace state
- Never duplicate the same roster file unnecessarily when a shared source-of-truth file already exists
- Removing a file from one account should only delete the physical file if no other account references it

## Non-Goals For V1

- Direct CalDAV sync
- Arbitrary hospital parser configuration
- Editing roster files
