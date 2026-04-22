# Roster to Calendar Tool PRD

## Goal

Build a local tool that ingests two roster formats for a selected doctor and exports a single `.ics` calendar file for Apple Calendar.

## Supported Sources

1. MMC workbook export
2. DDH FindMyShift Excel export

## Core Workflow

1. Upload the MMC workbook
2. Upload the DDH workbook
3. Detect names present in both sources
4. Auto-select the doctor when there is only one common name
5. Show a dropdown when multiple common names are found
6. Preview the normalized event count and date span
7. Export a single `.ics` file

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

- If one common doctor exists, auto-select and hide the dropdown
- If multiple common doctors exist, show a dropdown
- If no common doctors exist, show a validation error
- Show a preview count before export
- Keep event descriptions minimal

## Non-Goals For V1

- Direct CalDAV sync
- Multi-tenant cloud hosting
- Arbitrary hospital parser configuration
- Editing roster files
