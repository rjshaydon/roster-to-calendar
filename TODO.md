# Roster Converter TODO

## Launch Blockers

- Configure Cloudflare server-side storage for production: create/bind `ROSTER_STORE` KV to the Pages project and redeploy. Without this, account creation, shared repository data, and cross-device persistence cannot work.
- Verify account creation in production after `ROSTER_STORE` is available: creator creates a user, enters that account, links roster names, logs out/in, and sees persistent data.
- Implement name-claim conflict workflow: claimed names shown greyed out, user can report conflict, creator can transfer/alias/reject, and both users receive an in-app message.

## Account And Admin

- Creator admin can create accounts, enter them, delete them, and return to creator account.
- Creator admin can view all users by real name/email, claimed roster names by site, uploaded files, repository files, unresolved name conflicts, and unresolved roster-version conflicts.
- Creator admin can jump into a user calendar for debugging and then return to their own calendar.
- Creator admin feature: "When am I working with Dr X?" to search overlap with a selected clinician.
- Creator admin feature: "Who else am I working with this shift?" to show other rostered clinicians for a selected shift.
- Standard users can edit email/password/account details and delete their own account.
- Add in-app messaging for disputes and admin decisions.
- Add optional outbound email later through Gmail for dispute/admin notifications.

## Roster Repository

- Store each unique roster once, with user uploads becoming references to repository files.
- Detect duplicates by exact hash and normalized content fingerprint, not just filename.
- Retain older/superseded roster files for creator audit/download while hiding them from active calendar generation.
- Add version-resolution logic using final/version filename markers, modification dates, upload dates, and completeness.
- Add admin workflow for low-confidence roster version conflicts.
- Rebuild affected user calendars when a new active roster is added to the repository.

## Name Detection

- MMC Excel/PDF detection should include consultants, CMOs, registrars, HMOs, ENPs, and AMPs, excluding interns from self-service calendars.
- Add equivalent role-aware detection as new hospitals are added: Casey, VHH, and MCH.
- Improve fuzzy matching between account real names and roster names, including nickname/abbreviation variants such as `Abirama Thanikasalam` vs `Abi THANIKASALAM`.
- If no calendar events are found for a user, always offer roster-name selection from repository names: unclaimed first, claimed greyed out, and “my name is not listed”.

## Sources And Parsing

- MMC Excel import: working for current sample.
- MMC PDF import: initial support exists; bare time-only cells are flagged for review because the PDF text may not contain a shift code.
- DDH FindMyShift spreadsheet export: working for current sample.
- FindMyShift `webcal://` subscription URL ingestion: not built yet.
- For FindMyShift URLs, store the URL as a private secret-like value, refresh on login/app entry, and do not display the full token back to users.
- Add Casey, VHH, and MCH parsers/rules when sample rosters are available.
- Make parser warnings reviewable and editable before export.

## Calendar Editing

- Persist imported-event edits and custom events server-side per user.
- Preserve manual overrides when a newer roster version changes an imported event, and present conflicts for later resolution.
- Add per-event reset to imported roster details.
- Keep asterisk only when imported event differs from its original date/time/location/title.
- Improve drag-to-nudge time behaviour to continuously increment/decrement in 15-minute and then accelerated hourly steps.
- Support event drag/drop, copy/paste, delete, reset, and multi-day custom events reliably on desktop and mobile.

## Calendar UI

- Calendar console skin is the default.
- Fix and maintain responsive banner alignment, fixed left pane, and full-week calendar width.
- Settings should be a popover, not inline.
- Add hospital selector only when the user has roster data at more than one hospital.
- Add per-hospital default car park/location settings.
- Add shift colour settings and keep the category list aligned with real shift types.
- Make mobile interactions touch-friendly, including file removal, event editing, and sticky export where appropriate.
- Tighten the mobile layout substantially; the current mobile version needs dedicated responsive design rather than desktop layout compression.

## Export

- Export `.ics` for Apple Calendar, Google Calendar, and any calendar app that imports `.ics`.
- Add export controls for date range, hospital/site, leave types, and locations.
- Consider future subscribed calendar feed support after account/repository architecture is stable.

## Tests And Fixtures

- Keep fixtures for each supported roster type and hospital.
- Add regression tests for new parsers and known edge cases.
- Add tests for account persistence once server-side storage is abstracted enough to run locally.
- Add tests for repository duplicate/version handling.
