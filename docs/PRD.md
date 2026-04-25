# Roster Converter PRD

## Product Goal

Roster Converter is a multi-account web app that stores hospital roster sources in a shared repository, matches roster names to authenticated users, and generates each user's personal calendar as an Apple Calendar-compatible `.ics` file.

The app must move beyond per-browser local storage. Accounts, roster files, parsed shifts, name claims, disputes, and admin decisions are server-side product data.

## User Types

### Creator / Admin

- Email: `rhaydon@gmail.com`
- Has unrestricted retained roster history.
- Can view users, uploaded files, repository files, roster name claims, sites, roles, disputes, and unresolved roster-version conflicts.
- Can inspect another user's generated calendar for debugging.
- Can return from another user's calendar view back to their own calendar.
- Can resolve name-claim disputes.
- Can choose the active version of ambiguous roster files.
- Can eventually encode future version-selection rules after resolving an ambiguity.

### Standard User

- Creates an account with email, password, and real name.
- Can upload roster files.
- Can claim roster-name variants detected from repository files.
- Can see only their own generated roster.
- Can dispute a roster-name claim owned by another user.
- Can retain an active roster window of up to six months.

## Front Entrance

Unauthenticated users see a front entrance page rather than the roster workspace.

The entrance page must:

- Describe the app briefly.
- Provide login for existing users.
- Provide account creation for new users.
- Require email address, password, and real name during account creation.
- Explain that the real name should match or resemble at least one roster name.

After login or account creation:

- The app checks the repository for roster names likely matching the user's real name.
- The app shows detected candidate names grouped by site.
- The user confirms which names belong to them.
- If confirmed names produce roster events, the user is shown their calendar preview.

## Accounts

### Account Creation

- Users may create an account before uploading any roster.
- If repository data already contains likely matches for the user's name, the system should present those matches during onboarding.
- New users start with no claimed names until they confirm detected names or upload a roster that contains their name.

### Login

- Existing users log in with email and password.
- Login restores the user's server-side workspace and regenerates their current personal roster from repository data.
- Logging out does not delete uploaded files, claims, edits, or custom events.

### Account Uniqueness

- One account maps to one real person.
- A person may have many roster-name variants across sites.
- A roster-name variant should belong to only one person unless Admin explicitly allows a temporary duplicate alias during dispute resolution.

## Identity And Name Claims

### Real Name

- Real name is free text.
- Users may enter a preferred name, nickname, or formal name.
- The system does not require exact real-name matching before account creation.

### Roster Name Variants

Examples:

- `Richard HAYDON`
- `HAYDON, Richard`
- `SMITH, John`
- `Jane Small`

Roster name variants are stored as claims with:

- normalized name key
- display name as shown on the roster
- site
- role/grade if detected
- claiming user
- status
- source roster evidence

### Claim Rules

- If a roster-name variant is unclaimed, the user can claim it automatically.
- If a roster-name variant is already claimed, the user can submit a dispute.
- The app should show the claimant that the name is already taken and offer a dispute path.

### Disputes

If Jane attempts to claim a roster name already owned by John:

- Jane can submit a dispute.
- John receives an in-app message on next login.
- Admin receives the dispute.
- Email notifications should be supported later, but in-app messaging is required first.

Admin dispute actions:

- Transfer the roster-name claim from one user to another.
- Temporarily duplicate/alias the name.
- Reject the dispute.
- Message both users.

## Sites

Initial supported sites:

- `MMC`
- `DDH`

Planned supported sites:

- `Casey`
- `VHH`
- `MCH`

Each site has:

- canonical code
- display name
- roster source types
- location rules
- role/grade extraction rules
- name extraction rules
- shift parsing rules

Users may work across multiple sites.

The app must maintain:

- doctors/clinicians by site
- doctors/clinicians across multiple sites
- site-specific roster-name variants for each user

## Roles

Model roles now, even if not all roles receive calendar previews at launch.

Supported roles:

- Consultant
- Registrar
- HMO
- CMO
- ENP
- AMP

Interns:

- Parse and retain in backend data where present.
- Exclude from self-service calendar previews initially.
- Preserve enough data to support future "team-mates on shift" features.

Future roles may be added without changing core account or repository architecture.

## Roster Repository

### Repository Purpose

All recognized roster files are added to a central repository unless they are duplicates.

Users do not own roster files. Users gain access to their own shifts wherever their claimed roster names appear in active repository files.

### Upload Flow

When any user uploads a roster:

- Detect roster type and site.
- Compute an exact file hash.
- Compute a content fingerprint independent of cosmetic spreadsheet changes where possible.
- Check whether the file already exists in the repository.
- If the exact file or equivalent content already exists, discard the uploaded duplicate and create a reference/audit record.
- If the file is new, store it once in the repository.
- Parse roster names, roles, sites, date ranges, shifts, and source metadata.
- Rebuild affected users' generated rosters.

### Deduplication

Duplicate detection should use:

- exact file hash
- normalized roster content fingerprint
- source type
- site
- date range
- roster version metadata when available

The goal is to avoid storing dozens of identical or cosmetically changed roster files.

### Source Of Truth

Each repository file has status:

- active
- superseded
- duplicate
- rejected
- needs admin review

Only active files contribute to generated user calendars by default.

Superseded files are retained for audit/history and downloadable by Admin.

## Roster Version Handling

Multiple versions of the same roster may exist.

Version signals include:

- spreadsheet modification date
- upload date
- filename date
- filename version markers such as `V2`, `V3`, or `FINAL`
- roster date range
- completeness of shift coverage
- source-specific metadata

If the system can determine the latest version confidently:

- Mark older versions as superseded.
- Keep older versions for audit/history.
- Generate calendars from the newer active version.

If the system cannot determine the latest version:

- Flag a roster-version conflict for Admin.
- Admin selects the active version.
- Admin may record a future rule to help resolve similar conflicts.

## DDH / FindMyShift

DDH currently uses FindMyShift.

Accepted DDH sources:

- individual "my shifts" exports
- full DDH roster exports if available
- subscription URL input if available

Source-of-truth rule:

- Individual DDH my-shifts exports are acceptable until a full DDH roster covers the same individual and time period.
- Once a full roster exists for the same period, it becomes the preferred source of truth.

The app should allow DDH users to provide a subscription URL as an alternative to file upload where supported.

## Personal Roster Generation

Every time a user logs in:

- Authenticate account.
- Load the user's claimed roster-name variants.
- Query all active repository files where those names appear.
- Rebuild the user's generated personal roster.
- Compare regenerated events with user edits/custom events.
- Show "new shifts found" before applying new roster changes.

Default calendar view:

- Current full Australian medical term.
- If future roster data is available for the next term, extend preview to include that term.
- Past rosters are available via date range settings.

Term rules:

- Term 1 starts on the first Monday of February and runs 13 weeks.
- Term 2 starts on the first Monday of May and runs 13 weeks.
- Term 3 starts on the first Monday of August and runs 13 weeks.
- Term 4 starts on the first Monday of November and runs 13 weeks.

## User Edits And New Roster Versions

Users may edit imported events and add custom events.

If a newer roster version changes an event that the user manually edited:

- Preserve the user's manual edit by default.
- Create a conflict for later user review.
- Allow the user to reset the event to the latest roster value on an event-by-event basis.

Custom events remain user-owned and are not overwritten by roster imports.

## Calendar Output

- Export one `.ics` file containing the user's generated personal roster.
- Timezone: `Australia/Melbourne`
- Event descriptions remain minimal.
- Locations are included according to user settings.

## Existing Parsing Rules

### Title Rules

- MMC events: `MMC: <normalized shift>` unless disabled in settings.
- DDH events: `DDH: <normalized shift>` unless disabled in settings.

Examples:

- `0800-1730 AAC` -> `MMC: Amber AM`
- `0800-1730 AAR` -> `MMC: Amber AM Float`
- `1430-0000 PCC` -> `MMC: Clinic PM`
- `1430-0000 PSSR` -> `MMC: SSU PM Float`
- `Clinical Support` -> `DDH: CS`
- `Orange PM (on-call)` -> `DDH: Orange PM`
- `onsite CS` -> `DDH: CS onsite`

### Locations

MMC on-site:

- Label: `MMC Car Park`
- Address: `Tarella Road, Clayton VIC 3168, Australia`

DDH on-site:

- Label: `DDH Car Park`
- Address: `135 David St, Dandenong VIC 3175, Australia`

Off-site:

- No location

### Time Rules

If a source contains an explicit time, that time overrides canonical defaults.

MMC defaults:

- AM regular: `08:00-17:30`
- AM SSU: `07:30-17:30`
- PM any: `14:30-00:00` next day

DDH defaults:

- AM regular: `08:00-18:00`
- AM SSU: `07:30-17:30`
- AVAO AM: `07:30-17:30`
- PM regular: `15:00-00:00` next day
- AVAO PM: `14:30-00:00` next day

All-day special cases:

- MMC `CS`: all-day, off-site unless explicit time is provided
- MMC `CSO`: all-day, MMC location unless explicit time is provided
- DDH real shift labels without a time row: all-day
- `PHNW`: all-day single-day event
- `Annual leave`: seven-day all-day event, Monday through Sunday
- `Conference leave`: seven-day all-day event, Monday through Sunday

Overnight rule:

- Any shift ending at `00:00` ends at midnight on the following calendar date.

## Admin Dashboard

Admin dashboard must include:

- users by full name and email
- claimed roster names by site
- uploaded files, including accepted repository files and discarded duplicates
- repository files
- superseded roster versions
- downloadable audit/history files
- unresolved name conflicts
- unresolved roster-version conflicts
- clinicians by site
- clinicians across multiple sites

Admin user list:

- Show full names and email addresses inline.
- Show sites where each user works.
- Show how each roster displays the user's name by hovering over the site, rather than listing every alias inline.

Admin can open another user's calendar for debugging and then return to Admin's own calendar.

## Privacy

Standard users:

- See only their own extracted roster.
- Do not see other doctors' names by default.
- Cannot browse repository files.
- Cannot view other users' calendars.

Admin:

- Can inspect users, claims, repository files, and generated calendars for debugging and dispute resolution.

Future feature:

- Allow users to see colleagues working the same shift.
- This is explicitly deferred.

## Data Storage Requirements

Use Cloudflare server-side storage for:

- accounts
- password credentials
- uploaded source files
- file hashes/content fingerprints
- parsed roster records
- roster versions
- user/person records
- roster-name claims
- disputes
- admin messages
- user settings
- event overrides
- custom events
- generated roster snapshots or cache

Recommended implementation:

- Cloudflare D1 for relational entities and claims.
- Cloudflare R2 for uploaded source files.
- Optional KV for small session/cache data only.

Browser local storage may be used only as a temporary cache, not as the source of truth.

## Implementation Direction

The current local/browser persistence model is not sufficient for v2.

The v2 rebuild should introduce these entities:

- Account
- Person
- Site
- Role
- RosterNameAlias
- RosterFile
- RosterFileVersion
- RosterUpload
- ParsedRosterEntry
- NameClaim
- NameDispute
- AdminMessage
- UserEventOverride
- CustomEvent
- GeneratedRoster

## Open Questions

1. Email delivery provider:
   Which service should send dispute/login/admin emails later? Cloudflare Email Workers can receive email, but outbound transactional email usually needs a provider such as Resend, Postmark, Mailgun, or SendGrid.

2. Initial account verification:
   Should users verify email addresses before uploading rosters, or can verification wait until later?

3. Admin rule training:
   When Admin resolves a roster-version ambiguity, should the resulting rule be a simple pattern rule, or just saved as a note for now?

4. Subscription URLs:
   For DDH/FindMyShift subscription URLs, what exact feed formats are available: `.ics`, CSV, Excel, or authenticated web links?
