# Doctor names and possible duplicates UX plan

## Purpose

Replace the current implementation-shaped Identity review screen with a
plain-language, full-width workspace for understanding doctor records and
reviewing possible duplicate roster names.

The system should find plausible name variations with little effect on normal
application performance. It must make no identity change until an authorised
person confirms the result, except for narrowly defined formatting updates that
have independent evidence that the roster names already belong to the same
person.

This plan follows the durable transaction, audit, account-conflict, source-data,
snapshot, and exact-reversal rules in
[`doctor-identity-merge-plan.md`](doctor-identity-merge-plan.md). It supersedes
that plan's user-facing workflow and its requirement for a name-derived
canonical database key. The older plan must be aligned with the internal ULID
and visible Person ID model below before implementation.

## User promise

An administrator should be able to answer one ordinary question:

> Could these roster names belong to the same doctor?

They should not need to understand database keys, source keys, redirects,
similarity scores, account IDs, or snapshot jobs.

The normal outcomes are:

- **Yes, same doctor**
- **No, different doctors**
- **Not sure yet**

The interface supplies the evidence, proposes a safe result, previews the
effect in ordinary language, and keeps technical details available without
placing them in the normal path.

## Problems being corrected

The current prototype exposes a raw merge form requiring comma-separated
person IDs, optional source aliases, and account email addresses. It also shows
numeric scores and repository reason codes. These controls are useful for API
testing but are not suitable for routine administration.

The small current data set can also make duplicate discovery appear broken.
`Jay WEERARATNE` and `Jayantha WEERARATNE` are already an approved relationship,
so they must appear as two roster names for one doctor, not as a new possible
duplicate and not as an unexplained empty result.

Candidate discovery currently depends too heavily on approved person aliases.
It must also examine roster names that no account or person has claimed yet.

## Product decisions

1. **Use ordinary language.** The interface says doctor, roster name, main
   display name, combine records, different doctors, and undo. Technical terms
   remain in API documentation, logs, and an optional Technical details area.
2. **Doctor records and suggestions are distinct.** Existing approved roster
   names are shown on a doctor record. Only unresolved relationships appear as
   possible duplicates.
3. **People are selected through names and cards.** No normal task asks a user
   to type a database key, source key, or account email.
4. **The first decision is whether the names describe the same doctor.** Target
   records, alias movement, redirects, and database operations are derived by
   the system whenever possible.
5. **Discovery is suggestion-only.** A scan does not move roster names, combine
   people or accounts, change source events, or rebuild calendars.
6. **A name alone is not an identity credential.** Equivalent punctuation,
   spacing, case, titles, or Unicode forms can be normalised for comparison, but
   two otherwise independent records are not automatically combined solely
   because their normalised names match.
7. **Every administrator change is previewed and reversible when dependencies
   allow.** Candidate resolution and its identity change are committed together.
8. **Clinicians may confirm only a uniquely eligible, unclaimed roster name for
   their own established doctor record.** They cannot search other people,
   choose a target, move an owned name, or combine accounts.

## Identity and visible Person ID model

### Immutable internal identity

Every doctor has an immutable internal key:

```text
person:<ULID>
```

This key is created once, never edited, never recycled, and hidden from the
normal interface. Aliases, accounts, operations, snapshots, and foreign keys
refer to it. Combining records preserves one internal key and records the other
as a merged historical identity according to the durable merge rules.

### Visible Person ID

Each doctor also has a readable, unique Person ID, for example:

```text
kularatne-aeshan
```

The interface may display and copy this value under record details. It is a
human-facing reference, not the database identity key.

Store readable references in a separate mapping table such as
`roster_person_references` with:

- readable reference
- internal person ULID
- state: `current` or `historical`
- operation that created or replaced it
- creator and timestamps

The current readable reference is unique. Historical references continue to
resolve to the same internal person and are never reassigned to someone else.
Changing a visible Person ID is an advanced, audited reference change; it never
moves aliases, accounts, or roster events and never changes the internal ULID.

Generate a suggested reference from the main display name. Resolve collisions
with a clear suggestion and allow an authorised administrator to choose another
readable value. Surname-boundary questions are shown only when generation is
ambiguous, not during every combination.

## Interface language

Use these user-facing terms consistently:

| Technical concept | Interface wording |
| --- | --- |
| Identity review | Doctor names |
| Canonical person | Doctor record |
| Preferred display name | Main display name |
| Alias | Roster name |
| Candidate | Possible duplicate |
| Merge | Combine records |
| Reject candidate | Different doctors |
| Defer candidate | Not sure yet |
| Activity/history | Change history |
| Reverse operation | Undo this change |
| Scan/audit | Check for possible duplicates |
| Canonical ID | Internal key, never shown normally |
| Public slug/reference | Person ID |

Do not display numeric similarity scores, raw reason codes, version tokens,
redirect terminology, or snapshot vocabulary in the normal workflow. Translate
them into short evidence and warning statements.

## Information architecture

Create a full-width **Doctor names** workspace rather than placing the workflow
inside the existing narrow Admin modal.

It has three primary views:

| View | Purpose |
| --- | --- |
| Doctors | Find a doctor and understand all roster names attached to them |
| Possible duplicates | Review unresolved suggestions, with a pending count badge |
| Change history | Understand and, where safe, undo earlier changes |

Scanning is not a fourth primary destination. In Possible duplicates, show:

- last successful check time
- whether new roster names have been checked
- a **Check now** action with a required narrow scope
- current progress or a clear deferral/error message

Detailed run counters, cursors, budgets, and diagnostic history belong under
**Advanced scan details**.

Opening the workspace should load only summary counts and the selected view.
People, candidates, and history are fetched separately with server-side search
and cursor pagination.

## Doctors view

### Search and list

Provide one search field with the hint:

> Search by doctor name, roster name, hospital, email, or Person ID

Search is server-side and matches:

- main display name
- exact and normalised roster-name spellings
- source/hospital
- linked account email, for authorised administrators only
- current and historical visible Person IDs

Return a small cursor-paginated page, initially 25 records. Debounce text input
and do not run candidate matching as part of search.

Each compact result shows:

- main display name
- hospitals represented
- number of roster names
- account count when relevant
- event coverage summary
- latest meaningful identity change

Do not show an `Active` badge on ordinary records. Show exceptional state in
plain language, such as **Combined into Jayantha WEERARATNE**. Keep the internal
ULID hidden. Place the visible Person ID in expanded details rather than making
it compete with the doctor's name.

### Doctor detail

The full record groups exact source spellings by hospital and shows where and
when each roster name appeared. Approved relationships are described in normal
language, for example:

```text
Jayantha WEERARATNE

DDH   Jayantha WEERARATNE   Seen Feb–Aug 2026
VHH   Jay WEERARATNE        Seen Mar–Jul 2026

These roster names were confirmed as the same doctor on 31 Aug 2026.
Person ID: weeraratne-jayantha
```

Searching either `Jay` or `Jayantha` opens this one record. Nothing suggests
that further action is outstanding.

### Actions

Primary actions are:

- **Add a roster name**
- **Check for possible duplicates**
- **Change main display name**
- **View change history**

Advanced actions are:

- **Move one roster name**
- **Change visible Person ID**
- **View technical details**
- **Undo a previous change**

The roster-name picker shows eligible unclaimed names by default. If an
administrator searches for a name already attached elsewhere, show it disabled
with a plain explanation of its current doctor/account rather than silently
hiding it.

## Possible duplicates view

### Queue

Show unresolved cases by default. Provide useful filters without requiring
them:

- hospital/source
- likely match or needs caution
- account conflict
- overlapping-work warning
- not-sure and different-doctor history
- check run/date

Each row shows names and hospitals, not person IDs. Do not expose a number such
as `score 78` or codes such as `exact-given-surname-one-letter`. Translate the
evidence:

```text
Why this was suggested
• The given name is the same.
• The surnames differ by one letter.

Things to check
• These names both appear at DDH.
```

Highlight changed letters or tokens, but never rely on colour alone. Include
screen-reader text that states the difference.

Where pairwise suggestions overlap, group them into one review case. For
example, A↔B and B↔C should show all three related roster names so the
administrator does not make inconsistent decisions across separate rows.

### Empty and progress states

Distinguish these outcomes:

- **No unresolved possible duplicates.** Existing confirmed roster names are
  available under Doctors.
- **No roster names were available in this scope.** The system had no inputs to
  compare; this is not the same as finding no matches.
- **Check in progress.** Show progress and allow the user to leave the page.
- **Check paused safely.** Explain that it reached its performance limit and
  will continue, or provide a Continue action.
- **Check deferred.** Explain, for example, that a roster import is currently
  running.

An already confirmed pair links to the doctor record and relevant change
history rather than becoming another case.

## Review flow

Use a full-width comparison page with a persistent Back action. Preserve queue
filters and position when the reviewer returns.

### 1. Ask the human question

The heading is:

> Could these roster names belong to the same doctor?

Show:

- exact roster spellings and highlighted differences
- hospitals/sources
- first and last event dates and event counts
- all related roster names already attached to either record
- account state without unnecessary personal details
- reasons the names were suggested
- cautions such as overlapping shifts, conflicting stable identifiers, or
  different claimed accounts
- relevant previous decisions and what evidence has changed

The actions are:

- **Yes, same doctor**
- **No, different doctors**
- **Not sure yet**

There is no one-click combination from the queue.

### 2. If they are the same doctor

The system derives the safest result:

- if one side is an established doctor and the other is unclaimed, attach the
  unclaimed roster name to the established doctor;
- if neither side has a doctor record, create one with an immutable internal
  ULID;
- if both sides have doctor records, retain the older or more complete record
  unless a durable conflict requires administrator choice;
- suggest the main display name from the most complete source spelling;
- suggest a visible Person ID from that name;
- include all eligible roster names and the appropriate accounts by default.

The standard flow does not ask the administrator to choose a technical target
or tick individual aliases/accounts. A genuinely partial move is a separate
advanced **Move one roster name** workflow and is not presented as an ordinary
duplicate resolution.

Let the administrator change the proposed main display name. Put the visible
Person ID and target-record choice under Advanced options unless there is a
collision, ambiguous name boundary, or record conflict that requires input.

### 3. Show the impact in ordinary language

The server returns a deterministic preview. Present its meaning, not its raw
rows:

```text
These records will be combined under Aeshan KULARATNE.

2 roster names across 2 hospitals
43 historical roster events will continue to appear under this doctor
1 linked account
No source roster records will be changed
Calendar updates will run in the background
```

An expandable **Exactly what will change** section may list doctor records,
roster names, accounts, visible Person IDs, historical references, and affected
calendar/profile owners.

If multiple accounts would become associated with one doctor, interrupt the
normal flow with a separate confirmation naming the accounts and clearly
stating the consequence. Server-side permission, ownership, version, and
conflict checks remain authoritative.

### 4. Confirm and record the result

Use the primary button **Combine doctor records**. An optional note may be
entered under “Add a note for change history.”

Commit the candidate outcome, identity mutation, visible-reference changes,
audit record, and version checks as one transaction. The case must not remain
pending after a successful combination.

After success, show a receipt:

- resulting main display name
- roster names combined
- account effect
- background calendar-update status
- **Undo this change**, while exact reversal is still dependency-free
- link to the resulting doctor record

### Different doctors

Use the action label **Confirm they are different doctors**, not Reject.

Offer optional quick reasons:

- I know these are different people
- Conflicting accounts or identifiers
- Work records show they cannot be the same person
- Names are similar by coincidence
- Other

Keep the decision suppressed until material evidence changes. If it resurfaces,
state exactly what is new.

### Not sure yet

“Not sure yet” removes the case from the immediate queue without pretending a
decision was made. By default, return it only when relevant evidence changes.
The administrator may optionally choose a review date. Do not create an
indefinite, unexplained deferred queue.

## Change history and undo

Describe operations as human events:

- “Combined Jay WEERARATNE with Jayantha WEERARATNE”
- “Changed the main display name”
- “Moved one roster name to another doctor”
- “Changed the visible Person ID”
- “Undid an earlier combination”

Show actor, time, note, and a short before/after summary. Keep operation IDs,
internal ULIDs, stored rows, and version data under Technical details.

When exact reversal is possible, offer **Undo this change**. When a later change
depends on it, explain what must be undone or reassigned first using doctor and
roster names rather than operation IDs. Never perform a partial best-effort
reversal.

## Individual clinician experience

This remains a narrow Account feature, not access to Doctor names.

Show a card only when all eligibility rules pass:

> We found another roster name that might be yours.
>
> VHH: Jay WEERARATNE
>
> Is this also you?

Actions are **Yes, this is me**, **No, this is not me**, and **Ask the
administrator**.

Eligibility requires:

- a currently authenticated account;
- an existing active doctor record for that account;
- an unclaimed roster name, or one already attached to that same person;
- a persisted potential-match signal rather than request-time broad searching;
- a unique best match with no competing person/account candidate;
- no ownership by another person or account;
- explicit self-confirmation.

The endpoint cannot accept a target account/person, move an owned name, combine
people or accounts, edit a display name or Person ID, or reverse history.

Record a successful action as `self-confirmed`, show it in administrator change
history, and make it undoable under the normal dependency rules. “No, this is
not me” suppresses the prompt for that account and records useful evidence for
administrator review; it does not silently declare that two people are
different globally.

A formatting-only spelling may attach without a prompt only when independent
continuity evidence already establishes the same person, such as the same
account/person or a stable source identifier. Normalised name equality alone is
insufficient.

## Source identity index

Candidate discovery needs a durable input representing every distinct roster
identity, including names not claimed by an account and not yet attached to a
doctor.

Add a source-identity table or equivalent materialised index containing:

- source/hospital
- immutable source doctor key
- exact latest display spelling
- first-seen and last-seen dates
- event count
- source/import watermark
- nullable resolved internal person ULID
- active/last-observed state
- feature version and updated timestamp

Populate or update this index as part of successful roster import finalisation.
Do not manufacture an approved doctor record merely so an unclaimed roster name
can participate in matching.

Feature rows should be created when a source identity changes, not discovered
for the first time halfway through a scan. Handle removed or superseded source
identities explicitly so stale features do not continue generating cases.

## Candidate discovery engine

### Comparison features

Precompute explainable features such as:

- punctuation/case/Unicode-normalised name
- name tokens and token count
- surname and surname prefix
- given names and initial
- phonetic surname key
- source and department
- first/last event dates
- established person/account or stable external identifier, where available

Use these features for candidate generation, not automatic authority.

### Bounded matching

Compare only new or changed source identities. Use indexed blocking keys to
find a small neighbourhood and never perform request-time all-pairs matching.

Use a stable composite cursor such as:

```text
(feature_updated_at, source_type, source_doctor_key)
```

A timestamp alone is insufficient because a page boundary could contain many
rows with the same timestamp.

For common names or broad blocks:

- measure block cardinality before comparison;
- add stronger keys when a block is too broad;
- apply deterministic ordering and resumable paging;
- record a plain-language budget deferral rather than silently discarding
  candidates with an arbitrary limit.

### Evidence and warnings

Useful positive evidence includes:

- exact given name with a small surname difference
- uncommon exact/near surname with a given-name variation
- known nickname or shortening, as evidence only
- phonetic similarity
- shared verified account metadata or stable identifier
- compatible hospital/department history
- prior administrator decisions

Warnings or blockers include:

- overlapping shifts that make the identity doubtful
- different claimed accounts
- conflicting stable identifiers
- a competing person with comparable evidence

Hospital equality is context, not automatically a warning. Overlap and account
conflicts should be described specifically.

### Candidate lifecycle

Persist states such as:

- pending
- confirmed same doctor
- confirmed different doctors
- not sure yet
- superseded
- stale because the underlying identities changed

Store an evidence fingerprint. A different-doctor decision remains suppressed
until material evidence changes. When records are combined, atomically resolve
the reviewed case and supersede every other pair that now resolves to the same
internal person.

## Scan behaviour and performance limits

Scheduled and manual checks use the same incremental engine.

The scheduled check remains infrequent, timezone-aware for Melbourne,
lease-protected, paused during roster import/reprocessing, and resumable.

Initial batch limits remain deliberately conservative and are tuned in Preview:

- no more than 250 changed source identities per invocation
- no more than 15 seconds of work
- an explicit comparison/candidate budget
- explicit D1 row-read limits

A check may only read source identities/features and create or update candidate
and run rows. It performs no snapshot reads, invalidation, rebuilds, roster
parsing, source-event writes, feed work, or R2 writes.

Manual checks require one scope:

- one doctor or roster name
- selected hospitals
- one roster upload
- source identities changed within a date range
- unresolved identities, still processed in bounded pages

Do not offer an unlabelled “scan everything” button. Return a run ID promptly,
continue in bounded background invocations, and allow the user to leave the
page. A paused run must have an explicit continuation mechanism rather than
depending on repeated browser refreshes.

## API and repository changes

Add card-oriented, cursor-paginated queries rather than returning raw tables:

- doctor summary search
- doctor detail
- possible-duplicate counts and cases
- case detail with translated evidence data
- change-history pages
- eligible roster-name picker
- scan summary and run progress

The picker returns exact source spelling, hospital, ownership state, event
coverage, eligibility, and a user-readable explanation when blocked.

Preview/commit requests accept stable selected entities from cards and the
reviewed case ID. Raw internal person IDs and account emails are never typed by
the user, though internal keys remain in authenticated API payloads.

Add visible Person ID lookup and history through the separate reference table.
All repository reads accepting a visible reference resolve it to the immutable
internal ULID.

Keep mutation APIs versioned, transactional, audited, and administrator-only,
except for the constrained self-confirmation endpoint. Preserve account-conflict
checks, flattened merged-person resolution, exact reversal, and scoped
post-commit rebuild rules.

## Loading and responsiveness

- Fetch summary counts when the workspace opens.
- Fetch only the selected primary view.
- Use cursor pagination; do not load the complete doctor list or all history.
- Debounce server search and cancel/ignore stale responses.
- Cache safe summary data briefly where useful.
- Never start candidate matching from search, account load, tab navigation, or
  ordinary doctor-detail reads.
- Show skeleton/loading, empty, error, stale-preview, and background-update
  states without replacing the user's current context.

## Accessibility and responsive behaviour

- The comparison is usable at desktop width and becomes a clear sequential A/B
  layout on narrow screens.
- All actions are keyboard accessible with visible focus.
- Opening and closing detail panels restores focus correctly.
- Differences are stated in text and do not rely on colour.
- Status badges include readable text and sufficient contrast.
- Touch targets meet mobile sizing expectations.
- Screen readers receive meaningful names for each roster spelling, hospital,
  warning, and outcome action.
- Back navigation preserves filters, pagination cursor, and queue position.

## Implementation phases

### Phase 0: align the durable model and prototype the language

- Update the older merge plan to use immutable internal ULIDs plus visible
  Person ID references.
- Define reference creation, collision, historical lookup, and non-reuse rules.
- Create low-fidelity full-width layouts for Doctors, Possible duplicates,
  comparison, confirmation, and Change history.
- Validate terminology and the three outcome choices before building APIs.

### Phase 1: source identity and reference foundations

- Add/backfill the source-identity index, including unclaimed roster names.
- Add/backfill immutable internal ULIDs and the visible Person ID reference
  table without breaking existing references.
- Generate comparison features at import finalisation.
- Implement composite watermarks, stale-feature handling, and query-plan tests.

### Phase 2: Doctors workspace

- Add server-side search, pagination, detail, visible-reference lookup, and
  linked-name evidence.
- Build the full-width workspace and Jay/Jayantha approved-record example.
- Remove the raw merge form and numeric/repository-shaped presentation.

### Phase 3: discovery and possible-duplicate queue

- Run bounded matching across claimed and unclaimed source identities.
- Add case grouping, evidence translation, filters, lifecycle rules, accurate
  run telemetry, and empty/progress states.
- Verify that approved same-person names are excluded or superseded.

### Phase 4: simplified review and undo

- Implement the three human outcomes.
- Derive the target and full-record combination by default.
- Add the human-readable impact preview, conflict branches, atomic case/merge
  commit, success receipt, and immediate Undo action.
- Keep partial roster-name movement as a separate advanced workflow.

### Phase 5: constrained clinician confirmation

- Show only persisted, uniquely eligible unclaimed roster names.
- Enforce the no-other-person/account boundary server-side.
- Record yes/no evidence and expose administrator history and undo.

### Phase 6: Preview measurement and usability verification

- Exercise Jay/Jayantha, Aeshan spelling variants, Toby formatting variants,
  common surnames, ambiguous matches, account conflicts, and overlapping shifts.
- Measure search latency, page sizes, D1 reads, block sizes, comparisons, run
  duration, candidate quality, and background job counts.
- Confirm that checking for duplicates causes zero snapshot jobs, R2 writes,
  feed rebuilds, or roster mutations.
- Conduct task-based testing with people who do not know the identity data
  model, then refine wording and layout before production proposal.

## Acceptance criteria

### Comprehension and usability

- A non-technical administrator can explain within ten seconds that Jay and
  Jayantha are already confirmed as one doctor.
- They can resolve Aeshan's spelling variation without seeing or entering an
  internal person key, source key, or account email.
- The first decision is visibly Same doctor, Different doctors, or Not sure.
- They understand what will change before combining records.
- They understand that Different doctors suppresses a suggestion rather than
  deleting either doctor or roster history.
- They can return from a review without losing queue position or filters.
- Core tasks work with keyboard-only input, a screen reader, and a phone-width
  layout.

### Identity behaviour

- Internal `person:<ULID>` keys are immutable and hidden from normal UI.
- A visible Person ID resolves through a separate mapping and may be changed
  without changing the internal person.
- Historical visible Person IDs continue resolving and are never reassigned.
- Existing approved roster names appear on one doctor record, never as a new
  duplicate suggestion.
- Candidate confirmation and the resulting identity operation commit together.
- Combining records preserves every original roster name and source event.
- Every administrator mutation is previewed, version-checked, audited, and
  exactly reversible when no later dependency blocks it.

### Examples and safety

- `Aeshan KULARATNE` and `Aeshan KULURATNE` produce a review case and never
  combine automatically.
- `Toby O BRIEN`, `Toby OBRIEN`, and `Toby O’Brien` are treated as formatting
  equivalents, but independent records are not combined solely by name
  normalisation.
- A clinician may attach only a uniquely eligible unclaimed roster name to
  their own existing doctor record.
- The clinician action cannot combine people/accounts or take another owner's
  roster name.
- Multi-account, stable-identifier, and overlap conflicts receive specific,
  prominent handling.

### Discovery and performance

- Unclaimed source identities participate in candidate discovery without being
  turned into approved doctor records.
- Search and view loading use indexed server-side pagination and never trigger
  matching.
- Candidate generation uses indexed blocking and a stable composite cursor,
  not an all-pairs join or timestamp-only paging.
- Broad blocks are re-blocked or deferred explicitly rather than truncated
  silently.
- Interrupted checks resume without skipped identities or duplicate cases.
- Manual checks return promptly and continue in bounded background work.
- Checks respect runtime, identity, comparison, candidate, and D1 read budgets.
- A check creates no snapshot/rebuild jobs, R2 writes, feed work, or source
  roster changes.

## Recommended initial visual flow

The first implementation should test this sequence:

```text
Doctor names
├─ Doctors
├─ Possible duplicates (2)
└─ Change history

Possible duplicates
┌──────────────────────────────────────────────────────────────┐
│ Could these be the same doctor?                              │
│ Aeshan KULARATNE · DDH     Aeshan KULURATNE · VHH            │
│ The given name is the same; the surname differs by 1 letter. │
│                                                              │
│ [Review]                                                     │
└──────────────────────────────────────────────────────────────┘

Review
Could these roster names belong to the same doctor?
[Yes, same doctor] [No, different doctors] [Not sure yet]

Confirmation
These records will be combined under Aeshan KULARATNE.
2 roster names · 2 hospitals · 43 events · 1 account
Original roster records will not be changed.
[Combine doctor records]
```

This is a starting interaction model, not a fixed visual design. Refine it after
the first working full-width implementation and task-based usability testing.

## Out of scope

- Automatic approval of misspellings, nicknames, initials, abbreviations, or
  independent records sharing only a normalised name
- Rewriting original roster names or historical roster events
- Deduplicating source event rows as part of a doctor-record combination
- Joining or deleting user accounts
- Allowing clinicians to browse doctor records or possible duplicates
- Self-service person/account combination, name correction, or Person ID change
- Request-time or continuous whole-database similarity scans
- Production migration, roster reprocessing, scheduled activation, or snapshot
  rebuilding without separate approval
