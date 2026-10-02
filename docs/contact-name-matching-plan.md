# Universal contact name matching implementation plan

Status: implementation completed locally following approval to proceed; not committed, pushed, or deployed.
Prepared: 2 October 2026.

## Implementation checkpoint

The shared matcher now handles explicit alternate names (including `Thisun (Tea)`), longer-name misspellings and transpositions, shortened names, initials, reordered components, accents, and non-Latin names. Supplied surnames must agree; a contradictory surname cannot be discarded to obtain a first-name match. All sites use the same matcher, with existing site/shift and VHH active-event gates retained.

Tentative allocations have an editable asterisk, original sheet name, explanation, confidence score, and confirm/reassign/return-to-review controls. Review entries show the leading alternatives. Conflicting duplicate names or handsets require contact-sheet correction; the save path rejects a selection that would remain unusable. Identity-scoped social-alias configuration exists but ships empty: unrelated social names cannot be learned from guesses.

### Shipped evidence table

| Name evidence | Name score | Tentative |
| --- | --- | --- |
| Exact full name | 100 | No |
| Reordered full name or matching supplied components | 99 | No |
| Exact given name plus surname initial | 98 | No |
| Approved identity-scoped alias | 98 | Yes |
| Full-name components with spelling/alternate evidence | 96 | Yes |
| Exact given name | 94 | No |
| Inferred given name plus surname initial | 94 | Yes |
| Existing recognized shortened given name | 92 | Yes |
| Qualified prefix, internal component, or accepted single edit | 90 | Yes |
| Accepted two-edit longer name | 88 | Yes |

Explicit parenthesized or quoted alternatives mark the result tentative even when the matching variant is exact. Prefixes need at least four characters and at least half the roster token. Spelling comparisons require both tokens to have at least five characters; one edit also requires similarity of at least 0.8. Two edits require both tokens to have at least ten characters and similarity of at least 0.85. Tokens longer than 48 characters are not fuzzily compared.

Stream agreement adds 3 and grade agreement adds 2, capped at 100. Automatic allocation requires name score at least 88, total score at least 90, and a lead of at least 12 over the next plausible eligible identity. The provisional 15-point lead was replaced with 12 so an exact longer name can be separated from a two-edit alternative; short or equally plausible alternatives remain unresolved. This choice passes synthetic boundary tests but is not calibrated on real confirmed allocations. Scores are ordinal confidence scores, not probabilities.

All tentative proposals use the same complete comparison state. Competing guesses remain for review; one guess cannot remove another candidate and create an allocation cascade. Strong specific-name matches and human confirmations retain precedence. VHH's stricter duplicate-holder handling is preserved.

### Persistence and cost evidence

The existing resolution integer now explicitly represents assigned (1), cleared (0), and rejected (-1). Its schema has no boolean constraint, so no migration is needed. API/R2 payloads expose an explicit `decision`; rejection suppresses rematching for the same contact key and sheet date, while clearing permits automatic matching again. The two existing allocation/history writes are atomic, with revision checks and no audit record for a losing revision race.

R2 publication merges per-contact revisions and uses conditional writes to prevent older concurrent saves from resurrecting rejected guesses. Publication failure reports the already-saved decision, preserves it locally through refresh, and offers an explicit retry. There are no automatic SQL retries.

- Automatic parsing, scoring, explanations, and allocation: **zero D1 statements, writes, or network requests**.
- Existing shared refresh: **36,000 unchanged refreshes tested with zero D1 reads and zero writes**. Request frequency remains unchanged.
- Explicit persistence helper: one indexed read and two row writes, unchanged.
- Successful shared confirmation/reassignment/rejection action: three indexed reads and two row writes, unchanged from the existing assignment path. Clear: two indexed reads and two row writes.
- Publication adds a bounded R2 compare/merge read on explicit saves and up to three R2-only conditional attempts if saves compete. No candidate triggers a database lookup.
- Revision conflicts and explicit retries remain exceptional human actions; no migration, remote database operation, extra endpoint, polling, or scheduled job was added.

### Validation and limits

`scripts/test-contact-name-matching.mjs` covers 14 positive name scenarios, contradictory surnames, short-name gates, ambiguity, shuffled inputs, duplicate events/handsets, competing guesses, browser/server evidence parity, active VHH expiry, correction controls, real local SQLite persistence, actual save-action execution, concurrent publication, injected failures, and query budgets. UI rendering uses the actual application functions; a local browser preview verified the asterisk explanation and confirm/reject/reset interactions.

Passing regression checks: syntax, existing contact allocations, contact sync, shared contact cache, contact access, facility access, restoration mutations, database costs, and client request budgets. The full fixture suite's pre-existing automation failure at line 922 remains outside this change.

No representative held-out, clinician-confirmed contact dataset was available. The fixtures establish mechanics and safety boundaries; they do not establish real-world accuracy or the reduction in review volume. Unknown unrelated social names still require review unless explicitly supplied on the sheet or added to the identity-scoped alias configuration. Production deployment remains a separate step.

## Goal and boundaries

Reduce contact allocations needing review when clerks use misspellings, shortened names, or explicit alternate names. Automatically allocate a phone only when name evidence is sufficiently strong, the leading candidate is clearly separated, and there is no competing allocation. Show tentative matches with an asterisk and allow correction.

Use one deterministic matcher for MMC, MCH, DDH, and VHH. Casey can use the same matcher when its feed is restored; do not enable or add that feed as part of this work. Preserve roster streams, staff identity, calendar events, expiry rules, access controls, service-phone behaviour, and existing cosmetic changes.

Ambiguous entries must remain for review. Being the best of several poor candidates is insufficient. Scores are confidence scores, not calibrated probabilities.

## Hard D1 and request budget

- Automatic name parsing, scoring, ambiguity checks, and rendering add zero D1 reads and zero D1 writes.
- Use only roster rows, contacts, and resolutions already delivered through the current loading/cache paths.
- Add no endpoint, directory lookup, history search, scheduled job, polling frequency increase, database migration, or persisted automatic-match log.
- Preserve the existing token-based contact refresh route, which reads published R2 objects before D1 authentication.
- Only explicit human confirmation, correction, or rejection may use the existing resolution save path. Compare its SQL operations and row counts with the baseline; reuse or reorder existing reads rather than adding lookups per candidate.
- No remote database operations are required for implementation or validation.

## Baseline implementation and observed gaps

- `public/static/contact-allocations.js`: pure shared matching module, imported by browser code and server correction validation. Existing methods include exact names, aliases, surname initials, first names, prefixes, and internal name tokens.
- `public/static/app.js`: constructs roster assignments, renders allocations, and handles temporary correction UI. Only manual matches currently have editable asterisk controls.
- `functions/api/state.js`: validates correction targets and rejects contacts or clinicians already automatically matched. VHH's `Current` contacts are not supported by the existing period-based correction logic.
- `functions/_lib/facility-contact-cache.js`: publishes and serves contacts and daily resolutions from R2.
- `functions/_lib/d1-calendar.js`: existing indexed resolution reads, bounded saves, revision conflicts, and audit history.

Read-only reproduction established:

| Contact name | Roster name | Current result |
| --- | --- | --- |
| Thisun (Tea) | Tea GUNASEN | Unmatched |
| Jacquline | Jacqueline MOREL | Unmatched |
| Alex SMITH | Alex JONES | Incorrect first-name match |
| Pat | Pat FINN and Patrick EXAMPLE | Ambiguous, correctly unmatched |

The current matcher consumes assignment indexes sequentially. The new matcher must handle distinct clinician identities and competing contacts explicitly, avoiding contact-order-dependent fuzzy guesses.

## Matching contract

### 1. Eligibility and identity

Keep facility/area, operational date, period, handover, and expiry checks. VHH requires a current dated extract and an explicitly timed roster event active at the matching instant; all-day and finished shifts remain ineligible. No score can override these checks.

Deduplicate candidate people by source and doctor key within the applicable matching context. Multiple event rows for one person must not create multiple competing candidates or permit duplicate allocation. Preserve explicit service-contact rules and their existing priority. Scope duplicate phone/name conflicts to the appropriate site and period/current context, not globally across the day.

The candidate roster must remain complete for ambiguity checks; provisional guesses cannot remove alternatives and make subsequent guesses appear more certain. Human-confirmed and genuinely unambiguous strong allocations can reserve targets, but weak names must never match merely because only one unallocated clinician remains.

### 2. Name parsing

Retain raw contact and roster names for display and correction. Generate comparison forms that normalize case, Unicode composition, accents, whitespace, punctuation, and name separators without discarding non-Latin text. Do not assume every last token is a surname or every internal token is a given name.

Recognize explicitly parenthesized or quoted alternate names, while excluding notes, grades, times, and phone numbers. For `Thisun (Tea)`, evaluate `Thisun` and `Tea` as alternatives. A unique exact `Tea` match is sufficient evidence for a marked inferred allocation; a second eligible Tea requires review.

Support surname initials and compound names. A clearly supplied conflicting surname must not be ignored in favour of matching first names. Preserve useful existing aliases and prefix matches only where the revised ambiguity and contradiction checks pass.

### 3. Evidence and scores

Use bounded Damerau-Levenshtein comparisons for spelling changes and transpositions, alongside exact tokens, explicit alternatives, initials, and approved aliases. Apply stricter gates to short names; do not fuzzily equate three-letter names on one edit alone. Do not infer unrelated social names through edit distance.

Keep separate name evidence and contextual evidence. Stream agreement is a small positive bonus; missing streams are neutral. Different streams do not defeat strong names. A clearly compatible grade can supply another small bonus, but mixed labels such as SMS/SR are not strict exclusions. Context alone cannot make a weak name eligible for automatic assignment.

Record the method, confidence score, uncertainty flag, and explanation in the in-memory result. Keep the top alternatives for review. Do not expose these scores as percentage probabilities.

Initial score/lead discussion used 90/100 and a 15-point lead. These are provisional, not validated settings. Before activating tentative matches, define and document the exact evidence table, short-name gates, contextual bonus cap, absolute threshold, and separation threshold. Validate them against held-out confirmed examples. A weak match with no runner-up must still fail the absolute evidence requirement.

### 4. Assignment and precedence

Build candidate evidence deterministically, without relying on contact-sheet order. Reserve eligible human-confirmed resolutions and non-conflicting strong matches before tentative assignments. If confirmed data conflicts with another strong entry, retain the conflict for review rather than silently overriding confirmation.

For a tentative allocation, require adequate name evidence, an adequate total score, a clear lead over the next eligible person, and no similarly strong competing contact claiming that person. Apply all tentative decisions from the same comparison state. Do not recursively consume guesses to create artificial uniqueness. A global assignment algorithm must permit unmatched contacts; it must not force allocation to maximize coverage.

Preserve existing special service-phone rules. Roster streaming remains authoritative; a matched phone never moves a clinician between roster groups.

### 5. Correction and rejection

Distinguish strong automatic, tentative automatic, and manually confirmed allocations. Tentative matches are editable; they must not trigger the current blanket "safe automatic match" rejection. Human confirmation takes precedence over subsequent guesses.

Reuse the existing daily resolution mechanism, expected revisions, and R2 publication. Opening an explanation or viewing alternatives requires no request. Saving an explicit choice uses the existing save action.

Provide a rejection that returns the contact to review and suppresses its tentative rematch for the same contact key/date. Do not simply clear the allocation and let the next render recreate it. Determine whether an existing inactive resolution can represent this suppression without changing manual-clear semantics; if necessary, propose a minimal explicit decision field before implementation rather than silently repurposing persisted state.

For VHH, share active-event eligibility between automatic matching and correction validation; the literal `Current` label cannot be checked against AM/PM/Night. Expiration must invalidate confirmed as well as inferred allocations. A confirmation does not carry a phone beyond the rostered finish time.

Ensure browser and server validation use equivalent candidate names, grade/stream evidence, context, and matching time. Current browser-built assignments and server-built assignments differ; add parity fixtures to prevent correction disagreements. Do not fetch extra parser metadata to achieve parity.

### 6. Display

Show tentative phone allocations with an asterisk and an accessible explanation such as "Automatic tentative match — check allocation". Explain the original sheet name and matching evidence on click, with confirm, reassign, and return-to-review controls.

Existing manually assigned numbers may retain their asterisk, but use distinct wording indicating human assignment. Show unresolved candidates ordered by confidence with reasons. Preserve staff sorting, seniority acronyms, swing times, and omission of empty phone messages.

### 7. Approved social names

Use identity-scoped approved aliases, not broad global substitutions for unrelated names. A first-release configuration can be bundled with the app and shared by browser/server, requiring no D1 lookup or alias management screen. Do not automatically learn aliases from guesses or daily corrections.

An alias editor and publication into existing R2 roster data are separate future work. The first implementation must not depend on a full clinician directory or identity-merge project.

## Implementation sequence and checkpoints

1. Add meaningful failing fixtures for the observed examples, context boundaries, short-name ambiguity, contradictions, duplicates, and competing contacts. Capture correction/request/D1 baselines using local mocks.
2. Add structured name parsing and improve deterministic evidence checks. Fix contradictory surnames before widening matches. Preserve reviewed existing alias cases.
3. Specify the scoring table and implement bounded scoring and simultaneous conflict decisions. Validate determinism under shuffled contacts/events and duplicate event rows.
4. Implement correction precedence and rejection semantics in the shared module, existing save validation, and published resolution payload. Include VHH active-event parity.
5. Add tentative-match display and correction controls. Test actual rendered phone markers, explanations, and preserved grouping/times.
6. Evaluate proposed thresholds offline against confirmed multi-site examples. Report incorrect allocations separately from reduced review counts. Synthetic examples establish mechanics, not real-world precision. If representative examples are unavailable, keep uncertain rules conservative and state that calibration remains unverified.
7. Review the complete diff, query/request evidence, and threshold decisions before enabling or deploying this functional change. Deployment requires an implementation-stage user request; earlier cosmetic deployment authorization does not by itself authorize publishing this new matcher.

## Acceptance evidence

- `Thisun (Tea)` matches unique eligible `Tea GUNASEN`; duplicate Tea candidates remain unresolved.
- Plausible longer-name typos and transpositions resolve only with adequate strength and separation.
- `Alex SMITH` cannot match `Alex JONES` through first name alone; `Ama` is not guessed as `Arnav` without an approved identity alias.
- Stream agreement helps discriminate plausible candidates; changed/missing streams do not force or prevent otherwise valid strong matches.
- Ambiguous Pat/Qing/Ann, conflicting contacts, duplicate events, unsupported names, and absent roster staff remain appropriately unresolved.
- Shuffling inputs preserves decisions; tentative allocations cannot create a cascade of forced matches.
- Confirm, reassign, reject, refresh, and stale-revision conflicts work; rejected guesses do not immediately reappear.
- VHH stale sheets, handovers, finished shifts, all-day events, duplicate holders, and midnight/offset cases retain current restrictions.
- Automatic matching performs zero database/network calls. Existing refresh volume retains zero D1 activity. Explicit correction SQL/request counts do not increase beyond the documented bounded baseline without an explained, reviewed change.
- Run `npm run check`, `npm run test:contact-allocations`, `npm run test:contact-sync`, `npm run test:facility-contacts`, `npm run test:facility-contact-access`, and `npm run test:facility-access`, plus meaningful new correction/parity/budget tests where existing coverage is insufficient.
- The full fixture suite currently has an independently confirmed pre-existing automation assertion failure at `scripts/test-fixtures.mjs:922`. Record it separately; do not alter unrelated automation code to make this work pass.

## Handoff and reasoning effort

This is a functional change with subtle assignment, correction, and expiry interactions. A detailed plan reduces ambiguity but does not remove the need to reason through code. Sol on Low can be trialled for narrow implementation steps with exact fixtures; it is not a demonstrated reliability guarantee for the complete change.

Use Sol on Medium as the practical default for implementation. Use High to review evidence thresholds, assignment competition, correction/rejection semantics, browser/server parity, and D1 budget evidence. No model configuration is changed by this plan.
