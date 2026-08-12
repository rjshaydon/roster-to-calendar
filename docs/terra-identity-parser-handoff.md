# Terra handoff: account identity merging and parser convergence

## Objective

Correct the account-calendar identity bug demonstrated by Jessica McQuillian, preserve roster spelling variants as durable aliases, and make manual uploads and automatic imports use one roster interpretation.

This work must be completed in reviewable phases. It must not deploy, reprocess production rosters, modify remote D1/R2 data, or resolve the production unknown shift-code backlog without separate approval.

## Repository and working-tree safety

- Repository: `Roster to calendar tool`
- Start from the current `main` branch and create a `codex/` prefixed working branch.
- Preserve unrelated existing untracked files, including `.cursor/`, `docs/facility-overview-plan.md`, and `new_app_description/`.
- Do not overwrite, stage, or commit unrelated user changes.
- Use a separate commit for each completed phase.
- Run `npm run check` and `npm run test:fixtures` after each material phase.

## Confirmed current behaviour

### Jessica McQuillian

Jessica's account has two legitimate roster claims:

- DDH: `JESSICA MCQUILLIAN`
- MMC: `JESSICA MCQUILLAN`

The underlying D1 events exist under both roster keys. The account subscription implementation in `functions/api/feed.js` queries every claimed key, so it is intended to combine them. The account snapshot implementation in `functions/api/state.js`, specifically `buildDerivedAccountSnapshot`, groups claims by exact identity and selects only one group. It can therefore show only one spelling/site in the application calendar.

The browser error `Roster files need to be re-uploaded so they can be parsed into D1.` originates in `public/static/app.js` when `parseCurrentRosterForm` cannot hydrate browser-local files. It is not evidence that D1 roster events are absent. A claimed cloud account should remain on the D1 calendar path and should not fall back to this misleading re-upload instruction.

### Identity system

The existing canonical builder conservatively merges one-character surname variations when given names match, the names occur in different files, and working events do not conflict. However, canonical identities are rebuilt from current roster state and are not a durable administrator-approved alias registry.

Account claims remain source-specific roster identities. Multiple claims should contribute to one user calendar, irrespective of exact spelling.

### Parser duplication

Manual/browser parsing uses `public/static/roster.js`. Automatic/server parsing uses `functions/_lib/roster.js`. They are separate copies and have known semantic differences:

1. Browser parsing disables whole-week leave inference for FindMyShift-format DDH data; server parsing does not.
2. Browser parsing has DDH seniority look-ahead; server parsing does not.
3. Browser parsing recognises additional Paediatric HMO headings (`ED HMO'S`, `ED HMOS`, `HMO'S`, `HMOS`); server parsing does not.
4. Casey explicit-time unknown labels are treated differently: server marks them unknown with a warning; browser accepts them without that warning.

The manual implementation is the initial behavioural baseline because it is older and has received more roster-specific refinement. The Casey difference must be documented and deliberately resolved rather than silently copied.

## Phase 1: focused regression baseline

Keep this phase small. Its purpose is to make the confirmed failures executable before changing behaviour.

### Work

- Add a fixture/test account with two claims whose surname spellings differ by one character and whose events exist at different sites.
- Reproduce and characterize the current account snapshot split: selecting/building one identity omits the other identity's events. Keep the committed test suite green; the desired merged-calendar assertion may be activated with the Phase 2 fix.
- Test that the subscription-feed event-key selection includes all claims.
- Add normalized browser-versus-server parser comparison helpers for existing fixtures.
- Record the four known parser differences as explicit failing or expected-difference cases so they cannot be lost during convergence.

### Acceptance

- A focused characterization demonstrates the Jessica-style omission for the expected reason, while the committed test suite remains green.
- Existing unrelated fixtures still pass.
- Test output distinguishes event occurrences from distinct unresolved shift codes.

### Commit

Use a phase-specific commit containing tests and test helpers only.

## Phase 2: merge all account claims into one calendar

### Work

- For non-creator accounts, make `buildDerivedAccountSnapshot` query all unique keys from all valid account claims, rather than selecting one exact-spelling claim group.
- Return one account-facing doctor option using the account real name (or stable preferred display name) with every claim represented as an alias/site membership.
- Ensure selecting any claimed alias from the creator switcher resolves to and opens the same claimed account.
- Keep creator/owner doctor switching behaviour unchanged.
- Make application calendar aggregation and `.ics` subscription aggregation use equivalent claimed-key semantics.
- Deduplicate exact duplicate events using the established event identity without collapsing genuinely different shifts.
- Prevent claimed cloud accounts from falling back to browser-local roster parsing when a D1 snapshot is missing or rebuilding.
- Replace the re-upload message in this context with an accurate D1 empty/building/indexing message. Do not remove the message from genuine manual-upload workflows where re-upload is actually required.
- Invalidate/rebuild relevant account snapshot cache descriptors when claims change.

### Acceptance

- Jessica's DDH and historical MMC events appear in one application calendar.
- Either spelling routes to the same account calendar.
- Application and subscription paths use the same claimed roster keys.
- Existing exact-name, multi-site accounts retain all their events.
- A claimed identity with zero events receives an accurate empty-calendar response and no re-upload instruction.
- Creator and unclaimed doctor-profile views remain functional.

### Commit

Use a separate phase-specific commit.

## Phase 3: durable aliases

### Design requirement

Introduce a stable person identity and persistent source-specific aliases. Do not destructively rewrite historical roster events. Events may retain their original `doctor_key`; calendar queries should expand a canonical person into approved aliases.

An acceptable schema should express the equivalent of:

- Stable canonical person identifier and preferred display name.
- Alias: source type, roster doctor key, display spelling, canonical person identifier.
- Provenance: automatic or administrator-approved.
- Confidence/review state where applicable.
- Created/updated timestamps and administrator identity for manual approval.
- Uniqueness preventing one source/key alias from belonging to multiple people.

### Work

- Add a forward-only migration; do not delete or rewrite existing roster-event history.
- Seed durable aliases conservatively from current claims/canonical aliases.
- Preserve approved aliases when an old roster becomes inactive.
- Use automatic aliasing only for high-confidence transformations already considered safe: title/case/punctuation/order normalization and the existing conservative one-character surname rule.
- Require administrator confirmation for nicknames, broader misspellings, or ambiguous matches.
- Ensure account claiming attaches to the stable person while retaining source-specific aliases for diagnostics.
- Provide administrator-visible alias provenance and a safe way to add/remove an alias association.
- Treat alias removal as unlinking; do not delete underlying roster events.
- Detect and reject alias ownership conflicts.

### Migration cases

- Merge Jessica's DDH and MMC spellings into one canonical person.
- Preserve Aeshan Kularatne/Kuluratne as aliases only if supported by existing stored identities or explicit approval; do not fabricate an unsupported alias.
- Do not create speculative nickname mappings such as Mike/Michael automatically.

### Acceptance

- Historical events remain discoverable after the source roster becomes inactive.
- Approved aliases survive canonical-index refreshes.
- Similar but ambiguous names are not silently merged.
- Existing account and subscription tokens remain valid.
- Migration is idempotent and rollback-safe at the application level.

### Commit

Use a separate migration/identity commit. Do not apply the migration remotely.

## Phase 4: parser parity and one source of truth

### Part A: parity before consolidation

- Parse every existing fixture through both browser/manual and server/automatic implementations.
- Compare normalized doctors, aliases, memberships, events, dates, times, titles, locations, seniorities, leave, issues, conflicts, and unknown-code diagnostics.
- Port the more complete manual behaviour for FindMyShift leave, DDH seniority look-ahead, and Paediatric HMO headings to automatic parsing.
- Investigate the Casey explicit-time difference and document the chosen result. Prefer preserving a valid explicit shift while still surfacing a useful review warning if that matches product intent.
- Produce a concise before/after parser delta report.

### Part B: consolidate

- Establish one shared roster-parsing core.
- Keep environment-specific file/dependency loading in thin browser and server adapters only.
- Do not leave two hand-maintained copies of roster interpretation logic.
- If a generated browser artifact is used, add a deterministic generation command and a CI/test assertion that it is current.
- Add a parser version to processed roster metadata or equivalent revisioning.
- Ensure retained active source files can be reprocessed when the parser version changes, but do not trigger remote reprocessing in this task.
- Make canonical/alias refresh failure visible and prevent an import from being represented as fully healthy when identity refresh failed.

### Acceptance

- Both entry points produce identical normalized output for all supported fixtures.
- CI fails on future parser drift.
- Manual upload remains available as a recovery path but no longer has separate roster rules.
- Automated Adult, Paediatric, and FindMyShift processing all call the shared interpretation.
- Existing tests pass and new parity tests cover the known divergences.

### Commits

Prefer one commit for intentional behaviour reconciliation and a second for structural consolidation.

## Phase 5: preview verification and production proposal

This phase is report-only unless the user gives explicit production approval.

- Prepare a preview deployment or local production-equivalent verification.
- Verify Jessica, Aeshan, representative exact-name accounts, and multi-site accounts.
- Run retained roster sources through the new parser without writing results to production.
- Produce a before/after report of added/removed/changed events and changed unknown-code counts, grouped by facility, seniority, and code.
- Identify which active roster files would require reprocessing.
- Provide a rollout and rollback sequence, one source at a time.
- Stop and request approval before deployment, remote migration, snapshot rebuild, or roster reprocessing.

## Unknown shift-code backlog: coordination boundary

The previously reported `8,538` unresolved items must be re-audited after parser convergence. Confirm whether this number represents occurrences, stored issue rows, or distinct facility/seniority/code combinations.

Work that may proceed in parallel while these phases are implemented:

- Read-only extraction and classification of the existing unknown population.
- Grouping by facility, seniority, normalized raw code, source file, and frequency.
- Identifying obvious duplicates, stale/inactive-file issues, formatting noise, and already-known codes incorrectly reported as unknown.
- Preparing candidate interpretations with evidence from explicit roster times and surrounding labels.

Work that must wait until parser parity is complete and the retained files have been shadow-reparsed:

- Saving new production parser rules.
- Bulk marking unknowns ignored or resolved.
- Reprocessing production roster files.
- Treating the current count as the final remediation workload.

The final shift-code remediation should operate on the post-parity, post-shadow-reparse dataset so rules are written once against the permanent parser.

## Required handoff report from Terra

At the end of each phase, report:

- Files changed.
- Behaviour changed.
- Tests added and results.
- Known risks or unresolved decisions.
- Whether any fixture event output changed, with a summarized delta.
- Confirmation that no production data, deployment, or remote reprocessing occurred.

Stop for user direction if a phase requires a product decision that materially changes calendar output.
