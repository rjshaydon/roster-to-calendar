# Batch 2 identity interface preview — 6 October 2026

The identity review interface has been reorganised around finding a person and editing their details. Changes are restricted to the isolated synthetic preview at https://identity-batch2-acceptance.roster-to-calendar.pages.dev/; production flags and storage are unchanged.

## Behaviour

- Search runs after 400 ms of idle typing. Old search responses cannot replace newer results. Queries use indexed preferred-name prefixes, exact IDs/account emails, or exact roster keys across five indexed sites. Pages contain at most 25 records. Retired identities are hidden from ordinary name results and remain accessible through ID searches and history.
- Saving reloads the current results and relevant person details. Returning to the page refreshes outside an active edit, at most once per 30 seconds. There is no continuous directory polling. Calendar-publication status checks stop after three attempts.
- Person names open an adjacent detail panel (stacked below on mobile), with roster names, accounts, editing and history. IDs are under advanced details. Multiple selection has Clear selection and a contextual merge action.
- Separate this roster name selects an existing destination through bounded prefix search or creates a new person. Other roster names and account links remain assigned to their current people. All operations retain server-side previews, conflict checks, account confirmation and exact reversal.
- Edit person combines preferred-name and optional ID correction. Name interpretation uses comma precedence, distinct capitalised surname words, then the final-word fallback. Given names and surname are editable. IDs are proposed in surname/given-name order; the person prefix is added automatically. Existing IDs stay unchanged during ordinary name correction. Reserved ID conflicts offer the existing record for inspection.
- Suggestions prepares missing registry entries and finds/displays candidates in one explicit action. Both stages retain 25-identity batches and a 250-identity stop per stage. Preparation is skipped for five minutes after a completed preparation in the current session. Longer work can be stopped/resumed. There is no preparation scan during search or login.
- Dismissal has immediate Undo and a bounded Dismissed suggestions view. Restoration guards the exact evidence fingerprint and changes no identity assignments.
- History uses human-readable actions, affected roster/name/account summaries, dates and undo information. Administrator email and successful refresh details are hidden. Pending/failed updates remain visible, including failed reversal updates. Recent history remains capped at 25 operation records.
- Reasons use presets plus an optional note; Other requires a note. Review screens show roster-name ownership before and after, with explicit identifiers when correcting an ID.
- Versioned entry assets ensure previously opened previews receive the new module.

## Verification

Passed identity operation tests, name interpretation tests, bounded identity tests with 100,001 historical events, and Creator login containment tests. Added checks cover dismissed-suggestion retrieval, stale restoration rejection, dismissal undo, indexed all-site roster-key search, human-readable history, separation to an existing person without moving the source account, and exact separation reversal.

Browser acceptance verified automatic search; opening a person; separation to an existing person, saving and automatic detail updates; undoing separation; single-step suggestions; dismissal/undo; selecting a suggested pair; Clear selection; comma-separated name parsing; optional ID generation; and mobile layout at 390 × 844 with no horizontal document overflow. Mobile override was reset after review.

## Review data

Search for **Preview Doctor**. Alpha and Alfa remain separate and have an outstanding suggestion. Alpha deliberately contains **MMC — Preview Doctor Beta**, attached to the wrong person, to demonstrate separation. **Preview Doctor Beta** exists as an available destination. The wrong attachment was restored after acceptance testing so the user can repeat the test. Synthetic login credentials are unchanged.

No production migration, roster download, historical shift rewrite, real identity merge, or scheduled automation activation was performed. The separate pending Batch 2 work (Creator-selected subscription alias expansion, richer audit filters and production activation) remains outside this interface batch.

## Review refinements

Suggestion cards use **Review** and **Not the same person**. Review opens the suggested people’s details without opening a merge form. The unrelated people list is hidden while reviewing; **Close person** restores it. Editing and saving also show the results again. **Merge selected people** is available separately in the search-results selection area and opens its own editor.

Editing now uses the full panel width, with search results underneath instead of a narrow left column. Read-only copies of the person details and paired comparison are hidden during edits/review-before-save. History is a collapsed disclosure; failed calendar updates remain flagged outside it. Separating a roster name changes ownership, whereas rejecting a suggestion changes no existing links.

## Single-pane browsing and explicit comparisons

The surface now shows exactly one full-width pane: People when browsing, or the person/comparison/editor when open. No empty Person details placeholder is rendered. This also applies after Suggestions. Closing a person returns to People.

Name contains matches literal characters anywhere in the preferred name, case-insensitively, including surname and substrings spanning name parts. It uses the existing name index to seek through fixed 250-record windows (maximum eight per request); larger directories return a resumable search cursor. Empty browsing reads only a small page. No roster/event-history query or new schema/index is required. Tests cover Richard Haydon via Haydon, cross-name substrings, literal percent signs, pagination, indexed seek plans, and bounded continuation in a 2,500-person fixture.

Comparison panels ask Is this the same person? and offer Same person, Different people and Open this person. Same person opens the existing merge editor/review; it does not immediately mutate data. Different people dismisses the specific suggestion without changing existing links. Each account label names the person owning those account links; comparison does not imply that accounts are already linked to one another.


### Focused suggestion review

Suggestion cards now offer Same person, Different people, Review in that order. Same person loads only the suggested pair into the existing merge review, with no immediate commit. The current discovery service emits pairwise suggestions, so each card currently contains two people; separate cards are not treated as one group.

Opening a person or merge editor hides the search and suggestion area while preserving the explanatory text. Closing the person or cancelling the merge restores the browsing area and existing suggestions. Each comparison names the matched person AND the open person above the matched person's details.
