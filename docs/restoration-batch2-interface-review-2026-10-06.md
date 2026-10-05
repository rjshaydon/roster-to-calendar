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
