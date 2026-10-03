# Clinician access and Working together handoff

Saved on 3 October 2026 on branch `codex/clinician-access-history`.

Continue Working together restoration from this branch's clinician-access commit,
or merge/cherry-pick that commit into the restoration branch before editing the
overlapping files. Both features touch `functions/api/state.js` and
`public/static/app.js`; starting from the older main branch will create avoidable
conflicts. This checkout is shared across conversations: avoid switching its
branch while another conversation is editing it. Use a separate worktree if
independent branches are needed. A push of this branch is a saved checkpoint;
production release and migration remain pending.

## Contracts to preserve

- Enabled SMS and CMO accounts have all-site visibility. Other clinicians have
  the set of hospitals supported by their linked identities' active roster
  memberships during the current term, including verified locum memberships.
- Historical Working together uses server-authorized hospital/date segments for
  the requested period. Current hospital access must not substitute for past
  entitlements. Future trainee terms do not establish historical access.
- `queryFacilityOverviewTogetherContext` supplies the historical hospital choices,
  authorized doctor directory, scope revision, and missing-data information.
  The browser must not use the global doctor picker or pinned doctors to populate
  this directory. Cached results require fresh authorization and a matching scope.
- `queryFacilityOverviewWorkingTogether` accepts an empty doctor selection for
  date-first searching. Results are clipped to authorized hospital/date segments,
  including overnight shifts. Related cached roster insights apply the same
  scope before identifying colleagues.
- Access supports `all`, `site`, `sites`, and `denied`, with `facilityKeys` for
  multiple sites. A linked user may have historical search access despite having
  no current hospital entitlement. These are version 2 permission/cache contracts.

See `docs/at-a-glance-clinician-access-plan.md` for implementation details.

## Validation and remaining rollout work

Clinician access, facility access, snapshots, contact access, cached roster
insights, facility rollout, D1 account budget, and client request budget suites
passed. Syntax and whitespace checks passed. Three larger suites have unrelated
failures independently reproduced on the original main commit: fixtures
(automated correction identity assertion), facility maintenance (launch flag
expectation), and facility materialization (import SQL missing `name` column).

Migration `0038_clinician_access_scope.sql` adds an identity/term index and
permission-cache invalidation triggers. It has not run in production. No
application deployment has been performed. The 11:34 am AEST account analytics
sample was valid with 8,293 reads and 16 writes, but the quota gate requires a
second sample and a two-hour passive baseline. Recheck at 12:20 pm AEST on
3 October, then follow `docs/clinician-access-release-preflight-2026-10-03.md`.
The baseline report is `/private/tmp/clinician-d1-usage-first.json`.

Do not release unfinished restoration changes during the migration follow-up.
Inspect the live schema and migration ledger, estimate index cost, and establish
a recovery point only after budget admission. Apply only reviewed migration 0038;
the normal migrations command can also apply other pending files. Verify the
ledger and quota after migration. Historical data publication/backfill remains
a separate rollout consideration.
