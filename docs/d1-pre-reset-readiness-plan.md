# D1 pre-reset readiness plan

## Purpose

Use the remaining time before the next D1 allowance reset to remove uncertainty
without consuming a D1 row, changing Cloudflare runtime state or risking another
outage. The outcome is a reviewed evidence pack and a precise reset-day runbook,
not an early attempt to resume At a glance.

The reset is at 00:00 UTC (10:00 AEST on 8 September 2026). Cloudflare Analytics
is treated as unsettled for 15 minutes, so a report generated before 00:15 UTC
cannot provide evidence about the new quota day. No D1 request will be used to
test whether the reset occurred.

This plan is subordinate to
[`d1-account-quota-safe-rollout-plan.md`](./d1-account-quota-safe-rollout-plan.md).
Its work is local or control-plane read-only. It does not authorise a push,
deployment, variable change, migration, application request or remote D1 SQL.

## Non-negotiable state until the reset-day gates pass

- Keep Production at the verified fail-closed `a293251` deployment.
- Do not push local commit `6bda8b7` merely to make it available online; pushing
  `main` may start an automatic Production deployment.
- Keep every optional reader, builder, bootstrap, roster automation and
  maintenance switch disabled and every allowlist empty.
- Do not open At a glance or deliberately invoke colleague queries as a test.
- Use the Account Analytics Read token only with Cloudflare GraphQL Analytics.
- Do not use `wrangler d1 execute --remote`, an application endpoint, a remote
  migration, bootstrap or database console query.

## Workstream 1 — finish the incident evidence without D1

1. Save one final settled account-wide Analytics report covering the complete
   incident day. Record the query interval, generation time, daily totals,
   five-minute totals, per-database totals and query-fingerprint totals.
2. Preserve the raw sanitized GraphQL response outside the repository so a
   later interpretation can be checked without another API call. Never store
   the token or request authorization header.
3. Confirm that the daily and five-minute aggregates reconcile. Keep the
   fingerprint shortfall explicit; do not manufacture attribution for the
   8.23 million missing rows.
4. Produce a minute-by-minute correlation table for Production Pages requests,
   D1 rows read/written, deployments and GitHub workflow runs from 04:20–05:00
   UTC. Label temporal correlation separately from proven causation.
5. Record which historical data Cloudflare does not expose—particularly Pages
   request paths and the missing D1 fingerprints. This prevents tomorrow's
   observations being treated as proof of yesterday's exact route.

Deliverable: an immutable incident evidence folder under `/private/tmp` plus an
updated evidence ledger if the settled totals differ from the current report.

## Workstream 2 — complete the account and caller inventory

This work may inspect configuration and provider control-plane metadata, but it
must not contact a database binding.

1. Reconcile `config/d1-database-inventory.json` against a read-only Cloudflare
   D1 database listing. If the present token cannot list databases, leave the
   inventory incomplete and record the missing permission rather than widening
   the Analytics token or guessing.
2. For every database, record its ID, environment, Pages/Worker binding, current
   deployment and retained preview/branch callers.
3. Audit the repository's Pages configuration, Worker configuration and GitHub
   workflows for every D1 binding, scheduled trigger and remote command.
4. Confirm that no external scheduler, queue consumer, old Worker, local cron or
   retained preview can call Production. Unknown callers keep the gate at STOP.
5. Distinguish a database inventory from a caller inventory: knowing both
   database IDs is insufficient if an old deployment or browser loop remains
   unidentified.

Deliverable: a complete, evidenced inventory. `complete` remains `false` until
every account database and caller has been verified.

## Workstream 3 — local request-storm investigation

The 7 September burst coincided with 395 Production requests in 15 minutes.
Before any online experiment, investigate how the deployed client could create
that amplification.

1. Compare the last known stable client, the deployed incident client and the
   current fail-closed client. Concentrate on startup, login, profile switching,
   Admin Files, `calendarStoreStatus`, At a glance, visibility/focus handlers,
   retries and service-worker cache transitions.
2. Build a trigger matrix showing, for each user action and lifecycle event:
   endpoint, retry limit, polling interval, cancellation/coalescing behaviour,
   expected D1 statements and whether a hidden tab can call it.
3. Search specifically for duplicated event listeners, recursive refreshes,
   overlapping initialization, retry-on-503 loops, stale JavaScript calling a
   newly deployed API, and requests started by rendering error/unavailable
   states.
4. Reproduce each suspected sequence against the isolated local SQLite database
   with request and statement counters. Test one tab first, then a small scripted
   multi-tab amplification. Do not simulate an elaborate hospital; reuse the
   existing deterministic fixtures and synthetic history database.
5. Establish separate results for:
   - login and ordinary calendar opening;
   - Creator profile switching;
   - Admin Files opening while status summaries are unavailable;
   - At a glance opening and each tab transition;
   - contact refresh at 60 seconds; and
   - hiding, restoring and repeatedly reloading a tab.
6. For every invoked local SQL statement, capture `EXPLAIN QUERY PLAN`, rows
   examined/returned estimates and invocation count. An indexed query repeated
   hundreds of times is still an incident risk.
7. If a concrete fault is found, implement its fix only on a new local safety
   branch based on `6bda8b7`, add a regression test with a hard request/query
   ceiling, and run focused local tests. Do not push it before a separate review.

Deliverable: a route/trigger matrix, local reproduction report and—only if
supported by evidence—a locally committed fix. “Could not reproduce” is a valid
result and must not be converted into confidence.

## Workstream 4 — validate the gate and prepare tomorrow's artifacts

1. Re-run the budget-gate unit tests and all zero-D1 guard tests locally.
2. Test the command offline with fixtures for missing credentials, unknown
   database, incomplete inventory, missing fingerprints, delayed buckets,
   counter decrease, quota thresholds and a burst between samples. Every case
   must return STOP.
3. Confirm that the checker uses the settled cutoff and cannot accidentally
   treat the previous UTC day as the new day. Do not collect a reset-day first
   sample before 00:20 UTC.
4. Prepare, but do not execute:
   - exact Analytics commands and output paths for each sample;
   - the expected UTC and Melbourne timestamps;
   - a configuration read-back checklist;
   - the one-file MMC inspection payload;
   - the immediate non-D1 rollback command; and
   - a results worksheet containing GO/STOP reasons and cumulative deltas.
5. Review the staged/local diff for secrets and confirm unrelated user files are
   excluded. Record the local commit and Production commit separately.

Deliverable: commands that can be copied exactly tomorrow, with no improvised
IDs, timestamps, flags or estimates during the rollout.

## Workstream 5 — review checkpoint before the reset

Before stopping for the night, report:

- what is proven about the request storm and what remains unknown;
- whether the database and caller inventory is genuinely complete;
- whether the local gate and zero-D1 tests pass;
- the Production and local commit IDs;
- confirmation that no remote state changed and no D1 request was made; and
- the earliest possible time for the first meaningful Analytics sample.

Any unresolved safety issue remains a written blocker. It is better to lose a
rollout day than to consume the quota needed for login.

## Reset-day runbook prepared tonight

All times below are UTC on 8 September 2026; Melbourne is UTC+10.

| Time | Action | D1 effect |
| --- | --- | --- |
| 00:00 / 10:00 AEST | Reset boundary. Make no probe request. | None |
| 00:20 / 10:20 AEST | First account-wide Analytics sample, observing settled data only through approximately 00:05. | None |
| 00:35 / 10:35 AEST | Second account-wide sample; compare database totals, fingerprints and burn rate. | None |
| 01:00 / 11:00 AEST | Passive sample and ordinary-traffic attribution. | None |
| 02:20 / 12:20 AEST | Two-hour-plus passive-baseline assessment using settled data through approximately 02:05. | None |

At 02:20 UTC, continue only if the account inventory is complete, all samples
are internally consistent, every material increase is attributed, ordinary
usage projects safely inside its reserve and the budget command says GO. A STOP
result means no D1 canary that day.

If GO is achieved, the next action is still not execution. First review and
explicitly deploy the safety-controls commit, read back the effective variables
without D1, and confirm all capabilities remain closed. Only then can the
separately approved one-file inspection sequence in Phase E be considered.

## Completion criteria for pre-reset work

Pre-reset preparation is complete when:

- the final incident evidence is saved and discrepancies are explicit;
- all account databases and callers are verified, or the missing inventory item
  is clearly documented as a blocker;
- the Production request storm has been investigated locally with bounded
  request/query counts;
- the budget and zero-D1 guard suites pass;
- tomorrow's commands, timestamps, file ID, rollback and worksheet are ready;
- no secret is present in Git or generated output; and
- no Production/Preview D1 operation or Cloudflare runtime mutation occurred.

## Implementation checkpoint — 7 September 2026

- Production now runs safety commit `2af89c0`. The earlier instruction to keep
  Production at `a293251` was superseded by the reviewed safety deployment.
- Before reset-day observation, implement the additional fail-closed plan in
  [`at-a-glance-zero-d1-maintenance-plan.md`](./at-a-glance-zero-d1-maintenance-plan.md)
  so stale and current clients stop before At a glance authentication or access
  reads.
- Final settled incident Analytics is stored under `/private/tmp` with `0600`
  permissions; daily and five-minute totals reconcile exactly.
- The account inventory is complete: two roster databases plus the separately
  bound `acem-exam-tutor-db`. The account has one Pages project and the retained
  roster deployment classes are recorded.
- Local browser tracing reproduced 11 requests per Creator login. The safety
  revision reduces this to six and removes automatic console writes.
- Missing and malformed legacy/emergency flags now default closed. By stream no
  longer repeats failed metadata immediately.
- The budget checker now rejects pre-settlement UTC rollover and distinguishes
  valid early evidence from permission to proceed.
- The exact reset-day commands are in
  [`d1-reset-day-runbook.md`](./d1-reset-day-runbook.md).
- Broad On shift rollout remains blocked because 36,000 unchanged contact
  refreshes must first move off repeated D1 authentication.
