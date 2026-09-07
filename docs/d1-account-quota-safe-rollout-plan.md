# Account-wide D1 quota safety and controlled rollout plan

## Status and authority boundary

Prepared 7 September 2026 after a protected single-file bootstrap inspection
was rejected by Cloudflare because the account had already exhausted its daily
D1 row-read allowance. This plan supersedes every earlier rollout instruction
that treats a per-database Metrics chart, a successful small D1 query, or more
than 50% apparent headroom as sufficient authority for Production work.

Writing this plan does not authorise runtime changes, Production queries,
deployments, configuration changes, backfills, bootstraps or feature
activation. Until its prerequisites are implemented and separately approved,
make no Production or Preview D1 request merely to test availability.

The application must remain in the verified fail-closed state:

- `FACILITY_SHARED_ROLLOUT_ACTIVE=false`;
- `FACILITY_SHARED_EMERGENCY_PAUSED=true`;
- `FACILITY_LEGACY_READS_PAUSED=true` permanently;
- every shared builder and reader disabled;
- all facility build and reader allowlists empty;
- roster-status summaries disabled;
- roster, contact, queue and advanced-maintenance writes disabled; and
- the independent watchdog disabled.

At a glance and Creator roster-status diagnostics may remain unavailable.
Reduced optional functionality is safer than exhausting D1 and preventing
login or ordinary calendar access.

## What failed on 7 September

The rollout used the Production database's custom Metrics interval beginning
at 00:00 UTC. The displayed database figures were approximately 195,000 rows
in the regional chart and 407,000 in the summary tile. Both appeared to leave
more than 90% of the five-million-row allowance.

That evidence was not an account-wide admission-control counter. Cloudflare's
documentation distinguishes per-database Metrics from total account billable
usage under **Billing > Billable Usage**. D1's free allowance is enforced for
the account, across its databases, and analytics may arrive after the work
that generated them. A later token-protected inspection request immediately
received the account quota error.

The failed request did not execute a bootstrap. It failed before returning a
plan and wrote no compact data. Its intended first lookup was an exact primary-
key read, so it cannot plausibly explain millions of rows by itself. The
unattributed remainder must be identified from settled account-wide analytics;
it must not be guessed from the visible Production query list.

The incident also exposed two independent control problems:

1. An automatic Pages Git deployment updated code but did not apply the changed
   Pages variables. An explicit Wrangler deployment was required before the
   downloaded effective configuration changed.
2. Read-only bootstrap inspection currently requires opening the broad
   advanced-maintenance path and lowering the shared emergency pause. A safe
   inspection should never grant execution authority.

## Authoritative Cloudflare facts

- The Workers Free allowance is five million D1 rows read and 100,000 rows
  written per account per day, resetting at 00:00 UTC.
- Rows read means rows scanned, not rows returned. DDL and index creation can
  contribute both reads and writes.
- Per-query `meta.rows_read` and `meta.rows_written` are exact for that query.
- Per-database Metrics and query fingerprints are available through D1's
  GraphQL analytics datasets.
- Cloudflare documents total account D1 usage under Billing > Billable Usage,
  but the current Workers Free account exposes only R2 there. Account-wide D1
  GraphQL is therefore the available programmatic evidence.

References:

- <https://developers.cloudflare.com/workers/platform/pricing/#d1>
- <https://developers.cloudflare.com/d1/observability/metrics-analytics/>
- <https://developers.cloudflare.com/d1/observability/billing/>
- <https://developers.cloudflare.com/analytics/graphql-api/getting-started/authentication/api-token-auth/>

## Safety model

No single measurement is treated as a complete safety guarantee. Production
work requires all four layers below.

### 1. Query-shape safety

Every permitted canary query must have an indexed predicate and a small hard
result ceiling. No gate may rely on remaining quota to make an unbounded query
safe. A single request must be incapable of consuming a material fraction of
the daily allowance.

### 2. Account-wide budget evidence

Admission uses all-database usage from 00:00 UTC to the present, reconciled
against Billing > Billable Usage when Cloudflare exposes D1 there. On the
current Workers Free account that page exposes only R2. In that documented
case, admission uses settled account-wide GraphQL totals and requires their
daily, five-minute and query-fingerprint aggregates to reconcile. Per-database
charts are diagnostic only.

### 3. Capability isolation

Inspection, execution, publication, reading and routine ingestion have
separate default-off controls. Enabling one cannot enable another.

### 4. Immediate non-D1 rollback

Closing a canary uses Pages configuration and deployment controls, not a D1
query. The rollback remains available even after D1 is exhausted.

## Phase A — establish account-wide observability

Complete this phase without querying application databases.

1. Enumerate every D1 database in the Cloudflare account through the control
   plane. Record database ID, name, environment, Pages/Worker binding and known
   caller. Include Production, Preview, abandoned experiments and any database
   belonging to another project in the same account.
2. Inventory every possible caller: Pages Production and Preview deployments,
   Workers and cron triggers, GitHub Actions, watchdogs, queues, browser polling,
   local Wrangler commands using `--remote`, and retained branch previews.
3. Create a narrowly scoped Cloudflare API token with only **Account Analytics:
   Read**. It must have no D1 Edit, Pages Edit or Worker Edit permission. Store
   it outside the repository and, if automation needs it, as a dedicated secret.
4. Add a read-only budget command that queries `d1AnalyticsAdaptiveGroups`
   from 00:00 UTC to the current time without a `databaseId` filter, groups by
   database ID, and sums `rowsRead`, `rowsWritten`, `readQueries` and
   `writeQueries` across every returned group.
5. Query `d1QueriesAdaptiveGroups` for the same interval to attribute the top
   query fingerprints by rows read and written for each database. Parameters
   remain unavailable and must not be reconstructed from logs.
6. Compare the summed GraphQL total with Billing > Billable Usage when D1 is
   exposed. Record both, their timestamps and any stated reporting delay. When
   this Free account exposes no D1 product, record that exact limitation and
   require the GraphQL daily aggregate to reconcile with settled five-minute
   buckets and summed query fingerprints. Do not proceed while totals
   materially disagree, attribution is incomplete or any database/caller is
   unidentified.
7. After analytics have settled, reconstruct 7 September from 00:00 UTC through
   the 04:55 UTC quota failure. Attribute the missing usage by database and query
   fingerprint. If account totals still cannot explain enforcement, open a
   Cloudflare support/community report with timestamps and sanitized evidence;
   do not run another Production experiment to investigate it.

Phase A gate: all databases and callers are named, the 7 September consumption
is attributable, and an account-wide budget report can be produced without a
D1 query. If any of these are false, rollout remains blocked.

## Phase B — make the budget gate fail closed

Implement and test locally before any further Production canary.

1. Produce a machine-readable budget report containing:
   - UTC interval and generation time;
   - account total rows read and written;
   - per-database totals;
   - top query fingerprints;
   - unknown/unattributed usage;
   - observation age, aggregation/attribution differences and any available
     Billing reconciliation difference; and
   - explicit `GO` or `STOP` with reasons.
2. Treat missing credentials, API errors, pagination truncation, stale data,
   an unrecognized database, an unattributed spike, disagreement between
   GraphQL aggregates or disagreement with available Billing as `STOP`.
3. Take two account-wide measurements at least ten minutes apart. Use the higher
   total and calculate the intervening burn rate. Never subtract usage.
4. Reserve four million daily reads and 80,000 daily writes for ordinary app
   availability. Optional rollout work may use at most one million reads and
   20,000 writes on a day, including schema/index work and verification.
5. Do not begin optional work unless:
   - account reads are below 500,000 and writes below 10,000;
   - at least two hours of post-reset passive operation have elapsed;
   - no query fingerprint has examined more than 10,000 rows per invocation;
   - the projected ordinary-use total remains below the four-million reserve;
   - the proposed gate plus a 100% contingency remains inside the optional
     budget; and
   - both measurements reconcile internally and with Billing when D1 Billing
     is available.
6. Stop optional work for the day at 750,000 reads or 15,000 writes, or on any
   unexplained increase, even if Cloudflare has not enforced its limit.

These conservative thresholds are intentional. This service has no paid-plan
escape route and login depends on D1 availability.

Phase B gate: local tests prove that absent, stale, partial or contradictory
analytics always return `STOP`; no test contacts D1.

## Phase C — separate inspection from execution

Replace the current broad maintenance coupling before another canary.

1. Add `FACILITY_BOOTSTRAP_INSPECTION_ENABLED=false`. Inspection may be enabled
   while the shared emergency pause remains true.
2. Add `FACILITY_BOOTSTRAP_EXECUTION_ENABLED=false`. `execute:true` must return
   before D1 unless this setting is explicitly true.
3. Add an exact `FACILITY_BOOTSTRAP_FILE_ALLOWLIST`, containing one file ID—not
   merely an ED. Both inspection and execution reject every other file before
   D1.
4. Keep `ROSTER_ADVANCED_MAINTENANCE_ENABLED=false` for inspection. Execution
   requires the dedicated execution switch, advanced maintenance, the exact
   file allowlist, source allowlist and an unexpired matching plan revision.
5. Inspection performs only the four existing exact-key metadata/marker reads.
   It must never read `roster_events`, `roster_file_doctors` or any collection.
6. Add a separate inspection and execution workflow. The inspection workflow
   can never send `execute:true`. The execution workflow requires an exact plan
   revision and a separately approved manual dispatch.
7. Capture HTTP status and body safely before parsing. A Cloudflare quota error
   must be reported as such, not obscured by a secondary `jq` failure.
8. Make workflow output include deployment commit, effective non-secret flags,
   exact file ID, estimates and plan revision. Never output credentials.
9. Require an explicit Wrangler deployment for configuration changes and then
   download/read back effective Production variables. An automatic Pages build
   is not configuration evidence.

Phase C tests must prove zero D1 calls for every disabled, malformed, wrong-file,
wrong-source, expired-plan and unauthorized request. A supplied revision whose
underlying compact state has changed requires only the four bounded inspection
reads before returning `bootstrap-plan-changed`; detecting that state change
without reading its markers is impossible. Enabling inspection alone must
never authorise `execute:true`.

## Phase D — reset-day passive baseline

After a future 00:00 UTC reset:

1. Do not perform a test D1 query merely to see whether the reset happened.
2. Produce account-wide GraphQL measurements near 00:10, 00:20, 01:00 and
   02:00 UTC while all optional functionality remains paused. Reconcile with
   Billing only if Cloudflare begins exposing D1 there.
3. Attribute ordinary login, calendar, contact and colleague-query fingerprints.
   Any legacy Staff/status query, whole-file event count, unindexed scan or
   unexpected Preview activity is a blocker.
4. Calculate rows per request and the observed hourly burn rate. Distinguish
   actual rows from estimates and rows returned.
5. If the Phase B gate does not pass at 02:00 UTC, make no rollout request that
   day. Diagnose from analytics only.

Phase D gate: two hours of attributable, low, stable account-wide usage and all
Phase B thresholds pass.

## Phase E — one-file MMC canary

The selected first file remains:

```text
source: mmc
source id: monash-adults
file: AdultTerm3.2026.xlsx
file id: automation:monash-adults:47c0951dd2582465d59b19a9
```

Its read-only metadata inspection on 7 September examined two rows and wrote
zero. It is active, its retained source exists, and neither compact coverage nor
status summary exists. That historical result does not authorise execution.

Run the canary as separate gates:

1. Record a passing account-wide budget report and effective fully paused
   Production settings.
2. Deploy only the inspection switch plus the exact MMC/file allowlists.
   Keep emergency pause true and all execution/read/ingestion settings false.
3. Read back the effective settings. Run the protected inspection exactly once.
4. Record its per-query metadata. Expected maximum: four exact-key reads, no R2,
   no writes, and no event/doctor-table query.
5. Disable inspection and empty the file allowlist through an explicit
   configuration deployment. Wait at least 30 minutes for analytics to settle.
6. Reconcile account-wide totals. Stop if the observed delta is unexplained or
   exceeds 100 rows read or zero rows written.
7. Separately approve execution. The reviewed ceiling is approximately 27,021
   examined rows, 750 compact mutation statements, conservatively 2,250 indexed
   row writes, one summary row and no R2 operation.
8. Open execution for the exact file only, execute once with the fresh plan
   revision, and immediately redeploy all maintenance settings closed.
9. Verify from returned per-query metadata and, after the analytics delay,
   account-wide totals. Do not issue a D1 verification query while the delta is
   unexplained; the executor's result and analytics are the first evidence.
10. Only after the delta reconciles may one small exact-key verification query
    confirm that file's summary and compact revision.

No user test is required during bootstrap. The user will be asked to test only
after a shared publication exists and the Creator-only reader is explicitly
enabled for one hospital.

## Stop and rollback rules

Immediately close all optional controls, without querying D1, on any of:

- a quota error;
- account totals that are unavailable, stale or inconsistent;
- a database or caller not present in the inventory;
- a query not present in the approved fingerprint list;
- more than 10,000 rows examined by an ordinary request;
- any inspection write;
- any event/doctor-table read during inspection;
- execution beyond its declared ceiling;
- optional daily budget thresholds; or
- inability to prove the exact deployed configuration.

Rollback is an explicit Wrangler configuration deployment followed by a
control-plane read-back. Do not rely on an automatic Git deployment. Do not use
a D1 query to prove rollback.

## Current evidence ledger

- `3ee4bff`: bounded status code deployed.
- Migrations `0026` through `0031`: applied individually; no migration remains
  pending. New compact tables were empty immediately afterward.
- `9927321`: paused Admin Files state clarified.
- `8fc5f4f` / deployment `9ce41977-0b6e-45a1-bfbc-741f9003b0f8`: MMC
  maintenance configuration explicitly deployed.
- 7 September 2026 04:55 UTC: protected inspection failed with Cloudflare's
  account daily row-read-limit error before returning a plan; no bootstrap
  execution was requested.
- `a293251` / deployment `5ac63a19`: full pauses restored and effective
  Production variables read back without D1.
- Account-wide GraphQL later reported 9,064,320 Production rows read and 1,550
  rows written from 00:00 through approximately 06:30 UTC; Preview contributed
  no reported usage. Of the reads, 8,870,469 were assigned to 04:00–04:59 UTC,
  concentrated between 04:35 and 04:49. Query fingerprints explained only
  837,929 daily reads, leaving 8,226,391 unattributed. This discrepancy is a
  hard rollout blocker.
- The Free account's Billable Usage dashboard exposed only R2 products. The
  implemented checker therefore supports a named Free-plan Analytics-only mode
  but does not weaken reconciliation: daily, five-minute and query-fingerprint
  totals must agree before a sample can be valid.

The next authorised activity is Phase A analytics and caller inventory only.
No further D1 request is permitted until Phases A–D pass.

## Implementation checkpoint — 7 September 2026

Implemented locally, without querying D1 or changing Cloudflare state:

- a read-only GraphQL account-budget command and offline fail-closed tests;
- account, per-database and query-fingerprint aggregation;
- available-Billing reconciliation or explicitly documented Free-plan mode,
  observation settling, daily/five-minute/fingerprint reconciliation,
  two-sample/burn-rate checks, conservative thresholds and contingency;
- an explicitly incomplete database/caller inventory that prevents `GO`;
- separate inspection and execution controls, an exact one-file allowlist,
  ten-minute plan expiry and separate manual workflows;
- default-off Production, Preview and local-development configuration; and
- safe workflow capture of HTTP status and non-JSON quota responses.

The focused quota, materialisation, rollout, local-isolation, synthetic-cost
and representative fixture suites pass. A credential-free budget invocation
returns `STOP` as designed.

Not yet complete and not authorised:

- a control-plane-complete D1 database/deployment/caller inventory;
- creation/use of the narrowly scoped Account Analytics Read token;
- live reset-day account-wide Analytics reconciliation, plus Billing only if
  Cloudflare exposes D1 there;
- attribution of the 7 September incident;
- the reset-day passive baseline; and
- any deployment, configuration change, D1 inspection, bootstrap, publication
  or reader activation.

Therefore the rollout remains blocked at Phase A. None of the locally complete
work is evidence that Production can safely be exercised yet.
