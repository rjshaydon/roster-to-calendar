# D1 database and caller inventory

Status: **incomplete — rollout blocked**. Updated 7 September 2026 without an
application D1 query.

## Known databases

| Database | ID | Environment | Known binding/caller |
| --- | --- | --- | --- |
| `roster-converter-calendar` | `237d0d52-3a7c-4e02-8648-9f4dedbc1cb0` | Production | `roster-to-calendar` Pages Production `ROSTER_DB` |
| `roster-converter-calendar-preview` | `541d7779-a43a-4603-a4ab-c95e9b4b015d` | Preview | `roster-to-calendar` Pages Preview `ROSTER_DB` |

This list is deliberately marked incomplete. Wrangler's current OAuth session
was rejected by the D1 list control-plane API on 7 September. No database may
be assumed absent until a read-only account inventory succeeds and its result
is reconciled with account-wide Analytics and Billing.

## Known callers

- Pages Production and Preview API functions bind `ROSTER_DB` through
  `wrangler.toml`.
- Browsers call the state, account-context, subscription and facility overview
  APIs. Their database access is executed by Pages, not directly by browsers.
- The GitHub Actions workflows `process-monash-rosters.yml`,
  `facility-bootstrap-canary.yml` and `facility-bootstrap-execute.yml` can call
  token-protected Production endpoints when manually or automatically run.
- `worker/roster-queue-watchdog.js` can call roster automation endpoints, but
  its deployed configuration was observed with `ROSTER_AUTOMATION_ENABLED=false`.
- Local development uses a local D1 binding and strips remote Cloudflare
  credentials. Direct operator use of `wrangler d1 ... --remote` remains a
  separate manual capability and is forbidden during this rollout unless a
  specific step is approved.
- Retained Pages branch deployments may bind the Preview database. They must be
  enumerated through the control plane before the inventory becomes complete.

## 7 September correlation evidence

- Account-wide GraphQL attributed all reported usage to the Production database
  and none to Preview.
- No roster-processing GitHub Action ran during the 04:35–04:50 UTC burst.
- Two manually dispatched bootstrap inspection workflows ran at 04:53 and
  04:55 UTC, after the burst and after quota exhaustion; neither executed a
  bootstrap.
- Production and Preview deployments of `3ee4bff` existed before the burst.
  Several later Production-only configuration deployments occurred after it.
- Analytics showed repeated Production application-state query fingerprints
  during the burst. The fingerprint dataset accounts for only a small fraction
  of the account metric, so it cannot yet identify the dominant operation.
- Pages Functions Analytics recorded 395 Production requests during the three
  damaging five-minute buckets, compared with 4–10 requests in preceding
  buckets. Historical Analytics does not include their request paths.

## Completion gate

Set `config/d1-database-inventory.json` to `complete: true` only after all D1
databases, Pages deployments, Workers, schedules, queues and retained previews
are accounted for. Until then `npm run d1:budget` must return `STOP` even when
all supplied usage numbers are low.
