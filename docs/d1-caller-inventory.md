# D1 database and caller inventory

Status: **control-plane inventory complete; retained Production deployments contained**. Updated 8 September 2026 without
an application D1 query. Runtime rollout remains blocked by the incident's
unattributed usage and the reset-day passive-baseline gate.

## Known databases

| Database | ID | Environment | Known binding/caller |
| --- | --- | --- | --- |
| `roster-converter-calendar` | `237d0d52-3a7c-4e02-8648-9f4dedbc1cb0` | Production | `roster-to-calendar` Pages Production `ROSTER_DB` |
| `roster-converter-calendar-preview` | `541d7779-a43a-4603-a4ab-c95e9b4b015d` | Preview | `roster-to-calendar` Pages Preview `ROSTER_DB` |
| `acem-exam-tutor-db` | `46829fee-f518-4d23-a992-d050f53c4c41` | Production | `acem-exam-tutor` Worker `env.DB` |

The read-only `wrangler d1 list --json` control-plane request succeeded at
08:50 UTC and returned exactly these three account databases. The third
database had been absent from the original inventory. `wrangler versions view`
confirmed the active `acem-exam-tutor` Worker version binds it as `env.DB`.
The account has one Pages project, `roster-to-calendar`.

## Known callers

- Pages Production and Preview API functions bind `ROSTER_DB` through
  `wrangler.toml`. The control plane listed retained deployment URLs from
  `main`, `codex/at-a-glance-d1-optimization` and
  `codex/durable-doctor-identity-aliases`; a request to any retained URL is a
  possible caller even when it is not the active Production deployment.
- The `acem-exam-tutor` Worker binds the account's third database. Its D1 usage
  must be included in every account-budget decision even though it is unrelated
  to the roster app.
- Browsers call the state, account-context, subscription and facility overview
  APIs. Their database access is executed by Pages, not directly by browsers.
- The GitHub Actions workflows `monash-roster-sync.yml`,
  `facility-bootstrap-canary.yml` and `facility-bootstrap-execute.yml` can call
  token-protected Production endpoints when manually or automatically run.
- `worker/roster-queue-watchdog.js` can call roster automation endpoints, but
  its deployed configuration was observed with `ROSTER_AUTOMATION_ENABLED=false`.
- Local development uses a local D1 binding and strips remote Cloudflare
  credentials. Direct operator use of `wrangler d1 ... --remote` remains a
  separate manual capability and is forbidden during this rollout unless a
  specific step is approved.
- Retained Pages branch deployments bind the Preview database, while retained
  Production deployment URLs bind Production. They are covered as caller
  classes and must be re-enumerated if the Pages project changes.

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

## 8 September retained-deployment audit

A control-plane-only `wrangler pages deployment list` found 16 retained
Production deployments and seven retained Preview deployments. Every retained
Production deployment has its own public hash URL and belongs to a commit that
predates the ordinary-login identity containment. The oldest listed Production
deployment is `1a8ee1fb-ac75-4078-83fe-ea155674be8e` (`cff238d`); the active
deployment is `def76e30-d612-4de0-8d99-37e7baf52a25` (`af951d0`).

This confirmed that switching the active Production alias did not make older
Production functions unreachable. It did not prove that an old URL was called
during either incident. After contained commit `18ba7a1` was deployed and its
effective settings were read back, all 16 listed pre-containment Production
deployments were deleted by exact ID. Two deployments of `18ba7a1` remain: the
Git-triggered deployment and the explicit configuration deployment. Git history
retains every deleted version. Preview deployments bind the Preview database
and are not candidates for the Production database burst; they remain for
separate housekeeping and cannot consume Production database rows.

## Completion gate

The database/caller discovery gate is complete as of the timestamp above. The
budget command still returns STOP for an unknown database ID, so any later
account addition invalidates the inventory automatically. Inventory completion
does not override incomplete Analytics attribution, thresholds or passive-day
requirements.
