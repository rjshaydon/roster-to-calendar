# D1 account-budget gate usage

This command reads Cloudflare Analytics only. It does not connect to an
application D1 database. Keep all rollout and maintenance controls closed while
collecting samples.

## Prerequisites

1. Complete `config/d1-database-inventory.json` from the Cloudflare control
   plane and set `complete` to `true` only after every database and caller is
   identified.
2. Create an API token containing only **Account Analytics: Read**. Do not give
   it D1, Pages or Workers edit permission and do not save it in this repository
   or paste it into shell history. Expose it to the command as
   `CLOUDFLARE_ACCOUNT_ANALYTICS_TOKEN` using the operator's secret manager.
3. If **Billing > Billable Usage** lists D1, record its rows read, rows written
   and the UTC time at which you observed them. On the present Workers Free
   account the dashboard exposes only R2, despite Cloudflare's D1 documentation.
   Use the explicit Free-plan mode below; never substitute zero for unavailable
   Billing data.

## First sample

```sh
npm run d1:budget -- \
  --billing-reads 12345 \
  --billing-writes 123 \
  --billing-observed-at 2026-09-07T02:00:00Z \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --output /private/tmp/d1-budget-sample-1.json
```

The first sample returns `STOP` with `second-sample-required`; that is expected.
It is usable as the prior sample only when `sampleValid` is `true`. Refresh the
Billing figures immediately before each command. A Billing observation older
than 15 minutes is rejected.

When D1 is absent from the Billing dashboard, omit all three `--billing-*`
values above and use exactly:

```sh
  --billing-unavailable-reason free-plan-dashboard-omits-d1
```

That mode relies on account-wide GraphQL only and therefore applies additional
checks: daily totals must reconcile with settled five-minute buckets and with
the sum of query fingerprints. Missing attribution always returns `STOP`.

## Second sample

Wait at least ten minutes, then run:

```sh
npm run d1:budget -- \
  --previous /private/tmp/d1-budget-sample-1.json \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --output /private/tmp/d1-budget-sample-2.json
```

Only an explicit `GO` authorises consideration of the separately approved next
gate. `STOP`, missing output, an API error, an unknown database, incomplete
query attribution, or conflicting available Billing data authorises no D1
work. The command itself never deploys, changes configuration, runs a migration
or calls an application endpoint. Analytics are queried only through a settled
cut-off 15 minutes behind the report generation time.
