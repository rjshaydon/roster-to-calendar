# D1 account-budget gate usage

This command reads Cloudflare Analytics only. It does not connect to an
application D1 database. Keep all rollout and maintenance controls closed while
collecting samples.

The command refuses to query an interval before the 15-minute settlement window
has entered the current UTC quota day. A run before 00:15 UTC returns `STOP`
with `analytics-settlement-window-before-utc-day`; it cannot mistake the
previous day for a completed passive baseline.

## Prerequisites

1. Complete `config/d1-database-inventory.json` from the Cloudflare control
   plane and set `complete` to `true` only after every database and caller is
   identified.
2. Prefer an API token containing only **Account Analytics: Read**. If the
   current Cloudflare interface offers only permission templates, its **Read
   only** template is acceptable when the broader read scope is recorded. The
   token must contain no edit permission and must not be saved in this
   repository or pasted into shell history. The checker uses
   `CLOUDFLARE_ACCOUNT_ANALYTICS_TOKEN` when present, then falls back on macOS
   to the existing Keychain service `roster-d1-account-analytics`. Codex must
   execute the checker with host access; a sandboxed Keychain failure is a
   tooling failure and must not be handed back to the operator as missing setup.
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
  --raw-output /private/tmp/d1-budget-raw-1.json \
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

That mode relies on account-wide GraphQL only. Daily totals must reconcile with
settled five-minute buckets; those authoritative totals drive quota, burn-rate
and projection decisions. Query fingerprints are independently sampled and
therefore are not required to sum to the quota total. They remain mandatory for
query-shape review: a truncated response, unknown database, unreviewed canary
query or fingerprint averaging more than 10,000 examined rows returns `STOP`.

Controlled Production work additionally requires route/request attribution.
The canary's operation ID and request-local statement/row counters must match
the reviewed action and ceiling. Its five-minute bucket must remain below the
greater of four times the passive-baseline maximum or the declared canary read
ceiling plus 10,000 rows; any bucket above 100,000 reads is a hard `STOP` for
the current At a glance rollout. Baseline mode has no controlled operation and
therefore does not require an operation ID; post-action canary mode does.

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
daily/timeline reconciliation, an unknown or unsafe query fingerprint, missing
controlled-request attribution, an unexplained bucket spike, or conflicting
available Billing data authorises no D1 work. A difference between summed
fingerprint rows and the authoritative usage total is reported as sampling
diagnostics and is not, by itself, `STOP`. The command itself never deploys,
changes configuration, runs a migration or calls an application endpoint.
Analytics are queried only through a settled cut-off 15 minutes behind the
report generation time.

## Post-action canary sample

Baseline sampling uses the default mode above. After one separately authorised
controlled action, export that request's `roster_api_invocations` record to a
mode-`0600` JSON file with this shape:

```json
{
  "requestId": "the-captured-cf-ray-or-request-id",
  "d1Statements": 18,
  "d1RowsRead": 27000,
  "d1RowsWritten": 200,
  "d1Limit": 768,
  "d1MetadataComplete": true
}
```

Store the exact reviewed SQL strings that may first appear during the canary in
a second JSON file as either an array or `{ "queries": [...] }`. Then run:

```sh
npm run d1:budget -- \
  --mode canary \
  --previous /private/tmp/d1-budget-pre-canary.json \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --canary-read-ceiling 27021 \
  --request-id the-captured-cf-ray-or-request-id \
  --request-attribution /private/tmp/d1-canary-request.json \
  --reviewed-fingerprints /private/tmp/d1-canary-queries.json \
  --output /private/tmp/d1-budget-post-canary.json
```

Canary mode fails closed when the request evidence is absent or mismatched, its
counters exceed their ceilings, a newly appearing query is not in the reviewed
manifest, or the settled canary bucket exceeds its derived envelope. Files may
contain sensitive operational metadata and must not be committed.

`--raw-output` is optional and stores the GraphQL response with file mode
`0600`. It contains neither the API token nor an authorization header.
