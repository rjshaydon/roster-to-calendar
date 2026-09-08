# D1 reset-day passive runbook — 8 September 2026

## Boundary

This runbook observes Cloudflare Analytics only. It must not call an application
endpoint, run D1 SQL, deploy, migrate or change a variable. Reset is at 00:00 UTC
(10:00 AEST). Production contains zero-D1 maintenance safety commit `a146029`;
record the exact active documentation-only descendant, if any, before sampling.
Keep all optional capabilities closed.

Run commands from the repository root. The checker first uses
`CLOUDFLARE_ACCOUNT_ANALYTICS_TOKEN` when already present and otherwise loads
the existing read-only token from the macOS Keychain service
`roster-d1-account-analytics` without printing it. Codex must run the command
with host access because its default sandbox cannot read the user's Keychain.
The operator does not need to repeat token setup or export it manually.

The report's `credential.source` field must be `environment` or
`macos-keychain`; a null value is a checker-access failure, not evidence of zero
D1 use.

Every report is expected to say STOP until at least 02:20 UTC. Stop immediately
if a file is missing, `sampleValid` is false, an unknown database appears, daily
and five-minute totals differ, or usage has an unexplained increase.

## 00:20 UTC / 10:20 AEST — first settled sample

```sh
npm run d1:budget -- \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --raw-output /private/tmp/d1-2026-09-08-0020-raw.json \
  --output /private/tmp/d1-2026-09-08-0020.json
```

Expected blocking reasons include `two-hour-passive-baseline-required` and
`second-sample-required`. The report must nevertheless have `sampleValid:true`.

## 00:35 UTC / 10:35 AEST — first burn-rate comparison

```sh
npm run d1:budget -- \
  --previous /private/tmp/d1-2026-09-08-0020.json \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --raw-output /private/tmp/d1-2026-09-08-0035-raw.json \
  --output /private/tmp/d1-2026-09-08-0035.json
```

## 01:00 UTC / 11:00 AEST — passive checkpoint

```sh
npm run d1:budget -- \
  --previous /private/tmp/d1-2026-09-08-0035.json \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --raw-output /private/tmp/d1-2026-09-08-0100-raw.json \
  --output /private/tmp/d1-2026-09-08-0100.json
```

## 02:20 UTC / 12:20 AEST — earliest admission assessment

```sh
npm run d1:budget -- \
  --previous /private/tmp/d1-2026-09-08-0100.json \
  --billing-unavailable-reason free-plan-dashboard-omits-d1 \
  --estimated-reads 27021 \
  --estimated-writes 2252 \
  --raw-output /private/tmp/d1-2026-09-08-0220-raw.json \
  --output /private/tmp/d1-2026-09-08-0220.json
```

Inspect only the compact decision fields:

```sh
jq '{generatedAt,decision,sampleValid,reasons,interval,effectiveUsage,burnRateRowsPerHour,projectedReads,billingMode,unattributed:.analytics.unattributed,databases:.analytics.databases}' /private/tmp/d1-2026-09-08-0220.json
```

Proceed no further unless the command explicitly says GO and the figures have
been reviewed. A STOP result ends optional D1 work for the day.

## If the passive gate says GO

GO authorises review, not a D1 request.

1. Review local branch `codex/d1-pre-reset-readiness`, including the code-level
   default-closed flags and request-budget tests.
2. Confirm that publishing the branch cannot expose a retained Preview URL to
   automated or human traffic. Prefer an explicit reviewed deployment sequence
   over relying on Git-triggered configuration.
3. After separate approval, deploy the safety code with every optional setting
   still closed.
4. Download the effective Pages configuration into a temporary directory and
   verify the commit, bindings and non-secret flags. Do not overwrite the
   repository configuration.
5. Take another settled Analytics sample. Only then consider Phase E's one-file
   read-only MMC inspection.

The inspection itself is separately approved and is never combined with
execution.
