# Subscription reliability implementation and verification

The external assembly implementation is prepared, with the production
`SUBSCRIPTION_ARTIFACT_MODE` set to `off`.
Existing subscriptions still use their existing request path. This is not full
Batch 2 acceptance or a claim that every historical 503 cause is fixed.

## Implemented together

- One indexed evidence query binds token ownership, account settings/overrides,
  date range, custom events and hospital locations, identity links/approved names,
  historical designations and active roster generations. Sentinels explicitly
  refuse oversized cache evidence; those subscriptions retain the complete
  existing dynamic path, with no historical truncation.
- GitHub/Node uses the existing renderer, event filtering/leave merging and UID
  algorithm. Exact authorized event exports use 64-row keyset pages. All pages
  must finish before publication; there is no partial calendar promotion.
- Private R2 calendars are content-addressed by token/view/input evidence.
  Ownership and inputs are checked again on every GET. Late publication after
  reassignment or an edit cannot make the previous calendar valid. Current
  settings/custom events/grades are preserved. No source roster is rewritten.
- Requests for missing/changed calendars coalesce in private R2; unchanged reads
  do not create work. The existing watchdog can dispatch the bounded runner.
  The `build` mode allows priming both existing subscription views before
  switching request reads to `ready`. Initially missing/stale objects retain complete current
  dynamic responses, never an empty stand-in or a previous identity's calendar.
- Exports and publication remain behind source priority and shared account
  reservations. Unknown outcomes stay held; explicit stale/yielded checkpoints
  release their dispatch lease. Logs contain bounded status/counts, never names,
  subscription tokens or event bodies. There is no new daily limit or paid plan.
- Grant reuse combines the existing unknown-receipt and unexpired-day checks in
  one compact query. Subsequent reservations still enforce ceilings and races.

## Verification performed

The integration test exercises the actual middleware, authenticated endpoint and
Node runner using isolated SQLite/R2 fixtures: 130 shifts across multiple pages,
old historical shifts, both URL views, complete ICS/UID/order parity, changed
settings/locations/custom events/names/roster generations, ownership reassignment,
revocation, stale exports, source-priority deferral before reservation, publication
failure/replay, evidence overflow fallback, coalescing and build-only priming.
Existing identity, cold-route, account-budget, external-audit, attribution,
watchdog and restoration regressions pass. SQLite costs are simulated, not billing
or native CPU proof.

An authenticated temporary Worker uses read-only production D1 and a separate
private R2 prefix. Both real subscriptions retain 128/154 events and their exact
baseline UID sets. Initial captured cached GETs use 1–2 ms CPU; the final fresh/repeated sample
uses 1–8 ms CPU (fresh samples 6/8 ms, warm samples 1–2 ms), one statement,
60/110 reads and zero writes, with complete native metadata. Read-only export
preparation uses 4–8 ms; event pages use 2 ms. External Node assembly again matches
both real UID sets. The combined grant reuse check uses 1–5 ms and one row read.
These export-stage measurements OMIT real reservation/settlement overhead; they
are not acceptance of the complete production publication callbacks.

## Remaining rollout gates

1. Measure complete callbacks with real reservations/settlement in isolated native
   acceptance, including fresh requests, publication failure/replay and large
   calendars. Split further if they exceed the Free allowance. Read-only stage
   measurements alone are insufficient.
2. Obtain normal optional write admission; enable build-only, prime existing views,
   and verify exact live fields/UIDs before enabling cached GETs. Deploy the
   watchdog's mode with the application mode. Avoid a first-request cold
   rollout while builders are deferred.
3. Complete cached-name registration and the external production identity audit,
   then verify actual changed-roster delivery for MMC/MCH/VHH. DDH next-term
   delivery already has evidence. A successful deferred Actions run is not delivery.
4. Verify real Pages CPU, receipt costs and subscription continuity; record final
   Batch 2 acceptance. Keep all current D1 safeguards.

At this pass's initial settled cutoff the account had approximately 1.05M reads /
48.6K writes. The current grant had only 477 writes of raw headroom, below the
unchanged 20K protection for source imports before optional work. MMC, VHH and
DDH have real queued imports; the sampled DDH Actions run explicitly deferred
for account allowance. These are operational gates, not a Codex usage shortage.
No reservation was manually released and no forecast/cap was overridden.

Deployment note: the first Git deployment compiled successfully but publication
failed because two additional text flags raised the total to 65 variables, above
the Free limit of 64 (secrets plus text). They are consolidated into one
`SUBSCRIPTION_ARTIFACT_MODE` with `off`, `build` and `ready` values. This preserves
the staged rollout while staying on Free; no existing secret/feature is removed.

## Complete callback canary and forecast attribution follow-up

A capped canary routes every calendar query to read-only production D1 and every
reservation/settlement mutation to the isolated identity database. Its private R2
prefix is separate from both live calendars and previous tests. No roster/person
fixtures are inserted. The test grant retains existing ledger allocations and
allows only 192 writes above the unchanged 20K optional protection; it cannot
change production allowances. The complete 128-event Node publication succeeds,
with native complete metadata: snapshot 11 ms (8 statements), event pages 9 ms
(7 statements each), publication 8 ms (6 statements). All known receipts settle.
An initial canary namespace error fails before reservation and is corrected.

Removing the repeated preparation evidence read still measured snapshot 11 ms
(7 statements), so success is NOT CPU acceptance. The protocol now pins the
small evidence snapshot separately from provisional alias preparation. Every
subsequent event page and publication still checks current ownership/input
revision. Regression tests inject an identity edit between pin and preparation
and reject the following page. Full optimized native verification remains needed.
Missing-start range exports also retain the original complete-history behavior.

Native CLI metadata explains an additional false recurring forecast component:
12,000 writes for the isolated 1,000-name seed and 5,200 writes / 9,190 reads for
its cleanup. Neither SQL file touched `roster_maintenance_receipts`. The immutable
proof records retain the actual response/query SHA-256 hashes, database ID and
observation time. They apply only to that database/UTC day, after the same settled
analytics cutoff; the combined receipt and CLI costs must remain below measured
per-database totals or admission fails closed. Actual account usage, unknown
reservations, caps and protected source room are unchanged. No grant or receipt
is refunded. Tests cover wrong day/database, cutoff ordering, reconciliation and
unchanged totals/reservations. The feature uses the existing preview-summary gate,
adds no new Worker variable, and is inert after 12 October UTC.
