# Clinician access release preflight — 3 October 2026

Decision: **STOP for migration/deployment now; recheck later today.**

Read-only Cloudflare account analytics collected at 11:34 am AEST, covering
settled usage from the UTC reset at 10:00 am AEST through 11:19 am AEST:

- Account-wide rows read: 8,293.
- Account-wide rows written: 16.
- Daily totals and five-minute totals reconcile exactly.
- Largest five-minute bucket: 1,901 reads / 6 writes.
- No sampled query exceeds the checker's 10,000-read average-per-query threshold.
- The live control-plane inventory contains exactly the three configured D1
  databases. Only the roster production database reports usage in this interval;
  absence of reported Preview/exam-tutor usage is not a claim about later work.

The existing checker marks this sample valid, but returns STOP solely for:
`two-hour-passive-baseline-required` and `second-sample-required`.
The account has ample observed quota headroom; it is not quota-exhausted.

Use this sample for a later measurement at least ten minutes apart. A recheck
around **12:20 pm AEST today** gives settled analytics covering the first two
hours of the UTC quota day. It must still reconcile and show an acceptable
burn-rate projection; elapsed time alone is not approval.

Before applying migration 0038, also perform bounded production preflight under
an admitted budget: confirm the migration ledger/schema, establish whether the
new index is already present, bound the compact membership table's size and
index-building read/write cost, and confirm a recovery point. Apply only the
reviewed migration, then verify the ledger and usage before application release.
No application database SQL, migration, schema change, or deployment was executed
for today's usage check. No migration-cost estimate has yet been established
from the live table.

Evidence files in this session:
- `/private/tmp/clinician-d1-usage-first.json`
- `/private/tmp/clinician-d1-inventory-live.json`

The analytics credential was obtained through the existing macOS Keychain
integration; no credential was written to these reports. Billing reconciliation
uses the project's documented Free-plan analytics-only mode, where D1 Billing
usage is unavailable. Current Cloudflare daily quotas are documented at:
https://developers.cloudflare.com/d1/platform/pricing/
