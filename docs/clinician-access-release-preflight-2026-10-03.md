# Clinician access release preflight — 3 October 2026

Current decision: **Migration 0038 applied and verified at 1:01 pm AEST.
Application release remains pending. The historical DDH pilot is queued;
a manual maintenance run requires approval because it can process other jobs.**

Initial decision at 11:34 am: STOP until a second sample and sufficient passive
baseline were available.

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

## Scheduled 12:20 pm AEST recheck and migration preflight

The second account-wide analytics report was generated at 12:20:47 pm AEST,
covering settled usage through 12:05:47 pm. It returns **GO**, `sampleValid=true`,
with no stop reasons. The saved 11:34 am report supplies the same-day baseline.

- Settled account usage: **11,506 reads / 31 writes**.
- Observed burn rate: approximately 4,151 reads / 19 writes per hour.
- Projected UTC-day usage: approximately 101,391 reads / 451 writes.
- Daily totals and five-minute totals reconcile; no expensive query fingerprint
  exceeds the checker's average-per-query limit.
- A new live inventory check again matches exactly the three configured databases.
  Production is on the current D1 storage backend and is approximately 80.7 MB.
  Rolling 24-hour database statistics are not used as quota-day totals.

Bounded read-only production SQL confirms:

- Migration ledger contains 0001 through 0037; 0038 is absent.
- All six prerequisite tables (including the migration ledger) have the expected
  columns. The new identity/term index and all ten proposed cache triggers are
  absent; there is no partially applied migration 0038 to reconcile.
- `facility_term_staff_contributions` contains **1,295 rows**, established with
  a 5,001-row capped count. `facility_access_sessions` contains **8 rows**,
  established with a 501-row capped count. Neither cap was reached.
- No current-day account-budget entry or maintenance reservation receipt is
  present at the time checked. This does not exclude new work starting later.
- Preflight SQL consumed **1,523 reads / 0 writes** according to response metadata.
  No clinician names or roster events were retrieved.

Cost assessment deliberately budgets **13,950 reads and 13,950 writes** for
migration/index/verification work: ten times the compact row count plus 1,000
rows for metadata and verification in each metric. This is a conservative
planning estimate, not measured migration usage. Adding the 1,523 unsettled
preflight reads and a separate allowance of 10,000 reads / 1,000 writes for
unsettled traffic gives 25,473 optional reads / 14,950 optional writes. The
existing checker applies its further 2x contingency; projected traffic plus
these allowances remains below the 4,000,000-read / 80,000-write maintenance
stop thresholds. The resulting assessment is **GO** with no reasons.

Retrieved a Time Travel bookmark during preflight:
`00006216-00000000-000050f9-e376c41d413e6011dc76e86ba8db6b12`.
This records a recovery point; no restore was performed. Retrieve a fresh
bookmark immediately before any later migration. A full database restore would
also roll back intervening normal writes; prefer preserving the additive schema
and reverting application behaviour if a release needs rollback.

Local SQLite rehearsal used the exact committed migration and the inspected
table/index definitions with 1,295 synthetic membership rows. Applying the SQL
twice is idempotent, creates one index and ten triggers, preserves membership
rows and does not delete existing cache rows merely on installation. Each of
the ten intended staff/grade/SMS/activation mutations clears the cache; a no-op
activation preserves it. The identity query uses the new covering index.

**Outcome:** ready to apply only reviewed migration 0038 within today's admitted
budget, subject to a fresh check if timing or concurrent work changes materially.
This scheduled follow-up performed preflight only: no production DDL, migration,
data mutation, branch switch, merge or application deployment was executed.
The combined restoration code still needs implementation/validation before its
release, as described in `docs/at-a-glance-restoration-completion-plan.md`.

Additional evidence files:

- `/private/tmp/clinician-d1-usage-1220.json`
- `/private/tmp/clinician-d1-inventory-1220.json`
- `/private/tmp/clinician-d1-info-1220.json`
- `/private/tmp/clinician-d1-recovery-1220.json`
- `/private/tmp/clinician-d1-schema-1220.json`
- `/private/tmp/clinician-d1-preflight-counts-1220.json`
- `/private/tmp/clinician-d1-reservations-1220.json`
- `/private/tmp/clinician-d1-migration-estimate-1220.json`
- `/private/tmp/clinician-d1-migration-rehearsal-1220.json`

Current Cloudflare references:

- Pricing/DDL and index accounting: https://developers.cloudflare.com/d1/platform/pricing/
- Recovery bookmarks: https://developers.cloudflare.com/d1/reference/time-travel/


## Migration completion and restoration pilot

Following the user's instruction to continue, a fresh quota check at 12:59 pm
AEST returned GO with 25,152 account-wide reads and 31 writes. A fresh Time
Travel bookmark was recorded immediately before applying the migration.

Only the committed `0038_clinician_access_scope.sql` was placed in an isolated
migration directory. Its SHA-256 was
`dc210675b5c4aa756cb7d64f97bb1ec4c83ad097a91b96b22ae00ed334e8f922`.
The production ledger records application at `2026-10-03 03:01:04 UTC`
(1:01 pm AEST). Verification found the new index and all ten triggers;
the 1,295 staff contribution rows were preserved. Verification consumed
1,482 reads and zero writes. No other pending migrations were applied.

At 1:21 pm AEST, settled account analytics reported 32,271 reads and 1,340
writes across all configured databases. The pilot assessment remained GO with
an allowance of 180,000 reads and 500 writes plus the checker's 2x contingency.
The increase in writes includes the migration and ordinary traffic; it is not
an exact isolated measurement of migration cost.

A bounded DDH historical publication job for 4 May–2 August 2026 (91 dates)
was queued and verified pending. The queue write consumed three writes.
No other historical terms were queued. Automatic approval review rejected
starting the general `roster-maintenance.yml` workflow because it can process
other pending production jobs as well as the reviewed DDH pilot. That workflow
was not manually started. The existing scheduled maintenance may pick up the
queued pilot; completion has not been verified.

See `at-a-glance-restoration-checkpoint-2026-10-03.md` for application changes,
validation, publication gaps and remaining release work. No application deploy
or merge has occurred.

Private evidence (not committed):

- `/private/tmp/clinician-d1-before-migration.json`
- `/private/tmp/clinician-d1-recovery-before-migration.json`
- `/private/tmp/clinician-d1-migration-0038-apply.log`
- `/private/tmp/clinician-d1-migration-0038-verify.json`
- `/private/tmp/clinician-d1-history-pilot-budget.json`
- `/private/tmp/clinician-ddh-history-pilot-verify.json`
