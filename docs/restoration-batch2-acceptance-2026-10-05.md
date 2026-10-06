# Batch 2 acceptance — 5 October 2026

Status: implemented and committed on `codex/restoration-batch2`; isolated cloud Preview accepted. Production activation awaits the separate approval required by the doctor identity plan.

## Delivered feature set

- Creator review of preferred full names, canonical IDs, source aliases and account links.
- Read-only previews; versioned atomic merge, name correction, ID redirect, alias move and account linking; durable history and exact reversal. Later dependent changes block reversal rather than receiving a partial undo.
- Explicit named confirmation for multiple accounts. Accounts, subscription tokens, source shift rows and historical grades remain separate and unchanged.
- Indexed people search and bounded suggestions, durable rejection, obsolete-suggestion suppression, site-scoped manual audits and resumable progress. A single manual start processes up to 250 identities through short sequential requests and can be stopped after its current batch.
- Formatting-only automatic alias linking, with account conflicts and alias limits enforced. Spelling and nickname variations require Creator review. Existing approved associations are preserved.
- Exact affected-owner snapshot invalidation, retryable publication jobs, preferred names and approved aliases in personal/profile calendars and existing subscription feeds. Reassignment filters old cached shifts; an empty roster returns a valid empty subscription.
- Registration checks published metadata; unchanged metadata performs zero D1 work. Weekly suggestions use the indexed changed-feature queue, defer during imports, and begin Sunday 03:30 Melbourne time, with resumptions until 04:30. Candidate auditing creates no roster imports, snapshot jobs or R2 writes.

## Safeguards

One operation: at most eight existing people, 32 aliases, 16 accounts, 64 redirects; a calendar has at most 16 approved aliases. Transactions have a 400-statement ceiling. Discovery processes at most 25 identities and stops between identities after ten seconds. Each blocking group has at most 24 neighbours. All mutation/publication requests use existing account-wide maintenance admission and reservations in production. Runtime schema creation is absent.

The public identity revision is an opaque wake-up marker. It can wake unaffected open tabs; database snapshot invalidation remains limited to exact affected owners. Phone calendar providers retain their own subscription-fetch schedules.

## Verification

Passed: identity operation suite; bounded identity with 100,001 historical events; restoration mutations; Batch 1 freshness; watchdog; D1 quota; facility access. Identity tests cover injected transaction failures, stale previews, legacy-claim races, every edit type and reversal, reserved IDs, multi-account confirmation, rejected suggestions, indexed discovery, unchanged audit, feed continuity, unchanged historical grades and removal of earlier cached identities.

Cloud acceptance used a separate database (`741ce534-eafe-4aa7-9305-7b7ed5661979`) and bucket (`roster-identity-batch2-preview`), containing synthetic records only. The older Preview database has an incompatible parked identity schema and was left intact. No production migration or identity mutation was performed.

The existing synthetic subscription URL delivered both aliases after merging and removed the moved alias after exact reversal. Internal refresh retries completed. Registration and auditing processed 124 published aliases; the repeated audit examined zero identities. Browser verification confirmed Creator search, merge preview/approval, history, completed publication and reversed operations. Desktop layout was checked and corrected; mobile uses a single-column form layout.

Measured cloud D1 costs below include authentication where applicable; identity request metadata was complete:

| Request | Rows read | Rows written |
| --- | ---: | ---: |
| Merge preview | 28 | 0 |
| Merge commit | 51 | 33 |
| Exact reversal | 63 | 33 |
| Initial registration batch, maximum observed | 83 | 381 |
| Suggestion batch, maximum observed | 1,456 | 39 |
| Unchanged manual audit | 5 | 11 |

Registration batches took 4.8–5.0 seconds in the engine; suggestion batches took 3.8–4.0 seconds. The unchanged audit took 140 ms. Its small writes are run-history bookkeeping, not identity or roster changes. Existing-feed requests used 10–12 D1 statements with no writes; their row metadata was partial, so they are not presented as complete row-cost measurements.

Final settled account-wide safety sample: **GO**, 402,717 reads and 3,885 writes; projection approximately 598,359 reads and 14,049 writes for the UTC day. Analytics has a settlement delay; this sample includes only settled test usage. Additional synthetic test requests remain bounded by the measurements above.

## Production activation proposal

Production preflight found the original compatible identity schema and 48 people. Its migration ledger ends at 0038; the 0039 term policy is already applied (zero mismatched term rows), so record that verified baseline before applying 0040. The new migration adds identity metadata and indexes without roster/event backfill. Existing identity aliases and names are not overwritten by the migration.

After explicit approval: apply only migration 0040 with the existing schema/ledger baseline checked; release this branch; enable identity review and bounded registration; enable the weekly audit in both Pages and the existing watchdog; verify credentials and budgeted requests; initialize the published registry in bounded batches; check subscriptions and usage. Approve real spelling/nickname merges individually through the new interface.

Rollback: disable the three identity flags in Pages and the two watchdog flags, restore the preceding application release, and retain additive schema/history. Reverse any subsequently approved real identity operation through its exact preview before retiring that feature.

The implementation is ready for staged activation; real clinicians and their phone subscriptions have not been merged as a test. Richer historical evidence, upload/date-range audit filters, approved-alias expansion for Creator-selected subscription feeds (ordinary-user feeds are covered), and identity-dispute workflows remain follow-up work, with disputes/account lifecycle scheduled in Batch 6.
