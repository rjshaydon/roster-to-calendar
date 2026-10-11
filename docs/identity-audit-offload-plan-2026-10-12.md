# Identity audit offload — 12 October 2026

Status: comparison/export preparation implemented and tested. The external audit
protocol, automatic dispatch and production rollout are not enabled. Existing
production registration, audit and cached-suggestion behavior remain in use.
Cloudflare Free and the current D1 admission/reservation safeguards remain.

## Completed before the usage reset

- `scripts/identity-audit-plan.mjs` prepares complete suggestions in Node, sharing
  the existing evidence rules. It compares all relevant same-first-initial pairs,
  including dense groups and typos that cross the former surname-prefix buckets.
  No group is skipped and no output is silently truncated. The reviewed bounds
  are 2,000 names and 10,000 proposals; exceeding a bound fails explicitly without
  marking names audited or changing production data.
- The first external run must request a full audit with algorithm version
  `external-pairs-v1`. Previously audited fingerprints do not prove coverage under
  the new algorithm. Later runs can examine pending names against the full cache.
- Proposals have deterministic evidence fingerprints and are packaged in batches
  of at most five / 24 KiB. Feature-completion records are packaged separately,
  at most sixteen per batch. The root manifest contains descriptor hashes/counts,
  not a large embedded list of every name/completion record.
- `functions/_lib/identity-audit-snapshot.js` provides an internal read-only
  32-name keyset export. Revision and rows share a D1 transaction; a revision
  change between pages requires a restart. It reads cached features/identity links,
  with no roster-event history, grades or account-email reads. No new endpoint
  calls it yet.
- `functions/_lib/identity-suggestion-proposal.js` prepares the existing conditional
  upsert semantics for a future guarded transaction. Unchanged human rejections
  and their timestamps remain untouched; changed evidence is flagged for review.
  It is not connected to an enabled API and does not authorize an unguarded write.

Tests: `npm run test:identity-audit-plan` passes a 1,000-name indexed/read-only
export, 4,950 dense-group proposals, exhaustive pair-evidence parity, Jay/Jayantha,
Aeshan surname variants, same-person exclusion, unchanged no-op, complete initial
sweep, stable hashes, byte/count envelopes, stale/repeated cursor rejection,
registration inconsistencies and persistent human decisions. Existing identity
operation/registration regression tests pass.

A bounded production cached-name read followed by a LOCAL dry run examined all
1,013 cached features, compared 33,834 relevant pairs and prepared 97 proposals in
20 candidate batches in about 41 ms elapsed on the local computer. It used 3,040
D1 reads and ZERO writes. The 97 include existing/rejected proposals; they are
not 97 new unresolved suggestions. This timing is not a GitHub or Worker CPU
measurement. Nothing from the dry run was published to the app or public logs.

## Finish after the reset

1. **Pin and stage the audit protocol.** Add durable job/descriptor/receipt state
   through additive migrations. Authenticate internal export and write callbacks.
   Preserve source-import priority, optional-maintenance headroom, unknown receipts
   and bounded per-request costs. Pin the snapshot/algorithm/descriptor evidence;
   stage descriptors in bounded steps where necessary. Verify current identity
   revision and both aliases' feature fingerprints before applying proposals.
   No client-supplied SQL or automatic fuzzy merging.
2. **Finish names only after proposals.** Commit small proposal batches with their
   receipts atomically. Apply feature-completion batches only after every proposal
   batch has arrived. Mark the audit complete only after all feature receipts and
   guards succeed. Handle stale snapshots through explicit re-planning, and lost
   responses through durable replay; never skip names to obtain a success result.
3. **Connect the GitHub runner and dispatch.** Extend the existing identity workflow
   to export once, compute in Node, send bounded batches and finish. Registration
   stays separate. Coalesce due/name-change dispatches with a lease and bounded
   retries. Do not create jobs for unchanged checks. Use existing credentials;
   keep names/results out of public Actions logs, artifacts and Git.
4. **Keep cached UI results immediate.** The current cache already provides the
   initial suggestion list. Update it after committed results; preserve selections,
   rejected-suggestion review and existing account/calendar behavior.
5. **Acceptance and rollout together.** Test interrupted/stale/concurrent jobs,
   rejection races, full historical-name coverage, first algorithm migration,
   quota deferral and bounded cold/warm CPU. Enable the external path only after
   end-to-end verification, then confirm a real changed-roster delivery and
   unchanged subscription continuity before accepting Batch 2.

Medium effort remains suitable. No later restoration feature set is added here.
