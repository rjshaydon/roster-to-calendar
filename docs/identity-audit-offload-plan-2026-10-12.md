# Identity audit offload — 12 October 2026

Status: comparison/export preparation and internal durable checkpoint primitives
are implemented and tested. Migration 0044 has NOT been applied to production.
The authenticated API, runner integration, automatic dispatch and rollout are not
enabled. Existing
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

1. **Connect the checkpoint protocol.** Internal pin/apply/finish primitives and
   migration 0044 are prepared (details below). Authenticate internal export and
   write callbacks, add the lease/dispatch lifecycle, and apply the migration only
   after the complete protocol is reviewed.
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


## Additional checkpoint preparation before the reset

`functions/_lib/identity-audit-checkpoints.js` pins immutable manifests, applies
five-proposal or sixteen-feature batches, and records receipts/cursors in the
same transaction as the facts. It guards the current identity revision, each
alias's cached fingerprint, active-person links, primary roster-import priority
and cursor races. Feature completion cannot start before every proposal batch
commits; audit completion checks both receipt counts, unique feature coverage,
current fingerprints and the completion cursor. Missing/deleted evidence fails
closed. Duplicate callbacks perform no writes. These primitives only update
suggestions, audit flags and checkpoint tables, never roster grades/calendars.

Migration 0044 contains three empty tables for jobs, receipts and unique feature
checkpoints. It is prepared only: no production schema/data writes were performed
in this step and no public endpoint invokes the primitives.

`npm run test:identity-audit-checkpoints` passes injected receipt and completion
failures, complete transaction rollback, lost-response replay, immutable pins,
evidence/revision/cursor races, deleted evidence, concurrent human rejection,
source-import priority and quota deferral. SQLite metadata here is simulated and
does NOT establish Cloudflare billing or CPU costs. Native cold/warm CPU, runtime
statement budgets and reservation envelopes still need end-to-end verification.
Large descriptor pinning may need further subdivision after those measurements.

Completion also verifies the actual bounded cache size and the indexed pending
feature queue: a feature becoming pending without a global revision bump cannot
be silently missed. A full initial audit must include every cached name. These
checks are part of the final guarded transaction, not a stale preflight decision.

Checkpoint implementation is now ready for the authenticated API/runner connection
and native CPU testing after the reset. Migration 0044 remains unapplied and the
new path remains inactive; this additional preparation performs no production
reads/writes. Existing live auditing continues on its previous path.
