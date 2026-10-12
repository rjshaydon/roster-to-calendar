# Identity audit offload — 12 October 2026

Status: authenticated API, GitHub runner, dispatch lifecycle and cached UI updates
are implemented. The complete protocol passes local integration/regression tests.
Migration 0044 is applied to preview and production (three additive empty tables). Production enablement
remains gated by native full-size runtime/receipt verification and rollout checks.
Cloudflare Free and all existing D1 safeguards remain.

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

## Integration and native testing after the reset

The new authenticated `/api/automation/identity-audit` endpoint exports cached
names, pins the complete plan, applies bounded proposals/feature checkpoints and
finishes only after exact coverage. GitHub performs every name comparison; it
publishes no names, proposal bodies or credentials in public logs/artifacts.
Watchdog and creator checks coalesce through an R2 conditional-write lease.
Unchanged checks launch no workflow. Registration remains a separate pathway.

Each run currently compares the complete bounded cache, including already audited
names. This keeps plan hashes stable across partial feature completion/resumption;
audited flags are excluded from the snapshot hash, while all identity evidence
remains included and guarded. Changed evidence requires a new plan. Unchanged
human rejections survive the conditional proposal upsert. Confirmed bounded
pauses release the lease and preserve progress; unknown outcomes retain receipts
and the lease. A failed cache wake-up can replay an already-completed database job.

The UI continues to display cached suggestions immediately. While an external
check runs it refreshes the bounded saved page every 30 seconds only with the
section open and the document visible. Closing the section stops polling; editors
and selections remain intact. Deferred checks are described as deferred.

Initial native tests revealed 10–16 ms callbacks. Combining indexed evidence
guards/completion writes in native SQL and returning just the needed manifest
descriptor reduced captured callbacks to a maximum of 8 ms in the completed
124-name / 1,861-proposal dense preview run. No names were skipped. The larger
1,124-name test and production acceptance are still in progress. Test metadata
from SQLite is not billing evidence. Test Worker and synthetic fixture cleanup
are required before closeout.

`test:external-identity-audit` exercises the actual middleware, authenticated API
and Node runner: auth boundaries, dispatch coalescing/unchanged no-op, CAS races,
lost successful responses, completed-job wake failure, preserved rejection,
bounded yield/resumption and stale revisions. Planner/checkpoint, identity
operation, cached UI, request attribution, watchdog and account-budget regressions
also pass. Production flags stayed false through native acceptance. The activation commit
now enables the external path for production only; post-rollout verification
remains required.

Native acceptance: both the 124-name / 1,861-proposal and the 1,124-name /
1,861-proposal jobs completed with exact candidate/feature receipt cursors.
The optimized trace captured 565 successful callbacks, maximum 8 ms; fresh
full-size pin/start were captured. Not every callback was retained in the trace.
Actual combined preview receipts settled at 98,027 reads and 38,621 writes with
ZERO overruns or unfinished receipts. Synthetic seeding/cleanup are additional
small maintenance costs, not included in those callback totals.

Reservations now cover BOTH admission and transactional source-priority probes:
131,072 reads for write checkpoints and 65,536 for snapshot pages. These are
per-request held estimates, with unused capacity settled from actual metadata;
they do not change Cloudflare's daily allowance or discard unknown work.

Real production delivery evidence: workflow 38178709371 completed the DDH
2 November–31 January roster (109 doctors, 3,356 events). The recorded run
`f9338120-709a-4ea0-af36-a12f6f9344dd` has success/completion metadata.
Other site jobs were safe no-work runs, so they do not prove a fresh changed
workbook delivery. Production audit coverage, CPU traces and feeds still need
post-enablement verification before Batch 2 closeout.

Production activation deployed as `993abcf4` / commit `b40c2859`; watchdog
`e0583476` uses the external check route. Workflow `38201205790` safely deferred
before audit publication. Name registration is still active. Production admission
has zero optional write headroom despite low actual usage: the forecast had only
production receipt summaries, so one-off preview test costs were projected as
recurring ordinary account traffic. No limit or forecast was manually bypassed.

The follow-up adds a production binding used ONLY for the preview's fixed,
read-only receipt summary. The shared request meter covers both database reads.
Verified settled costs are checked against that database's analytics before they
leave the ordinary forecast; actual account totals remain untouched, late/unknown
costs remain held, and inconsistent/missing evidence fails closed. Preview/local
configurations do not enable this feature. No extra roster/identity mutation path
or daily-limit change is introduced. Native verification and the completed
production audit remain required. `test:account-maintenance-summaries` covers these
bounds, cross-database metering and restoration of forecast headroom solely from
verified one-off costs.

## Cold-route reliability follow-up

The forecast correction is deployed as `53a1b0a4` / Pages `8b756374`. A fresh
grant increased read headroom from 1.87M to 2.28M, but optional write headroom
remained zero. Actual usage/caps and held reservations were unchanged. No
production audit completion was claimed. Intermittent subscription resource-limit
503s and a 22 ms cold legacy registration sample remain acceptance failures.

The follow-up removes the feed's import of the full state handler for email
normalisation, preserving the same normaliser/re-export. Audit, term-date and
shift-launch formatters are initialised only when used and then reused. Cold
imports of feed/state/registration/account-budget now construct zero formatters.
The two sampled sessions are only 1,377 and 1,770 bytes: no large-session workaround
or data removal is justified by that evidence.

The GitHub registration runner now obtains account admission separately, as the
external audit already does, with a five-minute cache and fail-closed errors.
Server reservation/source-priority checks remain. Current tests pass cold import,
Melbourne/DST boundaries, registration admission expiry/error handling, identity
merge/reversal/subscription continuity, account summaries and attribution.
Native fresh/repeated request CPU and feed acceptance are still required.

First rollout `7d35531e` / Pages `87b3a725`: both sampled feeds returned their
exact baseline UID sets in the initial and eight repeated checks. Captured CPU
still ranged 10–30 ms for feeds and 17 ms for registration. Successful HTTP does
not establish Free-tier CPU reliability; Batch 2 remains open.

The next change consolidates subscription-token/account loading into one indexed
account lookup and exact site/name event/designation reads into two statements
in one D1 batch. It preserves active history, all aliases, historical grades,
per-site leave merging and token ownership. Existing fixtures verify complete
content parity, index plans, duplicate aliases and zero writes. Feed first-row
reads now retain billing metadata. Native repeat/cold CPU verification is pending.

Second rollout `15b4321d` / Pages `d781d1ed`: eight repeated subscriptions retained
exact baseline event IDs, with complete native D1 metadata and no writes.
Captured feed CPU improved to 7–17 ms but still exceeds the target in some samples.
The follow-up bundles independent hospital-location/custom-event reads into the
same bounded input batch. It preserves session fallbacks, personal custom events
and per-site identities. Active registration jobs now check source priority before
optional metadata/timezone preparation, while completed unchanged checks retain
their zero-D1 shortcut. These behaviors pass regression tests; native acceptance
of the final combined batch remains required.

The isolated, authenticated read-only feed Worker returned identical calendars
but still used 8–17 ms. A smaller Pages bundle alone is therefore not a sufficient
fix. It is deleted after diagnostics, with no roster or accounting mutations.
The next change consolidates active account/person/approved-alias resolution into
one indexed query, retaining the original retired/merged-person path and 16-alias
bound. This also reduces aliased On shift eligibility from five statements to two;
authorization/unapproved-alias regression checks pass.

Cloudflare documents a 10 ms Free HTTP limit with some per-isolate flexibility;
sustained excess can still terminate requests. Source:
https://developers.cloudflare.com/workers/platform/limits/#cpu-time
Thus successful HTTP responses above 10 ms remain insufficient for acceptance.
Final production sampling and full audit/registration completion remain open.
