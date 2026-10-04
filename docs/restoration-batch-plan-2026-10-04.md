# Remaining restoration and development batches — 4 October 2026

Status: Batch 1 implementation in progress. App-side code and tests are complete;
production deployment and live SharePoint Flow acceptance are tracked below. The user's agreed priority is **1 → 2 → 6 → 7 → 3 → 4
→ 5 → 8**, retaining the original batch numbers for continuity. This plan
supersedes earlier fifteen-minute freshness targets and batch ordering.

Baseline: production `ba5ee7ad` restored VHH history from 4 May, DDH five-minute
metadata polling and the four-site cached views. Existing account-wide admission,
source isolation, bounded imports and publication safeguards remain in force.

## Common delivery method

Implement each complete feature set, run one coordinated correctness/cost test
pass, deploy the reviewed batch, and perform representative live acceptance.
Avoid serial cosmetic releases or repeated full-suite runs without a new failure.
Retain independent feature/source stop controls inside each batch. A failing batch
can be disabled/reverted and its components isolated without stopping unrelated
working sites. Rollback cannot recover already consumed D1 allowance: retain
fail-closed reservations and stop before daily limits are reached.

Do not automatically restore broad SQL readers, global snapshot warm-up, runtime
DDL or Creator startup fan-out. Provide the requested outcomes through bounded
paths. Preserve clinical inputs, personal edits, existing subscription URLs and
the rollback release. Keep raw rosters and credentials outside Git.

## Batch 1 — roster freshness, next-term delivery and historical grades

### Metadata before retrieval

Every FindMyShift and SharePoint roster check must compare lightweight provider
metadata with the last successfully activated version of that exact site/file/
term window before downloading workbook content or running extraction/parsing.
Use Modified/LastModified plus a stable file identity and provider version/ETag
where available. Successful unchanged checks perform no download, parse, import,
publication, workflow dispatch or D1 status write. Avoid per-tick analytics calls.

DDH already applies LastModified and exact successful version/window checks.
SharePoint server ingestion deduplicates received versions, but that is after
the Flow supplies the file. Inspect each live Flow and put the metadata decision
before Get file content/Run script; server deduplication alone is insufficient.
Keep SharePoint change triggers as the fast path and add metadata reconciliation
for missed triggers. A file save with identical roster facts may still require
one extraction to discover the semantic no-op, but must not rewrite those facts.

Necessary exceptions are explicit: an initial import; a newly eligible term/file
not imported before; a user-authorised parser repair; and resuming failed work.
Resume queued/failed work using retained input where possible rather than fetching
the provider again. Preserve the previous valid active roster until completion.
A failed candidate must not become the successful-version watermark. Re-check
version stability around extraction and reject stale/concurrent deliveries.

### Timing and cost

Target retrievable, valid provider changes reaching the app, affected At a glance
views and subscription-feed responses within **five minutes** under normal
conditions. Start with a two-minute metadata interval, change triggers, prompt
bounded processing and a visible-browser revision check at most once per minute.
A five-minute check alone leaves no processing time. Measure the complete path,
especially GitHub dispatch/startup latency; adapt the runner if it cannot meet
the target with existing inexpensive infrastructure. Do not declare success from
metadata checks alone or silently introduce paid infrastructure.

Coalesce rapid saves and duplicate tabs. Hidden tabs stop polling, then check on
foreground return. An unchanged browser revision does not fetch/rebuild calendars
or re-run account discovery. Existing contact extraction remains at its agreed
five-minute frequency. Review HTTP feed caching so it does not add another five
minutes of server-controlled staleness. Phone calendar clients control their
own fetch schedule; the deadline applies to server/feed and visible-app freshness.

DDH's last measured unchanged tick used nine D1 reads and zero writes: at two
minutes this is about 6,480 reads/day before other sites. Measure each site's
actual costs, ordinary traffic and changed-import batches; include Power Automate
licence/action allowances, provider throttling, R2, Worker and GitHub limits.
Keep FindMyShift calls serial per API key, obey 429/backoff and never defeat an
admission stop to meet the timing target. Expose deferred/invalid/outage status.

### Next-term eligibility

Use the first day of the calendar month preceding the term's start month, in
Australia/Melbourne: 2 November 2026 → 1 October 2026; February → 1 January.
Continue checking the current term while checking eligible next-term input.
Identify each input/window independently so an unchanged team version does not
prevent the first fetch of a newly eligible term. Unpublished terms remain waiting
without repeated downloads at the same provider version.

Proposed calendar availability follows the same eligibility date once an
authoritative roster exists, replacing the fourteen-day visibility hold. Keep
this assumption explicit in release acceptance; importing and visibility must
not accidentally use different dates. Verify source selection, cached manifests,
identity suggestions, permission boundaries, personal calendars and feeds together.

### Historical grades

Grade belongs to the roster/site/term, not the person's permanent identity.
Historical views use the supplied grade for that roster; current directory grade
comes from the current term. No promotion, merge, alias correction or current
account update may rewrite past grades. Retain provenance and permit a reviewed
term-specific correction. Never infer missing historical grade from today's
provider staff list; mark uncertainty and retain authoritative historical inputs.

Correct both identified paths: manual grade overrides carrying forward despite
later roster evidence, and multi-term shared history combining overrides without
term-specific keys. Inspect parser, membership, correction publication, grouping
and historical permission readers together. Include a Junior Registrar → Senior
Registrar → SMS fixture across sites and terms, with an older manual correction,
merged identities, missing grade evidence and boundaries. Keep authorisation
policy separate from displaying a historical grade.

Acceptance covers unchanged ticks, rapid saves, missed triggers, term eligibility,
same timestamp/different provider version, concurrent arrivals, crash/resume,
budget deferral, open/hidden tabs, subscriptions and historical grade accuracy.
Use real needed provider changes for live checks, not synthetic clinical edits.
Export recoverable Flow definitions and inventory obsolete flows with this batch;
delete only after verifying replacements and the existing cleanup authorisation.

## Batch 2 — durable identities and account linking

Reconcile the parked alias branch and identity/merge plans with production.
Choose one stable ID design before migration; keep grade history separate.
Deliver preferred names, aliases, bounded duplicate suggestions, merge preview,
conflict checks, audit/history, exact reversal and preserved account/profile/feed
relationships as one feature set. Add bounded automatic linking/repair only for
unambiguous, permitted matches; ambiguous identity changes require review.
Prove concurrency, rollback, historical grade preservation, real subscription
continuity and scale costs. Candidate discovery must not scan every person pair.

## Batch 6 — disputes, messaging and secure account sessions

Build on stable identities: dispute submission, Creator queue, claim transfer,
temporary alias/rejection and in-app notifications. Implement secure session/
refresh-cookie authentication and account lifecycle integration together, with
upgrade/rollback compatibility for existing accounts and external subscriptions.
Add password recovery or additional email notifications only where requirements
and provider costs are agreed. Test complete user/Creator journeys and concurrency.

## Batch 7 — Director and reporting workflows

Existing non-clinical accounts, Director entitlement and At a glance access stay
available. First settle hospital/program scope, reporting/export/sharing and any
approval/delegation/editing requirements. Implement the agreed workflow as a batch
using the stable identities and session/access model from preceding batches.
Requirements are an external dependency, not permission to invent capabilities.

## Batch 3 — Casey and richer contacts

Obtain authoritative Casey input and join the same bounded delivery/freshness
contract. Restore its roster-backed views without inventing missing data. Add
clearly labelled on-call/non-rostered entries, switchboard instructions and
remaining contact matching acceptance. Preserve existing contact frequency/costs.

## Batch 4 — recovery and repository administration

Provide exact site/file/term dry runs, budgeted reconciliation, repair/rebuild,
facility bootstrap, version conflict review, checkpoints and recoverable outcomes.
Keep advanced capabilities scoped and normally closed; do not enable global
automatic rebuilding. Validate interrupted execution and unrelated-term preservation.

## Batch 5 — resilience and diagnostics

Deliver coalesced bounded background retries/backoff, recoverable publication
failures, persistent sampled/batched UI console history and clear freshness/error
status. Preserve current retry protections during earlier batches; earlier batches
must still handle their own errors safely. Test no retry storms or per-render writes.

## Batch 8 — optional work, polish and consolidation

Prioritise remaining proposals, mobile improvements, retention policy, test debt,
architecture tidying and documentation consolidation. Reassess deferred physical-
presence, any-pair matching and overview editing against real requirements.
DDH source subscription input is removed from the backlog: administrative
FindMyShift API access retrieves the team-wide shifts. Apple/Google subscription
outputs remain live and are unrelated to source ingestion.

## Effort and next step

Recommend **Medium** for Batch 1 and Batch 2 implementation: they change source
watermarks, concurrent delivery, calendar revisions, historical grades and identity
relationships. Low is suitable for routine cleanup and measured acceptance once
the design and regression checks pass. Higher effort is not currently required.
Pause after this plan/document cleanup; the next resumed implementation starts
with Batch 1 rather than another planning cycle.


## Batch 1 implementation checkpoint — 4 October

- Incorporated production `d318bf7b` and preserved the other chat's clinician
  visibility, startup and navigation changes; their regression checks pass.
- Added Melbourne preceding-month eligibility in provider windows and shared
  readers, plus exact-term grade corrections and removal of undated historical
  membership fallback. Historical DDH repairs must use retained term-specific
  evidence rather than today's live staff list.
- Added edge/R2-only opaque change fingerprints, coalesced visible-tab checks,
  safe calendar/view refresh, shorter feed caching and two-minute DDH polling.
  Unchanged browser checks contain zero D1 operations. Imports retain account
  admission and source isolation; processing is serial within each site.
- Added authenticated SharePoint pre-download checks, including unchanged,
  pending, invalid-candidate and budget-stop behavior. Added a private exported
  Flow package preparation tool; the MCH backup/update package is ready outside
  Git. Existing contact schedules are unchanged.
- Production admission passed at approximately 94,000 account-wide reads and
  570 writes. The compact visibility migration cost 12 reads and 12 writes.
- **Outstanding:** Mac unlocked for MMC/MCH/VHH export/import and live Flow
  validation; verify the actual Power Automate allowance before enabling
  two-minute metadata reconciliation, verify the VHH library/path from its
  export, then measure real changed-roster delivery across all sites. Five-minute
  all-site delivery has not yet been declared complete.
- Microsoft documents 6,000 daily requests for seeded/trial licences and 40,000
  for Power Automate Premium. Reconciliation must fit the verified existing
  licence alongside contact flows; do not silently add paid capacity.
  Sources: https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/faqs
  and https://learn.microsoft.com/en-us/power-platform/admin/api-request-limits-allocations.
