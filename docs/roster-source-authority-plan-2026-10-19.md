# Automatic/manual roster authority — 19 October 2026

Status: reviewed implementation plan only. No application code, migration, Flow,
production configuration or roster data changes are part of this planning step.

## Required behaviour

The user chooses the current working roster by successful ingestion, rather than
by an uploaded file's modification timestamp or a permanent manual override:

1. A successful automatic sync supplies the authoritative roster.
2. A successful manual upload replaces it for the exact site and term covered.
3. The next successful automatic reconciliation for that same scope takes over
   from the manual upload. It must verify the provider's currently selected file
   and revision after the manual activation; an old queued job finishing later
   is not sufficient evidence.
4. Failed, missing, ambiguous, incomplete or budget-deferred automatic work leaves
   the manual roster active. It must not erase working data or falsely announce
   that automatic authority has been restored.

“Latest” means this ordered handoff, not whichever background task happens to
finish last. Provider timestamps/ETags retain their existing within-source
ordering protections; local file modification times cannot establish precedence
between an emailed attachment and SharePoint.

An important consequence is intentional: a healthy automatic check may restore
the same provider version previously imported, even if the manual upload has
different shifts. A manual upload is a temporary fallback, not a correction
that pins or permanently overrides the provider. If only the current term is
available automatically, that does not supersede a manually uploaded next term.

For a manual-only next-term draft with no provider file, retain the draft until
that term's file becomes available and passes processing. SharePoint draft/final
filename selection continues to follow the existing approved-folder rules.

## Current implementation and gaps

- Bounded manual importing is enabled and uses separate `manual-mmc`,
  `manual-mch`, `manual-ddh` and `manual-vhh` source IDs. It stages input and
  activates through the existing complete-term replacement transaction.
- Staged replacement can retire overlapping active files across manual and
  automatic routes, while protecting unrelated terms and wider date ranges.
- `rosterDeliveryOrder` compares automated runs and provider files within an
  automatic source. It does not establish a cross-route authority generation.
- SharePoint's successful-version shortcut in `roster-check.js` returns
  `unchanged`; it does not check whether a manual upload now owns that term.
  FindMyShift has a similar unchanged-version shortcut.
- Patch planning requires an active baseline from its automatic source. It
  cannot silently patch an inactive automatic baseline while a manual roster
  remains the displayed source of truth.
- Source `activeFileId` is a latest-delivery pointer and can refer to another
  term. It must not be used alone to decide which term's roster to restore.

Relevant paths: `manual-roster-management.js`, `roster-import-staging.js`,
`roster-delivery-order.js`, `roster-import-protocol.js`, `roster-patch.js`,
`api/automation/derived.js`, `api/automation/roster-check.js`,
`api/automation/findmyshift-check.js`, and their publication/ingress paths.

## One implementation batch

### 1. Record exact authority and guarded handoffs

Add compact, indexed authority/provenance for each canonical facility/term scope:
active file/content revision, manual or automatic origin, authority generation,
activation reference/time and the applicable provider file/revision evidence.
Use additive schema changes after inspecting existing keys and query plans;
reuse existing records where they can provide the same guarantees.

Pin each manual import to its reviewed authority generation. Successful activation
updates authority, retires displaced files and queues publication atomically.
Partial uploads, previews and failed staging do not become authoritative.
A stale manual retry cannot silently replace a roster activated since review.
An intentional fresh upload can replace that roster under the normal rules.

Automatic work carries provider evidence and the authority generation it may
replace. At activation, enforce both the current generation and existing provider
ordering guards inside the transaction. If manual activation intervened, do not
relabel the old job as fresh: require a new provider confirmation against the
current generation, then reuse its retained input if the revision still matches.
Check unfinished imports, patch activation and completion/publication retries;
none may restore an earlier authority generation.

Scope authority by canonical term, not by first/last occupied shift or filename
alone. Retain existing completeness and wider-range protections. A partial or
multi-term input requires reviewed exact coverage; it cannot retire dates outside
its scope. Do not broaden accepted roster formats in this batch.

### 2. Integrate handoff with metadata-first reconciliation

After a trusted, complete provider metadata response selects the exact file:

| Current authority / provider state | Required action |
| --- | --- |
| Automatic; same successful revision | Existing read-only no-op; no download, parse, publication or D1 status write |
| Manual; file missing or provider check failed | Keep manual; existing useful failure/deferred reporting only |
| Manual; current provider revision already retained and prepared | Admit one guarded restoration from that exact retained snapshot, then publish only changed results |
| Manual; changed/new revision | Existing bounded retrieval/import, followed by guarded handoff |
| Either; work failed/deferred | Keep working authority and resumable progress |

The successful-version shortcut must distinguish “provider unchanged and already
authoritative” from “provider unchanged but manual fallback currently active”.
Allow at most one restoration per manual authority generation, with an exact
idempotency key; repeated checks/tabs/Flow deliveries must coalesce.

Prefer retained parsed facts/artifacts for unchanged provider content. Validate
their exact file/revision, completeness and parser compatibility. If only raw
input is usable, reprocess it outside Cloudflare through the existing runner.
Do not redownload an unchanged workbook merely to hand authority back. Missing
retained input must produce an explicit repair/deferred outcome; any exceptional
recovery fetch must be separately bounded, never a new recurring download loop.

Verify provider identity stability around retrieval and retain existing
five-minute coalescing, draft/final selection and stale-delivery rejection.
An unavailable or incomplete folder listing is not proof of current input.
Do not add a second Flow or change polling frequency for this feature.

### 3. Keep all displayed calendars on the selected authority

Make staged activation and any retained-snapshot reactivation select exactly one
source for the affected scope. Review automatic patch target selection and old
source/run completion callbacks so they cannot leave manual and automatic facts
both active, or move authority back to an older snapshot.

Publish affected calendar, At a glance, staff and subscription artifacts through
the existing bounded pipeline. Authority and publication generations must agree;
late publication must not overwrite newer output. Preserve current error and
freshness handling while publication finishes. Only changed semantic results
need rebuilding; an identical handoff may update compact provenance without
rewriting events or churning browser/feed revisions.

Keep historical grades, identities, account links, user edits and feed URLs
unchanged. Retain displaced source files under existing retention rules.
Show provenance in the existing Admin Files detail; no normal-status banners or
additional controls are needed. Report only actionable failed/deferred handoffs.

### 4. Run one coordinated correctness and cost pass

Extend the existing fixture suites, including actual reader/publication paths:

- Automatic A → manual M → changed automatic B: B wins for the exact term.
- Automatic A → manual M → successful unchanged-provider confirmation of A:
  A regains authority once, from retained input, without a provider download.
- Next-term manual draft with only current-term automatic checks: draft stays.
- Missing file, provider failure, incomplete metadata, parse failure and budget
  deferral: manual stays and no successful watermark advances.
- Automatic job observed before manual activation finishes later: no stale
  overwrite; fresh provider confirmation is required.
- Concurrent manual imports, new automatic revisions and delayed callbacks:
  guarded activation and publication preserve the latest accepted generation.
- Retired automatic baseline / semantic no-op / draft-to-final promotion:
  correct authority, no duplicate shifts and no unnecessary event rewrites.
- Lost replies, retries, worker interruption and recovery: one handoff, bounded
  receipts and no reactivation of an obsolete scope.
- Partial/multi-term inputs and overnight shifts: no loss of unrelated dates.
- Personal calendars, At a glance and subscription output converge correctly;
  historical grades, identity links and user edits are preserved.

Measure the new unchanged-authoritative path and one actual handoff. Keep
unchanged checks free of additional writes/downloads/parsing or dispatch, and
avoid roster-history scans. All handoff writes/rebuilds use existing account-wide
reservations, protected headroom and stop controls. No plan or paid-service change.
Run focused native CPU checks only where the changed code requires them; do not
repeat large synthetic cloud seeds or the completed Batch 6 acceptance campaign.

### 5. Release and acceptance

Deploy the shared code behind a scoped handoff control after additive migration
verification. Start with MMC and an actually needed manual draft if available;
do not manufacture clinical changes just to exercise takeover. Verify the manual
publication, then its next real successful automatic handoff. Apply the same
tested contract to MCH/VHH/DDH without changing their provider retrieval methods.

If automatic acceptance must await a provider file, report that gate honestly;
fixture success is not live replacement evidence. Rollback disables new handoffs
while preserving the currently working authority and existing automatic paths;
it must not blindly reactivate an older file or remove published manual data.

## Placement and effort

This is the first focused part of Batch 4, followed by its remaining recovery/
administration work, then Batch 5 and Batch 8. Batches 7 and 3 are deferred by the
user. Batches 1, 2 and 6 remain completed; this is an explicit authority refinement
and race-prevention requirement, not a reopening of their entire acceptance work.

Medium is appropriate for implementation against this plan. Review exact schema
and race guards before editing; pause for clarification only if actual input
coverage cannot safely identify the intended term. No implementation is authorized
by this planning step alone.

## Implementation checkpoint — 19 October evening

User authorised the combined handoff/checker remedial batch. Added migration 0055:
compact term generations, immutable run fences, provider observations and a
per-site handoff policy. Existing prepared manual selections are seeded from
compact catalogue rows; no shift rows are rewritten. Untouched automatic scopes
remain generation zero, preserving metadata no-op and ingress write costs.

Manual begins pin a generation. Activation compares it transactionally alongside
the existing active-file fence. Automatic metadata pins observations before
retrieval; runs carry that evidence. Late jobs cannot acquire a new generation
merely by retrying. Fresh metadata can create a distinct, idempotent recovery job
from retained bytes when an old job crossed a manual activation. Completion of a
retired file cannot move source pointers. Existing automatic provider ordering
continues alongside the cross-route generation guard.

Fresh unchanged-provider confirmation can restore a ready retained snapshot by
switching only control rows and queuing bounded publication. It verifies exact
raw/prepared metadata and full scope; unrelated terms and retired facts remain.
A changed version with identical bytes records its receipt before handoff so
budget deferral does not cause repeated downloads. Missing/incompatible prepared
snapshots return repair-required rather than guessing. Such a repair may need the
existing retained-input tools; this release does not silently refetch a workbook.

The CLI now calls the same accountMaintenanceHeadroom calculation and bounded
maintenance/preview summaries as production, subtracting live publication holds
for bulk headroom. It is explicitly read-only/advisory and cannot recover receipts,
modify grants or authorize writes. Unknown reservations and circuit stops remain.
The historical canary checks retain additional attribution gates; their write
ceiling now references the shared 100,000 policy rather than a literal 80,000.
Default reports are compact. Missing evidence fails closed.

Passed focused authority and diagnostic tests plus existing restoration mutations,
Batch 1, ingress idempotency, wire protocol, patch, retained-stage reuse, source
isolation, publication delivery, facility materialization, subscription artifacts,
personal artifacts, cold-start and maintenance-account-budget regressions.
Actual HTTP unchanged-provider recovery is covered with synthetic SQLite data.
The live read-only checker returned GO above 81,000 actual writes, with roughly
14,200 usable writes after forecasts, reservations and uncertainty. Local tests
are not native Cloudflare CPU/billing measurements.

No genuine manual fallback is awaiting restoration, so a live clinical takeover
must be observed at the next naturally needed manual upload/provider check.
Do not manufacture a roster replacement for acceptance. Keep that external
acceptance separate from deployed implementation and regression evidence.

Rollback: set roster_authority_policy.enabled=0 for an affected site to stop
handoffs from a manual selection while preserving the current active roster.
Keep generation guards and migration; do not reactivate an old deployment that
lacks those guards while manual authority rows exist. No staff-search code is
included. Existing source IDs in Admin Files remain the provenance record.
