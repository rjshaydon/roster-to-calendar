# Ordinary login and identity D1 remediation plan

## Purpose

This plan addresses the 8 September 2026 D1 burst that occurred while At a
glance, roster automation, contact automation and bootstrap controls were
closed. It protects ordinary login and calendar use before compact-fact rollout
or Doctor Names work resumes.

This document is a plan only. It authorises no Production/Preview D1 request,
migration, deployment, backfill or feature activation.

## Evidence and limits

The settled account-wide sample through approximately 12:35 AEST reported
1,336,466 rows read and 247 rows written. One five-minute interval accounted
for approximately 1,320,227 reads and 213 writes. Production Pages traffic was
small, so this was not a simple request-volume surge. Query fingerprints did
not reconcile with the aggregate: approximately 1.31 million reads remain
unattributed. Therefore no exact SQL statement may be claimed as proven.

The source audit nevertheless identifies unsafe ordinary-user mechanisms:

- an unclaimed or incompletely claimed account can trigger doctor discovery;
- discovery can fall back from compact/canonical records to roster-file doctor
  rows and historical event queries;
- a normal account save can rewrite profiles, claims, aliases and locations
  even when only transient UI state changed; and
- such saves can schedule snapshot warm-up work after the response, causing
  hidden D1 activity unrelated to a roster revision.

An observed `sqlite_master` fingerprint is not emitted by current code and is
evidence that retained older Production deployments must also be inventoried.
It is not proof that an old deployment caused the burst.

## Safety invariants

1. Login and ordinary calendar loading must never scan roster history.
2. Missing compact identity data must produce a friendly unavailable state,
   never a historical fallback.
3. An unchanged account/UI-state save writes no D1 rows and schedules no D1
   background work.
4. Identity/alias rows change only during explicit identity mutations or
   incremental roster ingestion—not during general account repair or login.
5. Snapshot work runs only for a changed roster/contact dependency and under
   its own default-off control and hard budget.
6. Disabled or malformed controls fail before D1.
7. At a glance entitlement records remain intact while its global maintenance
   gate is active.

## Stage 1 — contain ordinary request paths

Implement and test locally without remote bindings:

1. Add an independently controlled, fail-closed identity-discovery capability.
   Its Production default is off during remediation.
2. Keep authentication and existing claims available, but prevent account
   context preparation from calling historical discovery when compact identity
   records are absent. Return an explicit “identity linking is temporarily
   unavailable” state where selection is required.
3. Remove `queryRosterFileDoctors`, historical `queryDoctorEvents`, and
   historical seniority derivation from login/account-context and user-list
   request paths. These functions may not be runtime fallbacks.
4. Ensure stale clients cannot bypass the gate by calling older actions or
   supplying legacy request shapes.
5. Keep Creator user listing paused or limited to compact account summaries
   until its bounded replacement passes the same gates. This is defence in
   depth; it is not asserted as the 8 September cause.

## Stage 2 — make account persistence genuinely incremental

1. Separate account mutations by responsibility: profile, claims,
   entitlements, hospital locations and UI state.
2. Update only explicitly changed fields and facts. Do not delete/reinsert
   complete child collections during an unrelated save.
3. Compare stored and requested values transactionally; a semantic no-op must
   execute zero mutation statements.
4. Remove automatic durable-person/alias seeding from general login, repair and
   UI-state saves. Identity creation belongs to incremental ingestion or an
   explicit Creator identity action.
5. Preserve unrelated claims, roster contributions, current-term data and the
   14-day pre-term visibility rule.

## Stage 3 — isolate snapshot and background work

1. Do not schedule account snapshot warm-up after ordinary UI-state, profile or
   claim saves.
2. Key snapshot invalidation to explicit roster, access or contact revisions.
   An unchanged revision performs no D1 or R2 work.
3. Put every builder behind an independent default-off switch, exact scope,
   one-in-flight lock and per-run statement/row ceiling.
4. Prohibit historical identity discovery from snapshot preparation.
5. Record a bounded reason and dependency revision for every scheduled build so
   unexpected work can be attributed without querying D1.

## Stage 4 — inspect retained Production callers

Using control-plane and source information only:

1. Inventory current and retained Pages deployments, preview aliases, Workers,
   cron triggers, queues and external callers that can bind the Production D1
   database.
2. Map each callable deployment to commit, configuration and D1 binding.
3. Identify versions containing runtime schema inspection, historical identity
   discovery or unrestricted builders.
4. Prepare a reversible blocking/removal action for unsafe retained deployment
   URLs. Execute it only with explicit approval; preserve Git history and a
   known rollback deployment.

## Stage 5 — compact identity restoration design

The desired Doctor Names feature remains viable, but discovery moves to
ingestion time:

1. Maintain one immutable `person:<ULID>` plus user-facing person ID and
   editable display name in compact canonical records.
2. Index source identity and normalised blocking keys.
3. On a new or changed roster identity, compare only the small matching block
   and create bounded suggestions. Never perform all-pairs comparison or query
   roster events per candidate.
4. Store term staff/grade facts during ingestion. Do not rediscover them during
   login or At a glance use.
5. Let a Creator confirm, reject, edit or reverse suggestions through the
   planned non-technical Doctor Names workflow.

## Required local evidence

Use a deterministic synthetic database with more than 100,000 historical
events plus representative Excel/FindMyShift fixtures.

- Claimed, unclaimed and partially claimed logins perform no roster-event or
  roster-file-doctor scan.
- Missing canonical identity data returns the unavailable state with a bounded
  indexed account lookup.
- Repeated identical saves produce zero D1 writes and no `waitUntil` D1 work.
- UI-state changes touch only their intended row.
- Identity candidate generation examines only the indexed candidate block and
  never `roster_events`.
- Hidden tabs, stale clients and 50 simultaneously visible pages cannot invoke
  discovery/build fallbacks or stack requests.
- Query plans, statement counters and estimated rows examined are recorded;
  rows returned alone are not accepted as cost evidence.

Correctness fixtures separately cover name variants, swaps, sickness,
removals, overlaps and term boundaries. Performance and parsing correctness
must not be conflated.

## Rollout sequence

1. Implement and review containment locally; do not add compact migrations to
   the same release.
2. Deploy containment and default-off controls without contacting D1, then
   read back effective configuration through the control plane.
3. Wait for a fresh UTC quota day and observe at least two passive hours with
   no optional feature use.
4. Require reconciled account-wide daily, five-minute and fingerprint
   Analytics and a low, attributable ordinary-use burn rate.
5. Only then consider migration `0026` onward and the one-file compact bootstrap
   under the separate quota-safe rollout plan.

Any unexplained spike, historical query fingerprint, unexpected write or
unreconciled Analytics sample is `STOP`. Do not investigate it by issuing more
D1 queries.

## Post-containment observation — 8 September 2026

The contained Production deployment was active before the deliberately
isolated 15:15–20:30 AEST observation window. Settled account Analytics for
that interval reported 27,628 rows read, 20 rows written, 575 read queries and
8 write queries. The largest five-minute bucket was 1,658 rows read. This is
approximately 5,164 rows read per hour, or 124,000 per 24 hours if that exact
ordinary-use rate continued.

The two historical spikes visible earlier in the same UTC quota day occurred
before containment: 1,320,227 rows at 11:25–11:30 AEST and 950,145 rows at
13:55–14:00 AEST. Because those incidents make the full-day sample fail the
normal start and stop thresholds, this observation is evidence that immediate
containment is working; it is not a `GO` for optional features, migrations or
bootstrap. The fresh-day passive gate remains required.

## Restoration relationship

This remediation is a prerequisite for, not a replacement for, the features in
the feature restoration register. At a glance, automatic imports, contact
refresh, Admin Files and Doctor Names are restored only through their compact,
incremental implementations and their individual canaries.
