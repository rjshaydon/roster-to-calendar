# Core calendar synchronisation restoration plan

Prepared 13 September 2026 to restore current roster changes to users' personal
calendars before resuming At a glance.

## Authority and priority

This is the authoritative implementation and Production rollout sequence for
manual roster ingestion (FR-06) and automatic roster synchronisation (FR-07).
It changes the restoration order, but it does not weaken the account-wide D1
safety rules in
[`d1-account-quota-safe-rollout-plan.md`](./d1-account-quota-safe-rollout-plan.md).
The facility publication plan continues to govern At a glance only.

For avoidance of doubt, the document hierarchy from this date is:

- the account-wide quota plan governs admission, observation and stop limits;
- this plan governs FR-06/FR-07 implementation and roster-source rollout;
- the facility chunking plan governs later At a glance publication;
- the feature restoration register records current state and evidence; and
- earlier phase, handoff and remediation plans are retained as historical
  evidence, not competing restoration orders.

The priority is:

1. preserve login and existing personal calendars;
2. restore roster ingestion and personal-calendar synchronisation, one isolated
   source at a time;
3. restore the remaining roster sources by implementation class;
4. observe stable ordinary service; and
5. resume the separate At a glance publication and reader rollout.

At a glance publication is not a prerequisite for roster syncing. Throughout
calendar-sync restoration, At a glance maintenance, facility builders/readers,
contact ingestion, Preview, bootstrap, watchdog, Doctor Names and unrelated
maintenance controls remain closed. Existing At a glance entitlements are
preserved.

This document authorises planning, local implementation, focused local tests
and commits with all Production controls closed. It does not itself authorise
a Production D1 request, an enabled Power Automate flow, a remote migration or
an open Production write control. Each Production canary remains a separately
observed operational step.

## Current baseline

- `main` at `cea2a5a` contains source-isolated, bounded roster ingestion and the
  resumable facility publisher, but all ingestion, facility publication and
  reader controls remain closed.
- Automatic roster ingestion is paused by
  `ROSTER_AUTOMATION_WRITES_ENABLED=false`, an empty
  `ROSTER_AUTOMATION_SOURCE_ALLOWLIST` and
  `ROSTER_AUTOMATION_QUEUE_ENABLED=false`.
- Creator/manual roster mutation is independently paused by
  `MANUAL_ROSTER_WRITES_ENABLED=false`.
- All eleven known Power Automate roster, contact and bootstrap flows are off.
- The roster sources recognised by the application are:
  - `monash-adults` (MMC, SharePoint workbook);
  - `monash-paeds` (MCH, SharePoint workbook);
  - `vhh-active-medical-roster` (VHH, Office Script JSON extract); and
  - `dandenong-findmyshift` (DDH, FindMyShift API).
- Admin → Files uses its bounded compact status reader and is not a prerequisite
  for ingestion.
- Persistent request attribution and per-request D1 statement ceilings remain
  mandatory.
- Approximately 45 users subscribe to personal calendar feeds. Their periodic
  calendar-client refreshes are expected to produce continuous small,
  read-only traffic even when nobody has the web app open. That traffic belongs
  in the ordinary-service baseline; it is not evidence of a background writer
  unless its route, query shape or rate departs from the settled feed baseline.

## Problems that must be corrected before enabling a source

The existing controls are not sufficiently isolated for a safe one-source
rollout:

1. The master roster-write switch also governs manual file mutations. Opening
   automatic ingestion must not silently reopen Creator import, removal or
   repair actions.
2. `/api/automation/pending` lists queued work globally rather than for one
   exact allowed source. The processor can therefore pick up old work from a
   different source.
3. Dispatch leases are global and the processor currently requests several
   queued runs. A one-source canary must process one exact source and one job at
   a time.
4. The `failed` derived-save phase is not consistently subject to the exact
   source allowlist.
5. An unchanged provider version currently updates source bookkeeping. The
   no-change path must not rewrite events, staff facts, daily presence, cache
   objects or D1 status merely to record another poll.
6. Two VHH Production-labelled Power Automate flows exist. Only one may be the
   authoritative trigger; both must never be enabled together.

## Permanent design constraints

### Capability isolation

Introduce independent, default-off authority for:

- automatic roster ingress;
- automatic queue polling/dispatch and derived processing;
- manual roster mutation and advanced repair; and
- each exact source ID.

Enabling automatic ingestion must leave manual file mutation disabled. Every
automation route must reject a missing or disallowed source before D1 or R2.
Queue polling, claiming, dispatch lifecycle updates and derived phases must be
source-scoped. A global or omitted source is invalid while operating in
single-source mode.

### Idempotency and incremental writes

The provider version plus source ID and retained filename is the first
idempotency fence; the content hash is the second. For an unchanged source:

- do not store the payload again;
- do not create or update a sync run;
- do not rewrite the source status merely to record a poll;
- do not dispatch a processor;
- do not rewrite roster events, membership, coverage, daily presence, summaries
  or facility objects; and
- return a small `unchanged` response.

A changed source is staged and remains invisible until all derived chunks pass.
Completion compares the staged result with the currently active source and
updates only changed facts and dependencies. Failure leaves the previous
active roster visible. A retry must resume or replace the failed staging state
without duplicating active facts or dispatches.

### Roster semantics

- A complete source replacement removes facts absent from that source only
  within its authoritative hospital/term or declared FindMyShift window.
- Unrelated source and term contributions remain untouched.
- Overlapping replacements must not expose both the old and new contribution.
- SMS membership persists until a Creator removes it, including sabbatical,
  long-service or other extended leave.
- Non-SMS membership exists only for the hospital and roster term in which the
  person appears.
- Current-term data and the agreed 14-day pre-term visibility rule remain in
  force.

## Implementation sequence

### Gate 1 — document and test the closed baseline

1. Record the current commit, effective closed controls, disabled Power
   Automate flows and a settled account-wide Analytics sample.
2. Confirm a paused request to every roster automation route stops before D1,
   R2 and outbound dispatch.
3. Confirm ordinary login, an existing personal calendar and its calendar feed
   still work without enabling optional functionality.

Gate: the application is stable, the D1 budget report is reconciled, and no
unexpected caller or fingerprint is present. This gate uses no test ingestion.

### Gate 2 — implement true source isolation locally

1. Separate manual roster-mutation authority from automated ingress authority.
2. Require an exact allowed source on ingress, FindMyShift checking, VHH
   extraction, pending lookup, dispatch and every derived phase, including
   failure reporting.
3. Make queue listing and claiming source-specific. During a canary, return at
   most one run.
4. Pass the exact source through the GitHub workflow and processor. Reject a
   missing, unknown or mismatched source rather than falling back to a global
   queue.
5. Preserve bounded leases and deduplication, but prevent a dispatch for one
   source from consuming another source's work.
6. Keep every new control false or empty in committed Production and Preview
   configuration.

Gate: local tests prove that opening one source cannot read, write, claim,
process, fail or complete work for any other source, and cannot enable manual
roster mutation.

### Gate 3 — prove incremental ingestion locally

Use the existing representative roster fixtures for correctness and the simple
synthetic database for cost. Test each implementation class separately:

- Monash Excel: adults and paediatrics;
- VHH Office Script extract; and
- DDH FindMyShift window.

For each class prove:

1. first ingestion stages and completes within explicit per-request statement,
   examined-row, returned-row, D1-write, R2 and dispatch ceilings;
2. an identical provider version performs zero D1 writes and zero R2 writes;
3. identical content under the applicable idempotency rules does not duplicate
   events or queue work;
4. one corrected shift changes only the affected facts and dependencies;
5. swaps, sickness, removals, overlaps and term boundaries retain their existing
   behaviour;
6. SMS and non-SMS membership follow the stated term rules;
7. a failure at each chunk boundary preserves the previous active roster;
8. retries and duplicated callbacks are idempotent; and
9. no ingestion invokes facility publication while At a glance builders are
   disabled.

Record query plans and actual local statement/read/write metadata. Rows examined
must be distinguished from rows returned. Any query whose cost grows with all
roster history blocks Production rollout.

Gate: all focused safety, ingestion, parsing, queue-failure and quota tests pass.
Do not expand into unrelated UI regression testing.

### Gate 4 — deploy the closed implementation

1. Commit and push the isolation/incremental changes with every runtime control
   closed.
2. Allow the normal Production deployment, then explicitly deploy/read back
   effective configuration if variables are involved.
3. Verify the active commit, one Production deployment, zero Preview
   deployments and all Power Automate flows still off.
4. Verify credential-free and wrong-source automation probes stop before D1.
5. Wait for Analytics settlement and confirm ordinary usage remains stable.

Gate: the safety implementation is live but incapable of ingesting a roster.
Rollback is the preceding closed deployment and does not require D1.

### Gate 5 — one-source Production canary

Use `monash-adults` first because its parser and retained files already have the
strongest local incremental evidence. Before the canary, determine whether
`Sync Monash roster files` can submit MMC independently. If it cannot, do not
enable the multi-source flow; use a one-shot source-specific submission or
isolate the flow first.

The 13 September read-only inspection established that `Sync Monash roster
files` is not independently switchable by source. It watches the entire
`/Shared Documents/Medical Roster` folder and derives `monash-adults` only when
the filename begins with `Adult`; every other triggered filename is labelled
`monash-paeds`. The recurring flow must therefore remain disabled for the MMC
canary. Gate 5 will use a one-shot submission of the exact current Adult roster,
with `sourceId=monash-adults`, or a separately isolated copy whose trigger can
match only that file. No catch-all `else=monash-paeds` logic is permitted in an
enabled recurring flow.

1. Take two reconciled account-wide Analytics measurements at least ten minutes
   apart and apply the quota plan's admission thresholds.
2. Open only automatic ingress, queue processing and the exact
   `monash-adults` allowlist. Manual mutations and every other optional control
   remain closed. Read back the effective configuration without D1.
3. Submit the current MMC source exactly once. Do not schedule repetition.
4. Permit one source-scoped queued run to complete, then close ingress, queue
   and the allowlist immediately through configuration.
5. Confirm from the response, GitHub run and Analytics Engine attribution that
   only the expected route/source chain ran.
6. Wait for settled account-wide Analytics. Reconcile reads, writes and
   fingerprints before any browser test or second submission.
7. Test one affected personal calendar and its calendar feed. Do not enter At a
   glance.
8. Re-open the same exact canary controls and submit the unchanged MMC provider
   revision once. Close immediately and prove zero D1/R2 writes and no processor
   dispatch.

Gate: both the changed/current submission and explicit unchanged replay are
fully attributable, within their declared ceilings, and ordinary calendars show
the correct roster. Any unexplained delta is a stop, not permission to retry.

### Gate 6 — restore routine roster sources

Expand only after the previous source has settled:

1. `monash-paeds`, using the already-proven Monash Excel implementation;
2. `vhh-active-medical-roster`, after selecting exactly one authoritative VHH
   Production flow and leaving the duplicate off; and
3. `dandenong-findmyshift`, with its provider-version check and authoritative
   term/window semantics tested separately.

For the first run of each implementation class, repeat Gate 5's open once,
close, reconcile, browser/calendar-feed check and unchanged replay. After a
source passes, its routine trigger may be enabled while its exact server
allowlist remains in place. Never enable Preview roster flows during this
rollout.

The source restoration record must name the Power Automate flow, source ID,
deployment, observed interval, D1/R2 totals, result and rollback action. “All
flows enabled” is not an acceptable rollout step.

Gate: all four roster sources update personal calendars correctly; repeated
unchanged checks have no D1/R2 writes or queued processor work; daily usage
projects safely within the permanent ordinary-service reserve.

### Gate 7 — restore Creator manual roster operations

Manual import, removal and repair remain a separate capability. Restore ordinary
manual import only after automatic sources are stable, using a distinct control
and one-file canary. Destructive removal, bulk replacement and advanced repair
remain closed until their individual incremental and recovery tests pass.

### Gate 8 — resume At a glance restoration

Only after roster syncing is stable, return to
[`facility-publication-chunking-remediation-plan.md`](./facility-publication-chunking-remediation-plan.md):

1. plan MMC publication;
2. build its seven-day batches serially;
3. build one month at a time;
4. finalise atomically;
5. enable one Creator-only MMC reader with contacts off; and
6. expand by hospital and cohort only after settled evidence.

Contact ingestion and 60-second On shift contact refresh remain separate later
gates. The permanent legacy At a glance read path is never restored.

## Production stop and rollback rules

Immediately close automated ingress, queue and the exact source allowlist by a
non-D1 configuration deployment if any of the following occurs:

- a quota error, timeout, non-JSON error or unexplained request;
- a request for a source other than the current allowlist entry;
- more than one queued run claimed during a canary;
- an unexpected Power Automate, Preview, watchdog or scheduled caller;
- a query absent from the reviewed fingerprints;
- a query plan or invocation exceeding its declared ceiling;
- an unchanged input writing D1/R2 or dispatching work;
- an unexplained write, duplicate event or active-file ambiguity;
- incomplete Analytics attribution or totals that do not reconcile; or
- evidence that personal login/calendar availability is deteriorating.

Do not raise a ceiling, enable another source or retry until the failed interval
has settled and is understood. Rollback never depends on a D1 query.

## Power Automate restoration order

Keep all flows off until Gate 5 explicitly calls for one:

1. `Sync Monash roster files` only if it can be constrained to the exact canary
   source; otherwise use or create a one-shot source-isolated trigger without
   enabling its recurrence.
2. Select exactly one of `VHH Active Medical Roster to Production` and
   `Sync VHH Active Medical Roster to Production` after their definitions are
   compared. Keep the other off and record why.
3. Restore the DDH roster trigger only after its exact owning flow is identified;
   none of the contact flows substitutes for roster ingestion.
4. Preview, contact-allocation/contact-list and bootstrap flows remain off until
   their later restoration gates.

## Prepared Gate 5 operating packet

The following sequence is prepared in advance and must be followed without
combining steps:

1. Keep `Sync Monash roster files` disabled. Confirm all other Power Automate
   flows remain disabled.
2. Obtain two fresh, settled, reconciled account-wide samples at least ten
   minutes apart. The expected background delta includes ordinary personal
   calendar-feed refreshes from roughly 45 subscribers and must remain
   read-only and within the ordinary-service projection.
3. Deploy only these temporary Production values:
   `ROSTER_AUTOMATION_WRITES_ENABLED=true`,
   `ROSTER_AUTOMATION_QUEUE_ENABLED=true`, and
   `ROSTER_AUTOMATION_SOURCE_ALLOWLIST=monash-adults`. Leave manual roster,
   facility, contact, identity, bootstrap and watchdog controls closed.
4. Read the effective values back without D1. A missing, additional or
   differently cased source is a stop.
5. Submit the exact current Adult roster once through a one-shot action. The
   request must contain `sourceId=monash-adults`, the real filename, stable
   SharePoint version/ETag and modification time, and the matching file bytes.
6. Record the HTTP response, dispatch identifier and GitHub run. Only one queue
   record may be claimed and processed.
7. Immediately redeploy the three controls closed: writes false, queue false,
   allowlist empty. Read them back without D1 before observing the result.
8. After Analytics settlement, reconcile account totals, request attribution,
   D1 query fingerprints, D1 writes and R2 operations. The personal-feed
   baseline is allowed; any new unexplained route, write or burst is not.
9. Only after reconciliation, test one affected personal calendar and its feed.
   Do not enter At a glance.
10. Repeat the same open-once-close sequence with the identical Adult provider
    version. It must return `unchanged`, perform zero D1/R2 writes and dispatch
    no processor.

The rollback is configuration-only and never queries D1: set roster writes and
queue false and empty the source allowlist. Do not enable the recurring Monash
flow until Adults and Paediatrics have separate exact trigger conditions or
separate flows, and each source has passed its own canary.

### Gate 5 canary material prepared on 14 September 2026

The exact read-only SharePoint source selected for the first canary is:

- library folder: `/Shared Documents/Medical Roster`;
- filename: `AdultTerm3.2026.xlsx`;
- source ID: `monash-adults`;
- SharePoint version: `322.0`;
- SharePoint displayed modification: 14 September 2026 at 15:37 AEST;
- downloaded byte size: `565098` (displayed as 565 KB); and
- downloaded SHA-256:
  `91176f5a1fb5e3d77976b675feae4a4015d827d71a8de75d2ebfe8ec96dc9044`;
- modified by: Melissa Basic.

Safari confirmed that this exact workbook downloaded successfully to the local
iCloud Downloads folder as `AdultTerm3.2026-2.xlsx`; the `-2` is only Safari's
collision suffix and the submitted `fileName` must remain
`AdultTerm3.2026.xlsx`. The XLSX ZIP integrity check passed. An earlier 12:44
download is stale and must not be submitted. Do not use the displayed
modification time as the provider version. Immediately before the canary,
confirm SharePoint still shows version `322.0` and 15:37 AEST. If the current
version or modification time has changed, discard this prepared input and
obtain and fingerprint a fresh read-only copy.

`Sync Monash roster files` was turned off again before this preparation. It and
all other Power Automate flows must remain off. The current combined flow is
not the canary mechanism because it can also accept Paediatrics. Prepare a
manual, one-shot HTTP action for the exact file only, but do not save it as a
recurring trigger and do not submit it until steps 1–4 of the operating packet
have passed. Its fixed request fields are `sourceId=monash-adults`,
`fileName=AdultTerm3.2026.xlsx` and the XLSX content type; its ETag,
modification time and bytes must all come from the same current SharePoint
version.

Focused local preflight completed with no Production access on 14 September:

- `npm run test:source-isolation` passed;
- `npm run test:roster-idempotency` passed, including zero D1/R2 writes for an
  unchanged Monash provider version and for identical content under a new
  provider version;
- `npm run test:d1-quota` passed; and
- `npm run test:queue-failure` passed.

These tests prepare the canary but do not admit it. On return, first unlock the
Mac and locate the completed download. Then take the two fresh settled account
samples. Only if those samples pass may the temporary three-control opening and
single submission occur. Close the controls immediately after the response;
do not browse Admin, At a glance or another roster source during the interval.

#### Admission stop discovered after preparation

The 17:14 AEST account check returned `STOP`. Settled account Analytics recorded
1,000,186 rows read and six rows written in the 15:55 AEST bucket while all
Power Automate flows and roster controls were closed. Analytics Engine attributed
the burst at 15:57 AEST to an ordinary interactive session, principally the
automatically scheduled `queryRosterOverlapDoctors` insight warm-up and the
concurrent `loadCalendarEvents`/`loadAccountContext` path. The overlap request
reported 998,070 rows read. This is not roster ingestion and proves Gate 5 is
not currently admissible.

Do not open the canary controls or submit the prepared workbook until the
automatic insight warm-up is fail-closed or moved to the indexed compact daily
presence path, and the calendar/account-context burst has been separately
explained and bounded. A statement-count ceiling is insufficient because a
single permitted legacy event join can examine almost one million rows. After
remediation, deploy closed, collect a new settled passive interval and repeat
the two-sample admission gate from the beginning.

## Completion criteria

Calendar synchronisation is restored when every intended Production roster
source is source-isolated, routinely ingests changed provider revisions,
performs no fact/object writes for unchanged revisions, updates affected
personal calendars and feeds correctly, stays inside measured account-wide D1
budgets, and can be independently paused without disabling ordinary service.

At that point this plan is complete and At a glance becomes the next priority;
it is not silently re-enabled as a side effect of ingestion.

## Implementation checkpoint — 13 September 2026

- Gate 1 passed for local implementation. Two settled account-wide samples
  reconciled to `GO`: 75,025 reads, zero writes, projected 75,894 reads and no
  expensive fingerprint. The inventory was reduced to one Production
  deployment (`1890d569-33f3-414c-8587-9101c27dc921`, source `93feb08`) and zero
  Preview deployments. Because the sample immediately precedes the UTC reset,
  it does not authorise a later Production canary; Gate 5 requires fresh
  post-reset samples.
- Gate 2 is implemented locally. Manual roster authority is separate from
  automatic ingestion; pending work, dispatch claims, lifecycle callbacks,
  parser configuration, raw-file retrieval and derived processing require the
  exact allowed source; the workflow and processor handle one exact source and
  at most one queued run.
- Focused source-isolation, quota, request-attribution, rollout, queue-failure,
  local-isolation, facility-materialisation and full fixture tests pass.
- Gate 3 is complete locally. The three ingestion classes now return unchanged
  provider/content revisions with zero D1 and zero R2 writes. Changed Monash
  and VHH ingress creates at most four compact D1 rows and one R2 object before
  processing; repeating the queued revision writes neither store again.
- Automatic derived payloads are capped at 512 doctors, 25,000 events and
  5,000 issues before their first D1 statement. Completed/start/failure
  callbacks are idempotent, and a correction batch rollback preserves the
  prior active event.
- Disabled account-snapshot and identity-discovery controls now stop their
  post-ingestion fan-out before account, snapshot or canonical-identity reads.
  Shift-code issue delivery probes only the affected source/doctor claims,
  capped at 40, rather than listing every account.
- Migration `0032_roster_sync_source_indexes.sql` adds compact indexes for
  exact source/provider, source/queue and affected-claim probes. It performs no
  roster-event backfill. It is applied to Production: the remote operation ran
  three statements and reported 2,400 rows read and 1,175 index rows written;
  an exact schema query confirmed all three indexes.
- The synthetic cost fixture contains 10,000 sync runs, 10,000 dispatches,
  10,000 accounts, 20,000 claims and 109,200 roster events. Query plans use the
  exact source/version, source/hash, source/status and source/doctor indexes;
  no tested ingestion-control query scans global sync, account or roster-event
  history.
- Representative fixtures preserve Excel/FindMyShift/VHH parsing, swaps,
  sickness, removals, overlapping contributions, term boundaries, SMS
  continuity, non-SMS term membership and the 14-day visibility rule.
- Focused Gate 3 suites passed: `test:roster-idempotency`,
  `test:source-isolation`, `test:queue-failure`, `test:vhh-automation`,
  `test:facility-materialization`, `test:database-costs`, `test:fixtures`,
  `test:d1-quota` and `test:request-attribution`.
- Gate 4 deployment work is live at `cea2a5a`. Production has one deployment
  and zero Preview deployments; effective automatic/manual roster controls are
  closed; credential-free and wrong-source probes returned the pre-D1 paused
  response; all focused local suites pass; and all Power Automate flows remain
  disabled.
- Approximately 45 subscribed personal calendar feeds account for expected
  low-level read traffic while the app is otherwise idle. They must be measured
  as the ordinary-service baseline during every canary.
- The post-migration daily Analytics aggregate is intentionally not used to
  admit Gate 5: Cloudflare does not fingerprint the DDL and the same interval
  contains 882 additional unattributed reads. Although the scale is small and
  consistent with ordinary feed traffic, the counter will be allowed to reset
  before the two fresh Gate 5 admission samples.
- Read-only inspection confirmed `Sync Monash roster files` is a combined
  Adults/Paediatrics folder trigger with catch-all Paediatrics routing. It must
  remain off for the MMC-only canary; use the prepared one-shot exact-source
  path instead.
- The GitHub processor workflow is available, accepts an exact required source
  input, has no schedule, and had no run after 3 September 2026 when checked on
  13 September. Its five most recent runs were completed successful dispatches;
  no processor run was active during this preparation.
