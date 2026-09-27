# Full service restoration plan

Status: first scoped contact diagnosis and local correction authorised and
completed on 27 September. Deployment and the remaining restoration batches
have not been performed in this round.

## Authority and objective

This is the authoritative ordering for restoring the Production service after
the September 2026 D1 incident. It supersedes the serial rollout ordering in
the older quota, calendar-sync, At a glance and contact plans. Their permanent
safety constraints, parser requirements and recorded evidence remain valid.

The objective is to restore the user-visible service that existed on
31 August 2026 as quickly as possible without restoring the mechanisms that
scanned roster history, performed runtime schema work, retried blindly or
wrote unchanged state. Work proceeds in functional batches. A safe batch stays
enabled. A batch is split only if measured evidence identifies it as unsafe.

## Reconciled Production position — 27 September 2026

The repository and restoration register do not currently describe the same
state. The tracked Production configuration at `3161b23` has:

- automatic roster ingress and queue processing enabled for Monash Adults,
  Monash Paediatrics, Dandenong FindMyShift and the VHH active roster;
- cached At a glance readers enabled for MMC, DDH, MCH and VHH;
- cached contact readers enabled for DDH, MMC and MCH;
- contact-workbook ingress enabled for the exact DDH and MMC sources;
- automatic At a glance launch, ordinary manual roster mutations, colleague
  insights, identity discovery, broad Creator hydration, global snapshot
  warm-up and advanced maintenance disabled; and
- legacy historical At a glance reads permanently disabled.

The 27 September user test found no DDH contact details. Therefore a successful
Power Automate run is not accepted as proof that DDH contact publication works.
DDH and MMC/MCH contact automation are **not restored** until the correct
current-date overlay is present and visible in the application. The current
state of every external flow must be read back before changing it; Git cannot
prove whether a saved flow is On or Off.

Cached At a glance objects are also not yet kept current automatically after a
roster change because the metadata/day builders remain disabled. Existing
views can work while silently becoming stale. This is a restoration blocker.

## Features still requiring restoration

### Required for the core service

1. End-to-end automatic roster freshness for all four sources, including a
   changed source reaching personal calendars and subscribed calendar feeds.
2. Incremental publication of affected At a glance metadata/day objects after
   a roster change, without a whole-hospital or whole-term rebuild.
3. Automatic DDH and MMC/MCH contact publication and visible current-period
   contacts. VHH contact support remains a separate product task because the
   Production parser/source was never completed.
4. Automatic At a glance launch for entitled users and its existing
   visibility-aware 60-second contact refresh.
5. Explicit colleague tools: “Who else is working with me?” and “When am I
   working with…?”, served from compact daily facts rather than event joins.
6. Ordinary Creator manual imports, replacement and removal with incremental,
   source-scoped mutations.
7. Fast Creator switching and directory enrichment without roster-history
   discovery.
8. Bounded automatic doctor discovery/claim suggestions.
9. Cross-device calendar/session-setting persistence through a dedicated
   single-record API, not the former broad save path.
10. Creator contact corrections if the end-to-end contact review shows that
    they are not already operating on the compact overlays.

### Operational capabilities, restored only if needed

- A source-specific schedule for DDH checking if no existing external trigger
  provides routine automation.
- Advanced repair/bootstrap as short-lived operator actions with exact scope.
- A watchdog only if a demonstrated operational gap remains; never the old
  global poller.

### Work that follows service restoration

- Durable Doctor Names and identity aliases.
- VHH contact-list discovery and implementation.
- Other planned product additions that were not working Production features on
  31 August.

## Mechanisms that must not return

- legacy At a glance queries over roster-event history;
- runtime schema inspection or `CREATE ... IF NOT EXISTS` on user requests;
- automatic colleague-insight warm-up;
- global snapshot warm-up after ordinary saves;
- one D1 write per displayed status message;
- repeated blind login/calendar retries;
- duplicate same-open metadata requests; or
- automatic Excel actions that edit their own SharePoint trigger workbook.

Restoring the user outcome does not mean restoring its unsafe implementation.

## Restoration method

### Independent releases

Group independently ready capabilities into one reviewed release and one
Production deployment. A capability requiring substantial implementation must
not delay another that is ready. Contacts, roster readiness and At a glance
publication below are independently releasable; their numbering is priority,
not a requirement to wait for every preceding item. During
recovery, use one deployment mechanism; do not create both a GitHub deployment
and a manual deployment for the same commit. Keep the active deployment and at
most one verified predecessor until the batch passes, then delete older
deployments through the control plane. Git history is the long-term rollback
record.

All external callers must use the stable Production domain, never a hashed
deployment URL. Inventory Power Automate, GitHub Actions, schedules and deploy
hooks once at the start; subsequently inspect only changes and unresolved
items. Deployment cleanup is release hygiene, not a fresh prolonged gate for
each feature. The incident record does not prove that deployment count caused
the read bursts; it separately records an approximately 998,000-row colleague
query. Preserve that distinction.

### Evidence and stop rules

Do not wait for long passive intervals after every small action. Use live
request attribution and the existing request-local statement ceiling while a
batch is exercised, followed by one settled account check for the batch.

Before enabling a capability, prove query cost locally with representative
scale, effective indexes, bounded input and query plans. A returned-row LIMIT
does not bound examined rows. The current middleware stops excess statements
but observes row counts only after execution; it cannot prevent one expensive
statement. Telemetry thresholds below are detection/containment, not a hard
pre-execution row budget. Do not enable a known unbounded path to discover its
cost in Production.

Retain the account-wide quota checker's admission rules and reserves: four
million reads and 80,000 writes reserved for ordinary availability; optional
work limited to one million reads/20,000 writes including 100% contingency;
start below 500,000 reads/10,000 writes and stop optional work at 750,000
reads/15,000 writes. Use the existing reconciled two-sample baseline and freshness
rules; reuse valid evidence rather than restart a passive observation period
for each feature. Missing/stale telemetry blocks new Production experiments,
but independent local implementation continues. Changing budget thresholds
requires an explicit documented revision, never silent relaxation.

Rollback the batch immediately, without D1, if any of these occurs:

- an unexpected query fingerprint or caller appears;
- a request reaches the statement ceiling;
- a request examines more than 10,000 D1 rows unless that exact bounded
  maintenance request was declared in advance;
- a five-minute bucket exceeds the existing checker's allowed baseline/action
  envelope, or 100,000 total reads (not merely incremental reads);
- unchanged input writes D1 or R2, dispatches processing, or republishes a
  shared object; or
- disabling the batch does not stop it before D1.

If a batch fails, close its flags/flows, retain the last valid data, divide
only that batch by capability or source and rerun the smaller halves. Continue
with independent safe batches rather than blocking the whole restoration.

### User checkpoints

Stop for the user only when credentials/access are required, a clinical
interpretation is needed, a visible result cannot be verified otherwise, or a
rollback/production decision is required. Do not stop merely to announce that
an internal test, deployment or analytics wait succeeded.

### Execution discipline and checkpoint

Maintain one short current checkpoint in this document: observed deployment
and flags; exact Flow IDs/states and observation times; completed acceptance
evidence; unresolved failures; next executable action. Link detailed evidence
instead of rereading the entire incident history. Distinguish repository
configuration, deployed configuration, and observed behaviour.

For each failure, inspect the saved Flow definition and actual action outputs,
HTTP result, normalized date/count/hash, published object and reader result.
Identify the first divergence, reproduce it locally where possible, then fix
that stage. Prefer structured definitions/run outputs over repeated designer
navigation when supported. Do not record tokens or contact details in logs.
After a failed fix, use new diagnostic evidence before trying another change.

Run focused meaningful checks for changed paths once; rerun only after relevant
changes or new failures. Reuse passing evidence for unchanged code. Perform
independent local work during telemetry settlement. Avoid repeated unchanged
polling. Combine necessary user visual checks into one concise list.

Use Medium effort for unresolved diagnosis and coupled ingestion/cache changes;
Low is suitable for a bounded correction with a reproduced cause and clear
acceptance checks. Request High only for a specific unresolved architectural or
data-integrity decision. Keep reports concise at every effort level.

## Batch 0 — truthful inventory and deployment hygiene

This batch is control-plane/read-only except for deleting superseded
deployments after their targets are verified.

1. Read back the active Production commit and effective flags.
2. Export or inspect the exact saved definitions and On/Off state of every
   roster/contact flow. Record trigger, filename filter, endpoint and source ID.
3. Inventory Pages Production and Preview deployments, Workers, schedules,
   GitHub workflows and deploy hooks.
4. Confirm every external application endpoint is the stable Production
   domain and no flow references a hashed/deleted deployment.
5. Select one deployment pipeline for the remainder of recovery.
6. Keep the active deployment and one verified predecessor; delete older
   callable deployments. This is hygiene and caller containment, not a claim
   that deployment count itself consumes D1.
7. Correct the feature register to the read-back state.

Gate: relevant callers, active configuration and rollback switches are known.
No application data request is needed. Reuse verified inventory and inspect
deltas. Do not let unrelated housekeeping delay a proven correction.

## Batch 1A — restore current contacts

1. Trace DDH's latest relevant run from trigger to visible current-period
   contacts using the diagnostic sequence above. Absence of contacts invalidates
   the end-to-end completion claim; it does not by itself identify the failed
   stage or prove that no extract was stored.
2. Verify the known MMC/MCH configuration mismatch against deployed settings:
   ingress admits `mmc-shift-allocations`, but committed
   `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST=ddh` makes
   `contactPublicationEnabled()` return false for MMC. Correct contact
   publication scope without opening metadata/day builders.
3. Check transport decoding, operational date/expiry, populated contact count,
   semantic replay and missing-overlay recovery locally. An unchanged complete
   publication writes zero D1/R2; recovery of a missing derived overlay is a
   separately identified bounded repair, not an ordinary unchanged refresh.
4. Exercise DDH and MMC/MCH as one ready release with source-specific rollback.
   Use current source retrieval where possible; do not require repeated user
   edit-and-undo cycles when existing run evidence or a controlled replay is
   sufficient.

Gate: correct current-period contacts visibly reach DDH, MMC and MCH; unchanged
complete publication is a no-op; retrieval causes no recursive source edits;
visible refresh is zero-D1. A green Flow run alone cannot pass this gate.

## Batch 1B — roster automation and upcoming-term readiness

Proceed independently of contact failures.

1. Verify each source's actual recurring trigger, exact filename/folder routing,
   endpoint and most recent accepted provider revision. DDH scheduling belongs
   here: if missing, implement one source-specific bounded coalesced schedule.
   An enabled endpoint or successful manual refresh is not automatic syncing.
2. Test each ingestion class locally with a representative full new term,
   correction, unchanged replay, filename/version replacement and interrupted
   retry. Include supersession, unrelated contributions, current-term retention,
   14-day pre-term visibility, and personal/feed cache revision propagation.
3. Trace the configured 1,250 incremental-fact ceiling through staging and
   promotion. Prove a realistically sized new term completes within bounded
   resumable work; do not simply increase/remove the guard. If unsupported,
   implement bounded chunking with atomic activation and retry deduplication.
4. Verify Production source freshness and existing changed-run evidence; use
   one budgeted current-revision replay only where evidence is missing. Never
   send synthetic shifts to Production. Do not wait for a roster writer's next
   edit to progress: local changed/new-term evidence and current Production
   delivery evidence establish readiness. Record a later genuine edit as
   additional confirmation, without claiming it has already happened.

Gate: all four sources have working recurring triggers and current delivery
evidence; representative new-term ingestion, feeds and browser refresh pass
locally. Clearly label any live new-term observation still pending.

## Batch 1C — keep At a glance current

Release separately if this needs more work than 1A/1B.

1. Trace successful ingestion through compact facts to metadata/day publication;
   prove which disabled controls prevent freshness before changing them.
2. Publish only affected facility/date/term dependencies with bounded work and
   coalesced revisions. Test source corrections, deletion, a new term and
   duplicate delivery. Preserve the previous valid revision on failure and
   expose a clear stale/updating status with a bounded retry path.
3. Retain protected facility access, SMS continuity, non-SMS term membership,
   locum access and the 14-day pre-term rule. Never revive broad event scans.

Gate: changed roster facts update the affected shared views, unchanged imports
perform no publication work, and failures retain clearly dated valid data.

Expected normal costs remain those already demonstrated: unchanged ingestion
does not write; an authenticated cached At a glance request uses a few indexed
rows; visible contact refresh is R2/session-only with zero D1; contact
publication changes only a small contact source/overlay.

First operational milestone: current contacts visible, roster triggers verified
and new-term ingestion proven locally. Each completed release stays enabled
while remaining work continues.

## Batch 2 — user-facing feature restoration

Group ready items behind independent flags. If colleague insights require a
rewrite, release the other ready items first; never flip the old insight path
on for a trial. Items already live require focused regression evidence only:

1. automatic At a glance launch for entitled users;
2. visible-tab-only 60-second contact refresh and immediate refresh on reopen;
3. explicit colleague insight actions rebuilt on compact daily presence;
4. Creator contact corrections on the small contact overlay; and
5. the current multi-site access policy: SMS/CMO see all enabled sites; other
   grades see current-term membership sites plus genuine locum sites only.

The four cached At a glance tabs and all four facility selectors are already
enabled and remain part of this batch's regression check. No feature may fall
back to legacy event-history SQL when an artifact is absent.

Gate: one Creator and representative all-site and site-scoped accounts complete
the user journeys; hidden tabs issue no polls; repeated unchanged use stays
inside the declared request costs.

## Batch 3 — Creator and account workflows

Group independently ready items into a release with separate rollback flags.
Do not let identity/discovery implementation block a finished session-setting
or manual-import correction:

1. ordinary manual roster upload/replacement/removal, excluding bulk repair;
2. compact Creator directory enrichment;
3. bounded doctor discovery/claim suggestions generated when identities
   change, not on every login;
4. a dedicated one-record API for cross-device session settings; and
5. explicit bounded account repair, never an ordinary-login side effect.

Gate: unchanged saves/imports write zero; manual roster changes affect only the
declared source scope; ordinary and Creator login remain cache-first; switching
does not query the directory again.

## Batch 4 — operational completion

1. Confirm the source schedules established in Batch 1B remain healthy.
2. Keep bootstrap and advanced repair normally Off; document the exact
   time-limited runbook rather than enabling them permanently.
3. Restore a watchdog only if measured operations show it is necessary after
   the source triggers are working.
4. Run the final end-to-end matrix and reconcile one settled account-wide
   sample.
5. Update every register entry to Restored, Replaced permanently, or Deliberate
   future work. Remove obsolete ordering statements from older plans.

Gate: the Production experience matches 31 August user functionality, except
where an unsafe mechanism has been replaced by an equivalent bounded one.

## Completion and next work

Service restoration is complete only when automatic roster and supported
contact updates are proven end-to-end, personal calendars and subscriptions
update, At a glance remains current, all intended user and Creator actions are
available, and every unchanged path is a no-op at the data layer.

After that point, rebase and review the durable Doctor Names branch against the
stable Production baseline. It must not be mixed into the restoration batches.

## Current checkpoint

### 28 September morning release admission

- Account analytics at 08:14 AEST, settled through 07:59 AEST: 162,444 reads,
  225 writes for the UTC quota day; projected reads approximately 181,000;
  no expensive fingerprints. Checker GO against the previous valid sample,
  using the documented `free-plan-dashboard-omits-d1` billing mode.
- Production readback remains `baec779b-db7a-43a3-b5bc-53a1ca8cd17e` with
  publication allowlist `ddh`. The workbook-ingress regression and diff check
  passed again. User authorised commit/push; release changes only the contact
  publication allowlist plus tests/documentation. No roster/day builders opened.
- After deployment verification: restore/check the MMC contact flow and obtain
  one current workbook submission. Then one combined DDH/MMC/MCH On shift
  check, against contact names actually present in today's source sheets;
  verify AM/PM as applicable. A missing date's source is not a reader failure.
- Do not advance to the next restoration batch before reporting this result.

- Active Production readback: `baec779b-db7a-43a3-b5bc-53a1ca8cd17e`, commit
  `3161b23283f003b42216c6370022a7d67d127a26`. Git automatic Production deployment
  is enabled; avoid a second manual deployment of the same change.
- Power Automate: DDH contact Flow `67ee065a-c230-4d1d-9a20-d9425f9cacdf` On;
  MMC contact Flow `64d9fad7-3462-4b6a-bb6e-95ce1951f4fb` Off. Monash Adults,
  Monash Paediatrics and automated VHH roster flows show Enabled. DDH scheduling
  remains to be verified under 1B.
- DDH run `08584110963160113529091988067CU09`: HTTP succeeded at 22:04:50 AEST.
  HTTP inputs/outputs are protected; protections retained. Retained current-day
  R2 revisions start at 14:26:06 AEST, after the user's 13:59 check. This
  supports absence of a current publication at that earlier test; no earlier
  manifest history was available to prove the exact reader response. Current
  manifest points to the 18:53 revision with 11 populated AM contacts and 14 PM
  contacts. The local reader loads it as available with 25 populated phone
  contacts. Browser matching/display still needs one combined visual check.
- Analytics credential and Wrangler login work. At 12:26 UTC the settled
  sample through 12:11 UTC showed 59,373 reads, 44 writes, maximum five-minute
  reads 3,615 and no expensive fingerprint. Valid sample; checker STOP only
  for missing second sample. This is not deployment admission.
- MMC/MCH defect reproduced using tracked Production flags: ingestion succeeds
  but publication excludes MMC/MCH. Local correction expands the publication
  list to `ddh,mmc,mch`; metadata/day builders, advanced maintenance and Preview
  remain closed. Not deployed.
- Regression exercises workbook ingress through published MMC/MCH/DDH readers,
  existing-source overlay publication and zero-write subsequent replay. It
  failed before correction and passed afterward. Focused ingress, rollout,
  contact-sync and workbook suites passed.
- No Production D1 query, flow edit, deployment, push or deployment deletion was
  performed. Diagnostics used analytics/control-plane reads and specific R2
  objects. Contact details were not printed or added to Git.
- Next: fresh budget admission, one release of the tested configuration,
  controlled current MMC extract publication and routine Flow restoration,
  followed by a combined DDH/MMC/MCH visual check. Reuse passing evidence;
  do not repeat transport experiments or restart the full inventory.
