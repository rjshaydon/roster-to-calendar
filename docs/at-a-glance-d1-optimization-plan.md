# At a glance D1 read-reduction and shared-cache implementation plan

## Status and safety boundary

This document is an implementation plan only. Writing it does not authorise a production deployment, a production D1 migration, a production backfill, or re-enabling roster automation.

The protected historical production base is `aa9eed8` (`Stop runtime D1 schema
migrations`). As of 7 September 2026, GitHub `main`, the rollout branch and the
active Pages Production deployment identify `95bc9ad`. Migrations `0024` and
`0025` are applied; `0026` through `0030` remain pending. Runtime schema
inspection and repair must remain absent from every request path. All schema
changes described here must use explicit numbered migrations.

Development and performance testing must use local D1, local R2 substitutes, and production-sized synthetic or safely exported fixtures. No test, benchmark, query-plan check, cache warm-up, or backfill may point at production Cloudflare resources.

Before resuming Phase 7, complete the
[immediate D1 safety plan](./at-a-glance-immediate-d1-safety-plan.md). Production
legacy At a glance reads must be paused permanently. Until a hospital's shared
objects are published and its shared reader is explicitly enabled, that view
must return `preparing` or unavailable.

The Creator roster-status path has a separate proven broad-read query and is
currently blocked before D1 while roster writes are paused. Implement and pass
the [calendar store status remediation plan](./calendar-store-status-d1-remediation-plan.md)
before restoring roster automation; enabling writes must never reactivate that
legacy status query.

## Implementation handoff and restoration

Use [Sol's handoff and restoration checklist](./at-a-glance-sol-handoff-and-restoration.md) alongside this plan. It identifies confirmed temporary pauses, controls that must remain, source-by-source restoration, missing local harness dependencies and required deployment evidence. Restoring automation is part of the eventual rollout, not an automatic consequence of finishing code changes.

## Review outcome — 5 September 2026

The direction is correct: maintain one shared published view per ED/date or ED/term, compare incoming roster content with its existing contribution, and reuse unchanged objects across accounts. Opening the page must only authorise and read these views.

This revision strengthens seven areas: partial replacements and overlapping files; stable content comparison; recovery between D1 activation and R2 publication; date-dependent access; contact expiry without new uploads; whole-account resource budgets; and a rollout that cannot fall back to the known expensive queries.

This is a document review against local `aa9eed8`, the accessible messages in **Review doctor identity UX plan**, and current Cloudflare documentation. The task reader returned empty content for the four newest turns, so their detailed discussion could not be independently checked. The figures below are retained from the original plan, not newly measured production results. No production resource was queried or changed during this review.

## Objective

Make **At a glance**, especially **On shift**, suitable for routine use by most users while remaining comfortably inside Cloudflare's free D1 allowance.

The target design treats roster-derived information according to how often it changes:

- **Roster coverage, staff membership and roster grades** are calculated when a roster is ingested or activated.
- **On shift data** is built once per ED and date, then shared by all authorised users.
- **Designations and grade overrides** are small, separately versioned overlays.
- **Contact details** are a small, independently refreshed ED/date overlay.
- **User access** is determined at login or when access inputs change, not recalculated from term-wide events on every request.
- **Browser display choices** such as stream, seniority and clinical-support filtering are applied locally.

Opening At a glance must not calculate staff membership, roster coverage, grades, or a full-year view from raw `roster_events`.

## Current code and incident evidence

### Work already performed at ingestion

The roster ingestion and activation path already:

1. receives the parsed doctors, grades and events in memory;
2. writes one `roster_file_doctors` entry per doctor/file;
3. writes derived `roster_events`;
4. creates `roster_daily_presence` rows;
5. records continuing SMS membership; and
6. rebuilds daily presence when a staged roster becomes active.

This means the application already knows the facts needed to produce coverage, staff and daily ED summaries without rediscovering them during an At a glance request.

### Wasteful read-time work

The linked conversation reports two incident pathways: repeated whole-roster imports causing write amplification (the earlier incident report cites 272,288 rows written in 24 hours), followed by repeated runtime schema operations during ordinary requests. It records removal of the latter in `aa9eed8`. These are historical conversation reports, not fresh measurements. The local code confirms `ensureCalendarSchema` now performs no database work. This is distinct from the remaining expensive Staff and coverage queries, whose SQL is still present. Do not describe the latter as the sole proven cause of all outages or infer that caching alone guarantees prevention.

Cloudflare also [announced enforcement of the free D1 daily limits from 1 September 2026](https://developers.cloudflare.com/changelog/post/2026-09-01-d1-free-tier-limit-enforcement/). This is relevant timing, not proof of the application's individual query attribution.

The original plan records these production D1 insights (measurement window and query fingerprints must be attached to the implementation baseline):

| Read path | Observed rows read | Cause |
| --- | ---: | --- |
| ED staff membership | 668,326 in one execution | A correlated `EXISTS` search of roster events for candidate membership rows |
| File/ED coverage | 425,472 across six executions | Repeated `MIN`/`MAX` aggregation over active roster events |
| Additional coverage | 152,790 across six executions | Another repeated full-history coverage aggregation |
| On shift core roster query | About 109 per ED/date execution | Already date- and ED-scoped; worth caching but not the main incident source |

The browser also polls contact details every ten seconds while On shift is visible. Each poll currently repeats account authentication and can recalculate roster-derived access. This is not safe at multi-user scale even though each individual query is smaller than the Staff query.

### Production confirmation — 6 September 2026

The controlled rollout initially deployed the new system inert while
`FACILITY_LEGACY_READS_PAUSED=false` still allowed compatibility fallback. A
small Creator test across MMC, MCH, DDH and VHH raised Wrangler's rolling
24-hour total from 143,110 to 5,260,178 rows read. Cloudflare's one-hour query
insights attributed 3,440,682 reads to 11 legacy ED Staff membership queries,
averaging 312,789 rows examined per execution. This is direct production
evidence that the legacy compatibility route cannot be used even for a small
canary.

The immediate correction is fail-closed availability: disable legacy reads
before further migration work. The migration ledger update for `0024` did not
cause the spike.

### Current cache limitations

- On shift has no persistent shared ED/date cache.
- Staff responses are not shared by ED/term revision.
- By stream metadata is cached only in the current browser memory and is lost on reload.
- Contact polling discovers the latest object through D1 on each refresh.
- Existing account and doctor snapshot infrastructure is user/profile-oriented and is not the correct cache identity for shared ED views.

## Design principles

1. **Build on change, read many times.** Build roster-derived views during ingestion and publish after successful activation, never when a user opens the page.
2. **Share by data identity, not user identity.** The base cache key is ED/date/revision or ED/term/revision. Authorisation remains user-specific, but the roster payload does not.
3. **Keep independently changing data separate.** Roster, staff metadata, manual overrides and contacts have separate revisions and cache objects.
4. **Publish complete generations.** Build immutable objects first, then switch the manifest using a conditional write. D1 and R2 do not share a transaction; durable recovery must bridge the activation/publication gap.
5. **Keep reads read-only.** Missing or corrupt objects return a previous complete version or `preparing`; ordinary requests never build or repair them. A controlled ingestion/repair job owns all rebuilds.
6. **Prefer R2 for shared read models.** D1 remains the relational source of truth; R2 and the browser serve repeated roster reads without spending D1 rows.
7. **Use one schema for all EDs.** Do not create a database table per ED, department or user.
8. **Measure rows, not just response time.** Every relevant test must record or simulate rows read and examine the query plan.
9. **Fail closed in Production.** A missing switch, object, permission or
   publication returns `preparing` or unavailable. Production never falls back
   to historical At a glance SQL, including during canary or rollback.

## Target data model

Names below are illustrative; implementation should follow the repository's established naming conventions.

### 1. Roster-file coverage

Store these values on `roster_files` or in a small `roster_file_coverage` table:

- `file_id`
- `source_type` (normalised ED)
- `coverage_start`
- `coverage_end`
- `active`
- `content_revision`
- `staff_digest`
- `daily_digest`
- `updated_at`

Coverage is calculated directly from the incoming parsed events while those events are already in memory. It is never recalculated with `MIN` or `MAX` over the production event table during a user request.

An index must support the exact active-coverage lookup, for example ED plus active state plus coverage dates. The final index must be chosen using `EXPLAIN QUERY PLAN` against a production-sized local fixture.

### 2. Materialised ED/term staff

Maintain one compact logical staff set keyed by:

`source_type + term_start + stable_roster_identity`

Each entry should contain:

- roster doctor key and, when available, durable person ID;
- display name;
- roster grade;
- membership source;
- provider staff ID when supplied;
- first and last applicable dates;
- current roster revision; and
- whether the person is regular staff or an event-derived locum.

Choose one stable roster identity contract before implementation; do not alternate between doctor keys and person IDs depending on availability. Existing doctor keys can remain the source identity, with optional durable-person links in a separately versioned identity map. Record each contributing source/file and its effective dates so removing one file does not remove a doctor still supported by another file. Preserve date-specific grade changes within a term when the current UI needs them.

Explicit grade and designation overrides remain separate. They take precedence at display time and must not rewrite roster-supplied facts.

### 3. ED cache manifest

Store one small canonical manifest in R2 per ED. It contains:

- monotonic publication generation and contributing source revisions;
- actual covered date intervals, including gaps and explicit empty days;
- current staff revision by term;
- the content hash/revision for each available date;
- schema/version identifiers for cached payloads; and
- build completion time, source observation time, and last successful publication time (with distinct meanings).

Suggested key:

`facility-overview/v1/{ed}/manifest.json`

### 4. Shared daily roster snapshot

Store one immutable compressed R2 object per ED/date/content hash:

`facility-overview/v1/{ed}/days/{date}/{day-content-hash}.json.gz`

It contains only the fields required by At a glance classification and display. The snapshot is shared by every authorised user. The browser derives streams, seniority sections and clinical-support visibility from it.

Unchanged dates keep their exact existing key even when another date receives a new roster revision. Hash canonical uncompressed content, then compress for storage. Keep provenance and volatile import timestamps outside the content hash.

Do not create separate per-user or per-stream snapshots. If a future ED/day payload proves too large, measure before introducing a finer partition.

### 5. Staff, override and contact objects

Use separately versioned R2 objects:

- `facility-overview/v1/{ed}/staff/{term}/{staff-revision}.json.gz`
- `facility-overview/v1/{ed}/overrides/{term}/{override-revision}.json`
- `facility-overview/v1/{ed}/contacts/{operational-date}/{contact-revision}.json`

Separating these objects prevents a contact update or designation edit from rebuilding every daily roster snapshot. Add immutable identity/name-map objects where those values are projected rather than embedded in the base events.

The roster manifest points to staff, day and catalogue objects belonging to one complete generation. Contacts and manual overlays have independent small pointers; contact refreshes must not rewrite the roster manifest. Each overlay carries its scope, revision and any required identity/extract compatibility information. The browser pins a manifest for each render and rejects late responses from an older request/account/date.

Keep the hot manifest bounded to the supported active window; partition historical manifests by term or month before they become an ever-growing history download. Derive the By stream catalogue (unique shift patterns and coverage) during ingestion too; file min/max dates alone do not replace `queryFacilityOverviewCatalog`'s grouped event catalogue.

## Ingestion-time processing

### Confirmed roster lifecycle and sources

User clarification, 5 September 2026:

- Only changed roster facts should be updated when a revised roster arrives.
- In the month before a new term, each site may issue a draft roster and revise it multiple times.
- Each term is approximately three months. At a glance may expose that term from 14 calendar days before its actual start date. From that point, the operational expectation is a finalised roster with subsequent changes mainly for swaps, sickness and emergency leave. This records the user's product rule, not an independently verified legal requirement.
- DDH uses FindMyShift.com. MMC, MCH and VHH use Excel workbooks. Casey automatic updates are not yet connected because SharePoint access is pending.
- Formats differ between EDs, but the revision behaviour is similar. Site-specific adapters should produce the same normalised comparison input and use one shared change/publication pipeline. Casey can join that pipeline when its source is available; this plan does not depend on connecting it first.

This confirms incremental update behaviour, not that providers deliver only changed cells. The system may need to read and parse a complete revised workbook or provider extract to discover its changes. Compare that normalised source snapshot with the previously accepted snapshot for the same scope, then persist only additions, modifications and confirmed removals. Do not rebuild unrelated dates or accounts.

Keep current-term and next-term contributions separate: repeated next-term draft revisions must not replace the current term.

**Confirmed visibility rule:** `visibleFrom = actual term start date minus 14 calendar days`, evaluated using Australia/Melbourne local dates. At that boundary the entire available term becomes eligible for At a glance, subject to existing ED/account permissions. Use the application's actual term boundaries rather than assuming 90 days, adding a fixed three-month duration, or applying a rolling 14-day limit to individual shifts.

Earlier drafts may be ingested, compared and cached centrally in private storage, but must not be delivered through At a glance before `visibleFrom`. Apply this rule on the server to manifests, coverage/catalogue, staff, day and range endpoints, including direct requests for known object keys; hiding a selector alone is insufficient. Do not introduce a separate manual finalisation approval requirement. Reaching the date does not prove the provider has finalised the file: preserve genuine source status if supplied and do not label an unconfirmed draft as provider-finalised.

Store `termStart`, `termEnd` and `visibleFrom` with the shared term metadata. The reader checks the clock against that metadata; opening the window must not reparse the roster, rewrite event rows or rebuild unchanged snapshots. The next revision check must expose the eligible prebuilt term even if no import occurs at the boundary and no content hash changes. Include visibility state in the revision response/ETag and bound any cached response to the next visibility transition. Refresh an already-open browser at the transition or within the agreed refresh interval. Current-term data remains available alongside the newly eligible next term.

If the roster first arrives after the window opens, publish it through the normal change pipeline when available; show unavailable/preparing until then. Continue accepting later corrections, even if more extensive than the expected swaps or leave changes. “Finalised” is an operational expectation, not an immutable-data flag. This visibility rule applies to At a glance; it does not change other calendar or export visibility rules.

### Step 0: define replacement scope and stable comparison

An incoming file is a contribution, not automatically the complete ED or term. Persist its provider/source identity, declared date coverage, staff-group scope, revision/order and whether it is a full replacement or a patch. Recompute affected views from the **resulting active set** after existing supersession rules, retaining contributions from unrelated files. Missing staff/shifts imply deletion only inside an explicitly authoritative replacement scope. Preserve continuing SMS membership under the existing rule; absence from the latest file alone does not end it.

Coverage inferred from event min/max is insufficient for an intentionally empty date, a deleted last shift or a file spanning gaps. Retain declared coverage and actual occupied dates separately. An explicit empty snapshot means “covered, no shifts”; absent coverage means “no published roster”, not an empty working day.

Use two comparisons:

- Source byte hash/provider version can avoid parsing a previously accepted source only when scope, parser version and effective dependencies also match and that source is still the accepted version. A delayed older extract must not replace a newer one.
- Canonical semantic hashes compare sorted, normalised roster facts, preserving meaningful duplicates and changes to times, names, grades and locations. Exclude file IDs, upload times, row ordering and generated event IDs. Include schema/parser/rule versions where they affect the output. Identical bytes under a new parser may require rebuilding; a renamed/reordered but equivalent workbook should not republish every day.

Compare before relational replacement as well as before R2 writes. A true unchanged import must not delete/reinsert `roster_events`, membership or daily presence. A bounded audit/status record is allowed. For changed imports, implement fact-level relational differences as well: insert new shifts, update changed shifts, delete confirmed removed shifts, and update only affected daily-presence and membership rows. Existing whole-file replacement and whole-file daily-presence rebuilding do not meet this requirement for routine revisions. Replace those paths with staged change sets and bounded, recoverable activation; cost their index writes too. Cache diffing alone is insufficient. Stable provider shift IDs may help where available; Excel comparison must not depend on row positions, file IDs or generated import IDs. A change of doctor/time may be represented as a removal plus an addition when no stable shift ID exists. Publish both sides of a shift swap together in the affected ED generation.

All mutation paths must use this publication pipeline: automatic and manual imports, staged promotion, deletion, overlap trimming, rollback, reparsing and repairs. Contact, override and identity mutations use their corresponding overlay publishers.

### Step 1: derive summaries before D1 writes

While the parser's doctor and event collections are in memory:

1. calculate file coverage start and end;
2. normalise the incoming staff list and grades;
3. associate each doctor with the applicable medical term or terms;
4. add event-derived locums that are not in the formal membership list;
5. group events by ED and roster date;
6. calculate a stable digest for each daily group; and
7. calculate a stable digest for each ED/term staff set.

This computation does not consume D1 rows, but still costs CPU and memory. Prefer the existing parser/queue runner (`scripts/process-roster-queue.mjs`) for bulk hashing/compression; measure its actual runtime and the receiving Worker request limits. Do not move a whole-roster rebuild into a Pages request or assume `waitUntil` makes it durable.

### Step 2: compare small summaries

For each affected ED/term:

1. read only the current compact staff summary or its digest;
2. if the digest is unchanged, perform no membership writes;
3. otherwise compare the incoming and stored maps by identity;
4. insert genuinely new staff;
5. update only changed names, grades or metadata;
6. end or remove membership only when the complete resulting source set and continuity rules support removal; and
7. preserve manual overrides separately.

The comparison must never scan `roster_events`.

### Step 3: compare daily content

Compare each incoming ED/date digest with the existing manifest:

- unchanged date: retain the existing object;
- changed or new date: write a new versioned R2 object;
- removed date: omit it from the new manifest while retaining the old versioned object for safe rollback/expiry;
- unrelated ED/date: do nothing.

### Step 4: activate and publish

Continue staged activation, but explicitly implement a recoverable publication state machine. A single manifest write is atomic; the D1 and R2 changes together are not.

1. Acquire durable per-ED publication ownership with a generation/fencing token and bounded lease. A JavaScript `Map` only coordinates one Worker instance and is insufficient. A small indexed D1 coordination/outbox record is acceptable on this mutation path.
2. Validate source ordering, stage the input and compute the intended post-supersession active set. Write immutable changed objects and a complete candidate manifest referencing reused objects as well.
3. Verify required objects and record their hashes, previous generation and operation ID. Re-check ownership and supersession when committing.
4. Commit relational activation plus a durable pending-publication record atomically where possible. Existing multi-step deletes, overlap trimming and daily-presence rebuilds require staging or resumable checkpoints; they are not already one atomic operation. Do not advance publication until required derived state is ready.
5. Conditionally replace the R2 manifest only if its prior ETag/generation matches the recorded base. A stale publisher must never overwrite a newer publication. Conditional failure means reconcile/rebase; never blindly retry the old pointer write.
6. Mark the publication complete idempotently. Do not allow a later ED activation to bypass an unresolved publication. On crash/retry, resume from the durable record and recognise a pointer already successfully published.
7. Garbage-collect unreferenced objects in bounded maintenance batches only after the recovery, rollback and browser-cache retention windows. Never delete objects still reachable from retained manifests or pending builds.

If staging/building fails, the previous complete generation stays visible. If D1 activation succeeds but R2 publication fails, At a glance continues using the previous generation while a bounded repair job retries publication. Record this lag and expose last-published time; other D1-backed calendar views may already be newer. Do not claim cross-product atomicity or silently rebuild on a reader. Terminal failure stops further activations for that ED and raises an actionable operational error.

Test lease expiry and old workers resuming, not just simultaneous happy-path uploads. [R2 conditional writes](https://developers.cloudflare.com/r2/api/workers/workers-api-reference/) support pointer preconditions; [R2 consistency](https://developers.cloudflare.com/r2/reference/consistency/) does not make a separately cached manifest instantly fresh.

## Read-time processing

### On shift

The target request flow is:

1. verify a signed session and ED permission;
2. read the small ED manifest from cache/R2;
3. return the shared daily roster object for the selected date;
4. return or separately fetch the current contact object;
5. return the small override object if its revision changed; and
6. let the browser classify and render streams.

A warm request uses zero D1 roster-event reads. A missing object never invokes the old query or fills the cache on the request path. Return the retained complete generation with its publication time, or `preparing` with `Retry-After`. A deduplicated repair job with a predeclared ED/date and row budget may use the bounded query. Do not run one repair per viewer or automatically retry indefinitely.

### ED staff

Read the ED/term staff snapshot and the small override object. Never query `roster_events` from this action.

If the staff snapshot is missing, return a controlled `preparing` response; only the controlled repair pipeline may rebuild it. Do not fall back to the current correlated Staff query.

### By stream

Use the ED manifest for available coverage and stream metadata. For a requested date range, retrieve the relevant daily snapshots and filter/classify them in the browser.

Do not automatically build or retrieve an entire year. The normal initial range should remain small. A deliberately requested long range can be assembled from already versioned daily objects, with concurrency and maximum-range limits.

### Working together

Prefer the existing `roster_daily_presence` index or the shared daily snapshots. Query only the selected people and requested dates. Do not reintroduce an event-to-event range join over all staff.

## Contact refresh redesign

The ten-second loop must not perform D1 discovery, password verification and term-wide access calculation every time.

Preferred flow:

1. Contact ingestion writes a predictable latest R2 object or updates the ED/date contact manifest.
2. The browser sends a conditional request with its current revision or `ETag`.
3. After authorisation, an unchanged response returns no payload and performs zero D1 reads. Implement a dedicated authenticated GET for standard `ETag`/304 handling; current actions are credential-bearing POSTs and cannot simply start returning 304 to the existing JSON client.
4. A changed response returns only the small contact object.
5. Manual contact corrections update a separate small correction object/revision.

Also:

- stop refreshing when At a glance is closed or the document is hidden;
- do not refresh live contacts for historical/future dates unless required;
- use at most one combined revision check per visible browser session every 60 seconds, with jitter and cross-tab coordination; refresh immediately on reopening or manual request and stop polling hidden documents; and
- retain the previous object only while valid, show its source time, and mark an unsuccessful refresh explicitly.

One small revision response should cover the selected roster day, staff/override dependencies and contact overlay. This also lets an already-open page notice roster replacements; it must not poll only contacts and leave the roster stale indefinitely. Manual refresh and reopening check revisions, never rebuild data. Revalidate on visibility return with throttling; stop on logout and use bounded exponential backoff on 429/503/network errors.

Contact validity can change with the clock even when no new extract arrives. Publish `validFrom`, `validUntil`, operational date, previous-night eligibility and a `nextTransitionAt` derived from the existing `contact-allocations.js` rules. Recompute applicability locally at those transitions. Never keep displaying an expired phone assignment as current because its content hash is unchanged. Distinguish unchanged content from a successful fresh source observation, since freshness evidence can itself extend validity.

Corrections must be tied to the extract/contact identity and expected revision; a new extract must not inherit a correction solely because a row number or display name happens to match. Preserve existing overnight carry rules, expiry checks and optimistic-concurrency safeguards.

## Access and authentication redesign

At present, repeated At a glance calls send account credentials and can derive ED access from term-wide roster events. Replace that with a short-lived signed session or equivalent server-verifiable credential containing:

- account identity;
- role;
- permitted ED scope;
- access revision; and
- expiry and the next date/rule boundary that requires access recomputation.

Recompute access only when:

- the user logs in or renews the session;
- account permissions or identity links change;
- a relevant roster revision changes; or
- the token expires; or
- the Australia/Melbourne date, week or term changes in a way that changes the existing today/this-week/next-shift access rules.

Materialise compact doctor/site/date access facts during roster changes. Session renewal evaluates these facts rather than term-wide raw events. Preserve Creator/Owner, non-clinical, SMS, site-only, denied and entered-account behaviour from `resolveFacilityOverviewAccess` and its fixtures; do not broaden access to make caching easier.

A signed token containing `accessRevision` is not revocable by itself. Before enabling zero-D1 requests, implement an authoritative small account access/revocation record outside D1 (for example private R2) and compare it with a bounded cache lifetime. The agreed maximum permission-revocation delay is 15 minutes; fail closed after that freshness bound when access cannot be checked. Account/identity/permission writes must publish the access revision durably. Existing feature-disable controls must also take effect within the bound. Include these R2 reads and session renewals in cost tests.

Implement this once as a narrow common At a glance authentication middleware before switching the shared endpoints. Otherwise the existing `/api/state` credential checks still spend D1 rows before the cached action runs. Keep passwords out of cache keys and object URLs.

If a persistent access record is needed, it must be a small indexed record keyed by account and roster/access revision. The shared ED data itself must remain user-independent.

## Browser caching

Use persistent browser storage, preferably IndexedDB, keyed by ED/date/revision and ED/term/revision.

When a user opens At a glance:

1. verify a still-valid session, ED scope and access freshness, then render an authorised matching local snapshot;
2. request only the small current manifest/revision;
3. retain the local data when the revision matches;
4. download changed objects only; and
5. show publication/source times and distinguish cached from stale; do not render persisted data after the access validity window expires. Offline access beyond that window needs an explicit product decision.

Browser caches must be cleared or namespace-versioned when the payload schema changes. Sensitive shared responses must still be authorised before delivery; they must not be placed in an accidentally public CDN cache.

## Invalidation matrix

| Change | Recalculate or invalidate |
| --- | --- |
| Identical accepted roster received again | No data rebuild; bounded audit only, unless parser/rules/scope changed |
| New/replacement roster for one ED | File coverage, changed ED/date snapshots, affected ED/term staff diff and ED manifest |
| One corrected date | That day object plus any affected staff, coverage, catalogue or access summary; reuse all other unchanged hashes |
| New locum shift | The affected ED/date object and ED/term staff snapshot |
| 14 days before the next term starts | Expose that term's existing shared snapshots to authorised users; change visibility/revalidation state, with no roster rebuild or event writes |
| New term begins | Select the already prepared term and renew date-dependent access; published historical versions stay immutable, but corrections may publish a new historical revision |
| Roster-supplied grade changes | Affected ED/term staff entry and staff snapshot |
| Manual grade/designation override | Override object only; optionally staff presentation cache |
| Contact extract changes | ED/operational-date contact object only |
| Manual contact correction | Contact correction object only |
| Doctor display-name/identity decision | Identity/name map, affected staff/access/correction dependencies; rebuild daily objects only if affected values are embedded |
| Account permission change | That account's access/session revision only |
| File deletion, overlap trim or rollback | Recompute affected resulting source contributions, including removals and empty days |
| Parser/classification rule change | Only affected derived representations; never silently reuse hashes made under incompatible rules |
| Contact expiry, shift change or local date rollover | Re-evaluate time validity and date-dependent access even without an import |

## Query and cache budgets

These are acceptance limits, not aspirations:

| Operation | Required steady-state budget |
| --- | ---: |
| Warm On shift roster load | 0 D1 roster-event rows |
| Unchanged contact refresh | 0 D1 rows |
| ED coverage lookup | 0 D1 rows on the shared manifest path; indexed compact reads during builds only |
| Warm ED staff load | 0 D1 roster-event rows |
| Cold staff materialisation | No correlated event subquery; bounded to the affected ED/term |
| Access check after login | 0 roster-event rows |
| Browser revisit with unchanged revision | Manifest/revalidation only; no D1 roster read |
| Cache miss | 0 request-time builds; one deduplicated bounded repair job per affected revision |

Additional guardrails:

- No At a glance request may run `MIN`/`MAX` coverage aggregation over `roster_events`.
- No At a glance Staff request may query `roster_events`.
- No production request may execute schema DDL or `PRAGMA table_info`.
- Long date ranges must be explicit, bounded and assembled from dated objects.
- Cache builders must use durable publication ownership; reader concurrency cannot trigger duplicate builds.
- D1 result metadata and query fingerprints must be logged without patient, staff-contact or credential data.

## Whole-account capacity and outage containment

Zero D1 roster reads is necessary but not a complete free-plan budget. Current [D1 allowances](https://developers.cloudflare.com/d1/platform/pricing/) are 5 million rows read and 100,000 rows written per day, resetting at midnight UTC; index maintenance contributes writes. Measure both against all databases and execution surfaces on the account. Melbourne reset time changes with daylight saving, so operational reporting must use the UTC quota day rather than a rolling 24-hour chart.

[Workers Free](https://developers.cloudflare.com/workers/platform/limits/) allows 100,000 requests/day and has CPU/subrequest limits. [R2 Standard](https://developers.cloudflare.com/r2/pricing/) includes 10 GB-month storage, 1 million Class A operations and 10 million Class B operations/month; it is not an unlimited free cache. Conditional/304 responses save transfer, but requests and backing object reads still count.

Illustrative capacity calculation, before login, calendar subscriptions, other endpoints, retries and background work:

| 100 simultaneous viewers, 8 hours/day | 10-second polling | 60-second polling |
| --- | ---: | ---: |
| One revision request per interval | 288,000/day | 48,000/day |
| R2 reads if each check reads two objects, 30 days | 17.28 million/month | 2.88 million/month |

These are arithmetic examples, not measured traffic. Ten-second polling already exceeds the daily Worker request allowance in this example. Internal edge caching may reduce R2 operations but does not remove the authenticated Worker request itself. Size a combined revision check and auth lookup against the user's actual peak viewer-hours; do not deploy separate always-on pollers for each overlay.

Before rollout, fill in a capacity worksheet for ordinary days, peak imports, doubled traffic, retries and cache-cold conditions. Include account login/state, calendar feeds, roster ingestion, contacts, automation, preview deployments, maintenance, indexes and migrations. Provisional go/no-go target: projected normal/peak planned work uses at most 50% of each relevant allowance, retaining the rest for other traffic and recovery. Adjust only using measured evidence and explicit capacity assumptions, not a claim that the free plan is guaranteed.

Define a small shared operational control that can stop optional builds and expensive paths before D1 use. Existing environment pauses remain available as an emergency control; a browser flag or isolate-local circuit breaker is insufficient. At 50% actual usage or a forecast approaching the daily limit, investigate and pause optional work; at 70%, stop optional database reads/builds and retain shared published reads subject to valid authorisation. Monitor outside the hot D1 path, rate-limit ingress and cap repair retries. Metrics are delayed, so these thresholds supplement pre-budgeted batches and endpoint limits; they cannot guarantee protection from one unbounded query. Never fall back to broad SQL on quota/cache errors.

`FACILITY_LEGACY_READS_PAUSED=true` is a permanent Production invariant, not a
late-rollout switch. Add a code-level fail-closed default before enabling any
shared reader so a missing or malformed variable cannot restore the legacy
route. Retain a legacy implementation only for explicitly isolated local
correctness comparison until it can be removed.

`EXPLAIN QUERY PLAN`, mock call counts and returned-row counts are useful but do not prove Cloudflare billable row usage. Record local measurements as estimates, then validate actual `meta.rows_read`/`rows_written` during the separately approved, tightly bounded rollout. A SQL `LIMIT` alone does not bound rows scanned. Logging must not write one D1 telemetry row per request.

## Implementation phases and gates

### Phase 0: local safety harness and baseline

- Confirm the deployed production version and binding inventory at implementation time, then branch from that baseline, not unfinished identity work. Include Pages production/preview, separately deployed Workers, GitHub schedules, manual paths and other clients sharing D1. Local Git HEAD alone does not prove deployment state.
- Preserve all unrelated and untracked files.
- Confirm local network isolation prevents accidental Cloudflare D1/R2 access.
- Build a production-sized local fixture containing at least the current number of files, doctors and events, plus headroom.
- Capture current row-read estimates and `EXPLAIN QUERY PLAN` output for Staff, coverage, On shift, access and contacts.
- Add tests that fail when a prohibited broad query reappears. Include the whole endpoint before action dispatch, authentication, error handling and disabled-feature paths. Cost explicit migrations and index creation before approval too. The handoff identifies `/api/automation/pending` as an existing pause gap with mutation during listing, and local-safety scripts missing from this baseline; address both before relying on the harness or pause controls.

Gate: no remote writes, migrations or deployments; baseline report reviewed.

### Phase 1: coverage and staff materialisation

- Add an explicit migration for compact coverage and ED/term staff state.
- Extend ingestion to calculate coverage and staff/daily digests in memory.
- Implement set comparison and diff-only event, daily-presence and staff writes, using staged change sets; routine revised imports must not rewrite the whole file.
- Preserve manual override precedence and locum behaviour.
- Remove event-table dependency from the new Staff and coverage repository functions.

Gate: production-sized local tests demonstrate the row budgets and correct term transitions. Do not switch the UI yet.

### Phase 2: access/session materialisation

- Introduce short-lived signed sessions or compact access records.
- Remove term-wide roster access derivation from every At a glance request.
- Implement and test the revocation, date-boundary and entered-account behaviour before enabling shared reads.

Gate: repeated authorised calls perform no roster-event access queries and unauthorised ED access remains impossible.

### Phase 3: shared staff and metadata reads

- Build versioned staff and manifest objects in the local R2 substitute.
- Change Staff and metadata actions to read only materialised data.
- A missing materialisation must return `preparing` rather than call the old broad queries.
- Use the shared access/session middleware from Phase 2 for every cached endpoint. Verify all ED, All EDs and entered-account permission rules.

Gate: repeated local Staff/metadata requests perform zero `roster_events` reads.

### Phase 4: shared On shift day snapshots

- Generate ED/date snapshots and per-date digests during ingestion.
- Publish complete generations with durable recovery and conditional manifest writes; test the D1/R2 failure gap.
- Return last-published or `preparing` on misses; exercise the separate deduplicated repair path for missing dates.
- Make browser filtering derive all existing stream and seniority presentations from one ED/day payload.

Gate: concurrent readers across simulated independent Worker instances reuse one published object and perform zero D1 roster reads; misses do not build. Crash injection at each publication step proves recovery without lost or mixed generations.

### Phase 5: contact overlay

- Publish predictable, versioned contact objects.
- Replace D1 discovery in the refresh loop with revision/`ETag` revalidation.
- Stop or slow polling in inactive contexts.
- Keep contact corrections separate and preserve existing allocation safety checks.

Gate: an hour of simulated visible On shift polling with unchanged contacts performs zero D1 reads after authentication.

### Phase 6: browser persistence and range views

- Add IndexedDB snapshot storage with schema versioning.
- Revalidate manifests rather than downloading unchanged data.
- Rework By stream and Working together to use daily snapshots or bounded daily-presence lookups.
- Preserve all current UI behaviour and terminology.

Gate: reload, offline/stale handling, account switching and revision changes cannot expose another user's unauthorised data.

### Phase 7: controlled production rollout

Production rollout requires separate approval after all local gates pass.

1. Complete the immediate D1 safety plan: deploy
   `FACILITY_LEGACY_READS_PAUSED=true`, verify the effective setting without an
   At a glance request and keep it true permanently.
2. At the start of a fresh UTC quota day, review a costed migration/backfill and
   deployment plan against current-day read/write budgets. A rolling 24-hour
   figure is not the current-day quota counter. Apply explicit migrations only;
   never runtime schema setup.
3. Keep all shared builders/readers and automation disabled initially. Apply
   the remaining migrations individually, starting with `0025`; `0024` is
   already applied and must not be repeated.
4. Populate materialisations from retained roster sources or the normal next ingestion, avoiding a broad production D1 backfill.
5. Validate one ED and one term without exposing a reader that lacks published
   objects.
6. Enable the new read path for the Creator and that ED only. Every other At a
   glance route remains unavailable, never legacy.
7. Inspect D1/R2 metrics, query fingerprints and correctness. Any
   `roster_events` or `roster_file_doctors` access from an At a glance request
   fails the gate immediately.
8. Expand one ED at a time only after its shared objects exist.
9. Old expensive SQL must be unreachable from every Production endpoint,
   including blocked, missing-object and error paths. On failure, roll back to
   the previous published generation or disable the affected view; never
   restore the old Staff/coverage query as a fallback. Remove dead code after
   parity is proven.
10. Restore automation only using the linked handoff runbook after bounded ingestion tests, source-specific gates, backlog reconciliation and approval. Restore the independent watchdog last; retain permanent quota fixes and separate maintenance controls.

Any production backfill must declare its maximum expected rows read and written before execution, stop between bounded batches, and remain below the agreed capacity budget. Create small new tables first where feasible; a “migration only” index over a large table may itself consume substantial quota.

Do not declare success after one cheap live request. Observe at least seven consecutive UTC quota days, including a realistic roster replacement and contact refresh cycle, and record daily totals, top query costs, projection against peak traffic, publication lag and errors. Longer observation or a representative bounded exercise is needed if no import occurred. Verify the deployed code, flags and resource bindings at each rollout stage. The plan does not itself establish a monitoring automation or authorise production changes.

## Test requirements

### Correctness

- Every existing On shift fixture produces the same recognised working events, ordering, stream classification and contact attachment.
- Staff lists preserve regular staff, continuing SMS logic, locums, manual designations and grade overrides.
- Term boundaries do not leak a new grade into an old term or vice versa.
- Re-uploading identical content performs no materialisation change.
- Correcting one date changes only that daily object and genuinely affected summary/access dependencies.
- Activating a staged roster publishes all related objects together.
- Failed ingestion leaves the previous manifest and cached views usable.
- Full replacements, patches, overlapping sources, deletion of the last shift, missing doctors, continuing SMS and historical corrections produce the correct resulting active set.
- Row reordering, renamed files and repeated uploads reuse hashes; parser-version changes and meaningful duplicate changes are detected.
- Late/out-of-order roster or contact extracts cannot overwrite newer accepted content.
- Multiple next-term draft revisions preserve current-term data and reuse unchanged dates.
- The next term is unavailable via all At a glance endpoints at 15 days before its start and becomes eligible at local midnight 14 days before; direct object requests cannot bypass the rule.
- With no import at the visibility boundary, an already-open page discovers the whole eligible term through clock-aware revalidation, with zero event writes or snapshot rebuilds. Test 304 handling and cached manifests across the boundary.
- Actual term dates, Melbourne daylight-saving transitions, a late first roster and corrections after the visibility boundary are handled without assuming a fixed 90-day term or locking finalised data.
- A sickness correction or shift swap updates only changed facts, affected daily presence and dependent summaries; both sides of a swap appear together. Assert writes as well as final rendered output.
- Normalised DDH provider and MMC/MCH/VHH Excel fixtures exercise the same diff/publication contract despite different source formats.
- Midnight, week/term boundaries, overnight shifts, contact operational-day transitions and Melbourne daylight-saving changes match existing semantics.
- Contact corrections cannot attach to an incompatible new extract; contact expiry removes invalid allocations without requiring new content.
- An already-open page learns of a changed roster as well as contacts within the agreed freshness bound.

### Performance

- Run production-sized and doubled-size fixtures locally.
- Assert exact repository calls made by every endpoint.
- Use `EXPLAIN QUERY PLAN` to reject full scans and correlated scans on hot paths.
- Record cold-build and warm-read row counts separately.
- Simulate concurrent users requesting the same uncached ED/date and prove no request builds; only one fenced repair owner can publish a generation, including lease expiry/retry scenarios; resumptions are idempotent.
- Simulate contact polling for one hour and 100 users without D1 growth when the contact revision is unchanged. Count Worker requests, R2 GET/HEAD/PUT operations, auth-record refreshes and CPU as well; include background tabs, multiple tabs, errors and retry bursts.

### Security

- Authorisation occurs before a shared object is returned.
- Cache keys never include passwords.
- Browser storage is cleared or isolated on logout/account switching.
- A user restricted to one ED cannot obtain another ED's shared snapshot by editing a request.
- Signed access state expires and revocation takes effect within the agreed bound, including cached-token and offline cases.
- Shared storage is private, authorisation precedes object reads/304 responses, and stale in-flight responses cannot repopulate another account's UI after logout/switching.

## Confirmed product details

1. Incremental update behaviour, site sources and the 14-day pre-term visibility rule are confirmed above. Verify each adapter's actual extract coverage and how cancelled shifts are represented using existing parser contracts and fixtures. Ask for a source-specific example only if those leave ambiguity.
2. Refresh contact allocations every 60 seconds only while On shift is visible, immediately on reopening or manual refresh, and never from a hidden document.
3. Design and test for 50 simultaneously visible browser views, with hard request budgets, jitter, cross-tab coordination and backoff rather than treating 50 as a guaranteed maximum audience.
4. SMS membership is enduring and is removed manually by a Creator; absence from a term extract can represent LSL, sabbatical or other extended leave and must not remove membership. Non-SMS staff membership is department- and term-specific: trainees and other non-SMS doctors count as staff only when they appear on that department's roster for the term, and may disappear and return in a later term.
5. Do not render downloaded roster or contact information after authorisation expires. Permission revocation takes effect within at most 15 minutes. Expired contacts disappear immediately according to their validity metadata. If roster refresh fails during valid authorisation, retain the last valid roster snapshot with a clear last-updated/stale warning.

## Implementation reference points

- `functions/_lib/d1-calendar.js`: `queryFacilityOverviewStaff`, `queryFacilityOverviewCatalog`, `queryFacilityOverviewRange`, `queryFacilityOverviewOnShift`, and `queryFacilityOverviewAccessEvents`; these identify existing query behaviour to replace or preserve.
- The same file: `replaceDerivedRosterFile`, `promoteVerifiedStagedRosterFile`, `trimDerivedRosterFileOverlap`, `deleteDerivedRosterFile` and daily-presence helpers; all need publication/invalidation coverage and write-cost accounting.
- `functions/api/state.js`: authentication before action dispatch, `resolveFacilityOverviewAccess`, `loadLiveContactListForOnShift`, entered-account authorisation, contact correction and roster mutation actions.
- `public/static/app.js` and `public/static/contact-allocations.js`: polling lifecycle, account changes, roster classification, operational-date and contact validity semantics.
- Existing regression entry points in `package.json`: facility access, contacts, roster fixtures, queue failures and D1 quota guards. Extend them with meaningful cost/failure fixtures rather than treating their current pass as evidence for the unimplemented design.

## Definition of done

This work is complete only when:

- Staff membership, roster grades and coverage are ingestion-time facts;
- the old Staff correlated event query and read-time coverage aggregations are no longer reachable;
- On shift data is shared by ED/date/revision rather than built per user;
- unchanged contact refreshes consume zero D1 rows;
- access checks do not rescan roster events;
- browser and R2 caching reuse unchanged revisions;
- all local cost, correctness and security gates pass;
- production rollout is separately approved and performed in bounded stages, with the restoration ledger completed or explicit remaining pauses recorded; and
- the observation gate and whole-account worksheet show headroom for D1 reads and writes, Worker requests/CPU and R2 operations/storage under realistic projected traffic. No claim of guaranteed outage prevention substitutes for these measurements.
