# Shared colleague comparison implementation plan

Date: 14 October 2026. Baseline: commit `63d88bb`.

Status: implementation authorized on 14 October 2026 after the planning pause.

The aim is to make My calendar's “When am I working with…?” and At a glance's “Working together” use the same prepared roster data and browser cache. Cached searches should display promptly, repeated searches should avoid unnecessary authenticated requests, and server work should depend on the selected people and dates rather than every shift in a department. At a glance remains enabled per user.

## Evidence and scope

Both features currently read published R2 rosters rather than searching D1 for roster events. They nevertheless download and process entire facility ranges before selecting people. My calendar performs colleague discovery followed by a separate selected-colleague request. Working together requests its directory before showing its persistent cached result and then processes the roster before checking `cachedRevision`.

Small production samples from 14 October, using Richard and Abi and a 15 October–1 November range:

| Search | Sum of server durations | Sum of CPU | D1 rows read | D1 writes |
| --- | ---: | ---: | ---: | ---: |
| My calendar, uncached | 2,865 ms | 187 ms | 6 | 0 |
| Working together, uncached | 1,569 ms | 54 ms | 12 | 0 |
| Working together, repeat | 1,451 ms | 47 ms | 12 | 0 |

These are a few request samples, not browser interaction timings or a capacity test. My calendar's existing five-minute warm cache can avoid both comparison requests. Working together has a better sampled cold path, but continues to spend CPU on unchanged results. Neither complete implementation should simply replace the other.

The first release covers these two selected-person features across MMC, MCH, DDH and VHH. Other consumers of `queryRosterInsights`, including “Who”, and searches with no selected doctors retain their existing behavior unless an adapter change is required and separately verified. Casey follows its existing publication readiness policy. Personal calendars, subscriptions and On shift must continue to pass their existing checks.

## 1. Pin behavior and access contracts

Add characterization cases for both existing interfaces before replacing their data path. Preserve their visible results and defaults:

- “When” matches the same hospital and roster date. It does not currently require the shifts to overlap in time. Clinical support remains optional with its existing default.
- “Working together” requires actual time overlap at the same hospital. Preserve multiple-person, pair and all-person modes, its existing clinical-support handling, date clipping and overnight behavior.
- Preserve each feature's existing date-range and selection limits: do not widen them merely because the shared reader supports a larger range.

Move common working-event classification and date/interval normalization into a small shared module with two explicit matching modes. Use sorted intervals grouped by hospital/date for time intersection; avoid multiplying every shift by every other shift. Keep interface-specific presentation separate. Characterize discovery's current overnight and clinical-support candidate behavior as well as final matches, so a new index does not lose an eligible colleague.

Keep the permission checks distinct: My calendar insights eligibility and full At a glance eligibility remain separate. Server authorization determines permitted historical segments, publication visibility and identities. Use source-qualified doctor keys and approved aliases; avoid treating identical names at different sites as one identity without an approved relationship.

Verify Creator/God mode and doctor-profile requests explicitly carry the viewed subject. The current My calendar requests authenticate as Creator during impersonation but omit `targetEmail`; the new adapter must resolve the intended subject and enforce that subject's permitted scope. Browser cache partitioning includes both the acting session and viewed workspace.

## 2. Add a shared selected-person reader and API contract

Create a small comparison service under `functions/_lib/`, based on `loadPublishedDoctorSourceRange` in `doctor-source-artifacts.js`. Reuse the existing immutable per-doctor artifacts and small catalogue shards, including their parent-manifest fencing, term visibility, coverage gaps and missing-data checks.

Introduce one comparison API contract with context, discovery and selected-person operations. Existing endpoint names can delegate where needed for compatibility. The response carries authorized segments, a scope revision, data revisions, access expiry, missing/preparing status and only the requested facts. Authenticate and resolve access once per request; reuse request-local decoded R2 reads.

For a selected-person search:

1. Validate the feature, subject, date range, sources and bounded selection.
2. Resolve permitted historical segments and source-qualified selected identities.
3. Read publication metadata and the selected people's artifact pointers.
4. Calculate the comparison revision from the query, authorized scope, selected identities and their immutable digests, plus coverage/visibility state.
5. Return `unchanged` before calling `readEvents()` when the revision matches and the generation is complete.
6. Otherwise load only selected people's events, filter by authorized segments and return normalized facts.

Include carry-in night shifts at the start boundary using the bounded lookback required by existing attendance rules. Apply the existing entitlement rules to those events; fetching a lookback must not expand access. Test month, term and year boundaries explicitly.

A peer-only change must invalidate the pair; an unrelated person's change should not force selected event bodies to reload. Missing artifacts or a publication race are preparing/unavailable states, never an authoritative empty result. Do not cache incomplete generations as complete. Keep the old route available behind a rollout switch during preparation rather than automatically performing a costly broad scan when the new artifacts are unavailable.

## 3. Prepare a compact colleague-discovery index outside Workers

Selected-person artifacts alone cannot cheaply discover every possible colleague. Extend the existing Node/GitHub source publication process to produce a compact index partitioned by site and month, containing source-qualified identities, display labels and date-presence information sufficient for the existing discovery rules. Include working/clinical-support and overnight distinctions required by the characterization cases.

Build this from the already published doctor/month facts. Compact date masks or small date lists are appropriate; choose the simpler representation after measuring fixture size and query cost. Keep full events in the existing per-doctor store and calculate only requested comparisons. Precomputing every pair would grow quadratically and is unnecessary.

Publish content-addressed index objects and revision pointers with the source generation. Extend the automation upload/read whitelist, readiness check and promotion checks for the new index. Handle existing generations without it by rebuilding externally even when their parent manifest is unchanged. Do not make existing personal calendars or subscriptions unavailable merely because this additional derivative is being backfilled.

Discovery reads only relevant monthly summaries and returns eligible names, not everyone's events or an unrestricted index to the browser. Clip discovery to authorized date/site segments. No additional D1 roster-event searches or per-user background rebuilds are needed.

## 4. Use one browser data service and bounded cache

Create a shared browser comparison service backed by memory and IndexedDB, building on the existing snapshot-cache utilities. Both interfaces consume its normalized facts. Cache facts by person/site/range as well as assembled query results, so a different presentation mode or compatible selection can reuse loaded data. Reuse a wider date range only when its authorized scope covers the requested subset.

Cache identity includes acting session, viewed subject, authorized segment/scope revision, source-qualified identities, dates, publication/identity policy and schema version. Store selected-person revisions separately from discovery-index revisions. Do not depend only on the viewer's personal-calendar revision.

Use these freshness rules initially:

- Results may display immediately only after the exact feature and historical scope have been verified in the current session and remain within server access expiry. Persistent cached permission metadata alone is insufficient after reload or account switching.
- Reuse recent authorization/context for at most 60 seconds, bounded by server expiry and feature state. Never extend a server expiry in the browser. Clear/invalidate on logout, account/profile switch, entitlement updates, permission denial and incompatible identity changes.
- Reuse the existing visible-view publication fingerprint polling and its cadence. A publication change triggers a conditional authenticated check of the selected digests; an unrelated edit need not reload their event bodies. Polling must pause when hidden and must not create a second independent loop for these interfaces.
- Manual refresh and return after expiry force validation. On a transient refresh error, retain only a still-authorized cached result with a clear refresh status; a denial clears it. Missing coverage stays visible.

Deduplicate identical in-flight requests. A view closing should detach its subscriber without cancelling a request still needed by the other view. Old responses must not update a changed subject, selection or range. Bound cache entry count/size and prune least-recently-used entries; fall back to memory if IndexedDB fails. Store no credentials.

Show valid cached content before directory/network work and retain it during background validation. An unchanged response should not rebuild the result DOM or rewrite the persistent event payload. Hospital, matching-mode and clinical-support changes should be local when the loaded authorized facts cover them.

Combine context/discovery with selected facts when the saved selection is already known, or fetch independent pieces concurrently. A first visit that requires discovery to choose a default colleague can still require a second request. Do not prefetch every colleague or every site on login. Ordinary cold searches must use no more authenticated requests than today; fresh identical warm searches should use none.

## 5. Validate correctness and cost before enabling broadly

Extend existing doctor-artifact, cached-insight, snapshot-cache, facility-access, clinician-switching, calendar-refresh and client-request-budget suites. Add service integration tests that assert actual reads, events and request counts, not merely function names or source text.

Required cases:

- Same roster day with non-overlapping shifts; real time overlap; different hospitals; clinical support; leave/unknown/all-day entries; duplicate events and aliases.
- DDH/VHH overnight carry-in, Melbourne daylight-saving changes and cross-month/term/year searches.
- Multiple-person pair/all-person modes, empty selections, current/historical entitlements, and unchanged existing date limits.
- Normal login, Creator account, God mode and doctor profiles; permitted/denied features; no cached-result leak after scope or subject change; access expiry and revocation on validation.
- Unchanged comparison reads zero selected event bodies and zero full-facility roster bodies. Changed peer refreshes the comparison; unrelated peer leaves selected data reusable.
- Interrupted upload, missing index/body, stale parent manifest and retry do not activate or cache a partial generation.
- Identical concurrent searches share one request; rapid filter changes and account switching cannot display a stale response; cache failure remains usable.

Use synthetic fixtures with approximately 100, 500 and 1,000 staff, realistic shifts, a full permitted annual window and the largest existing selection. Measure decoded bytes, object reads, subrequests, memory and Node/browser computation as the dataset grows. Selected searches must scale with selected people, not total staff; discovery scales with compact relevant summaries. Local timings establish structural behavior, not production Worker CPU.

Then repeat matched cold/warm searches in Safari and obtain native Cloudflare samples with authentication and attribution intact. Record time to useful results, total UI time, median/p95 where sample size supports them, CPU per invocation, D1 rows/statements/writes, R2 body reads, response bytes and failure outcomes. Use the same people, dates, sites and CS settings when comparing.

Acceptance requires exact result parity for each matching mode, no D1 roster-event queries, zero user-search D1 writes, bounded authorization reads that do not grow with roster size, and no request-budget violations. Target routine conditional requests below the Free plan's 10 ms CPU allowance, with headroom where practical; investigate persistent excess rather than relying on occasional successful over-limit requests. Confirm the account's effective plan because `wrangler.toml` specifying 30,000 ms does not establish Free-plan capacity. All paths must remain within effective subrequest/memory limits, including maximum selection. A fresh cached reopen should display within about 100 ms on the tested browser, as a measured target rather than a promise for every device.

Project daily comparison usage at 500 users × 20 searches: at the observed 6–12 D1 rows per complete search, that would be 60,000–120,000 rows/day before cache savings. This is an illustration, not a total-app forecast; use the new measured cost plus existing background/account/automation usage for the final budget. Include Worker requests and R2 operations, not just D1.

Cloudflare currently documents 10 ms HTTP CPU, 50 subrequests and 100,000 daily requests for Workers Free, with flexibility for occasional CPU excess; sustained excess may terminate execution. D1 Free includes 5 million rows read/day and 100,000 rows written/day. Sources: [Workers limits](https://developers.cloudflare.com/workers/platform/limits/) and [D1 pricing](https://developers.cloudflare.com/d1/platform/pricing/). These limits are shared with other app activity.

## 6. Roll out and report

Implement in this order: behavior/access tests; shared selected-person reader; externally prepared discovery index and backfill; shared browser cache and both adapters; integrated performance verification. Keep each step reviewable, but deliver the two interfaces as one coherent change.

Prepare the source indexes before enabling the new route. Check four-site generation completeness and parity, then enable through a narrow switch and verify the signed-in production flows. Avoid running old and new readers on every normal user request; comparisons for acceptance should be deliberate bounded tests. Preserve a switch back to the previous comparison route if needed without reversing roster publications or the earlier restoration work.

When implementation is authorized, run the relevant checks, commit/push/deploy under the existing authorization, and finish with the measured before/after results and remaining limitations. Do not call the app safe for hundreds of concurrent users on the strength of a few production samples; distinguish fixture scale checks, limited live checks and any actual capacity testing.

Planning ends here. No application implementation or deployment is part of this planning pass.
