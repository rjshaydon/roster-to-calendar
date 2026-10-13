# Shared colleague comparisons — acceptance record

The two selected-person features now use prepared per-doctor roster facts and a shared browser cache. My calendar still matches roster days; Working together still matches actual time overlaps. Their feature permissions remain separate.

The server compares metadata before reading events, accepts per-person revisions and reads only changed people. Identical fresh searches and new pairings of already cached people avoid authenticated requests. Fresh authorized broad results can be narrowed locally. Cache authorization is memory-only, subject/scope-specific and bounded to 60 seconds and server expiry. Missing coverage remains explicit when partial results are cached or acknowledged unchanged; an unfinished publication is unavailable and cannot become an empty authoritative result.

Discovery uses monthly date masks partitioned by roster term. This preserves previous-term overnight attendance without exposing an unreleased term. External preparation uses the existing GitHub workflow, resumable person receipts and an immutable completion certificate. No per-user build or D1 roster-event query was added. Larger existing UI selections use requests of at most three roster keys to bound Worker subrequests.

The production stage first failed because adding a rollout variable exceeded the Workers Free limit of 64 text variables/secrets. Omitting the empty `ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256` value is equivalent to its existing runtime default; the source allowlist, hard fact limit and explicit pinned-hash support remain intact. A regression assertion confirms that equivalence. The retry deployed successfully as `de114ebc`, initially with the comparison reader disabled. Preview remains disabled.

The external jobs `38467895411` and `38469808327` completed successfully. Read-only downloads confirmed matching native parent ETags and completed v2 indexes for MMC (10 months), MCH (10), DDH (12) and VHH (9).

## Validation

Eighteen relevant suites/checks passed: shared comparisons, doctor artifacts, cached insights, facility access, clinician switching, calendar refresh, overview defaults, client request budgets, snapshots, subscriptions, materialisation, delivery reliability, source isolation, all-staff On shift, restoration batches/mutations, database cost checks and JavaScript syntax.

The new suite executes the actual authenticated endpoint adapters and UI loader/cache, including God mode target isolation, stale responses, authorization expiry, unchanged DOM/storage behavior, cached hospital narrowing, per-person reuse/deltas, DST and term-boundary carry-in. It also covers publication failure and absent indexes. Older restoration harnesses needed missing globals and an incorrectly placed rejection assertion corrected; their intended checks now execute.

Synthetic annual fixtures with 100, 500 and 1,000 staff each returned 384 shifts for two selected people using exactly two event bodies. Conditional checks used zero event bodies. The compact discovery index for one month was approximately 5.4, 27.4 and 54.9 KB respectively. Local selected-reader execution was approximately 1 ms; this is **not** Cloudflare CPU or browser response time.

Live publication parity was verified using read-only downloads from the existing Cloudflare API into restricted local files, without creating another endpoint. The comparison window was 1 October 2026–31 January 2027, including the released upcoming term. Two selected identities were sampled per site.

| Site | Selected shifts | Shift-field parity | Eligible colleagues | Discovery parity | Old R2 gets | New R2 gets | Conditional event bodies |
| --- | ---: | --- | ---: | --- | ---: | ---: | ---: |
| MMC | 9 | Exact | 112 | Exact | 7 | 5 | 0 |
| MCH | 21 | Exact | 65 | Exact | 7 | 5 | 0 |
| DDH | 40 | Exact | 136 | Exact | 10 | 5 | 0 |
| VHH | 49 | Exact | 52 | Exact | 10 | 5 | 0 |

Get counts describe calls to the R2 wrapper, excluding native version-tag heads. Most importantly, full searches use only the selected event bodies; conditional searches decode neither selected event bodies nor full-facility rosters. These limited samples and fixture scale checks do not establish capacity for hundreds of concurrent users.

## Remaining live checks

The Mac was locked, preventing signed-in Safari interaction and end-to-end production timings. No claim is made that every authenticated invocation is below the Free plan's 10 ms CPU allowance. Actual native CPU/D1 samples and signed-in Safari cold/warm behavior should be checked once the Mac is unlocked.

Automatic approval review rejected an optional expiring, key-protected comparison Worker because it would add an internet-facing path to private roster data without explicit authorization. That Worker and its secret were never created. Verification instead used existing read-only API downloads. Private artifacts remain restricted locally; only aggregate evidence is recorded here.

The optimized production reader is enabled only after these readiness and parity checks. Rollback is setting `ROSTER_COMPARISON_ENABLED` to `false` and deploying; roster publications, calendar artifacts and subscription delivery need not be reversed. Reloading the app picks up the final versioned browser module.
