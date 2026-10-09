# On shift for all roster-linked staff

Status: implemented and tested locally, 9 October 2026. Production release authorized.

## Intended behaviour

Every signed-in clinician linked to a published working shift can access On shift from one hour before that shift until one hour after it ends, regardless of their At a glance toggle. Keep the existing Melbourne-time policy, leave exclusions, midnight handling, and active-shift preference.

Automatically open On shift at login/session restoration within that window, showing the shift's ED and roster date. Open it once per account transition; a delayed response must never override subsequent manual navigation. Users can return to My calendar and reopen On shift while eligible. No background polling to discover that an upcoming shift has entered its window in this first release.

The existing At a glance toggle continues to govern the broader overview: ED staff, Working together, By stream, history, and ordinary ED/date navigation. On shift-only users see a separate On shift navigation label and only their eligible ED/date. Do not change stored permission flags or bulk-enable accounts. Existing enabled users and non-clinical Directors keep their current behaviour.

“Everyone” means roster-linked accounts on sources with published shared-reader coverage. An unlinked account cannot establish its shift. Currently the configured shared-reader sources are MMC, DDH, MCH and VHH; Casey requires publication/reader enablement before the same feature can be offered there. Do not enable legacy D1 roster reads to fill that gap.

## Findings in the current implementation

- `public/static/shift-launch-policy.js` already calculates the desired window.
- `launchClinicalOnShiftWorkspace` already opens the appropriate ED/date, but depends on full overview entitlement.
- `queryFacilityOverviewLaunchWindow` and `queryFacilityOverviewOnShift` reject accounts whose overview toggle is off.
- The existing startup eligibility helper filters shared published roster data by account claims. Approved identity aliases must also be respected when extending it.
- Every ordinary state API request authenticates using `loadAccountMirror`, which performs an indexed account query with claim/state joins. One request is not necessarily one D1 row read.
- Opening On shift currently also requests available terms. Restricted users do not need that request; its membership lookup can use D1.
- Daily rosters already have an R2 reader and a local revision cache. Contact polling already uses a short-lived signed token before D1 authentication. Token expiry stops contact polling on authorization failure; it does not automatically renew indefinitely.

## Implementation approach

1. Separate full overview entitlement from temporary On shift eligibility in the server and client contracts. Represent the eligible facility, roster date, start and end explicitly. Keep broader API actions behind the existing entitlement checks.
2. Extract a shared server eligibility helper using trusted account claims and approved identity aliases, the published roster, and the existing timing policy. Reuse alias resolution within a request. Never use private/custom calendar events or a client-supplied shift as authorization.
3. Reuse eligibility evidence available during existing login/calendar hydration when authoritative and current. Avoid adding a blocking all-site roster load to fast login. If eligibility cannot be established there, permit at most one deduplicated authenticated startup fallback per account transition, as the app already does. A negative authoritative result must not trigger another fallback.
4. Allow On shift requests for an eligible account even when the overview toggle is off. Validate the requested ED/date against the server's current shift window before reading shared data. For restricted accounts, bypass full overview access/membership resolution; use only bounded account/identity lookups and R2 roster data. Reject All EDs, arbitrary dates, and legacy roster query fallback.
5. Adjust navigation, page-opening guards, controls and account switching to distinguish the two access modes. Restricted opening must skip term menus, stream metadata, staff directories and other overview requests. Use the existing local roster cache and revision checks, scoped to the viewed account and ED/date.
6. Preserve contact refreshes without D1 reads. For restricted access, issue contact authorization scoped to the eligible ED/date and cap expiry at the shift-window end as well as the existing short token lifetime. Update token validation as needed without changing ordinary overview access. Refresh after expiry only through an explicit authorized reload in this release; do not introduce an automatic D1 renewal loop.
7. On restricted-window expiry, stop refreshes and prevent further access to the restricted roster view, including cached reopening. If still on that page, return to My calendar. Account changes clear eligibility, data and tokens; Creator impersonation follows the viewed user's access.

## D1 and performance budget

Prefer the existing authenticated API architecture for the initial release. A new general roster-token/session system is unnecessary given permission for a modest increase in D1 reads.

Target added startup traffic: one On shift roster request for eligible users, plus at most one eligibility fallback if existing hydration cannot supply the result. Off-shift users must never fetch the ED roster or term menus. Incremental D1 work must remain bounded to account and approved-identity lookups, independent of the number of staff in an ED roster.

Illustrative request volume: 500 users opening the app twice daily gives 1,000 opening sessions. If every session is eligible, this design adds roughly 1,000 authenticated roster requests and at most 1,000 eligibility fallbacks. These are request estimates, not D1 row estimates; identity queries and joined claims require measurement. Repeat explicit reloads add requests. A minute-by-minute polling design would scale with viewing duration and is deliberately excluded.

Contact refreshes must continue to consume zero D1 reads. Eligibility and daily roster fetching must perform no D1 roster-event scans, term membership queries, schema mutations, or maintenance. Bound R2 reads to linked, supported sources and the small date range needed for overnight shifts. Missing publications should show a preparing/unavailable state while leaving the personal calendar usable.

Before rollout, measure D1 statements and rows read, R2 operations, request count, first calendar paint, and On shift paint for cold/warm sessions. Compare with the existing baseline, including identity review enabled. Do not claim an exact row budget until the traced endpoint tests confirm it.

## Validation and release

- Test toggle-off eligibility and automatic opening; toggle-on behaviour must remain unchanged.
- Cover before/after boundaries, overnight shifts, DST, leave, overlapping shifts, aliases, missing publications and unsupported sources.
- Prove server rejection of forged ED/date requests, all other overview actions, expired windows and cross-account cache/token reuse.
- Test account switching, Creator impersonation, delayed startup responses, manual navigation and cached sessions.
- Extend meaningful request/D1 budget tests to exercise actual handlers: no term/metadata calls for restricted opening, at most one startup fallback, no D1 roster scans, and zero D1 contact refreshes. Check both unchanged and changed revisions.
- Run the existing shift-launch, overview-opening, access, contact-access, clinician-switching and request-budget suites, plus syntax checks.
- Put the new access path behind its own environment flag. Verify it in preview with eligible and ineligible accounts, then enable production through the normal release workflow. Monitor real request/row usage and startup timings. Disabling the flag restores the existing per-user overview behaviour without account migrations.

## Completion criteria

A clinician with At a glance disabled opens directly to their eligible On shift page; other overview views remain unavailable. Outside the window, their personal calendar opens normally. Hundreds of users increase bounded account-level request volume, not roster-wide D1 scans or continuous authenticated polling. Existing overview users and Directors retain their current access.


## Implementation and validation record

Implemented the separate client access path and server-side shift authorization. Restricted users receive fixed ED/date controls, an On shift navigation label, read-only phone allocations, and automatic opening after server confirmation. Window expiry and account transitions clear temporary access. Full At a glance users retain their existing views and navigation.

The first release uses one deduplicated startup fallback for restricted users; it does not add blocking roster/eligibility work to login or introduce a general roster-token system. Daily rosters stay on R2. Date-scoped contact tokens expire within 15 minutes or at the shift-window end, whichever is sooner. There is no automatic authenticated renewal loop.

`ON_SHIFT_FOR_ALL_ENABLED` is explicitly true in the production configuration and false in Preview. This is a local configuration change; production has not been deployed. Existing source allowlists remain in force.

Actual-handler tests with identity review enabled traced three D1 statements for the basic eligibility check and three for a restricted roster load. Approved linked identities use five statements per request. These are statement counts, not Cloudflare billed-row measurements. Increasing the day roster from two staff to 500 changed neither the D1 statement/result count nor the R2 operation count. Contact refreshes used zero D1 statements. No roster-event scans, full-overview membership queries, mutations, or maintenance ran on the restricted path.

Validation passed: the new all-staff handler/client suite, shift-launch boundaries and DST, overview opening/defaults, clinician access/switching, facility access, contact tokens/allocations, DDH night review, maintenance, rollout, snapshot cache isolation, client request budgets, and 36,000 shared contact refreshes. JavaScript syntax and Git whitespace checks also passed. Tests cover unsupported sites, missing publications, unapproved aliases, forged dates/EDs, direct broader-view denial, Creator impersonation, delayed navigation/account changes, cached reopening and expiry.

Two obsolete assertions in the maintenance test were corrected: production already enabled automatic launch before this change, and the page-opening implementation already used a separate asynchronous helper. Existing isolated client fixtures now provide the new full-access predicate.

Live browser/production timing and Cloudflare billed-row measurements remain release checks. The follow-up request authorizes commit, push and deployment; live user testing follows deployment.
