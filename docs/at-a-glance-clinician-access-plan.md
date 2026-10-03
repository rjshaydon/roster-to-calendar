# At a glance clinician access plan

Prepared: 2 October 2026 (Australia/Melbourne).
Intended implementation session: 3 October 2026.
Status: implemented locally on 3 October 2026; not deployed. See implementation and verification notes below.

## Intended behaviour

At a glance should make a clinician's existing roster access convenient. It
should not grant trainees access to unrelated hospitals.

All rules below require an enabled At a glance account. Existing Creator,
owner and explicitly enabled non-clinical access remains supported.

| Viewer | Ordinary At a glance access | Retrospective Working together |
| --- | --- | --- |
| SMS or CMO | All published hospitals | All published hospitals, subject to available history |
| Other clinician, one current-term hospital | That hospital | Hospitals where they had roster access during the searched period |
| Other clinician, several current-term hospitals, including locums | Every verified authorised hospital for this term | Hospitals where they had roster access during the searched period |
| Other clinician, no verified current-term hospital | No current hospital overview | Historical search remains possible where previous access can be verified |
| Account without At a glance enabled | No access | No access |

Use the Melbourne date to determine the present term. A former hospital claim
alone must not grant access in the present term. Future-term assignments must
not grant ordinary access before that term begins.

## Policy interpretation to build against

1. SMS and CMO share one all-site access policy. A CMO should not need an SMS
   roster entry or a non-clinical flag to obtain that access.
2. Multiple verified hospital assignments are legitimate access, not ambiguous
   evidence. Identity ambiguity or missing evidence still denies access to the
   affected hospital.
3. A locum assignment permits the roster access that comes with that assignment.
   Where the clinician already receives a hospital's whole term roster, At a
   glance may show that whole term. Where an explicit entitlement has narrower
   dates, respect those dates. Do not grant permanent access after a locum term.
4. Retrospective Working together may use former hospitals, but only for periods
   in which that clinician had roster access there. Being at DDH in Term 2 must
   not expose DDH's Term 3 roster after the clinician moves to MMC.
5. Historical search is not limited to shifts the viewer personally worked:
   it should let them identify a date, see who was on shift, and compare selected
   colleagues within their authorised hospital and period. This matches the
   intended convenience of an already accessible roster.
6. Historical access to a former hospital belongs to Working together. It must
   not make that hospital available in the current On shift, ED staff or By
   stream selectors.

The historical policy is the working interpretation of the requested case and
interaction lookup. Keep it explicit when reviewing the implementation.

## Existing implementation and gaps

- Current-site restriction was introduced in commit `5baa001e` on 25 August.
- `resolveFacilityOverviewAccess` in `functions/api/state.js` recognises SMS
  only and returns one `facilityKey` for other clinicians.
- Its compact reader uses exact current-term staff contributions and continuing
  SMS membership. Its legacy reader uses current-term working events.
- Server handlers already check the returned scope for Metadata, On shift,
  Staff, By stream, Working together and contact operations.
- The browser uses a fixed hospital label for restricted accounts. Snapshot
  scope keys also assume a single hospital.
- Working together currently uses the present access decision even when the
  request searches an older term. This does not express historical entitlement.
- `queryRosterInsights` and `queryRosterOverlapDoctors` separately check
  Insights permission and take hospital/doctor filters from the request. Review
  their callers and apply equivalent hospital/period restrictions where they
  provide trainee colleague data; changing only the At a glance tab would leave
  inconsistent policy across related views.
- Local role checks confirmed SMS gets `all`, whereas CMO, registrar, HMO and
  intern get `site` in both access implementations.
- The existing facility-access test reaches and passes its access assertions,
  then fails at an unrelated contact-menu source assertion. Record that baseline
  separately from new access tests.

## Implementation sequence

### 1. Establish reliable entitlement evidence

Trace how linked roster identities are verified and how term membership is
published. Check representative rotation and locum records before selecting the
data model. A name search, self-selected hospital filter, doctor selected for
comparison, or old account claim is not sufficient evidence by itself.

Represent an entitlement by verified account identity, hospital and applicable
period. Normal term membership can supply term boundaries; an explicitly
limited entitlement supplies narrower boundaries. If legitimate locums are
missing from roster membership, provide a Creator-managed hospital/period grant
with recorded provenance rather than broadening access based on guesswork.

Use effective grade evidence for the SMS/CMO policy, including supported grade
overrides. Review continuing SMS membership so obsolete evidence does not
silently override a later documented grade change. Decide whether continuing
CMO evidence is needed using the same rules as SMS, rather than requiring a
current working shift from a permanent CMO on leave.

Retain historical entitlement facts independently of roster-file replacement.
A corrected or superseded file must not accidentally erase a legitimate
previous rotation, while a corrected mistaken identity or assignment must
remove access. Missing historical evidence gives an unavailable/denied result,
not an automatic grant based on present access.

### 2. Replace the single-site access contract

Return a versioned scope containing an authorised hospital list, preferred
hospital, working-today status, current term, expiry and entitlement revision.
Keep all-site access distinct from an empty restricted hospital list.

Add a shared server helper that authorises hospital and date ranges for an
action. Ordinary overview actions use present-term entitlements. Retrospective
Working together resolves entitlements for each requested term or interval.
Both compact and any remaining legacy paths must produce the same policy.

Use indexed compact facts and bounded reads. Reuse the existing short-lived
access cache; key historical decisions by subject, period and entitlement
revision. Permission revocation must still take effect before cache retrieval.
Roster corrections, grade changes and entitlement changes must invalidate
affected decisions. Do not introduce broad event scans on login or refresh.

### 3. Enforce the scope across server routes

Update Metadata, On shift, Staff, By stream, Working together, contact reads and
contact corrections to use hospital-set membership instead of one-key equality.
Single-hospital and multi-hospital requests must be checked consistently. For
restricted users, “All my hospitals” means only the authorised set.

Historical requests spanning terms must be split into authorised hospital/date
segments before loading roster data. Filter returned rows, selectable colleague
names, metadata and coverage to those segments. Reject an explicitly requested
unauthorised hospital/period; never silently fall back to all hospitals.

Preserve Creator-entered account and doctor-profile behaviour: access comes from
the viewed account. Review the related Who/When and overlap routes so request
filters or arbitrary comparison identities cannot expand the viewer's scope.

Contact access tokens must remain limited to an authorised selected hospital;
historical search must not grant live contact permissions for a former site.

### 4. Update the interface and browser caches

- One authorised current hospital: retain the fixed label.
- Several: show a selector containing just those hospitals and “All my
  hospitals”. Prefer today's hospital where unambiguous; otherwise preserve a
  valid selection or use a stable default.
- SMS/CMO: retain “All EDs” and the available published hospital list.
- Working together: allow a historical date or range and derive hospital and
  colleague options from that period's entitlements. Support date-first lookup
  of who was on shift as well as comparison of remembered staff members.
- Show clear states for unavailable history and periods without authorised
  hospitals. A clinician with historical access but no present assignment must
  still be able to enter this historical workflow.
- Changing date ranges clears or revalidates hospital and colleague selections.
- Include subject, full hospital/period scope and entitlement revision in
  snapshot keys. Bump the snapshot schema as needed so old single-site caches
  cannot be reused under new rules. Clear stale scope on account switching,
  term change, expiry or revocation.

Keep the existing future-roster publication window. Historical access must not
accidentally expose unpublished upcoming rosters.

### 5. Validate locally, then prepare rollout

Use synthetic fixtures and representative retained roster structures. Required
behavioural cases:

1. Enabled SMS and CMO get identical all-site scope; disabled accounts get none.
2. Registrars, HMOs and interns access a verified current hospital and cannot
   retrieve another hospital by changing the request.
3. MMC rotation plus verified DDH locum permits both, while unrelated MCH is
   rejected. A past or future locum does not grant present-term access.
4. At term change, ordinary scope switches to the new assignment and former-site
   Working together remains available only in its authorised historical period.
5. A historical range crossing terms returns only the permitted hospital/date
   segments. Colleague names and coverage obey the same restriction as shifts.
6. A trainee with no current assignment can search verified previous rotations.
7. Date-first lookup and selected-colleague comparison work, including overnight
   shifts and Melbourne daylight-saving boundaries.
8. Unknown or conflicting identity evidence grants no additional hospital.
9. Creator account switching, permission revocation, grade changes and browser
   cache reuse cannot carry over another subject's or an obsolete scope.
10. Replacement/corrected rosters preserve legitimate history and remove mistaken
    entitlement; absent retained history is reported accurately.
11. Equivalent restrictions hold in related Insights/overlap and contact routes.
12. Access checks remain within the existing D1 request and account budgets;
    published roster retrieval continues to use the existing R2 approach.

Run the focused access, snapshot, contact-access, roster-insights, rollout and
quota checks appropriate to the final change, plus syntax checks. Separate the
known unrelated contact-menu test failure from access regressions.

Before any production rollout, inspect retained historical coverage and the
required migrations/backfill, validate with synthetic or controlled accounts,
and prepare the concrete diff and results for review. Use the existing rollout
and emergency pause controls; no deployment is part of today's planning task.

## Completion criteria

The feature is complete when CMO and SMS access match; other clinicians can
access all verified current-term hospitals including locums; and trainees can
search Working together at their former hospitals during authorised historical
periods. The same rules must hold in server responses, selectors, colleague
discovery and cached results, with focused behavioural checks passing.


## Implementation and verification — 3 October 2026

Implemented the requested access behaviour in the server and browser:

- SMS/CMO share all-site access, including effective grade overrides and CMO
  continuity from the most recent active roster membership when absent on leave.
  Present roster grades supersede obsolete continuing SMS evidence.
- Other clinicians receive a hospital set from current-term compact membership;
  verified locum sites are included. Live database authorisation uses the same
  compact policy even when a legacy payload reader is selected.
- Ordinary overview dates stay in the present term for restricted clinicians.
  Historical Working together and the related Who/When colleague readers use
  hospital/date segments from the viewer's membership in the searched terms.
- Working together fetches its historical hospital and staff options from the
  server. Global current staff pickers and pinned results cannot add names.
  “Who was on shift?” searches without a selected person.
- Historical-only accounts can open Working together. Missing published history
  is reported, and unavailable compact membership does not grant access.
- Browser snapshots use schema version 2 and hospital/date scope keys. Historical
  snapshots require a freshly verified period revision before reuse. Denied
  ordinary reads clear stale displayed data, and server responses refresh scope.
- Migration `0038_clinician_access_scope.sql` adds the identity/term lookup index
  and invalidates the small access cache after staff, grade, SMS membership or
  roster activation changes. It must be applied before production release.

Entitlement evidence is the existing linked roster identity plus active compact
term membership. Locums receive the same whole-term roster convenience as other
term members. There is no separate manual hospital-grant workflow or narrower
explicit entitlement storage in this implementation. A missing legitimate
membership should be corrected in the roster/identity data, rather than being
inferred from an arbitrary account hospital link.

Historical results depend on retained compact membership and published historical
rosters. No production data backfill, migration or deployment was performed.
Inspect retained coverage before release; do not promise history that has not
been retained or published. Existing identity linking is reused; this change is
not a redesign of account identity verification.

Validation:

- `npm run test:clinician-access`: passed synthetic API and browser-scope cases
  for CMO parity, multiple sites, locums, term changes, historical rotations,
  date-first lookup, denied direct requests, Creator-entered users, corrections,
  missing history, overnight/DST boundaries and the maximum 16 linked identities.
  Historical entitlement lookup is one bounded indexed query; a cold access
  lookup plus historical context remains within the 64-statement guard.
- `test:facility-access`, `test:facility-snapshots`,
  `test:facility-contact-access`, `test:cached-roster-insights`,
  `test:facility-rollout`, `test:d1-account-budget`, and
  `test:client-request-budget`: passed.
- Application/server syntax and diff whitespace checks: passed.
- `test:facility-materialization`: its overview read checks pass after updating
  fixture boundaries and scoped-All expectations for the new policy. The suite
  later fails in an unrelated roster import with `no such column: name`.
- `test:fixtures`: fails at the existing automated-correction source assertion.
- `test:facility-maintenance`: fails at the existing assertion that automatic
  launch is disabled in both production and preview.

All three broader failures were reproduced in an isolated archive of the
unchanged Git HEAD. They were not repaired as part of the clinician-access
change. Production configuration and existing untracked files were left intact.
