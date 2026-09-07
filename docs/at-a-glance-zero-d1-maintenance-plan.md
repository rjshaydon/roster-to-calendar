# At a glance zero-D1 maintenance gate plan

The user-visible capabilities temporarily affected by this plan and their
eventual restoration requirements are tracked in
[`feature-restoration-register.md`](./feature-restoration-register.md), entries
FR-01 through FR-04. Implementation must update those entries if its scope or
behaviour differs from this plan.

## Purpose

Temporarily make At a glance unavailable in a clear, non-technical way while
guaranteeing that attempts to open or operate it cannot consume D1 rows. This
is an additional containment layer around the rollout controls already active
in Production; it is not a substitute for the shared R2 implementation or the
account-wide quota gate.

Production currently runs commit `2af89c0`. Its legacy roster reads are
blocked, so At a glance cannot execute the former historical scans. However,
an At a glance request currently authenticates and resolves access in D1 before
the rollout router returns its unavailable response. This plan moves the
maintenance decision ahead of every database and object-store operation.

Implementation and deployment of this gate must not query D1, run a migration,
build a snapshot, invoke an application data endpoint, or change any rollout
allowlist.

## Product behaviour

Use normal language and keep At a glance discoverable so users understand that
it has not been removed.

The maintenance message is:

> At a glance is temporarily unavailable while we complete a reliability
> upgrade. Your roster and settings have not been removed. Please use My
> calendar for now.

For a non-clinical Director who has no personal calendar, omit the final
sentence and retain the normal account and sign-out controls. Do not expose D1,
Cloudflare, caches, feature flags or rollout terminology in the interface.

While maintenance is active:

- the desktop and mobile At a glance controls remain visible and identify the
  feature as temporarily unavailable;
- selecting either control opens one accessible maintenance panel and performs
  no network request;
- automatic On shift launch after login is suppressed;
- non-clinical Director login opens the maintenance panel instead of a blank or
  repeatedly loading workspace;
- tab changes, date changes, manual refresh, visibility changes and browser
  focus do not fetch At a glance data;
- the 60-second contact refresh timer is stopped and cannot reschedule itself;
- no cached roster or contact allocation is displayed; and
- the ordinary personal calendar and unrelated administration remain
  available.

Existing IndexedDB snapshots are left intact for possible reuse after the
feature returns, but the maintenance UI must not read or render them. They are
not a safe general fallback: they are device-specific, may be stale and may
contain contact information. The existing access-expiry validation remains in
place for eventual reuse.

## Phase 1 — dedicated fail-closed capability

Add one environment capability:

`FACILITY_OVERVIEW_MAINTENANCE_MODE`

Semantics:

- missing, empty, malformed or `true` means maintenance is active;
- only an explicit supported false value disables maintenance;
- Production and Preview configuration declare it as `true`;
- isolated local development declares it as `false` so the feature can still
  be exercised against local SQLite and local object storage; and
- changing it later is a separately reviewed control-plane operation, not an
  automatic consequence of enabling a builder or shared reader.

Implement the parser in the shared facility-rollout helper and test its truth
table. Do not duplicate string parsing in the request handler or client.

Include `facilityOverviewMaintenance: true|false` in the successful login
envelope. The client starts fail-closed before login and changes state only from
an explicit boolean returned by the current server. A response from an older
server that omits the field therefore keeps the UI unavailable.

## Phase 2 — pre-authentication server gate

Define one immutable set of maintenance-gated action names close to the start
of `functions/api/state.js`. Immediately after safely parsing the request body
and action, but before email/password validation, `hasCalendarDb`, expired
invite cleanup, authentication, target-account lookup or access resolution,
return the maintenance response when the capability is active.

Gate these operational reads:

- `queryFacilityOverviewMetadata`;
- `queryFacilityOverviewByStream`;
- `queryFacilityOverviewOnShift`;
- `queryFacilityOverviewContactList`;
- `queryFacilityOverviewStaff`; and
- `queryFacilityOverviewWorkingTogether`.

Gate these mutations as well, so a stale client cannot alter facility data or
force its validation reads during maintenance:

- `setContactAllocationResolution`;
- `setFacilityStaffDesignation`;
- `clearFacilityStaffDesignation`;
- `setFacilityStaffSeniorityOverride`; and
- `setFacilityStaffSeniorityOverrides`.

Do not gate `login`, personal-calendar actions, entitlement administration
(`setUserFacilityOverviewEnabled`) or the Creator's doctor-profile entitlement
lookup. Those are not At a glance data operations and must not be conflated
with this temporary gate.

The server response should be a stable JSON shape such as:

```json
{
  "ok": false,
  "unavailable": true,
  "maintenance": true,
  "retryable": false,
  "error": "At a glance is temporarily unavailable while we complete a reliability upgrade. Your roster and settings have not been removed."
}
```

Use HTTP 503 with `Cache-Control: no-store` and no `Retry-After` header. The
client must treat `retryable: false` as terminal rather than starting a retry
loop. The response must be constructed without reading D1 or R2.

This server gate is authoritative. It protects Production from stale service
workers, old JavaScript, direct POSTs, multiple tabs and clients that have not
yet loaded the new UI.

## Phase 3 — client maintenance state

Add a single client-side maintenance state, initially `true`. Update it from
the explicit login-envelope boolean and reset it to `true` on logout, failed
login, account-context reset or an unrecognised response.

Centralise the UI decision rather than placing guards in individual render
fragments:

1. `openFacilityOverview` checks maintenance before metadata or tab loading.
2. Automatic Director and clinical On shift launch paths check the same state.
3. All At a glance loaders and contact refresh scheduling fail closed as a
   defence in depth.
4. Entering maintenance aborts any in-flight At a glance request, stops the
   contact timer, clears in-memory contact information and renders the friendly
   panel.
5. The panel provides a My calendar action only when that account has a personal
   calendar.

Do not read the browser snapshot before this guard. Do not silently show stale
data beneath a maintenance banner. Do not add automatic probes to discover
whether maintenance has ended; the state is refreshed on the next normal login
or explicit application reload.

## Phase 4 — focused verification

### Server zero-D1 test

Exercise every gated action against instrumented bindings whose `prepare`,
`batch`, `exec`, `get`, `put`, `head`, `list` and `delete` methods count and
fail if called. For maintenance values that are missing, empty, malformed or
true, assert for every action:

- HTTP 503;
- the stable maintenance payload;
- zero D1 calls;
- zero D1 rows read and written;
- zero R2 calls; and
- no scheduled `waitUntil` work.

The test must prove that the gate precedes expired-invite cleanup as well as
authentication. Use direct handler tests; do not infer safety from source-text
ordering alone.

For an explicit false value, assert that the request passes this maintenance
gate and reaches the existing authentication/rollout controls. It need not make
a real D1 query in this test.

### Client request-budget test

With maintenance active and network calls counted, cover:

- desktop and mobile button selection;
- clinical login that would normally auto-open On shift;
- non-clinical Director login;
- each At a glance tab and date control;
- manual refresh;
- repeated hide/show and focus events;
- a simulated hour containing all 60 contact-refresh intervals; and
- 50 independent visible-page instances.

The acceptance result is zero At a glance requests, zero contact polls and no
timer growth. The maintenance panel must remain stable.

### Cache and privacy test

Seed a valid-looking local roster and contact snapshot. Confirm it is neither
read nor rendered while maintenance is active and that no phone number appears
in the DOM. Confirm expired cache protection still works when maintenance is
explicitly disabled.

### Regression checks

Run the focused facility-rollout, D1 quota guard, client request-budget,
facility access and syntax/fixture tests. Confirm personal calendar login and
opening remain functional locally. Do not broaden this into a full remote test
or run a Production application request.

## Phase 5 — commit and safe deployment

1. Review the diff for scope, secrets and accidental user-file inclusion.
2. Commit the implementation on `main` only after all focused gates pass.
3. Push `main`, allowing the existing Pages integration to deploy Production.
4. Verify through the Cloudflare control plane that the active Production
   deployment uses the new commit.
5. Download and inspect the effective Pages configuration without contacting
   D1. Verify all existing emergency, legacy, reader, builder, bootstrap,
   automation and maintenance switches remain closed. The new maintenance mode
   must either read back explicitly as true or be absent and therefore fail
   closed.
6. Do not test the deployed button or call `/api/state` before the reset. The
   direct local handler tests are the evidence for zero-D1 behaviour.

Rollback is the previous known Production deployment, selected through the
Pages control plane. Because the previous deployment already blocks legacy
reads, rollback cannot restore At a glance but it retains the earlier
containment posture. Do not roll back by reopening legacy reads.

## Interaction with reset-day work

This maintenance gate does not change the passive Analytics timetable or grant
permission to run bootstrap inspection. After the reset, ordinary traffic is
observed while At a glance remains in maintenance. A flat, attributable
account-wide baseline is still required before the separately approved
single-file inspection.

Removing the maintenance screen is not part of this plan. It requires:

- a GO result from the account-wide budget checker;
- explicit approval for the next canary step;
- shared artifacts and access controls proven ready for the selected cohort;
- legacy reads still permanently paused; and
- a reviewed configuration change setting the dedicated maintenance flag to
  false.

## Completion gates

The maintenance patch is ready to deploy only when all of the following are
true:

- every listed stale-client action returns before all D1/R2 access;
- every new-client entry path performs zero At a glance requests;
- automatic launch and all contact polling are suppressed;
- no cached roster or contact information is shown;
- the message is accessible and understandable without technical knowledge;
- personal calendars remain locally functional;
- all focused safety tests pass; and
- Production rollout, bootstrap and automation controls remain closed.
