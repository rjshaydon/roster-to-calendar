# Creator login/bootstrap D1 remediation plan

## Purpose

This plan addresses the 9 September 2026 production test in which one Creator
login, followed by one attempt to open the maintenance-disabled At a glance
workspace and a return to Calendar, coincided with a 1,623,542-row-read spike.
It is a containment plan for the current login/bootstrap path, not permission
to restore At a glance, run migrations, bootstrap compact data or resume any
automation.

No production application or D1 request is required to implement or test this
plan. The existing maintenance controls and feature entitlements remain as
they are.

## What the evidence establishes

- Before the controlled interaction, ordinary traffic was stable at roughly
  4,000–6,000 rows read per hour, with no large post-cleanup bucket.
- The 14:50–14:55 AEST bucket recorded 1,623,542 reads, 6 writes and 149 read
  queries. This aligns with the Creator login test.
- The At a glance maintenance interface correctly refused access. The tab then
  remained open and idle; the next settled sample added only 109 reads and no
  writes. The expensive activity was a one-off startup/navigation burst, not a
  continuously repeating idle poll.
- Newly visible fingerprints included two complete active
  `roster_file_doctors` listings, repeated status/revision counts, personal
  calendar reads and account lookups. These attributed only a small part of the
  spike. Approximately 1.61 million reads from the test remain unattributed by
  Cloudflare's query-fingerprint feed.
- Therefore the exact million-row statement is not yet proven. The safe
  conclusion is that the Creator startup request fan-out still reaches more
  work than authentication and cached calendar display require.

The retained-deployment cleanup remains valid evidence for a separate issue:
all pre-containment Production deployment hashes were removed. The 14:50 test
must not be attributed to those deleted deployments because the raw Analytics
delta did not add the earlier runtime-schema fingerprint.

## Current source paths requiring isolation

After a fast login response, the browser currently schedules or can schedule:

1. `loadAccountContext`;
2. calendar hydration and `bootstrapImports`;
3. `listUsers` for the Creator directory;
4. `calendarStoreStatus`;
5. doctor-switcher synchronization using remaining roster files; and
6. calendar/snapshot revalidation and possible deferred work; and
7. automatic colleague-insight warm-up after rendering a cached calendar.

The seventh path is particularly important: merely rendering the Creator's
saved calendar scheduled an overlap-doctor request with the historical
fallback allowed, even though the Creator had not opened “Who is working with
me?” or “When am I working with…?”. This is a plausible contributor to the
unattributed spike, but Analytics still does not prove it was the sole cause.

Some server controls already return a paused result before broad work. That is
not sufficient: a disabled feature must also not be requested automatically,
and one paused response must not trigger another discovery or synchronization
path in the browser.

## Safety invariants

1. Creator login loads only authentication, the Creator's compact account
   context and an already-published or locally cached personal calendar.
2. Creator login never lists all users, files, roster doctors or historical
   events as a side effect.
3. While At a glance is in maintenance, opening it and returning to Calendar
   performs zero D1 operations.
4. Admin data loads only after the Creator deliberately opens its specific
   surface. A paused surface returns its friendly maintenance state before its
   own data query.
5. An unchanged cached workspace performs no D1 refresh after authentication;
   an idle or hidden tab schedules no D1 work.
6. No login response schedules D1/R2 work with `waitUntil` unless a separately
   enabled, revision-driven builder has an explicit changed dependency.
7. Missing flags, stale clients and legacy request shapes fail closed before
   feature data is read.
8. No remediation may erase At a glance entitlements, roster data, cached
   snapshots or Doctor Names work.

## Stage 1 — deterministic local request trace

Build one focused test harness around the existing local Worker/browser test
infrastructure. Do not create a second application simulation.

Run these scenarios from a clean browser state against the deterministic
synthetic database already used for D1 scale tests:

- explicit Creator login;
- reload with a persisted Creator session;
- Calendar to disabled At a glance and back;
- 15 minutes of visible idle time using fake timers;
- hidden-tab transition and reopening; and
- deliberate opening of each Creator Admin surface, one at a time.

For every scenario record:

- ordered HTTP requests and their `action` values;
- every prepared D1 statement and binding class;
- whether the statement uses an index according to `EXPLAIN QUERY PLAN`;
- database rows examined as an estimate distinct from rows returned;
- mutation statements; and
- work scheduled after the response.

The harness must fail on any unlabelled statement or scheduled task. Existing
representative roster fixtures remain correctness tests; the synthetic history
is used only to reveal cost growth.

## Stage 2 — make Creator login minimal

1. Define an explicit minimal-login response contract containing only the
   authenticated compact account/profile data, entitlements, maintenance
   states, current calendar identity and exact snapshot/cache metadata.
2. Do not automatically call `listUsers`, `calendarStoreStatus`, roster doctor
   discovery, file restoration or doctor-switcher synchronization during
   Creator hydration.
3. Do not run `bootstrapImports` merely because the authenticated account is
   the Creator. Existing cached/published calendar data may be displayed; an
   unavailable current snapshot gets a clear manual-refresh state.
4. Keep `loadAccountContext` only if its local trace proves it is bounded and
   contains no data already returned by minimal login. Otherwise fold the
   necessary compact fields into the first response or defer it to a deliberate
   surface action.
5. Cancel queued hydration work when navigation changes, the tab becomes
   hidden or maintenance is reported. A response from an obsolete transition
   must not start follow-on requests.

## Stage 3 — gate Creator-only surfaces independently

1. Add or retain separate default-off controls for user directory enrichment,
   roster-file status, doctor discovery/switcher enrichment and snapshot
   validation. Check the relevant control before querying feature data.
2. Opening Admin → Accounts may request a bounded compact user page only; it
   must not calculate identity, seniority, facility access or roster history.
3. Opening Admin → Files while paused returns the existing maintenance wording
   without loading file, doctor, event or revision collections.
4. The disabled At a glance action must return its maintenance response before
   authentication-dependent facility work and the client must treat that as a
   terminal result.
5. Personal calendar refresh remains separate from every Creator/Admin path
   and is bounded by doctor identity and date range.

Every paused or malformed action gets a zero-feature-D1 regression test. The
small indexed account lookup needed to authenticate an otherwise authorised
request is reported separately and must not be mislabelled as feature cost.

## Stage 4 — local acceptance gates

The change is ready for review only when all of the following pass:

- Creator login issues no query against `roster_events`,
  `roster_file_doctors`, complete roster-file collections or complete account
  child collections.
- Creator login does not request `listUsers`, `calendarStoreStatus`,
  `listRosterDoctors` or any bootstrap/maintenance endpoint.
- Opening disabled At a glance and returning to Calendar adds zero D1
  statements after the initial authentication has completed.
- Fifteen visible idle minutes and a hidden-tab cycle add zero D1 statements.
- Repeated unchanged authenticated cache refresh adds zero D1 statements.
- Initial authentication/account context uses only indexed exact-key reads;
  no statement examines more than 500 rows and the complete minimal login
  examines no more than 2,000 rows in the scale fixture.
- A deliberately refreshed personal calendar is tested separately: it must be
  doctor- and date-bounded, indexed, examine no more than 10,000 rows, and its
  cost must not grow with unrelated roster history.
- No-op requests write zero rows and schedule no background work.
- Existing login, calendar, entitlement and maintenance-message correctness
  tests pass.

These are future acceptance gates. The 9 September Production result is a
failed baseline and is not expected to meet them.

## Stage 5 — release without spending D1 quota

1. Review the focused diff and test evidence. Keep migrations `0026` onward,
   compact bootstrap and all restoration work out of this release.
2. Commit the containment separately. A push/deployment may occur before the
   next reset only after review, because Pages deployment and configuration
   read-back do not require an application D1 request.
3. Read back the deployed commit and effective non-secret maintenance flags
   through the control plane. Do not open the application to verify it on the
   current quota day.
4. After the next UTC reset, collect the planned passive account-wide samples
   first. Do not log in as Creator until the passive period passes.
5. Perform one controlled Creator login with browser request capture enabled.
   Do not open Admin, At a glance or invoke colleague tools in the same test.
6. Wait for Analytics settlement, reconcile the delta and stop unless the
   result is small and attributable.
7. Only on a separate approved gate test disabled At a glance once and return
   to Calendar. Reconcile that delta independently; its target is zero D1 rows
   after authentication.

No result from this remediation authorises migration `0026` onward. The
quota-safe rollout resumes only after the Creator-login canary and the existing
passive baseline both pass.

## Stop conditions

Stop immediately, keep the tab closed and perform no further application D1
request if any of the following occurs:

- more than 10,000 rows examined by an ordinary request;
- any broad roster/account query during login;
- any D1 activity caused by disabled At a glance navigation or idle time;
- unexpected writes or post-response work;
- incomplete account/timeline reconciliation; or
- another unattributed burst.

Investigate a failed gate from the saved local trace and settled Analytics,
not by repeating the Production request.

## Local implementation checkpoint — 9 September 2026

Implemented without accessing Cloudflare or an application D1 database:

- added explicit fail-closed `CREATOR_STARTUP_HYDRATION_ENABLED` and
  `CREATOR_DIRECTORY_ENABLED` controls for Production, Preview and local
  development;
- forced contained Creator logins onto the minimal fast response even when a
  stale client asks for the former full response;
- stopped explicit and persisted Creator login before deferred account
  context, calendar hydration, import bootstrap, user-directory loading,
  roster status, doctor-switcher synchronization and snapshot revalidation;
- prevented cached Creator calendar rendering from scheduling the automatic
  colleague-insight warm-up;
- removed the pre-authentication flush of stale account state;
- gated paused user-directory, roster-status and roster-doctor actions before
  authentication, producing zero D1/R2 operations and no scheduled work; and
- retained a clear maintenance message without deleting accounts,
  entitlements, roster data or cached snapshots.

Focused evidence:

- the request-containment test reports one initial Creator login request and
  zero automatic follow-up requests;
- disabled directory, status and roster-doctor stale-client requests each use
  zero D1, zero R2 and schedule no work;
- a direct in-memory state-handler trace forces a stale full Creator login to
  the fast response, schedules no `waitUntil` work and contains no
  `roster_events`, `roster_file_doctors`, roster-file collection or account-
  directory statement; and
- the 109,200-event/10,000-account/20,000-claim SQLite fixture proves the
  remaining account and snapshot-registry lookups use exact primary-key
  indexes and do not grow with unrelated roster or account history.

This is local evidence only. The release and production canaries in Stage 5
remain outstanding.
