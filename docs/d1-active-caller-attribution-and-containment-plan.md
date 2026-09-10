# Active D1 caller attribution and containment plan

## Purpose

Identify the route responsible for the unexplained Production D1 bursts while
preserving login and My calendar. Remove dormant callers that have no present
user value, make every paused automation path incapable of reaching D1, and
capture enough non-D1 evidence to attribute any future burst.

This plan follows the 10 September Pages cleanup. The verified deployment
inventory at its start is one Production deployment (`bae86cea…`, source
`8ef0215`) and no Preview deployments. Preview reported no D1 use during the
investigated incidents, so its removal is containment rather than retrospective
causation. The 992,117-read 14:30 AEST burst remains unattributed.

## Authority and constraints

- Do not query D1, run migrations, bootstrap data, restore At a glance or
  re-enable roster/contact automation while implementing this plan.
- Preserve ordinary login and My calendar.
- Preserve source code, Git history, existing D1/R2 data and configuration
  needed to restore desirable features later.
- Every additional pause or removal must update
  `feature-restoration-register.md` in the same commit.
- A deployment is not complete until effective variables are read back through
  the control plane and obsolete Production/Preview deployments are deleted.
- Missing or malformed safety settings fail closed.

## Stage 1 — finish the live-caller inventory

1. Record the sole Pages Production deployment and both custom/public domains.
2. Record the separately deployed `roster-queue-watchdog`, its active version
   and its `*/15 * * * *` cron. Confirm it has no D1 binding.
3. Confirm GitHub roster workflows have no schedule and record their last run.
4. Inventory each Power Automate or other external flow that can call roster or
   contact endpoints: owner, trigger type, schedule, target URL and source ID.
   Do not record credentials.
5. Account for every Cloudflare Worker, queue, cron and Pages project in the
   account. Unrelated projects remain named but are not modified.

Gate 1: every known caller is either removed, explicitly paused at its source,
or identified as necessary ordinary user traffic. An unknown caller is `STOP`.

## Stage 2 — eliminate dormant scheduled and external callers

1. Delete the deployed `roster-queue-watchdog` Worker, which also removes its
   cron. Retain its source and configuration in Git and mark FR-10 as removed
   from Cloudflare but restorable last.
2. Pause the roster/contact Power Automate flows at their source. This requires
   the flow owner's explicit action; record the resulting state and time.
3. Leave GitHub manual workflows present but do not dispatch them. Verify no
   scheduled trigger exists.
4. Do not remove user entitlements or delete roster/contact data.

Gate 2: Cloudflare lists no roster watchdog or roster cron, and each external
flow is visibly disabled. If a flow cannot be inspected or paused, proceed only
with its target endpoint blocked before D1 and retain `STOP` for attribution.

## Stage 3 — make paused endpoints zero-D1 by construction

Add a narrow Pages middleware/shared preflight for `/api/automation/*`. While
the relevant capability is paused it must return before authentication, body
parsing, D1/R2 access, outbound dispatch or `waitUntil`. Preserve separate
controls for roster ingestion, contact ingestion, queue dispatch, bootstrap
inspection and bootstrap execution; enabling one must not enable another.

Retain the existing pre-authentication gates for maintenance-disabled At a
glance and paused Creator startup actions. Add focused tests using bindings that
throw on any access. Direct requests, stale clients, valid/invalid credentials,
missing variables and malformed variables must all demonstrate zero binding
access when their capability is closed.

Gate 3: all closed automation and maintenance routes pass zero-D1/R2/outbound-
request tests. Ordinary login and representative My calendar correctness tests
still pass locally.

## Stage 4 — add non-D1 route attribution

1. Bind a dedicated Workers Analytics Engine dataset to Pages and write one
   data point per API invocation during the investigation. Pages real-time
   Functions logs remain available for live checks, but are not persistent.
   Attribution must not use D1 or R2.
2. Emit one structured completion record per API invocation containing only:
   deployment revision, request ID, route, HTTP method, `/api/state` action,
   response status, duration, containment result, and caller class where it can
   be derived without identity data. Never log passwords, tokens, email
   addresses, roster content, telephone numbers or request bodies.
3. Add request-local D1 statement accounting around the Pages bindings. Record
   statement count, operation class and Cloudflare `rows_read`/`rows_written`
   metadata where the API returns it. Hash or classify SQL; do not log bound
   values. A route without complete metering must be marked incomplete rather
   than reported as zero.
4. Give each action a declared statement ceiling. Reject additional statements
   before execution. This supplements—but cannot replace—indexed query-plan and
   rows-examined tests because one bad statement can still scan many rows.
5. Provide saved/queryable views grouped by route, action, response status and
   five-minute interval so Pages invocations can be reconciled with account-wide
   D1 Analytics.

Gate 4: local tests prove sensitive values cannot appear in logs, statement
ceilings fail closed, logging failure does not trigger D1 work, and the
account-budget checker remains independent of application databases.

## Stage 5 — controlled Production release

1. Review the diff and focused test evidence. Commit code, tests, configuration
   and restoration-register changes together.
2. Deploy once using the explicit Wrangler configuration; do not rely on a Git
   build to apply variables.
3. Read back effective non-secret Production configuration without invoking an
   application route.
4. Confirm the Analytics Engine binding is present through the control plane.
   Confirm records during the later approved canary; do not use a D1-backed
   endpoint merely as a logging probe.
5. Delete every superseded Production deployment and any automatically created
   Preview deployment. Final inventory must again be one Production and zero
   Preview.

Any configuration mismatch, duplicate retained deployment or missing logs is
`STOP` and is resolved through non-D1 control-plane rollback.

## Stage 6 — fresh-day attribution sequence

After a UTC reset, leave At a glance and all automation closed:

1. Observe at least two passive hours using settled account-wide GraphQL D1
   Analytics and persisted invocation logs.
2. Reconcile five-minute D1 totals with Production route/action counts and
   request-local metadata. Do not infer causation from timestamps alone.
3. If passive use is safe, perform one ordinary login/My calendar canary, wait
   for analytics settlement, then reconcile again.
4. Do not perform a Creator, At a glance, Admin, colleague-query, ingestion or
   bootstrap canary during this sequence.

Success requires no unexplained spike, no action above its statement/row gate,
and reconciliation within the documented Analytics tolerance. A burst with a
recorded route identifies the remediation target. A D1 burst with no matching
invocation becomes evidence for Cloudflare; stop Production experiments and
submit the timestamps, aggregates and sanitized log export.

## Restoration sequence

Once the source is fixed and a fresh-day baseline passes, restore capabilities
only in the order already recorded in `feature-restoration-register.md`.
Power Automate sources return one at a time. The watchdog is restored last,
with coalescing and a hard request budget. Persistent privacy-safe invocation
logging and the one-current-deployment policy remain permanent safeguards.

## Implementation checkpoint — 10 September 2026

- Pages inventory began at one Production and zero Preview deployments.
- The `roster-queue-watchdog` Worker and its cron were deleted. Control-plane
  verification returned `10007`, confirming that the Worker no longer exists.
- A Pages middleware now blocks closed automation classes before touching D1,
  R2, authentication, request bodies or outbound services.
- A dedicated Analytics Engine binding, privacy-safe structured completion
  records and request-local D1 statement/row metering are implemented. This
  uses no D1/R2 rows or objects. Pages real-time logs retain the same records.
- Login has a 32-statement ceiling; common account/calendar actions have
  explicit ceilings and all other API actions have a 64-statement ceiling.
- Focused quota, Creator containment, facility maintenance/rollout, client
  budget and local-isolation tests pass. The Pages Functions bundle compiles.
- A local Pages integration check returned a paused contact request with zero
  D1 statements and an invalid login with one row read. Logged records contained
  no supplied email, password, token, source ID, telephone number or body.
- Source-side Power Automate inspection/pause remains outstanding because its
  configuration is outside this repository and Cloudflare account.
