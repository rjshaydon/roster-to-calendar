# Feature restoration register

## Purpose and maintenance rule

This is the authoritative register of application features and operational
capabilities paused, reduced, replaced or parked during the September 2026 D1
incidents. The safety work is temporary containment, not a decision that the
affected product features are unwanted.

Every future change that disables, hides, degrades, throttles or removes a
feature must update this register in the same commit. Every restoration must
record its evidence, approval, deployment and post-deployment observation here.
No entry may be removed after restoration; change its state to **Restored** so
the history remains auditable.

Current verified Production safety code: `8ef0215` on 10 September 2026,
deployment `bae86cea-390f-44e4-9d9e-d34a95c7854c`. The zero-D1 At a glance
maintenance gate and ordinary-login containment are active. Effective
Production configuration was read back after that deployment with
all facility readers/builders, roster and contact automation, roster status,
bootstrap and advanced-maintenance controls closed.

The safety commits did not intentionally delete Production roster files,
roster events, account data, facility snapshots or contact data. Because D1 is
currently quota-exhausted, this statement describes the behaviour of the
changes and recorded operations; database contents will be verified only under
the reset-day runbook.

## Status meanings

- **Paused** — the feature exists but its Production path is blocked.
- **Reduced** — the feature remains available with retries, polling or
  diagnostics deliberately limited.
- **Parked** — developed work is preserved outside Production and must be
  reintegrated deliberately.
- **Replaced permanently** — the old mechanism must not return; its useful
  outcome is provided by a safer mechanism.
- **Still live** — listed to prevent accidental removal or confusion with a
  paused feature.
- **Containment required** — an unsafe path has been identified and must be
  locally remediated before further rollout; this does not claim it caused a
  specific unattributed incident.

## Configuration control index

This index accounts for every Production/Preview safety control currently
declared in `wrangler.toml`, plus the separately deployed watchdog control and
the planned maintenance flag. An empty allowlist means no source is enabled.

| Control | Current safe value | Registered under |
| --- | --- | --- |
| `FACILITY_OVERVIEW_MAINTENANCE_MODE` | `true`; missing or malformed also fails closed | FR-01–FR-04 |
| `FACILITY_SHARED_ROLLOUT_ACTIVE` | `false` | FR-01 |
| `FACILITY_SHARED_EMERGENCY_PAUSED` | `true` | FR-01 |
| `FACILITY_LEGACY_READS_PAUSED` | `true` permanently | FR-01, FR-17 |
| `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST` | empty | FR-01, FR-12 |
| `FACILITY_SHARED_READER_SOURCE_ALLOWLIST` | empty | FR-01 |
| `FACILITY_SHARED_READER_COHORT` | empty | FR-01 |
| `FACILITY_ACCESS_MATERIALIZATION_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_METADATA_BUILD_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_DAYS_BUILD_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_CONTACTS_BUILD_ENABLED` | `false` | FR-04 |
| `FACILITY_SHARED_METADATA_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_DAYS_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_CONTACTS_ENABLED` | `false` | FR-04 |
| `ROSTER_AUTOMATION_WRITES_ENABLED` | `false` | FR-06, FR-07 |
| `ROSTER_STATUS_SUMMARY_ENABLED` | `false` | FR-05 |
| `ROSTER_AUTOMATION_SOURCE_ALLOWLIST` | empty | FR-07 |
| `ROSTER_AUTOMATION_QUEUE_ENABLED` | `false` | FR-07 |
| `ROSTER_ADVANCED_MAINTENANCE_ENABLED` | `false` | FR-06, FR-11, FR-12 |
| `FACILITY_BOOTSTRAP_INSPECTION_ENABLED` | `false` | FR-12 |
| `FACILITY_BOOTSTRAP_EXECUTION_ENABLED` | `false` | FR-12 |
| `FACILITY_BOOTSTRAP_FILE_ALLOWLIST` | empty | FR-12 |
| `CONTACT_AUTOMATION_WRITES_ENABLED` | `false` | FR-08 |
| `CONTACT_AUTOMATION_SOURCE_ALLOWLIST` | empty | FR-08 |
| `IDENTITY_DISCOVERY_ENABLED` | `false`; missing or malformed also fails closed | FR-21, FR-24 |
| `ACCOUNT_SNAPSHOT_BUILD_ENABLED` | `false`; missing or malformed also fails closed | FR-23 |
| Watchdog `ROSTER_AUTOMATION_ENABLED` | `false` | FR-10 |

## User-facing restoration register

### FR-01 — At a glance workspace

- **State:** Paused.
- **Includes:** On shift, ED Staff, By stream, Working together, multi-ED
  Director overview, facility/date selection and At a glance navigation on
  desktop and mobile. It also includes the compact staff-membership,
  file/hospital-coverage, additional-coverage and daily-presence facts needed
  to build those views without repeatedly deriving them from event history.
- **User effect:** Requests currently receive a controlled unavailable response
  instead of roster data.
- **Controls:** `FACILITY_SHARED_ROLLOUT_ACTIVE=false`,
  `FACILITY_SHARED_EMERGENCY_PAUSED=true`,
  `FACILITY_LEGACY_READS_PAUSED=true`, empty reader/build allowlists and all
  shared reader flags false.
- **Data/code preservation:** The interface and shared-object implementation
  remain in `main`. Legacy broad SQL is retained only as inaccessible code
  during migration; it is not the restoration path.
- **Restoration outcome:** Restore the complete user experience through compact
  ED/date, ED/term and range artifacts shared across users. A missing artifact
  must say unavailable/preparing and must never fall back to historical SQL.
- **Required evidence:** Completed compact bootstrap for the selected ED/term;
  shared access and cache tests; zero roster-event D1 reads on warm requests;
  account-wide budget GO; one Creator/one-ED canary; settled post-canary
  Analytics; then gradual ED and user expansion.
- **Permanent constraint:** `FACILITY_LEGACY_READS_PAUSED` remains true. The
  feature returns; the unsafe implementation does not.

### FR-02 — At a glance cached display during maintenance

- **State:** Paused in Production by the fail-closed maintenance gate.
- **Includes:** browser-local saved On shift, Staff, metadata and range views.
- **User effect:** No cached roster or telephone data is
  shown while global maintenance is active. Cached records remain stored and
  are not deleted.
- **Reason:** Browser snapshots are device-specific, may be stale and may
  contain contact information. They cannot be a general maintenance fallback
  without a current authorisation check.
- **Restoration outcome:** Resume the existing cache only after maintenance is
  explicitly disabled. Preserve the 15-minute maximum access-revocation delay,
  immediate removal of expired contact information and a clear last-updated
  warning on refresh failure.
- **Plan:** `at-a-glance-zero-d1-maintenance-plan.md`.

### FR-03 — Automatic On shift launch and contact refresh

- **State:** Paused in Production. Automatic launch and 60-second contact
  polling are explicitly suppressed while maintenance is active.
- **User effect:** Clinical users no longer receive the intended automatic live
  On shift workspace while At a glance is paused.
- **Restoration outcome:** Reinstate automatic launch for eligible clinical
  users and refresh contacts every 60 seconds while visible, immediately on
  reopening/manual refresh, and never from hidden tabs.
- **Required evidence:** Fifty visible pages for 12 hours produce 36,000
  unchanged refreshes with zero D1 rows after authorisation; one in-flight
  request per page; no stacked timers; R2/session costs within their separate
  budgets; permission expiry remains fail-closed.

### FR-04 — Live contact allocations and Creator corrections

- **State:** Paused as part of At a glance. Shared contact reads and publication
  are also disabled.
- **Includes:** MMC/DDH contact overlays, automatic roster/contact matching,
  unresolved-number review and temporary Creator corrections.
- **Controls:** `FACILITY_SHARED_CONTACTS_ENABLED=false`,
  `FACILITY_SHARED_CONTACTS_BUILD_ENABLED=false`, plus the facility emergency
  pause and empty allowlists.
- **Data/code preservation:** Contact matching, overnight carry, expiry and
  optimistic correction logic remain in the repository. Pausing does not
  intentionally remove saved extracts or corrections.
- **Restoration outcome:** Publish small independent contact overlays keyed by
  facility/date/extract revision; contact changes must not rebuild base roster
  artifacts.
- **Required evidence:** unchanged refreshes read zero D1 rows; unchanged
  extracts write zero D1/R2 objects; new extracts invalidate only their contact
  overlay; stale corrections cannot attach to a different extract; expiry and
  revocation tests pass.

### FR-05 — Admin → Files status

- **State:** Paused.
- **Includes:** Auto-sync source status, manual files for the current/next term,
  and earlier manual-file status.
- **User effect:** The interface explains that status is temporarily
  unavailable and that existing sources/files have not been removed.
- **Control:** `ROSTER_STATUS_SUMMARY_ENABLED=false`.
- **Data/code preservation:** The former historical counting path is not used.
  A bounded summary implementation exists in code; its deployment dependencies
  and migration state must be verified after reset.
- **Restoration outcome:** Restore the same useful status information from
  compact per-file/source summaries maintained during ingestion.
- **Required evidence:** Indexed reads scale with the number of retained files,
  not event history; requests coalesce; unchanged responses write zero rows;
  hidden tabs do not retry; account-wide budget GO and a read-only Creator
  canary pass.

### FR-06 — Manual roster import and file management

- **State:** Paused.
- **Includes:** upload/retain a source file, save its parsed calendar, remove an
  import, reconcile retained files, reset/reparse a derived file and replace
  the active file set. Reading an already retained raw file is not intentionally
  disabled.
- **Control:** `ROSTER_AUTOMATION_WRITES_ENABLED=false`; destructive/rebuild
  actions additionally require `ROSTER_ADVANCED_MAINTENANCE_ENABLED=true`.
- **Affected actions:** `uploadRawRosterFile`, `saveDerivedCalendarFile`,
  `removeRosterImports`, roster-removing account saves,
  `syncRosterRepository`, `resetDerivedCalendarFile` and
  `replaceActiveRosterFiles`.
- **Data/code preservation:** Existing retained files and derived data are not
  removed by the pause.
- **Restoration outcome:** Restore ordinary Creator imports and file management
  with genuinely incremental ingestion. An identical file/provider revision
  must be a no-op: no event, membership, daily-presence, summary or cache
  rewrite. A correction updates only affected facts and artifacts and preserves
  unrelated file contributions.
- **Required evidence:** representative Excel and FindMyShift correctness
  fixtures; swaps, sickness, removals, overlaps and term boundaries; exact
  per-import read/write ceilings; transactional failure tests; one explicit
  source/file canary; settled Analytics before expanding.

### FR-07 — Automatic roster synchronisation

- **State:** Paused.
- **Includes:** FindMyShift change checking, generic automated upload, VHH
  extraction, derived processing, queue dispatch and pending work processing.
- **Controls:** `ROSTER_AUTOMATION_WRITES_ENABLED=false`, empty
  `ROSTER_AUTOMATION_SOURCE_ALLOWLIST`, and
  `ROSTER_AUTOMATION_QUEUE_ENABLED=false`.
- **Data/code preservation:** Configured sources have not intentionally been
  removed. The UI's unavailable status is not evidence that they disappeared.
- **Restoration outcome:** Restore automation one source at a time after manual
  incremental ingestion is proven. Provider file identity/version must be the
  idempotency key; polling an unchanged source must not rewrite D1.
- **Required evidence:** source-specific no-change, correction, replacement and
  failure/retry tests; bounded dispatch; no duplicate queue entries; exact
  allowlist containing only the canary source; account-wide before/after
  samples. Do not enable every source in one change.

### FR-08 — Automated contact-list ingestion

- **State:** Paused independently of roster automation.
- **Includes:** receiving new contact extracts and publishing updated contact
  overlays.
- **Controls:** `CONTACT_AUTOMATION_WRITES_ENABLED=false` and empty
  `CONTACT_AUTOMATION_SOURCE_ALLOWLIST`.
- **Restoration outcome:** Enable one named source only after its provider
  revision, maximum payload, bounded retention and unchanged-input no-op tests
  pass. Restore separately from roster automation so either can be stopped
  without affecting the other.

### FR-09 — Durable Doctor Names and identity aliases

- **State:** Parked; not deployed on Production.
- **Includes:** durable person identity, approved aliases, editable preferred
  display names, merging, later editing/reversal, and low-cost duplicate
  suggestions.
- **Preserved work:** Local branch `codex/durable-doctor-identity-aliases` at
  `1e32e48` contains the later isolated implementation and database-cost gates.
  The remote branch currently points to the earlier reverted state `a06a93f`.
  The branch must not be deleted, reset or casually merged.
- **Restoration outcome:** Return to the non-technical Doctor Names workflow
  after the D1 safety rollout stabilises. Preserve immutable hidden
  `person:<ULID>` identities, user-facing IDs and editable display names.
- **Required evidence:** rebase/review against the then-current safe `main`;
  local-only migration and representative identity fixtures; indexed candidate
  blocking generated only when identities change; no whole-database or
  per-person repeated audit; bounded manual scope; transactional merge and
  exact reversal; no roster-event rewriting; explicit approval before any
  remote migration or rollout.
- **Plans:** the current local `docs/doctor-identity-merge-plan.md` (presently an
  untracked user document),
  `codex/durable-doctor-identity-aliases:docs/doctor-identity-review-ux-plan.md`,
  and the tracked `local-development-safety-plan.md`. Preserve the untracked
  plan and review it deliberately before any later commit; do not sweep it into
  unrelated safety commits.

## Operational capabilities paused

### FR-10 — Roster queue watchdog

- **State:** Removed from Cloudflare on 10 September 2026. Source and
  configuration remain in Git.
- **Control:** `ROSTER_AUTOMATION_ENABLED=false` in
  `wrangler.roster-watchdog.toml`.
- **Effect:** The scheduled Worker returns before dispatch or FindMyShift calls.
- **Restoration:** Restore last, after the individual automation endpoints and
  their allowlists are already proven safe. A health check must expose the
  state, overlapping ticks must coalesce and one tick must have a hard request
  budget.
- **Containment evidence:** Wrangler deletion succeeded and a subsequent
  deployments lookup returned Cloudflare `10007` (Worker does not exist). Its
  15-minute cron was removed with the Worker. No D1 endpoint was called.

### FR-11 — Advanced maintenance and repair

- **State:** Paused.
- **Includes:** full repository reconciliation, reset/rebuild/replace actions,
  and daily-presence repair.
- **Control:** `ROSTER_ADVANCED_MAINTENANCE_ENABLED=false` in addition to the
  roster-write pause.
- **Restoration:** This should remain a time-limited operator capability, not a
  normal always-on feature. Each use requires an exact target, dry-run plan,
  row/statement ceiling, account-wide GO decision and immediate non-D1 stop.

### FR-12 — Facility bootstrap inspection and execution

- **State:** Paused.
- **Controls:** `FACILITY_BOOTSTRAP_INSPECTION_ENABLED=false`,
  `FACILITY_BOOTSTRAP_EXECUTION_ENABLED=false`, and an empty exact-file
  allowlist.
- **Effect:** Neither inspection nor execution may touch D1. These are rollout
  tools, not end-user features.
- **Restoration:** Open inspection and execution separately for one exact file
  and one ED/term, for a short documented window. Inspection never grants
  execution. Close all controls immediately afterward whether it succeeds or
  fails.

## Safety-related behaviour reductions

### FR-13 — Login/calendar background retries

- **State:** Reduced, not disabled.
- **Change:** Post-login background calendar retries were reduced from four to
  one, and hidden tabs do not retry.
- **User value preserved:** Manual reload remains available and the foreground
  calendar load retains its bounded resource-limit retry.
- **Restoration decision:** Do not restore four blind retries. If evidence shows
  reduced resilience, replace them with coalesced revision notifications,
  exponential backoff and a hard per-tab/account budget.

### FR-14 — Automatic persistence of every UI status message

- **State:** Reduced.
- **Change:** Rendering a status message no longer automatically writes it to
  the D1-backed Creator console history. Error reporting and live in-session
  display remain separate.
- **Restoration outcome:** If durable UI history is still desired, restore it as
  sampled/batched telemetry outside the request-critical D1 path with retention
  and write budgets. Do not reinstate one database write per rendered message.

### FR-15 — Duplicate By stream metadata retry

- **State:** Reduced.
- **Change:** Opening By stream makes one metadata attempt rather than
  immediately repeating a failed request.
- **Restoration decision:** Keep the single in-flight attempt. Later resilience
  should use a user-initiated retry or bounded backoff, never a duplicate fetch
  in the same opening sequence.

## Permanent safety mechanisms

These controls are not temporary feature pauses and must remain in place when
features are restored:

- The account-wide GraphQL quota checker and its fail-closed GO/NO-GO rules.
  Dashboard billing summaries are not a substitute for D1 row metrics.
- A complete inventory of every D1 database, Worker, Pages function, scheduled
  job and other caller sharing the account allowance.
- Local development isolation: local databases and storage by default, with no
  inherited Production credentials or bindings.
- Numbered, reviewed migrations and no runtime schema inspection or DDL.
- Query-plan, estimated rows-examined, exact rows-returned and capacity tests;
  correctness fixtures remain separate from scale tests.
- Independent default-off capability flags and exact allowlists for readers,
  builders, imports, roster automation, contact automation and maintenance.
- Control-plane configuration readback after deployment, without exercising an
  application or D1 data route.
- A passive fresh-day baseline and settled account-wide measurements before
  and after each deliberately authorised canary.
- A tested non-D1 rollback/stop path that remains usable after quota exhaustion.
- Per-tab request budgets, request coalescing, one in-flight request per view,
  visibility-aware polling and no automatic retry storms.

Restoration evidence must demonstrate these protections, not remove or bypass
them for convenience.

## Unsafe mechanism replaced permanently

### FR-16 — Runtime schema inspection and creation

- **State:** Replaced permanently.
- **Old behaviour:** Ordinary requests inspected `sqlite_master` and attempted
  `CREATE TABLE/INDEX IF NOT EXISTS` work from multiple Worker isolates.
- **Replacement:** Numbered, reviewed migrations applied once through a
  controlled deployment step. Local disposable databases use the explicit
  initializer.
- **Restoration decision:** Never restore runtime DDL. The useful capability—an
  up-to-date schema—returns only through migrations.

### FR-17 — Legacy historical At a glance reads

- **State:** Replaced permanently.
- **Old behaviour:** Staff, coverage and other At a glance requests repeatedly
  scanned roster-event history.
- **Replacement:** Incremental compact facts and shared R2 artifacts keyed by
  the data identity rather than the viewer.
- **Restoration decision:** Never set `FACILITY_LEGACY_READS_PAUSED=false` in
  Production. Restore FR-01 through the shared reader only.

## Important features still live

These are recorded because anxious avoidance is not the same as a deployed
pause, and future maintenance work must not accidentally disable them.

### FR-18 — Personal calendar and ordinary login

- **State:** Still live, subject to the account-wide D1 allowance being
  available.
- **Note:** These paths remain the protected core service. Optional rollout
  work stops before their daily operating headroom is threatened.

### FR-19 — Colleague insight tools on calendar events

- **State:** Still live; not covered by the At a glance rollout flags.
- **Includes:** “Who else is working with me?” and “When am I working with…?”.
- **Current implementation:** Authenticated, date- and source-bounded D1 event
  queries (`queryRosterInsights` and `queryRosterOverlapDoctors`).
- **Required follow-up:** Measure their settled production frequency and local
  query plans before increasing use. If they consume material rows, move them
  to the same shared compact day/range artifacts. Do not disable them merely
  because they were not exercised during the incident investigation; any
  future pause must be added here first.

### FR-20 — At a glance entitlement configuration

- **State:** Still live.
- **Includes:** Creator configuration of whether an account has At a glance and
  doctor-profile entitlement lookup.
- **Note:** Entitlements remain stored even while the data workspace is paused,
  so restoration does not require recreating user settings.
- **Safety decision:** Do not disable individual entitlements during global
  maintenance. The pre-authentication maintenance gate already blocks every At
  a glance data/edit action regardless of entitlement. Editing user settings
  would add D1 writes, destroy desired configuration and provide no additional
  protection.

### FR-21 — Automatic doctor discovery during login/account loading

- **State:** Paused in Production behind a default-off control.
- **Risk:** Unclaimed or incompletely claimed ordinary accounts can fall back
  from compact identity data to roster-file doctors and historical event
  comparisons. This is a plausible high-cost path but is not proven as the
  source of the unattributed 8 September burst.
- **Containment:** Preserve login and existing claims, but fail closed with a
  friendly identity-linking-unavailable state when compact records are absent.
  No login/account-context request may scan roster history.
- **Restoration:** Restore automatic suggestions from bounded indexed candidate
  blocks maintained only when a roster identity changes.
- **Plan:** `ordinary-login-identity-d1-remediation-plan.md`.

### FR-22 — Automatic account repair and identity seeding

- **State:** Reduced in Production.
- **Risk:** A general account save can rewrite profiles, claims, aliases and
  locations, while durable identity rows may be seeded as a side effect.
- **Containment:** Split mutations by responsibility; semantic no-ops write
  zero rows; identity changes occur only during explicit identity actions or
  incremental ingestion.
- **Restoration:** Retain useful repair as an explicit, bounded, idempotent
  maintenance operation—not an ordinary login/save side effect.

### FR-23 — Snapshot warm-up after ordinary account saves

- **State:** Paused in Production behind a default-off control.
- **Risk:** Ordinary saves can schedule post-response snapshot preparation and
  hidden D1 work even when roster facts did not change.
- **Containment:** No snapshot warm-up follows UI-state/profile saves. Builders
  require a changed dependency revision, independent default-off control, exact
  scope and a hard per-run budget.
- **Restoration:** Revision-driven, coalesced rebuilds only; unchanged inputs
  perform zero D1/R2 work.

### FR-24 — Creator user directory identity/seniority enrichment

- **State:** Paused in Production while identity discovery is disabled.
- **Note:** `listUsers` is an explicit Creator action and is not a credible
  explanation for an incident when the Creator did not open the app. It still
  must not derive identity or seniority from event history.
- **Restoration:** Serve bounded compact account/term summaries, or show the
  enrichment as temporarily unavailable.

### FR-25 — Retained deployments with Production D1 bindings

- **State:** Contained at the Pages deployment layer. On 10 September, the
  complete control-plane inventory showed 262 callable deployments: eight
  Production and 254 Preview. Seven superseded Production deployments and all
  254 Preview deployments were deleted. Independent environment-specific
  listings then showed exactly one Production deployment (`bae86cea…`, source
  `8ef0215`) and no Preview deployments.
- **Evidence:** Analytics exposed a `sqlite_master` fingerprint not emitted by
  current code. Preview D1 reported zero rows during the investigated incidents,
  so retained Preview deployments were not proved to be their cause; deletion
  nevertheless removes their callable URLs and code as a future risk.
- **Containment:** Keep only the current Production deployment. After any future
  release is verified, delete its predecessor and any generated Preview
  deployment. Git history remains the source for rollback; rollback requires a
  deliberate fresh deployment with current fail-closed configuration.

### FR-26 — Automatic Creator login/bootstrap fan-out

- **State:** Containment deployed to Production as `a712068`; effective flags
  verified closed. The fresh-day passive gate passed, but the Creator canary
  failed after a later-settling 992,117-read bucket appeared. All related
  restoration remains blocked.
- **Observed behaviour:** On 9 September, one controlled Creator login plus one
  attempt to enter maintenance-disabled At a glance and a return to Calendar
  coincided with 1,623,542 reads in one five-minute bucket. The subsequent idle
  tab added only 109 settled reads and no writes.
- **Risk:** Creator hydration still queues user-directory, roster-status,
  doctor-switcher, calendar/bootstrap and snapshot work that is not required
  to authenticate or display a cached calendar. Cached-calendar rendering also
  scheduled a colleague-insight warm-up without an explicit user action.
- **Containment:** Minimal Creator login; no automatic Creator/Admin data fan-
  out; independently gated, on-demand surfaces; zero D1 for disabled At a
  glance navigation and post-login idle time.
- **Restoration:** Restore each on-demand surface only after its own bounded
  local cost test and production canary.
- **Plan:** `creator-login-bootstrap-d1-remediation-plan.md`.

### FR-27 — Persistent API invocation attribution

- **State:** Implemented for the next controlled Production deployment.
- **Effect:** Every API invocation records route, state action, deployment,
  status, containment result and request-local D1 statement/row metadata in a
  dedicated Workers Analytics Engine dataset and the real-time Functions log.
  Sensitive bodies, credentials, identities, roster data and contact details
  are excluded.
- **Safety:** Paused automation requests stop in middleware before D1/R2,
  authentication, body parsing or outbound work. Requests also have a hard D1
  statement ceiling; this supplements indexed query-plan limits.
- **Restoration:** Keep privacy-safe invocation attribution and statement
  ceilings permanently. Sampling may be reduced only after the incident is
  attributed and sustained safe operation is demonstrated.

## Restoration order

The order restores product value without reopening several D1 consumers at
once:

1. Keep the documented zero-D1 At a glance maintenance gate active and preserve
   all individual entitlements.
2. Deploy the ordinary-login, identity-save and snapshot-warm-up containment in
   FR-21–FR-25, then observe a passive fresh quota day.
3. Restore the bounded Admin → Files summary read only.
4. Apply separately approved compact-fact migrations and bootstrap one exact
   file/ED/term.
5. Restore At a glance shared reads to the Creator for one ED, with contacts
   still unavailable.
6. Restore shared On shift contacts and their zero-D1 visible-page refresh.
7. Expand At a glance by ED and cohort only after settled evidence at each
   step.
8. Restore manual incremental roster imports for one source.
9. Restore automatic roster and contact ingestion source by source.
10. Restore the watchdog only after all called endpoints are already safe.
11. Reintegrate the durable Doctor Names work on a fresh branch from the stable
    Production baseline.

FR-13 through FR-17 are not prerequisites to undo. Their original unsafe
mechanisms are intentionally excluded from restoration.

## Incident change ledger

| Commit | Change recorded here |
| --- | --- |
| `cff238d` | Paused quota-heavy roster automation and watchdog activity (FR-07, FR-10). |
| `8d0cf60` | Extended the write pause to manual roster mutations (FR-06). |
| `aa9eed8` | Removed runtime DDL from ordinary requests (FR-16). |
| `716ac43` | Permanently paused legacy At a glance reads (FR-01, FR-17). |
| `95bc9ad` | Blocked the expensive Admin → Files status path (FR-05). |
| `9927321` | Added accurate maintenance messaging in Admin → Files (FR-05). |
| `6bda8b7` | Separated and closed roster, contact, bootstrap, queue and maintenance capabilities (FR-01, FR-04, FR-07, FR-08, FR-11, FR-12). |
| `2af89c0` | Reduced automatic status writes, calendar retries and duplicate By stream requests (FR-13–FR-15). |
| `a146029` | Added the zero-D1 At a glance maintenance UI and pre-authentication gate (FR-01–FR-04). |
| Plan, 8 Sep 2026 | Recorded required ordinary-login/identity/save/warm-up containment and retained-deployment audit (FR-21–FR-25); no runtime change. |
| Local implementation, 8 Sep 2026 | Added default-off identity discovery and account snapshot-build controls; removed login history fallback and automatic repairs; made ordinary saves incremental with no snapshot warm-up. Not yet deployed. |
| `18ba7a1` / `a4305bca-5538-4583-88f8-655bcb687556` | Deployed ordinary-login containment with all safety controls read back closed; removed the 16 retained pre-containment Production deployments by exact ID (FR-21–FR-25). |
| Observation, 8 Sep 2026 | Settled 15:15–20:30 AEST Production usage after containment was 27,628 reads and 20 writes, with a maximum five-minute bucket of 1,658 reads. Full-day status remains STOP until a clean UTC quota day completes the passive gate. |
| Incident and cleanup, 9 Sep 2026 | The fresh-day gate found a 1,693,486-read burst at 10:25 AEST and a pre-containment schema fingerprint. Wrangler's 25-result listing had hidden hundreds of callable Production hashes. Removed 673 additional pre-containment Production deployments in bounded batches; an independent Production-only listing confirmed exactly four contained deployments remain. Optional features remain closed pending post-cleanup settled evidence. |
| Controlled Creator test, 9 Sep 2026 | Creator login plus one blocked At a glance navigation coincided with 1,623,542 reads in one five-minute bucket. The idle tab did not repeat the burst. Added FR-26 and blocked further Creator testing pending minimal-login remediation. |
| Creator-login containment release, 9 Sep 2026 | Deployed `a712068` to Production (`65d4ac5f-9d97-4464-94e8-8210aa37145a` and same-source Git deployment `b7d8d8e5-c53f-4aab-8951-3e10c3091bb5`). Control-plane read-back confirmed both Creator controls and all prior safety controls closed. No application or D1 request was made. |
| Creator-login canary, 10 Sep 2026 | Fresh-day passive usage passed after almost four settled hours. The immediate 14:08 Creator-login bucket used 64 reads and one write, but later settlement exposed 992,117 reads in the 14:30 bucket with almost no fingerprint attribution. This preceded the disabled At a glance test, whose 14:40 bucket used only 33 reads and six writes. The Creator/bootstrap gate is failed, not passed. |
| Deployment cleanup, 10 Sep 2026 | Removed seven superseded Production deployments and all 254 Preview deployments through the Pages control plane. Final listings showed one Production deployment (`bae86cea…`, source `8ef0215`) and zero Preview deployments. No application or D1 endpoint was called. |
| Active-caller containment, 10 Sep 2026 | Deleted the deployed `roster-queue-watchdog` and its cron; added pre-handler automation containment, persistent privacy-safe invocation records and request-local D1 accounting/ceilings (FR-10, FR-27). Local paused-contact smoke test used zero D1 statements; invalid login used one row read. Power Automate source-side pause remains an operator action. |

## Restoration record template

Append this block to the relevant entry whenever its state changes:

```text
State changed to:
Approved by:
Commit and deployment:
Flags/allowlist:
Local evidence:
Estimated D1/R2 cost:
Pre-change account usage:
Post-change settled usage:
Rollback/stop verified:
Known limitations:
```
