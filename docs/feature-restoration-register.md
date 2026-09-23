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

Current tracked Production code before the Creator/MMC reader canary:
`f81d7611` on 22 September 2026. The exact active deployment identity must
still be read back at every operational gate. Ordinary-login containment and
all permanent quota protections remain active. The Creator-self cohort may
read only the completed MMC shared cache; facility builders, contacts, other
facility readers, bootstrap and advanced-maintenance controls remain closed.

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
| `FACILITY_OVERVIEW_MAINTENANCE_MODE` | `false` in Production for the Creator/MMC canary; `true` in Preview; missing or malformed fails closed | FR-01–FR-04 |
| `FACILITY_OVERVIEW_AUTOMATIC_LAUNCH_ENABLED` | `false` in Production and Preview; missing or malformed fails closed | FR-03 |
| `FACILITY_SHARED_ROLLOUT_ACTIVE` | `true` in Production; `false` in Preview | FR-01 |
| `FACILITY_SHARED_EMERGENCY_PAUSED` | `false` in Production; `true` in Preview | FR-01 |
| `FACILITY_LEGACY_READS_PAUSED` | `true` permanently | FR-01, FR-17 |
| `FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST` | empty | FR-01, FR-12 |
| `FACILITY_SHARED_READER_SOURCE_ALLOWLIST` | `mmc` in Production; empty in Preview | FR-01 |
| `FACILITY_SHARED_READER_COHORT` | `all` in Production for MMC only; empty in Preview | FR-01 |
| `FACILITY_ACCESS_MATERIALIZATION_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_METADATA_BUILD_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_DAYS_BUILD_ENABLED` | `false` | FR-01 |
| `FACILITY_SHARED_CONTACTS_BUILD_ENABLED` | `false` | FR-04 |
| `FACILITY_SHARED_METADATA_ENABLED` | `true` in Production; `false` in Preview | FR-01 |
| `FACILITY_SHARED_DAYS_ENABLED` | `true` in Production; `false` in Preview | FR-01 |
| `FACILITY_SHARED_CONTACTS_ENABLED` | `false` | FR-04 |
| `ROSTER_AUTOMATION_WRITES_ENABLED` | `false` | FR-06, FR-07 |
| `MANUAL_ROSTER_WRITES_ENABLED` | `false` | FR-06, FR-11 |
| `ROSTER_STATUS_SUMMARY_ENABLED` | `true` in Production; `false` in Preview | FR-05 |
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
| `ROSTER_INSIGHT_READS_ENABLED` | `false`; missing or malformed also fails closed | FR-19 |
| Watchdog `ROSTER_AUTOMATION_ENABLED` | `false` | FR-10 |

## User-facing restoration register

### FR-01 — At a glance workspace

- **State:** MMC cached readers restored for entitled users. Other facilities
  remain paused.
- **Includes:** On shift, ED Staff, By stream, Working together, multi-ED
  Director overview, facility/date selection and At a glance navigation on
  desktop and mobile. It also includes the compact staff-membership,
  file/hospital-coverage, additional-coverage and daily-presence facts needed
  to build those views without repeatedly deriving them from event history.
- **User effect:** Entitled users can manually open MMC cached metadata, On
  shift, Staff and related base day views. Other facilities still receive the
  controlled unavailable response.
- **Controls:** Production uses `FACILITY_SHARED_ROLLOUT_ACTIVE=true`,
  `FACILITY_SHARED_EMERGENCY_PAUSED=false`, reader source `mmc`, cohort
  `all`, and shared metadata/day readers only. Contacts and every builder
  remain false, all build allowlists remain empty, and
  `FACILITY_LEGACY_READS_PAUSED=true` permanently.
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
- **Control:** `FACILITY_OVERVIEW_AUTOMATIC_LAUNCH_ENABLED=false` independently
  suppresses automatic launch while manual MMC readers are restored.
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

- **State:** Restored through the bounded compact summary reader.
- **Includes:** Auto-sync source status, manual files for the current/next term,
  and earlier manual-file status.
- **User effect:** Creator users can see Auto-sync and retained manual-file
  status without invoking the former historical counting path.
- **Control:** `ROSTER_STATUS_SUMMARY_ENABLED=true`.
- **Data/code preservation:** The former historical counting path is not used.
  The bounded summary implementation and required schema are deployed.
- **Restoration outcome:** Completed. Keep the useful status information on
  compact per-file/source summaries maintained during ingestion.
- **Required evidence:** Indexed reads scale with the number of retained files,
  not event history; requests coalesce; unchanged responses write zero rows;
  hidden tabs do not retry; the bounded Creator canary and settled observation
  remain recorded in the incident ledger.

### FR-06 — Manual roster import and file management

- **State:** Paused.
- **Includes:** upload/retain a source file, save its parsed calendar, remove an
  import, reconcile retained files, reset/reparse a derived file and replace
  the active file set. Reading an already retained raw file is not intentionally
  disabled.
- **Control:** `MANUAL_ROSTER_WRITES_ENABLED=false`; destructive/rebuild
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

- **State:** Partially restored. MMC, MCH and VHH are enabled; DDH remains
  paused after an isolated upstream HTTP 503 response.
- **Includes:** FindMyShift change checking, generic automated upload, VHH
  extraction, derived processing, queue dispatch and pending work processing.
- **Controls:** Production roster automation writes and bounded queue processing
  are enabled only for the exact allowlist `monash-adults,monash-paeds,
  vhh-active-medical-roster`. DDH, contacts, Preview and unrelated automation
  remain excluded.
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
- **Performance requirement (22 September):** ordinary login, Creator login
  and Creator switching to claimed/unclaimed calendars must render cache-first
  and use one bounded indexed background validation. Target first paint is two
  seconds with a valid cache and five seconds for a cold single-doctor build.
  A cold claimed-account switch observed at approximately 30 seconds and a
  Creator login at approximately 10 seconds remain performance defects even
  though subsequent cached attempts were faster.
- **Switcher requirement:** the Creator's complete claimed/unclaimed doctor
  directory remains available from every Creator-entered calendar. It is
  carried locally across switches and must not require a directory-wide D1
  query.

### FR-19 — Colleague insight tools on calendar events

- **State:** Paused behind the default-off `ROSTER_INSIGHT_READS_ENABLED`
  control. Automatic background warm-up is also disabled in the client.
- **Includes:** “Who else is working with me?” and “When am I working with…?”.
- **Incident evidence:** At 15:57 AEST on 14 September, an ordinary calendar
  render automatically invoked `queryRosterOverlapDoctors`; request telemetry
  reported 998,070 rows read by its legacy overlapping-event join.
- **Containment:** Both insight actions stop in middleware before D1 unless the
  explicit control is true. Calendar rendering never schedules a remote insight
  warm-up. Explicit use receives the existing unavailable presentation.
- **Restoration:** Replace the legacy event joins with indexed compact daily
  presence/shared artifacts, prove bounded plans and returned/examined rows,
  then restore explicit user actions first. Automatic warm-up is not restored
  unless it has a separately demonstrated zero-D1 cache hit.

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

### FR-28 — Automatic bounded personal-calendar refresh

- **State:** Restored as one automatic revision check after login while broad
  Creator startup hydration remains contained. The temporary manual refresh
  button was removed after its production canary served its purpose.
- **Problem addressed:** A successful roster replacement changes the active D1
  events immediately, but an existing browser/R2 calendar snapshot can remain
  on its previous revision when automatic snapshot warm-up is disabled.
- **Behaviour:** After the cached calendar paints, one authenticated,
  date-bounded, doctor-specific revision check runs only while the page is
  visible. A matching browser revision avoids the R2 payload and rebuild. If
  stale, the Creator path may rebuild only that visible calendar. Applying the
  returned server snapshot suppresses colleague-insight warm-up. Rehydrating
  already-saved custom events is always local-only and cannot schedule the
  former redundant full cloud save. It does not open Admin, At a glance, roster
  automation or a recurring poll.
- **Permanent constraint:** Keep automatic Creator fan-out and global snapshot
  warm-up disabled until separately restored under FR-23/FR-26. This explicit
  action must retain the API request ceiling and the indexed doctor/date event
  query.

### FR-29 — Browser-local session settings during D1 recovery

- **State:** Cloud persistence of ordinary calendar display/session changes is
  temporarily paused. Settings, filters, overrides and undo history continue
  to persist in the current browser workspace.
- **Reason:** Request attribution at 18:03 and 19:06 AEST on 15 September 2026
  proved that simply rendering a cached or refreshed calendar scheduled full
  `save` requests. One request per render wrote 12 rows; the following request
  attempted a batch larger than the permanent 64-statement ceiling and was
  blocked with zero writes. Both accompanying `loadCalendarEvents` requests
  succeeded with only 10–16 statements and 3–11 rows read.
- **Containment:** `rebuildClientPreview()` is render-only and
  `saveCurrentSessionState()` writes only to browser storage. Neither may call
  the cloud save queue. Explicit roster/account operations retain their own
  deliberate persistence paths.
- **Restoration:** Restore cross-device session-setting persistence only through
  a dedicated, bounded session-only API action. It must update a single account
  state record and must never resubmit roster imports or rebuild snapshots.

## Restoration order

The order restores the core personal-calendar service before At a glance and
does not reopen several D1 consumers at once. The detailed source-by-source
sequence and acceptance gates are authoritative in
[`core-calendar-sync-restoration-plan.md`](./core-calendar-sync-restoration-plan.md):

1. Keep the documented zero-D1 At a glance maintenance gate active and preserve
   all individual entitlements.
2. Preserve the deployed ordinary-login, identity-save, snapshot-warm-up and
   bounded Admin → Files protections.
3. Separate automatic ingress/queue authority from manual roster mutation and
   prove exact source-scoped processing locally.
4. Restore automatic roster synchronisation source by source: MMC, MCH, VHH and
   then DDH, with changed and unchanged canaries and settled evidence at every
   new implementation class.
5. Restore ordinary manual roster import separately; retain destructive and
   advanced repair gates until individually proven.
6. After roster syncing is stable, resume the resumable At a glance publication
   sequence for one exact MMC term.
7. Restore At a glance shared reads to the Creator for one ED, with contacts
   still unavailable, then expand by ED and cohort only after settled evidence.
8. Restore shared On shift contacts and their zero-D1 visible-page refresh.
9. Restore automatic contact ingestion source by source.
10. Restore a watchdog only if it remains useful after every called endpoint is
    already safe; do not restore the historical global poller.
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
| `66402b2` / `382049f6-edcb-4ef2-b12e-5d5081adcd01`, 10 Sep 2026 | Explicitly deployed active-caller containment after Analytics Engine was enabled. Effective configuration read-back showed maintenance/emergency/legacy, roster/contact automation and Creator startup controls closed and `REQUEST_ANALYTICS` bound to `roster_api_invocations`. A single payload-free paused contact probe returned 503 and logged zero D1 statements/rows. The predecessor was deleted; final inventory was one Production and zero Preview. |
| External-caller pause, 10 Sep 2026 approximately 19:20 AEST | Operator confirmed all eleven shared Power Automate roster/contact/bootstrap flows were turned off. Flow definitions were preserved. Restore individually under FR-07/FR-08 only after attribution and source-specific gates pass. |
| Monash Adults routine-ingress restoration, 15 Sep 2026 at approximately 21:04 AEST | Verified Power Automate flow `Sync Monash roster files` (`1dd007f9-868d-4fce-a0fd-ed03d413a4ac`), narrowed its SharePoint trigger from `Adult`/`Paeds` to `Adult` only, saved it with zero Flow checker errors/warnings, and enabled it after Production deployment `0373666` became active. Production permits only `monash-adults`, with queue processing enabled and a hard 1,250 incremental-fact ceiling; Preview, Paediatrics, manual roster writes, contacts, identity work and facility readers remain closed. A pinned content hash remains available for one-shot canaries, while routine future versions use the exact source allowlist and fact ceiling. Rollback is **Turn off** this named flow, then redeploy the preceding closed configuration if server ingress must also close. |
| Monash Paediatrics preflight, 15 Sep 2026 | Created the separate disabled Power Automate flow `Sync Monash Paediatrics roster files` (`0d191bae-d7c5-47f1-a1c9-e754cb288790`) with an exact `Paeds` filename trigger; the copied HTTP mapping therefore resolves only to `monash-paeds`. Downloaded the current read-only SharePoint workbook `Paeds - Term 3 2026.xlsx` (487,475 bytes; SHA-256 `c4e6b3bfe18528f9651ee699a7b58a7b7790faaa99dadafb4d978f9f06b8afaf`) and parsed it locally as MCH only: 79 doctors, 2,376 events and nine issues. Against the retained 12 Aug Production baseline it changes 345 event facts (178 additions and 167 removals), below the hard 1,250 incremental-fact ceiling. Source-isolation, unchanged-ingress and queue-failure tests passed. The server allowlist may be widened to `monash-adults,monash-paeds`, but the Paediatrics flow must remain off until one controlled submission and its request-attribution evidence pass. Rollback is **Turn off** the named Paediatrics flow and remove `monash-paeds` from the server allowlist. |
| Monash Paediatrics restoration, 15 Sep 2026 at approximately 21:31 AEST | Deployed `21ea99e` as Production deployment `c22cb1db-1eb9-4d42-97b2-b2baf81cc0a2`, widening only the exact automation allowlist to `monash-adults,monash-paeds` while retaining the 1,250 incremental-fact ceiling and all unrelated pauses. Enabled the isolated edit-trigger flow `Sync Monash Paediatrics roster files`. Updated the preserved instant flow `Bootstrap Paeds Term 3 roster` (`39e8cd85-b596-439f-85b8-eb13d6f5d941`) to unique provider version `bootstrap-term-3-20260915-c4e6b3bf`, ran it once against the exact reviewed SharePoint path, and immediately turned it Off again. Power Automate completed in four seconds and GitHub processor run `34963749546` completed all ingestion, parsing and completion steps successfully without a fact-budget failure. Routine Paediatrics changes now sync automatically; the bootstrap remains disabled. A settled account-wide usage sample and user-facing MCH calendar verification remain the observation evidence, not prerequisites for keeping the bounded flow enabled. |
| Chunked publication implementation, 13 Sep 2026 | Replaced the callable 91-day publication endpoint with explicit `plan`, seven-day `build-batch`, one-month `build-month` and atomic `finalize` modes. Staging remains invisible, retries are idempotent, finalisation is fenced by D1 ownership and R2 preconditions, and each mode has its own D1 statement ceiling. Production flags and Power Automate remain closed pending the serial canary in `facility-publication-chunking-remediation-plan.md`. |
| Calendar-sync-first plan, 13 Sep 2026 | Made source-isolated roster synchronisation the next restoration priority and moved At a glance publication/readers after stable calendar syncing. No runtime control changed. See `core-calendar-sync-restoration-plan.md`. |
| Pre-sync deployment cleanup, 13 Sep 2026 | Removed 25 superseded, contained Production deployments through the Pages control plane. The 25-result listing initially concealed one further predecessor, which was removed after re-listing. Final inventory: one Production deployment (`1890d569-33f3-414c-8587-9101c27dc921`, source `93feb08`) and zero Preview deployments. No D1 query was used. |
| Calendar-sync Gates 1–2, 13 Sep 2026 | Two settled account-wide samples reconciled at 75,025 reads and zero writes with no expensive fingerprint. Implemented local-only manual/automatic authority separation and exact-source, one-job queue processing. Focused and full fixture suites pass; Production flags and all Power Automate flows remain closed. |
| Calendar-sync Gate 3, 13 Sep 2026 | Proved local zero-write unchanged ingress for Monash, VHH and FindMyShift; bounded changed ingress, payloads and duplicate callbacks; indexed source/version, source/status and affected-claim probes; and stopped disabled identity/snapshot fan-out after ingestion. Representative parsing, correction, rollback, overlap, membership and term tests pass. Migration `0032` remains local only. Production and all Power Automate flows remain closed. |
| Gate 6 DDH containment and VHH preflight, 20 Sep 2026 | The isolated DDH provider check returned HTTP 503 with four statements, one row read and two compact status writes; no ingestion or processor ran and the retained roster stayed active. DDH is paused for a later provider window. VHH extraction, isolation and unchanged-input tests pass. The exact-workbook instant flow `Sync VHH Active Medical Roster to Production` is selected for one controlled canary; the folder-wide automated VHH flow remains off pending exact-trigger and retry review. |
| VHH routine restoration, 20 Sep 2026 | The controlled VHH canary parsed 46 doctors and 971 events successfully. Its derived phase used 41,057 D1 row reads and 2,128 writes within the request-local ceiling; Claire CHARTERIS's calendar loaded without error. An identical replay then used two statements, zero row reads and zero writes and dispatched no processor run. The automated flow `VHH Active Medical Roster to Production` now has concurrency one, an exact `Active Medical Roster.xlsx` trigger and no HTTP retries; Flow checker reported zero errors/warnings before it was enabled. The instant canary flow remains Off. Production permits only MMC, MCH and VHH roster sources; DDH remains excluded and paused. |
| VHH autosave-loop containment, 21 Sep 2026 | Editing `Active Medical Roster.xlsx` caused the automated VHH flow to submit successive SharePoint autosave versions approximately once per minute. Because each version had different content, content-hash idempotency could not suppress them. Several imports then rejected the new designation `Swing 12:30PM`. Turned off `VHH Active Medical Roster to Production`, removed `vhh-active-medical-roster` from the Production allowlist in `9ed96c80`, and verified no later processor dispatch. Existing VHH calendars remain available. Commit `0a8a88ac` generalises VHH Swing parsing: `Swing` defaults to 10:00 and an encoded start in 12- or 24-hour form produces a 9.5-hour `VHH: Swing` event. Do not restore routine VHH automation until the flow has a stability/debounce check that rereads the current file version after a delay and stops stale trigger versions before HTTP submission; then use one controlled changed canary and one unchanged replay. |
| Monash Paediatrics historical-file containment, 21 Sep 2026 | The enabled Paediatrics flow's broad `Paeds` filename condition submitted historical `Paeds - Term 1 2026.xlsx`. The parser correctly rejected it because its workbook dates crossed term/year boundaries; the retained active MCH calendar was not replaced. Turned off `Sync Monash Paediatrics roster files` (`0d191bae-d7c5-47f1-a1c9-e754cb288790`) and removed `monash-paeds` from the Production allowlist in `a17dd1be`. Existing MCH calendars remain available. Do not weaken the cross-term parser protection. Restore only after replacing the broad trigger with an explicit current/next-term source selection that cannot submit historical workbooks, followed by one controlled current-workbook canary and one unchanged replay. |
| VHH debounced routine restoration, 21 Sep 2026 | Added a five-minute stability delay to `VHH Active Medical Roster to Production`, reread the current SharePoint file metadata, and stopped stale trigger versions successfully before HTTP submission. Concurrency remains one and the trigger remains restricted to `Active Medical Roster.xlsx`. Flow checker reported zero errors/warnings. Reopened `vhh-active-medical-roster` in `26d217b6`; the controlled current-workbook canary parsed 47 doctors and 966 events successfully, and its identical replay dispatched no processor job. The instant canary is Off and the guarded routine flow is On. Rollback is **Turn off** the routine VHH flow and remove its source from the allowlist. |
| Monash Paediatrics guarded routine restoration, 21 Sep 2026 | Restricted `Sync Monash Paediatrics roster files` to the exact current workbook `Paeds - Term 3 2026.xlsx`, set concurrency to one, added a five-minute stability delay, reread current SharePoint metadata, and stopped stale trigger versions successfully before HTTP submission. Flow checker reported zero errors/warnings. Reopened only `monash-paeds` in Production commit `9bfcdff0` / deployment `74e4b516-91f6-419d-ba08-a0513beb2ce6`, retaining the 1,250 incremental-fact ceiling and all unrelated pauses. The exact-workbook instant canary succeeded in one second; because the workbook content was unchanged from the retained successful import, ingress idempotency dispatched no GitHub processor job. The instant canary is Off and the guarded routine flow is On. Rollback is **Turn off** the routine Paediatrics flow and remove `monash-paeds` from the allowlist. |
| `c8ac83e` / `00042ece-aa3a-4784-bffe-008e90381d7c`, 14 Sep 2026 | Settled account telemetry attributed the 15:55 bucket's million-row burst to automatic `queryRosterOverlapDoctors` during an ordinary session. Deployed a default-off pre-D1 server gate for both insight actions and disabled client background warm-up. A credential-free Production probe returned 503 `roster-insights-paused`; one current Production and zero Preview deployments remain. The first settled 07:22–08:15 UTC interval recorded 4,209 rows read by 23 calendar-feed requests, zero writes, and exactly one contained insight request (the deliberate probe) with zero D1 statements/rows. The controlled 18:31 AEST ordinary login used three reads; its incremental save used 12 reads and one write. By the fully settled 19:44 AEST cutoff, account usage had increased by only 3,584 reads and one write over the preceding 90 minutes, with no expensive query or second insight request. Roster automation remains closed; restoration is tracked under FR-19. |
| DDH and switcher restoration, 22 Sep 2026 | DDH completed a guarded import of 149 doctors and 3,651 events. Dennis CHUNG payroll-transfer shifts and Steve GUASTALEGNAME's Rover shift were verified. `90bec097` repaired duplicate retained-file reprocessing and false queued status; `79c5c7ec` carries the visible Creator doctor directory across claimed/unclaimed views without another D1 request. Cold Creator/claimed-account performance remains an explicit FR-18 acceptance item. |
| Gate 8 analytics correction, 22 Sep 2026 | Three settled read-only samples showed low account usage and no expensive fingerprint, but independently sampled query-fingerprint totals differed from the exactly reconciled daily/five-minute quota ledger. Revised and implemented the three-ledger checker locally; focused account-budget, quota and request-attribution tests pass. The first fresh revised baseline at 14:59 AEST is valid at 37,987 reads, 184 writes and a maximum five-minute bucket of 4,603 reads; it returns `STOP` only because a second sample is required. At a glance remains closed until that second sample returns explicit `GO`. |
| MMC exact-file bootstrap inspection, 22 Sep 2026 | After a settled revised baseline returned `GO`, Production inspection was enabled only for `mmc` and exact active file `automation:monash-adults:47c0951dd2582465d59b19a9`; execution and every At a glance reader/builder remained disabled. Initial runs returned `compactReady: true`, `statusReady: true`, raw source available and content revision `4cc784efe113e6642b56f963a713967594abbfb08866450fdc0d4312f65cc390`, but exposed that D1 `.first()` omits row metadata. Commit `e5d9ddc9` changed only these four exact probes to metadata-returning `.all()` calls and normalized the response Ray ID. Final workflow run `35695192406` captured request `a3ef43f81b8f0e96` with the same ready revision. No bootstrap execution is required. The temporary inspection/source/file allowlists were closed immediately; wait for settled exact attribution before any reader canary. |
| MMC inspection attribution gate, 22 Sep 2026 | Settled request `a3ef43f81b8f0e96` reconciled to exactly four D1 statements, four rows read, zero rows written, complete metadata and a 16-statement limit. The account-wide canary report `/private/tmp/d1-mmc-inspection-canary.json` returned explicit `GO`: 108,182 daily reads, 250 writes, a largest post-baseline five-minute bucket of 929 reads, no expensive fingerprints and no unreviewed fingerprints. This authorises only the next read-only plan for one MMC-term shared publication; it does not authorise bootstrap execution, publication writes or shared readers. |
| MMC Term 3 publication plan, 22 Sep 2026 | Read-only workflow `35698071882`, request `a3ef79baca4286e4`, planned MMC Term 3 (`2026-08-03`–`2026-11-01`) with revision `8923f979716e4603001a75a2259197e65fd12f6e299574e8d15b8298a98f947e`: 91 dates, 13 serial seven-day batches and four month objects. Planning declared seven D1 statements and zero writes. Its `eachBatch.maximumRowsExamined` value is 145,835 because it reuses the whole-operation conservative ceiling; that exceeds the 100,000-read At a glance hard stop and cannot authorise batch execution. The temporary maintenance/source window was closed immediately. Require settled request attribution and a genuine per-batch ceiling before any publication writes. |
| MMC per-batch estimate remediation, 22 Sep 2026 | Corrected the chunked plan estimate to count one compact planning pass, one publication-state row and only the batch's seven indexed ED/date probes: conservative maximum 53,167 examined rows per batch, below the 100,000 hard stop. Month/finalisation steps now publish their own bounded row ceilings. Exact publication-state probes use metadata-returning `.all()` calls so request attribution cannot silently report incomplete row counters. Focused materialisation, request-attribution and quota tests pass. Production remains closed; the correction does not authorise a batch until the preceding read-only plan request settles and passes the account-wide gate. |
| MMC publication-plan attribution gate, 22 Sep 2026 | Settled request `a3ef79baca4286e4` used seven statements, read 391 rows, wrote zero rows and reported complete metadata. The account-wide canary returned `GO` at 114,232 daily reads and 250 writes, with a largest post-baseline five-minute bucket of 1,418 reads, no expensive fingerprints and no unreviewed fingerprints. This authorises only MMC Term 3 batch 0 under operation revision `8923f979716e4603001a75a2259197e65fd12f6e299574e8d15b8298a98f947e`; later batches still require batch 0 completion and settled attribution. |
| MMC Term 3 publication batch 0, 22 Sep 2026 | Workflow `35700457227`, request `a3efa3c64aaf6e9e`, successfully staged seven day objects for 3–9 August under operation revision `8923f979716e4603001a75a2259197e65fd12f6e299574e8d15b8298a98f947e` (`pointerCount: 7`). The candidate remains invisible because no month assembly or finalisation has occurred and all shared readers remain disabled. The temporary maintenance/source window was closed immediately. Require settled exact request attribution and account-wide `GO` before batch 1. |
| MMC batch 0 settled cost, 22 Sep 2026 | Request `a3efa3c64aaf6e9e` completed with 16 statements, 1,062 rows read, one row written, complete metadata and a 20-statement hard limit. This is far below both its 53,167-row conservative ceiling and the 100,000-read hard stop. With daily usage still far below the five-million allowance, batches 1–12 may run serially in one MMC-only maintenance window with automatic stop-on-failure; shared readers remain disabled until all batches, months and fenced finalisation succeed. |
| MMC Term 3 consolidated publication, 22 Sep 2026 | After batch 0 measured only 1,062 reads and one written row, batches 1–12 ran serially with independent 20-statement limits and automatic stop-on-failure; every batch succeeded. August, September, October and November month objects then assembled serially, and fenced finalisation workflow `35713481875` succeeded for operation revision `8923f979716e4603001a75a2259197e65fd12f6e299574e8d15b8298a98f947e`. The complete MMC Term 3 shared cache now exists. The MMC maintenance/source window was closed immediately after finalisation and shared readers remain disabled pending one combined settled account-usage check. |
| MMC consolidated publication settlement and Creator-reader authorisation, 22 Sep 2026 | All 18 publication requests reconciled to 235 D1 statements, 15,487 rows read and three rows written, with complete request metadata; the largest request read 1,079 rows. The account-wide checker returned `GO` at 143,958 daily reads and 253 writes, with no expensive fingerprint and a maximum five-minute bucket of 60,905 reads, below the 100,000 hard stop. This authorises only Creator-self reads of MMC shared metadata/day objects. Contacts, builders, automatic launch, other facilities, ordinary users and legacy SQL remain disabled. |
| Creator MMC cached-reader canary, 22 Sep 2026 at 20:54 AEST | Creator login followed by MMC On shift and ED Staff completed without user-visible error. Each shared reader used one D1 statement and read three rows with zero writes. The settled account-wide checker returned `GO` at 152,071 reads and 257 writes, with the maximum five-minute bucket unchanged at 60,905. This authorises widening only MMC cached readers to entitled users; automatic launch, contacts, builders, other facilities and legacy SQL remain disabled. |
| DDH/MCH publication window opened, 23 Sep 2026 | A fresh-day second baseline returned `GO` at 141,181 reads and 302 writes, projected 141,763 reads, no expensive fingerprints and only eight additional reads across the sampling interval. Opened advanced maintenance solely for the manually dispatched chunked publisher with exact source allowlist `ddh,mch`. Readers remain MMC-only; contacts, automatic launch and legacy SQL remain disabled. The sequence must stop on the first failed plan, request ceiling or attribution gate and close immediately after finalisation. |
| DDH publication planning stopped, 23 Sep 2026 | The first and only DDH request (workflow `35838159050`, request `a3f83a267c44163e`) returned HTTP 500 with a non-JSON response while planning term start `2026-08-03`. No batch or MCH request was dispatched. The DDH/MCH publication window was closed immediately pending request attribution and diagnosis; MMC readers remain unchanged. |

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
