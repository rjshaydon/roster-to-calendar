# At a glance: Sol implementation handoff and restoration checklist

## Scope and reading order

Prepared 5 September 2026 from local `aa9eed8`, emergency commits `cff238d` and `8d0cf60`, and accessible incident messages in **Review doctor identity UX plan** (`01a05fe0-8683-7732-b01d-667b5060fc1e`). The four newest turns returned no contents. Cloudflare, GitHub and Power Automate live settings have **not** been inspected in this review. “Confirmed” below means confirmed in repository code/history; the conversation's reported deployed state is historical evidence, not current verification.

Read [the implementation plan](./at-a-glance-d1-optimization-plan.md) first, then this document. The plan owns architecture, product rules and acceptance gates; this document owns implementation handoff and restoration tracking. Read source code rather than assuming old setup documents or another task's completion report describe the current checkout.

The current request authorises planning only. Subsequent implementation should complete the local code and tests through plan Phases 0–6 and produce a concrete rollout package. Production changes remain subject to the plan's separate rollout approval. Nothing in this checklist calls for turning automation on now.

## Instructions for Sol

1. Inspect branch, untracked documents, applicable repository instructions and deployed-baseline evidence. Preserve the user's local databases, test records, unrelated work and unfinished identity branch. Include these two currently untracked plan documents in the eventual implementation changes so they are not lost on a branch switch.
2. Implement Phase 0 first. The `aa9eed8` `package.json` does **not** contain the `local:*`, `dev:local`, `test:local-isolation` or `test:database-costs` commands reported on the identity branch. The conversation identifies local-safety commits `b13270d` and `ea1c749`, and identity cost-test commit `1e32e48`. Inspect their diffs if available; selectively reuse relevant infrastructure without importing unfinished identity behaviour or undoing `aa9eed8`. Otherwise implement the equivalent isolated harness here. Do not run a reset/setup script against the user's existing local state.
3. Work in the plan's numbered phases: local safety/baseline → incremental ingestion and summaries → access/session handling → shared Staff/catalogue → shared day views → contacts → browser/range views. Keep new live paths disabled until their gates pass. Implement source-specific rollout controls before any restoration.
4. Before substantial edits, record concrete choices for stable source/shift identity, schema/indexes, canonical hashes, publication ownership/recovery, overlay compatibility, access revocation and resource budgets. Use existing naming and parser contracts. Ordinary technical decisions do not require asking the user; only unresolved product behaviour should block its dependent work.
5. Treat “only changed data” as a requirement for **D1 events, daily presence, membership and shared cache objects**, not only response caching. Diff complete incoming extracts where necessary. Identical imports perform no data rewrite; a single sickness correction must not cause a whole-file delete/reinsert. Include index write costs and removed shifts.
6. Every import, promotion, deletion, overlap trim, reparse and repair must go through the change/publication pipeline. New shared GET endpoints must bypass repeated password/term-event work through the common authorised session path. Cache misses do not invoke old SQL or build on a reader.
7. Add end-to-end tests for the actual handler call chain, not just the new repository helper. Keep tests of paused behaviour and permanent schema safety. Extend the existing regression scripts with local large-fixture cost estimates, clock boundaries, failure injection and concurrency cases.
8. At each phase, record changed files, tests run, measured/estimated costs, unresolved limitations and whether its gate passed. Add a concise implementation-status document rather than presenting partial work as complete. No feature toggle is a substitute for wiring every entry point correctly.
9. Before rollout, deliver exact commit/deployment targets, migration/backfill cost and batches, source allowlist, queue reconciliation steps, before/after configuration, monitoring thresholds and a tested pause/rollback procedure. Fill in the restoration ledger below. A successful Pages deployment does not update the independent watchdog.

## Product requirements already resolved

- Share ED/date and ED/term data centrally across authorised accounts; preserve account-specific permissions and preferences.
- DDH uses FindMyShift; MMC, MCH and VHH use Excel. Casey automatic SharePoint access is pending and is not a restoration prerequisite.
- Pre-term drafts may change repeatedly. Make a whole term visible in At a glance from 14 Melbourne calendar days before its actual start. Prepare data beforehand if available; the clock boundary changes visibility without a rebuild or another import. Current and next terms coexist.
- Later changes are normally swaps, sickness or emergency leave; do not reject legitimate larger corrections or lock “finalised” data.
- Contact polling interval and likely simultaneous viewer-hours are still unanswered. The plan uses a provisional 60-second interval for capacity design, not an agreed clinical freshness promise. Do not restore ten-second polling by default.
- Preserve current continuing-SMS rules. Offline access beyond valid authorisation and the proposed revocation/staleness bounds remain product decisions before those behaviours are enabled.

## Confirmed temporary shutdowns to restore selectively

| ID | Control and evidence | What is stopped | Required restoration and proof |
| --- | --- | --- | --- |
| R1 | `wrangler.toml`: `ROSTER_AUTOMATION_WRITES_ENABLED = "false"`; added by `cff238d`, extended by `8d0cf60` | Automated ingress, derived saving, dispatch, DDH checks, VHH extraction; also manual roster actions listed below | After incremental-write and publication gates pass, set the effective **production Pages** value explicitly to `"true"` in the approved release. Retain the guard code. Before this, implement/test a source allowlist and separate maintenance permissions so the global switch cannot expose unconverted routes. Verify deployed version and effective setting, then one accepted changed import and one unchanged repeat. |
| R2 | `wrangler.roster-watchdog.toml`: `ROSTER_AUTOMATION_ENABLED = "false"`; `worker/roster-queue-watchdog.js` returns before scheduling either request | Both queue recovery dispatch and DDH FindMyShift polling | Restore last, on the **separately deployed Worker**, after Pages/queue/provider gates pass. Set explicitly to `"true"`; verify its own version and `/health` reports `configured: true`, `paused: false`. Observe one scheduled cycle, both downstream results and bounded usage; health alone is insufficient. |
| R3 | `rosterWritesExplicitlyPaused` checks in `functions/api/state.js`, introduced by `8d0cf60` | Creator/manual import, removal, reset and repair actions | Convert and test each route before it is exposed. Routine upload can resume with R1 plus route permissions; heavy rebuild/repair stays separately controlled and requires a scoped job budget. Do not remove pause checks to regain UI functionality. |

**Flag detail:** automatic writes require an explicit truthy setting, but the current manual pause helper blocks only the literal `"false"`. Deleting the variable can therefore leave automation paused while unblocking manual writes. Use explicit values and test missing/false/true behaviour; do not restore by removing the variable.

R1 currently gates these five POST endpoints:

- `/api/automation/ingest`
- `/api/automation/derived`
- `/api/automation/dispatch`
- `/api/automation/findmyshift-check`
- `/api/automation/vhh-roster-extract`

R3 covers these `/api/state` actions:

- `syncRosterRepository`, `removeRosterImports`, `saveDerivedCalendarFile`, `uploadRawRosterFile`
- `resetDerivedCalendarFile`, `replaceActiveRosterFiles`, `repairRosterDailyPresence`
- The ordinary save path when `removedImportIds` is non-empty.

The global flag currently couples routine imports and advanced maintenance. Implement explicit per-source and maintenance gates, disabled by default, for the staged restoration below. Choose and document actual names during coding; no such separate controls are claimed to exist yet. Test that disallowed work stops before expensive repository operations. Authentication itself may still have a small accounted cost on manual actions.

## Dependencies to verify, not blindly “turn back on”

| ID | Surface | Repository evidence and required check |
| --- | --- | --- |
| V1 | GitHub `Process Monash roster queue` in `.github/workflows/monash-roster-sync.yml` | It has `workflow_dispatch`, **no scheduled trigger**, and serial concurrency. No repository evidence shows it was disabled in GitHub. Verify its enabled state, running jobs, workflow ref and token configuration; re-enable only if actually disabled. Preserve concurrency and terminal-failure reporting; do not add a second cron. |
| V2 | Queue and raw-source endpoints | `/api/automation/pending` is **not guarded by R1** in `aa9eed8`; `listQueuedRosterSyncRuns` also calls `supersedeObsoleteQueuedRosterSyncRuns`, which can write. The workflow calls pending before derived saving. Add pause coverage and move reconciliation to bounded mutation work; audit `/raw` and other automation routes too. A paused workflow must not keep reading or modifying D1 just because derived ingestion is blocked. |
| V3 | Existing SharePoint/Power Automate roster flows | Sources are `monash-adults` (MMC), `monash-paeds` (MCH), and `vhh-active-medical-roster` (VHH), per `automation-import.js`. Their external on/off state and connector retry backlog are unknown. Check existing flows, destination, trigger conditions and retries. Restore only flows actually paused, one source at a time; use their current supported transport, not a newly invented replacement flow. |
| V4 | DDH FindMyShift | Source `dandenong-findmyshift`; watchdog and Pages switches both matter. Verify configured API key/team and range overrides without exposing values. An unchanged provider check must not download/reimport a full roster, and checks must cover current-term corrections even during next-term preparation. The current implementation selects a term range; test this transition explicitly. The provider's early-fetch window is separate from the 14-day At a glance visibility rule. |
| V5 | Contact extracts/Power Automate | `/api/automation/contact-list-extract` is **not blocked by R1** in this checkout. Do not assume contacts were paused with rosters. Verify existing MMC/DDH clinician-extract flows and freshness independently. If a flow was externally disabled, restore it only after its overlay, expiry, deduplication and cost tests pass. |
| V6 | Production versus preview bindings | Verify effective Pages variables, D1/R2 bindings, GitHub workflow target/ref and base URL independently for each environment. The repository has preview-specific `VHH` settings; do not copy those into production. Leave preview automation paused unless separately needed and approved; preview usage still matters to the account budget. |
| V7 | Secrets and dispatch routing | Check presence/validity and destinations of `ROSTER_AUTOMATION_TOKEN`, `ROSTER_WATCHDOG_TOKEN`, `GITHUB_ACTIONS_TOKEN`, GitHub production/preview tokens and relevant provider-specific tokens. There is no evidence all were removed. Do not recreate or rotate working secrets as a blanket restoration step, log them, or dispatch work merely to discover configuration. |

The watchdog cron still exists at `*/15 * * * *` in the repository. Restore the enabled flag, not an additional schedule. `docs/roster-automation.md` contains an older five-minute description as well as a fifteen-minute DDH description; update that documentation to the chosen, tested fifteen-minute baseline during implementation rather than increasing frequency accidentally.

External flow status may require the user's tenant access at rollout time. Record any such verification as pending, not complete or assumed enabled. This planning audit did not call external control planes or send messages to other tasks.

## Permanent fixes and safeguards to keep

- **Never revert `aa9eed8`.** Ordinary requests must not inspect/repair schema or execute DDL. Retain explicit migrations and the zero-database-call schema regression. Do not revert `cff238d` or `8d0cf60` wholesale either; change approved settings while retaining guard behaviour and the invite-schema fix.
- Keep bounded doctor profile loads (`230e513`), removal of switcher snapshot fan-out (`549c427`), scoped revisions and query reductions. These are performance fixes, not services to resume. Replace an implementation only with a tested equal-or-better path.
- Keep the old full-workbook contact routes (`contact-list` and `contact-list-binary`) returning `410 Gone`; clinician-only JSON is intentional, not an incident shutdown to reverse.
- Keep local network isolation, missing-configuration fail-closed behaviour, explicit migration policy, queue deduplication, pause controls and bounded retries. Local execution must never be “restored” to production connectivity.
- Do not revive background identity auditing or deploy unfinished identity work as part of roster restoration. The conversation reports the audit was not running in production; no current production switch to restore was found. This is a separate workstream.
- Do not restore the ten-second browser poller, full-year event scans, runtime broad Staff/coverage fallbacks or whole-roster rewrites.
- Casey automatic integration is new work, not an automation disabled by the incident. Manual Casey support should follow the tested manual-import path.

## Ordered restoration runbook — only after implementation and rollout approval

1. **Capture state and budgets.** Record deployed Pages/Worker versions, effective flags and source settings, other account usage, external flow states and current UTC-day headroom. Keep incoming roster flows and the watchdog paused while preparing the release. Do not change a live external flow without the rollout's authorisation.
2. **Deploy the tested code/schema with switches off.** Confirm the schema/no-DDL guard and zero-work automation pauses still hold. Verify new source/maintenance gates. Deploy compatible queue-runner code too, and record the exact ref used by dispatch; old in-flight GitHub jobs must not submit incompatible payloads to the new receiver.
3. **Reconcile the backlog in declared batches.** Inspect retained input, queued/processing/failed runs, dispatch leases and external retries. Keep the newest valid source revision for each replacement scope, including separate current/next terms and non-overlapping files. Do not replay every missed revision or blindly empty the queue. Mark obsolete work with provenance; handle partial jobs via the new recovery protocol. Obtain a fresh extract where paused HTTP ingress never retained one. Bound cleanup writes as well as import writes.
4. **Enable one source and run a controlled sample.** With the watchdog and external automatic triggers still off, enable R1 behind the tested source allowlist. If accepted ingress automatically dispatches GitHub, that is part of this controlled sample: do not also start duplicate runs. Process one current source version within budget, check roster/staff/contacts/access parity, and resend unchanged content to prove no event/presence rewrites or cache rebuild. Confirm edits do not break normal calendar feeds and account snapshots.
5. **Restore ordinary manual actions.** Enable converted upload/save/removal routes after their individual tests. Keep advanced replace/reset/repair behind independent controls, existing confirmations and scoped job budgets. Do not run a maintenance job just to mark its button “restored”.
6. **Restore external source flows one at a time.** Only if actually paused, resume the matching MMC/MCH/VHH flows after their canary and correct destination are verified. Include retry-queue handling. Bring DDH checks through a controlled unchanged/changed cycle. Monitor each source before expanding the allowlist; source order should favour the smallest representative import.
7. **Verify or restore contacts separately.** Preserve working contact flows. For any actually paused flow, resume with current extracts, safe correction matching and expiry behaviour; do not replay expired contact history. Verify roster visibility timing never causes future contact assignments to be displayed as current.
8. **Enable the independent watchdog last.** Set R2 to true in its own approved deployment, preserve the existing fifteen-minute cron, verify `/health`, then inspect a scheduled queue check and DDH check. An empty queue and unchanged DDH roster should not create imports or workflow storms.
9. **Observe and close the ledger.** Apply the plan's seven-UTC-day observation gate including realistic changed/unchanged imports and contact cycles. Record actual read/write/request/object usage and publication lag. Mark restoration complete only for entries with evidence; explicitly list anything deliberately left paused and why.

### If usage rises or a publication fails

Pause R2 and the affected ingress/source controls; set R1 false as the broad emergency stop if necessary. Check running GitHub work and connector retries rather than assuming disabling future schedules stops in-flight requests. Workers already executing may finish a bounded batch: implementation must check pause state at safe checkpoints and preserve resumable state. Keep valid published snapshots and authentication available, subject to existing access freshness rules. Do not roll back to the expensive old reader or delete data to regain quota. Account for the time/usage needed to apply the stop itself.

## Required handoff evidence and restoration ledger

Existing local regression commands available at `aa9eed8` are `npm run check`, `test:d1-quota`, `test:facility-access`, `test:fixtures`, `test:contact-allocations`, `test:contact-sync`, `test:contact-workbook`, `test:vhh-automation` and `test:queue-failure` (each test name also prefixed with `npm run`). Inspect each script before running; extend with the plan's isolated integration/cost tests. Their current existence does not establish that the new design is implemented or safe. Do not run `shadow:active` as an ordinary local test; review its remote behaviour and scope separately.

Deliver a gate report containing: source and scale of local fixtures; request/SQL call counts and query plans; estimated D1 reads/writes including indexes; unchanged/one-shift/whole-new-term cases; R2 and Worker request projections; auth and clock-boundary parity; crash/retry tests; paused-path tests; current/new-term DDH refresh; and the exact commands/results. Distinguish estimates from eventual approved production measurements.

Fill this in as rollout work actually happens; never pre-tick it from configuration text:

| Item | Before/live verification | Approved change/version | Correctness + usage evidence | State / outstanding work |
| --- | --- | --- | --- | --- |
| R1 Pages writes and per-source gate | Pending | Pending | Pending | Paused in repository |
| R2 independent watchdog | Pending | Pending | Pending | Paused in repository |
| R3 routine manual actions | Pending | Pending | Pending | Globally paused in repository |
| R3 advanced maintenance | Pending | Pending | Pending | Keep controlled; no automatic replay |
| V1/V2 queue runner, pending/raw, backlog | Pending | Pending | Pending | Audit/fix before resume |
| V3 MMC/MCH/VHH external flows, each recorded separately | Unknown | Pending if needed | Pending | Verify tenant state |
| V4 DDH provider checks | Pending | Pending | Pending | Both gates required |
| V5 contact flows | Unknown | Pending if needed | Pending | Not gated by roster pause |
| V6/V7 bindings, refs and secret presence | Pending | Pending if needed | Pending | Production and preview separate |

## Suggested implementation prompt

> Implement `docs/at-a-glance-d1-optimization-plan.md`, using `docs/at-a-glance-sol-handoff-and-restoration.md` as the implementation and restoration checklist. Start with the actual current repository state and Phase 0; preserve local data and unfinished identity work. Complete the locally authorised phases and tests, including incremental database writes, shared immutable caches, recoverable publication, time-aware access/visibility, contact overlays and comprehensive pause coverage. Record passed gates and measured costs. Prepare a concrete production rollout/restoration package with exact versions and remaining product questions. Do not deploy, run remote migrations/backfills, or re-enable production automation until that rollout is separately approved. Keep all permanent quota fixes and guards.
