# VHH historical recovery and roster freshness

Prepared 3 October 2026, Australia/Melbourne. Originally a plan-only document.
The user subsequently approved implementation. VHH's missing historical input
has now been restored and its complete Term 2/3 shared views verified; DDH's
five-minute scheduler is deployed. See
`vhh-history-and-ddh-polling-checkpoint-2026-10-03.md` for release evidence.
The broader all-site fifteen-minute delivery work below is not yet accepted as
complete: SharePoint reconciliation, queue/publication latency and visible-app
revision refresh still need end-to-end verification and any necessary changes.

## Outcomes

1. Recover trustworthy VHH roster history from 4 May 2026, including the missing
   24 August–20 September period, using the downloaded workbook's hidden sheets.
2. Make committed provider changes reach the app's personal calendars,
   subscription-feed responses and affected At a glance views within 15 minutes
   under normal operating conditions. Keep the existing account-wide D1 safeguards.

The clock starts when a provider has saved and exposed a retrievable revision.
The target includes detection, settling, downloading/extraction, queueing,
activation and publication. A five-minute check interval leaves processing time;
a fifteen-minute interval alone cannot meet a fifteen-minute delivery target.
An already-open app also needs a bounded revision refresh. Third-party calendar
clients determine when they fetch a subscription; server/feed freshness cannot
guarantee Apple, Google or Outlook displays the revision within fifteen minutes.

Provider outages, rejected/invalid rosters and exhausted safety budgets must
preserve the previous valid roster and expose a delayed/failed freshness state.
Do not claim the target was met by recording a successful metadata check alone.
Casey remains deferred until an authoritative input exists; its eventual source
must join this same contract rather than receive an unverified polling switch.

## Evidence and current gaps

Current checked-out release is `1374f219`, containing the four-site access and
historical-publication restoration. Prior release acceptance verified Working
together and ED Staff through the actual readers. This plan adds missing input
and delivery guarantees; it does not reopen legacy SQL or global rebuilding.

Local read-only workbook inspection found:

| Workbook | Roster sheets and coverage | Relevant finding |
| --- | --- | --- |
| Downloads/Active Medical Roster (1).xlsx | Ten roster sheets, 20 blocks, 4 May 2026–31 January 2027 | Five hidden sheets cover 4 May–20 September, including `24.08-20.09` |
| Downloads/Active Medical Roster.xlsx | Seven roster sheets, 13 populated blocks, 4 May–1 November 2026 | Earlier copy; 24 August–20 September is still visible |

Recovery candidate: `/Users/rhaydon/Downloads/Active Medical Roster (1).xlsx`,
69,170 bytes, SHA-256
`bcebfcbf599e22cdba2e634cdb0ae9025d6b1fc3a5054a71c000f55f35528043`.
The other copy has SHA-256
`3b0666bbab09a07619850ccc5a3e4c9dacfdbc606e32cc50e3b7f988ff6740fa`.
Names and raw clinical rows are not copied into this plan.

The raw-workbook extractor reads hidden sheets and records their visibility.
`buildVhhDerivedRosterPayload` explicitly skips non-visible blocks. The legacy
Office Script also extracts visible sheets only. Thus hiding completed sheets
can explain missing history even without a failed trigger. The shutdown may
also have contributed; no historical incident cause is asserted from this file.

A memory-only dry run deliberately included hidden blocks without editing either
workbook or the parser. Existing designation/timing interpretation succeeded:

| Requested interval | Header dates | Rostered identities | Parsed events |
| --- | --- | --- | --- |
| Term 2: 4 May–2 August | 91 | 54 | 1,296 |
| 3 August–20 September | 49 | 48 | 669 |

These are feasibility counts, not a production import or a count of changes.
Some August dates are already published and must be reconciled, not duplicated.
The later workbook also differs in current/future content, so downloading it
does not authorise replacing newer live facts with this local copy.

| Site | Last verified mechanism | Remaining freshness work |
| --- | --- | --- |
| DDH | FindMyShift LastModified check inside six-hour maintenance | Restore frequent metadata checking and prompt processing/publication |
| MMC | Enabled SharePoint file-change roster flow | Verify latency and add bounded metadata reconciliation for missed triggers |
| MCH | Enabled SharePoint file-change roster flow | Same, retaining exact current/next-term filename selection |
| VHH | Enabled SharePoint file-change roster flow | Same, plus preserve archived coverage as sheets become hidden |

DDH has the confirmed schedule gap. The other three can receive changes promptly,
but none has a demonstrated end-to-end fifteen-minute guarantee. Teams-hosted
roster files use SharePoint metadata: stable file identity, Modified and
ETag/version are available through the existing connector path. This allows
cheap change checks before retrieving workbook content.

## Batch A: recover VHH history

1. Pin the candidate hash and inspect only recognised Shift Label blocks. Confirm
   real date cells, assignments, ignored teaching/student rows, swing shifts,
   explicit times and overnight endings. Check overlapping sheets/copies for
   conflicts. Treat hidden sheets as archived source material, not absent data.
2. Add an explicitly scoped historical extraction mode; normal live extraction
   must not suddenly ingest every hidden sheet on each autosave. Restrict this
   recovery to 4 May–20 September, split at the 3 August term boundary. Infer
   seniority within each term, avoiding evidence bleeding between terms.
3. Read the exact active VHH file/coverage and published objects through bounded
   indexed/R2 paths. Prepare one dry-run comparison: added, changed, identical
   and conflicting facts; recovered coverage; affected accounts/permissions;
   estimated reads, indexed writes and publication work. Do not scan all users
   or all roster-event history. Obtain a recovery point and pin the active set.
4. Choose the existing bounded import route after inspecting its ownership and
   replacement semantics. Produce complete authoritative term payloads if a
   replacement requires them: use archived cells for historical gaps and the
   latest authoritative retained/live source for dates already current. Preserve
   Term 4 and unrelated sources. Do not feed a partial file to a full-term
   replacement, layer duplicate active facts, or use download time as a newer
   SharePoint provider version. If safe composition is unsupported, implement
   and test a scoped historical contribution path before any production import.
5. Stage Term 2 and the affected Term 3 contribution with immutable provenance,
   term/date boundaries and source hashes. Activate only under account budget
   admission and an unchanged pinned active set. New live input invalidates or
   replans the affected part; it cannot be overwritten by the historical job.
6. Publish affected staff, monthly and daily objects with bounded continuation.
   Preserve every previously valid current/future pointer and existing clinical
   corrections. Historical membership must grant only the relevant past scope,
   never present-day hospital/contact access. No historical contact-sheet import.
7. Verify May 4, June/July boundaries, August 2/3 and August 23/24, September 20/21,
   plus representative overnight shifts. Check Working together, ED Staff,
   personal calendars and feed responses, with restricted and all-site subjects.
   Repeat the same recovery locally to prove zero fact duplication and no-op
   publication. Reconcile measured receipts with settled account analytics.

Keep SharePoint and both downloaded workbooks unchanged. Store any authorised
recovery payload privately with provenance; do not commit raw rosters to Git.
Use an exact recovery manifest, not the general maintenance backlog as an
implicit list of everything to import. A shared maintenance dispatch has broader
scope; reuse existing user authorisation only where it applies to the actual run.

## Batch B: frequent checks and prompt delivery

### Detection and source identity

- DDH: use a bounded Cloudflare scheduled check every five minutes. Query
  LastModified first; download only when the provider revision or required
  current/next-term window needs importing. Keep existing waiting-for-publication,
  invalid-report and unchanged-version safeguards.
- MMC/MCH/VHH: retain working SharePoint change triggers as the fast path. Add
  five-minute metadata reconciliation through the authenticated Microsoft
  connector, reusing existing flows where practical. Check only approved file
  identities and validated current/next-term selectors, at most two files per
  source. Never enumerate historical libraries or broadly match every Paeds file.
- Compare the provider revision with the last successfully activated revision
  and any in-flight revision before downloading. An unchanged tick performs no
  fact/status rewrite and dispatches no parser workflow. Use a small read-only
  admission/change-check response if existing ingress cannot check before bytes.
- A Modified timestamp is a change hint; file identity plus ETag/version and
  the existing content/clinical hashes supply deduplication. Coalesce rapid saves
  to the newest retrievable revision. Recheck metadata around retrieval, with a
  bounded settling period, so an autosave burst cannot create stale imports.
- Do not make Excel modify its own trigger workbook. Keep extraction read-only
  and retain small JSON extraction where proven; changes to contact flows are
  outside this batch. Their five-minute schedule stays independent.

### Processing, publication and recovery

- Reuse bounded staging, small-correction diffs, stale-provider rejection, atomic
  activation and affected-date publication. Do not fan out rebuilds to every
  staff account. Shared publication plus account-specific reads should update
  all staff whose claims refer to the affected source.
- A frequent bounded reconciler must detect stuck queued/processing/publication
  work and resume it promptly. The six-hour maintenance job remains suitable for
  slow housekeeping; it cannot be the normal continuation path for this target.
- Measure the existing shared `monash-roster-sync` concurrency queue first.
  A 35-minute historical maintenance run can block a normal update beyond the
  target. Separate historical work from routine delivery and use tested atomic
  source/file leases with account-wide reservations before allowing concurrency.
  Source/term dependencies and active-set fences still prevent overlapping work.
- Do not merely turn on the old global watchdog or change the six-hour workflow
  to fifteen minutes. Its source-agnostic kick is incompatible with exact-source
  controls, and detection intervals omit processing time.
- Prefer Cloudflare scheduled orchestration for the deadline; GitHub scheduled
  workflows can be delayed or dropped. Existing dispatched GitHub parsing also
  has queue/start latency. Benchmark it rather than promise a hard guarantee.
  If the measured queue and parsing path cannot meet the budget, redesign only
  that execution stage with a bounded promptly available runner before declaring
  the requirement fulfilled. Do not silently introduce paid infrastructure.
- Retain archived VHH coverage when a previously visible sheet becomes hidden.
  Routine imports own only their declared source/term/date ranges; hiding history
  is not an instruction to delete it. Archived corrections require an explicit
  scoped repair or a separately reviewed per-term content change mechanism.

### User-visible freshness

Suggested acceptance budget: detection/reconciliation within five minutes,
settling plus processing/activation/publication within eight more, and active
app refresh within the remaining two. Measure from provider modification through
the final readable revision; overlapping stages may reduce the total.

Add or verify a shared, cheap published revision check for visible calendars,
including Creator-entered accounts. Coalesce tabs/requests, pause hidden tabs,
check on foreground return, and fetch only the selected account when its relevant
revision changes. Check permission expiry before serving cached details. Do not
poll D1 for every staff member or rerun identity discovery on each refresh.

Feed requests should immediately return the activated valid data. Test a real
subscribed client separately and report its observed refresh interval. For
unpublished future terms, record waiting for publication and preserve the agreed
14-day visibility window; freshness does not override publication/access policy.

### Cost, acceptance and rollout

Five-minute checking is 288 cycles/day: four sources imply 1,152 source checks
before current/next-file multiplicity. Benchmark exact request/read costs.
Metadata-only checks must remain tiny indexed reads or provider/R2 reads, with
no repeated source-status writes. Never download/parse whole rosters every tick.

Keep the account ceilings of 5,000,000 reads/100,000 writes per UTC day and
existing 4,000,000-read/80,000-write maintenance admission with ordinary-service
headroom. Include request quotas, Power Automate runs, GitHub minutes, R2 and
provider throttling in the estimate. Reuse budget grants until their valid
refresh deadline; do not run full analytics independently on every source tick.
Reservations, measured settlement and independent source stop controls remain.

Validate the complete batch locally: unchanged ticks, rapid saves, stale/later
arrival, equal timestamps with changed version, failures/retries, duplicate flow
delivery, concurrent sites, missed trigger, busy history job, quota deferral,
term rollover, hidden-sheet rollover and permission/cache boundaries. Reuse the
focused import, source isolation, queue, publication, clinician access, client
request and budget suites; repair related failures if the touched path exposes one.

Release the tested batch once, verify effective settings/flows, then check one
genuinely needed changed delivery and an unchanged cycle at each available site.
Timestamp provider exposure, detection, queue start, activation, publication and
visible refresh. Verify a second missed-trigger recovery case and normal bursts
without artificial edits to clinical rosters. Pause only a failing source.
Observe actual traffic across a full UTC quota day and report deadline misses
as delayed, not successful freshness. Update the restoration register/checklist
with the measured limits and exact remaining provider/client dependencies.

## Implementation effort and next handoff

Extra High is unnecessary now. Low is suitable for the next bounded stage:
VHH extraction/recovery preparation, source/flow metadata audit and routine
verification against this plan. Use Medium for the scheduler/concurrency and
client-refresh implementation batch: it changes competing writers and an
end-to-end freshness contract. Return to Low for measured acceptance and rollout.
This is task-specific judgement; lower effort generally favours speed and token
usage, but tests and explicit limits remain necessary at either setting.

Execution order: Batch A local preparation and dry-run comparison; finish its
scoped recovery under admitted production costs; Batch B source audit and local
implementation; coordinated release and measured acceptance. No fresh planning
cycle is needed unless source provenance or runner latency changes the design.
Pause after delivering this plan until the user resumes implementation.

## References

- Repository: `scripts/vhh-roster-workbook.mjs`, `functions/_lib/vhh-roster.js`,
  `.github/workflows/roster-maintenance.yml`, `scripts/process-roster-maintenance.mjs`,
  `functions/api/automation/findmyshift-check.js`, and the 3 October restoration
  checkpoint. Older VHH documents contain superseded Preview/contact restrictions.
- [SharePoint connector metadata and triggers](https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers).
- [GitHub schedule delay limitations](https://docs.github.com/en/actions/reference/workflows-and-actions/events-that-trigger-workflows#schedule).
- [Cloudflare Cron Triggers](https://developers.cloudflare.com/workers/configuration/cron-triggers/).
- [Official OpenAI reasoning-effort guidance](https://developers.openai.com/api/docs/guides/reasoning).
