# At a glance restoration completion plan

> Completed on 3 October 2026: migration 0038, the four-site application release
> and all retained historical publications. Casey Terms 3 and 4 are deferred by
> the user. Earlier pending steps below are historical. See
> [the final checkpoint](at-a-glance-restoration-checkpoint-2026-10-03.md#final-release-and-verification--1620-aest)
> for deployment, acceptance, quota evidence and remaining limitations.

Prepared 3 October 2026 from clinician-access commit `3a3a1b8a` on
`codex/clinician-access-history`. This is the additional restoration plan for
the work now being coordinated in this conversation. It builds on the access
implementation; it does not replace or repeat it.

This request is for a plan. No application code, deployment configuration or
production data is changed by writing it. The separately scheduled 12:20 pm
Melbourne-time quota/migration-readiness check remains applicable. Its outcome
does not itself release unfinished application changes.

## Outcome and scope

Restore reliable Working together and ED Staff searches across the hospitals
the viewer is authorised to see, preserving the new SMS/CMO, current-term and
historical access rules. An unavailable hospital or missing period must not
break an otherwise valid combined search or be silently presented as complete.

Include the shared hospital selection, publication coverage, browser cache and
release checks that affect Metadata, On shift, By stream and related colleague
tools. Their existing working behaviour must survive this restoration.

Advanced administration (global reconciliation, rebuilding/resetting and broad
daily-presence repair) remains a separate workstream. Durable Doctor Names and
identity merging remain the parked product project. Casey automatic SharePoint
integration is also separate new work; existing/manual Casey roster support can
be restored through the bounded publication path without that integration.

## Starting position

| Area | State in the saved branch | Remaining work |
| --- | --- | --- |
| CMO/SMS visibility | Both receive all-site scope when At a glance is enabled. | Verify deployed behaviour during the combined release. |
| Multiple hospitals | Current active term memberships supply an authorised set, including verified locums. | Make combined requests tolerate authorised sites whose shared reader is unavailable. |
| Historical Working together | Server resolves hospital/date segments and a scoped staff directory; date-first searching is implemented. | Verify retained evidence and publication coverage; make partial/unavailable results accurate. |
| Working together hospital selection | All-site default uses enabled reader sources. Trainee membership sets and explicit requests may still include Casey. | Apply one consistent availability contract and preserve explicit-site semantics. |
| ED Staff, All hospitals | Server still expands All to a fixed five-source list, including Casey. | Derive its read set from authorised scope and enabled readers. |
| Production configuration in the repository | Shared readers allow MMC, DDH, MCH and VHH; Casey is excluded. Legacy reads are paused. | Confirm effective deployed flags and source sets before rollout. |
| Access migration | `0038_clinician_access_scope.sql` is committed, not applied. | Complete quota admission, schema/cost/recovery preflight and isolated migration. |

The screenshot describes the earlier limitations. Do not treat its CMO,
multi-hospital or historical-policy rows as unimplemented in this branch.
Repository flags describe the checked-in configuration, not proof of live state.

## 1. Make authorisation and availability explicit

Implement a shared server decision for each action, subject and searched period:

1. Determine authorised hospitals or hospital/date segments using the existing
   access policy. An explicit unauthorised hospital remains denied, before any
   roster/publication read. Neither rollout flags nor a staff comparison can
   expand that entitlement.
2. Determine operational availability using the reader cohort, emergency pause,
   source allowlist and action-specific metadata/day switches.
3. For All, read the authorised, enabled subset. Record authorised hospitals or
   segments excluded by rollout separately from missing published data. Never
   pass a mixed enabled/disabled set to the all-or-nothing reader gate.
4. For an explicitly chosen authorised but disabled hospital, show unavailable
   for that hospital. Do not substitute a different hospital or broaden the
   query to All. If no authorised reader is available, return an availability
   state with no roster rows rather than claiming there were no shifts.

Choose the smallest response extension that can express available, partial,
preparing and unavailable states, with hospital/date coverage and reasons.
Return such information only within the viewer's authorised scope. A trainee
must not discover other hospitals' staff or coverage through these messages.
Fail closed for missing configuration and preserve the emergency pause.

Use this decision for ED Staff and both Working together routes. Review the
same hospital-list construction in Metadata, On shift, By stream and related
Who/When/overlap readers so another All selector cannot retain the mismatch.
Keep contact source controls independent and current-site scoped.

## 2. Restore selectors and honest results

Have the server supply action/period availability. Avoid deriving SMS/CMO
hospital choices from only the selected doctor's sources, local roster files,
or an old browser preview. Do not reduce the underlying access scope when a
hospital is temporarily unavailable.

- ED Staff All returns staff from every enabled, authorised hospital for the
  permitted term, with a visible notice for any omitted authorised hospital.
  Restricted clinicians remain limited to the current term.
- Working together derives its hospital choices and staff directory from the
  freshly authorised historical context. Preserve date-first lookup, selected
  colleague comparison and searches by people with historical-only access.
- Preserve existing All EDs / All my hospitals wording, adding a plain-language
  coverage notice when results are partial. Unavailable authorised sites can be
  labelled unavailable; never make a partial result look like all-site coverage.
- Distinguish no shifts in complete coverage from missing roster history,
  publication in progress, rollout unavailability and permission denial.
- Revalidate staff/site selections after date, subject or permission changes.
  Clear old directory/results when their scope is no longer valid. Ignore
  cancelled or late responses from a previous selection.

Availability affects cached contents. Include the effective source set,
availability/coverage revision and publication revisions in cache validation,
alongside the existing subject and entitlement revision. A cached complete
answer must not survive a removed reader or changed coverage. Historical data
still requires fresh period authorisation before reuse.

## 3. Verify historical coverage and restore only recoverable gaps

Audit coverage in a bounded way once admitted by the quota policy. Start with
known source/term manifests and pointers; inspect only exact roster identities,
files and periods needed to explain a gap. Record entitlement evidence, staff
publication, shift partitions and first-day overnight coverage separately.

The current range reader can omit a missing month pointer while returning a
non-preparing result. Multi-source staff loading can likewise omit a hospital
without declaring partial coverage. Correct those completeness checks before
using their results as proof that historical searches are restored.

Propagate missing hospital/term/month/date coverage through both the context
and result APIs. Return available authorised results with a coverage notice;
when all requested coverage is absent, display unavailable history. Do not
cache an incomplete response as a complete unchanged result. Preserve term
publication visibility rules and overnight/DST clipping.

If retained active compact facts and source data support a missing historical
publication, prepare a bounded repair for that exact source/file/term or month.
Use the existing publication/continuation mechanisms with cost estimates,
idempotency and a recovery point. Handle corrected/replaced files according to
their actual provenance; do not reactivate obsolete files merely to grant
historical access. Missing trustworthy entitlement evidence remains denied;
missing recoverable shift data remains explicitly unavailable.

Historical repair is conditional on the audit. Do not assume all older terms
were retained, run a global backfill, or scan every doctor's event history.

## 4. Treat Casey as a separate availability step within restoration

First restore combined searches for enabled MMC/DDH/MCH/VHH readers, reporting
Casey as unavailable only to viewers authorised for it. This is a partial
restoration milestone, not completion of five-hospital visibility.

Then inspect Casey's existing/manual input and compact/publication readiness.
If valid data is retained, prepare bounded publication and read verification
using the established manual-import/publication path. Before adding Casey to
the relevant reader/build allowlists, demonstrate source/term parity, scoped
access, coverage completeness and acceptable costs. If input is missing, record
exactly which roster/period is needed; enabling a reader cannot create it.

Adding a roster reader does not automatically enable Casey contacts or new
provider automation. Keep those independent switches and data requirements.
Finish Casey when its data and budget permit; otherwise report the remaining
limitation explicitly in the release record and user-facing results.

## 5. Validate the combined change locally

Extend behavioural API/browser tests rather than relying on source-text checks:

- SMS and CMO All searches succeed with four enabled readers and Casey disabled;
  single-site and combined ED Staff results agree for those readers.
- MMC/DDH trainees see their authorised set only. An authorised MMC/Casey
  trainee receives MMC results with an accurate partial notice; a Casey-only
  trainee sees unavailable. Explicit unauthorised hospitals are denied without
  published roster reads; explicit authorised/disabled hospitals never fall back.
- Historical rotations, cross-term ranges, locums and historical-only accounts
  retain the committed access policy. Current ED Staff cannot use old access.
- Missing manifests, staff objects, month pointers/objects and overnight data
  produce accurate partial/unavailable coverage, including all-missing cases.
  A complete empty roster remains distinguishable from missing data.
- Revocation, grade/term changes, Creator subject switching, reader removal,
  emergency pause and late responses cannot reuse stale names or results.
- Restoring Casey expands available results only within existing entitlement.
  Normal On shift, By stream, contacts, personal calendars and subscriptions
  retain their expected behaviour.
- Invalid/oversized requests remain bounded. Cached reads and explicit searches
  obey the statement/request/account budgets; no polling, legacy fallback,
  login history scan or automatic global repair is introduced.

Run `test:clinician-access`, `test:facility-access`, `test:facility-rollout`,
`test:facility-snapshots`, `test:cached-roster-insights`,
`test:facility-contact-access`, `test:d1-account-budget` and
`test:client-request-budget`, plus syntax and diff checks. Exercise the actual
publication readers with representative synthetic coverage gaps. Run import,
publication and maintenance checks if their paths change.

The saved baseline has unrelated failures in fixtures, facility maintenance and
facility materialization, documented in the access plan. Reassess those failures
if this restoration touches the affected path; do not waive a release-critical
failure merely because it also existed before the change.

## 6. Migrate and release as one coordinated batch

1. At or after 12:20 pm on 3 October, rerun the account-wide settled D1 checker
   with the saved first sample. Follow the existing release preflight. Time alone
   does not establish GO; estimates must include concurrent normal traffic,
   outstanding work and proposed migration/publication costs.
2. Confirm live schema and migration ledger, index/trigger presence and a
   recovery point. Rehearse migration 0038 locally, then apply only its reviewed
   SQL if admitted and compatible. Avoid the normal command accidentally applying
   other pending migrations. Verify schema/ledger and usage afterwards. The
   migration can be completed independently of the unfinished restoration code.
3. Prepare the complete code/configuration diff and local results for the
   combined release. Keep production legacy reads paused. Use isolated local
   fixtures or a deliberately configured Preview; current Preview is paused and
   must not be mistaken for a working acceptance environment.
4. Release a reviewed, tested commit after migration readiness. Do not deploy
   the shared checkout if it contains unrelated or unfinished edits. Verify with
   controlled enabled SMS, CMO, trainee/multi-site and historical-only subjects.
   Check All, single-site, date-first, comparison and missing-history states.
5. Record effective flags, coverage limits, deployment/commit, schema revision
   and measured usage. Use existing request attribution and settled account
   analytics, allowing for their delay. Disable/revert the failing new reader
   behaviour if costs or results regress; preserve valid roster data, the
   additive schema and the legacy-read pause.

Continue on `codex/clinician-access-history` for this coordinated work. No new
branch is required. Retain the pushed access commit as the known checkpoint and
the handoff note as documentation; avoid concurrent edits to this checkout from
another conversation.

## Completion criteria

The enabled-site milestone is complete when ED Staff and Working together
combined requests work across all enabled authorised sites, report omissions
accurately, preserve the new privacy policy and pass the combined local/live
checks within budget. Full hospital restoration additionally requires Casey's
verified publication and reader enablement. Historical restoration is reported
against verified retained periods, with any unrecoverable gaps documented.

Supporting documents:

- `docs/at-a-glance-clinician-access-plan.md`
- `docs/clinician-access-release-preflight-2026-10-03.md`
- `docs/clinician-access-handoff.md`
- `docs/full-service-restoration-plan.md` (latest checkpoint first)
- `docs/at-a-glance-sol-handoff-and-restoration.md` (historical incident context)


## Implementation checkpoint — 3 October, afternoon

The additional restoration code has now been implemented and validated on
`codex/clinician-access-history`. Migration 0038 is applied in production;
the application changes remain undeployed. The detailed release record is
`at-a-glance-restoration-checkpoint-2026-10-03.md`.

Combined searches retain enabled authorised hospitals and explicitly identify
missing publication. Explicit hospital requests and emergency pause gates remain
strict. Publication completeness and browser cache changes are implemented.
Historical roster data exists for earlier terms at MMC, DDH and MCH, but its
published manifests mostly cover only the current term. A DDH May-term pilot
is queued; broader historical publication awaits pilot verification. Casey has
no published manifest or retained term coverage, so it requires roster input
before its reader can safely be enabled.

Starting the general maintenance workflow manually was rejected by automatic
approval review because it may process unrelated pending production jobs.
The workflow remains unstarted; its existing scheduled runs may process the
queued pilot. Advanced administration and Doctor Names merging remain outside
this restoration release.
