# Restoration status — 28 October 2026

This is the consolidated status as audited on 28 October. Dated plans and
checkpoints retain historical evidence; their earlier pending statements are
not current release gates when superseded by later acceptance below.

## Completed and deployed

- **Batches 1, 2 and 6:** accepted as complete. General password migration,
  secure sessions, optional passkeys, identity requests and private admin support
  are deployed. Password migration still occurs as existing users sign in; batch
  completion does not mean every dormant account has migrated.
- **All EDs staff-history search (PR 34):** merged and activated on 20 October.
  Both production history flags were verified true on 28 October. The published
  catalogue contained 710 people, prepared on 27 October; Casey was unavailable.
  Authenticated release acceptance and subsequent grade/grouping/search-retention
  acceptance are recorded in the linked checkpoints. Do not repeat initial
  seeding or activation based on the obsolete PR warning about 907 writes.
- MMC Term 4 publication and its 2 November On shift visual check completed.
  VHH historical coverage recovery and the 22 October workflow reliability fixes
  are deployed. DDH Clinical Support grouping was corrected on 28 October.

## Remaining work, in the agreed order

| Batch | Remaining scope |
| --- | --- |
| **4 — recovery/admin** | Scoped repair/reconciliation/bootstrap tools with dry runs, budget estimates, resumable execution and recovery controls. Existing handoff, checker and manual-import optimisations are deployed. |
| **5 — resilience/diagnostics** | Broader coordinated background recovery and sampled/batched persistent UI diagnostics. Existing bounded retries and safeguards remain active. |
| **8 — polish/consolidation** | Reported iPhone bottom-bar positioning issue, legacy test repairs, remaining polish and documentation consolidation. |
| **3 — deferred** | Casey automatic roster integration and richer contacts. |
| **7 — deferred** | Director/reporting refinements; requirements remain to be agreed. |

Keep the original numbering: priority is **4 → 5 → 8**, with 3 and 7 deferred.
Advanced global repair/bootstrap remains closed by design; completion requires
safe scoped tools, not simply enabling the old broad controls.

## Observations and evidence limitations

- Observe manual upload → automatic takeover and actual manual-import costs at
  the next genuinely required upload. Do not manufacture a clinical replacement.
- Patrick Tan's MMC Term 2 evidence and seven historical DDH grade summaries
  remain missing or inconclusive. Do not invent continuity or grades.
- The MMC Term 4 checkpoint did not download a live private subscription output;
  backend publication/readiness and artifact tests passed. Do not relabel that
  specific output check as completed.
- The old draft PR 6 (time formatting) needs reconciliation: some changes already
  exist; editor/console formatting remains a Batch 8 item. Do not merge it blindly.
- On 28 October the live deployment/configuration/cache metadata was rechecked;
  full authenticated UI acceptance was not repeated. Earlier UI acceptance is
  dated evidence, not a claim of a fresh exhaustive audit.
- Budget snapshots expire. The 28 October 21:23 Melbourne read-only check recorded
  2,106,571 reads / 29,161 writes and withheld new bulk read admission because of
  the forecast against the app's four-million-read safeguard. It did not report
  an actual daily quota breach. Always obtain fresh admission before maintenance.

## Evidence

- [Batch plan](restoration-batch-plan-2026-10-04.md)
- [Staff-history activation and authenticated acceptance](ed-staff-history-search-implementation-2026-10-20.md)
- [History grade/grouping/refresh final acceptance](ed-staff-history-refinement-plan-2026-10-20.md#final-acceptance-after-unlock)
- [Publication completion and Orientation visual acceptance](publication-delivery-repair-plan-2026-10-19.md)
- [Manual/automatic handoff](roster-source-authority-plan-2026-10-19.md)
- [Manual-import optimisation](manual-roster-import-optimization-2026-10-20.md)
- [Workflow reliability verification](workflow-reliability-2026-10-22.md)
- [Batch 6 staff release](restoration-batch6-requests-2026-10-18.md#staff-release-after-iphone-acceptance)
