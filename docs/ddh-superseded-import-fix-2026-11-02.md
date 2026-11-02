# Superseded roster patch handling — 2 November 2026

DDH workflows 41370610099 and 41371508297 parsed the Term 4 roster successfully
but failed with `page.rows is not iterable`. The derived endpoint can return
`completed:true,superseded:true` when a newer delivery replaces the inspected
input. The patch driver handled deferral but assumed every other page response
contained rows. Other checkpoints similarly lacked completed-outcome checks.

The patch driver now stops on completed outcomes at entry, startup, each read,
plan pinning and each staged batch, matching existing activation/rebase behavior.
The queue reports superseded work as skipped, not as an indexed zero-shift roster.
No subsequent old-job callback is issued. Genuine errors still fail; existing
source fences, daily safeguards and atomic activation are unchanged.

Tests cover supersession at every section, a continuation page, plan pinning,
staging, activation and finalization, plus an already-completed input. Existing
MMC/MCH real-fixture diff/parity and rollback checks pass. The actual queue
executable with synthetic transport confirms one superseded response exits zero
without reporting an import success or failure; transient/stale/write-loss and
queue failure tests continue to pass.

Before this fix, normal DDH workflow 41372426366 recovered: 119 doctors and
3,424 shifts indexed at 12:30 Melbourne; processing/publication completed at
12:34. The successful retry is recovery evidence, not live acceptance of the
new supersession branch. No clinical roster edit or test import is manufactured.

Release is through the scripts checked out from main by GitHub Actions. There
are no application/Worker changes, migrations, flag changes or production data
writes required to release the fix. Future jobs use the published revision;
already-running jobs retain their checked-out revision. Observe the next natural
supersession for live acceptance without reopening an obsolete roster job.
