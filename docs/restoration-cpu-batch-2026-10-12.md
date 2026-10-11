# Free-tier CPU reliability batch — 12 October 2026

Cloudflare Free remains required. This batch improves roster delivery and quota
admission; it does not accept all of Batch 2 or restore later feature sets.

## Changes

- The automation runner computes complete roster comparisons and compact staff,
  grade and stream metadata outside Pages. Workers validate pinned, typed facts
  in small resumable transactions. Existing calendar facts remain active until
  the final guarded transaction commits.
- Existing-term corrections inspect exact-file indexed pages, stage only changed
  facts, then atomically update events, presence, grades, coverage and publication
  jobs. No whole-term event JSON is parsed/hashed in the Worker. Oversized changes
  use the inactive bounded replacement path with its existing promotion fence.
- Automatic event batches contain at most 64 occurrences / 64 KiB. Compact and
  correction batches contain at most 16 entries / 24 KiB. The original 1,250-fact
  safety envelope, source ordering and shared account reservations remain.
- One bounded SQL preflight returns file/job/source/term evidence. Catalogue reads
  retain a 751-row sentinel per overlapping exact-source file; active-file lookup
  retains its 33-row sentinel. Physical event history is not scanned here.
- Read-only target lookup, initialization, atomic activation and completion
  bookkeeping are distinct callbacks. An interrupted completion cannot repeat
  committed facts. Uncertain derived responses are not retried in the same pass.
- Routine account admission requests database totals and timeline reconciliation,
  excluding SQL fingerprint diagnostics. The separate diagnostic CLI still asks
  for them. D1 aggregates bounded maintenance receipts into eight counters instead
  of hydrating thousands of objects in JavaScript. Midnight carry, reservations,
  unknown outcomes, recovery deadlines and allocation-race fences are preserved.
- Automation may reuse an existing unexpired same-day SQL grant. This neither adds
  headroom nor bypasses the unfinished-maintenance gate; every mutation still
  reserves its own cost atomically.

Migration 0043 was applied successfully to production after listing only that
pending migration. It adds three empty staging/receipt tables and performs no historical
backfill or index rebuild. Preview bindings and unrelated local files are kept.

## Validation

Real MMC (4,295 events) and MCH (2,363 events) workbooks pass wire transport,
compact-metadata/grade parity, unchanged/corrected/removed facts, source ordering,
indexed cursors, interrupted transactions, stable resume and publication tests.
Failures preserve existing calendars. Receipt accounting tests cover midnight,
unknown cost, reconciliation, 10,000-receipt sentinels, races and unchanged ceilings.

Isolated live tests use synthetic far-future rosters in the identity preview DB,
never production clinical records. Earlier experiments still used 12–21 ms for
combined setup/activation; they are not accepted as a CPU fix. Combining bounded
SQL preflight reduced the first observed request on a fresh final preview to
8 ms. Observed follow-ups used 3–7 ms, and both correction and next-term workflows
completed. The paced follow-up captured correction activation at 5 ms, next-term event
staging at 8 ms, activation at 7 ms and completion at 3–5 ms. A separate test at
the actual MMC size (148 doctors / 4,295 shifts / 72 batches), with synthetic
names and no activation, initially used 13 ms for initialization. Removing its
duplicate validation and using two typed SQL inserts reduced initialization to
10 ms and fresh lookup to 8 ms. This is at the allowance, not proof of ample
CPU margin. Tail sampling did not capture every callback: successful HTTP
responses alone are not evidence that every request fits 10 ms. Final CPU and production
acceptance must remain explicit in the acceptance record.

## Remaining acceptance / scaling work

1. Trace actual final activation and event staging, fresh and warm; verify natural
   production changed-roster delivery, publication and unchanged subscription UIDs.
2. Dense identity suggestion matching still runs within Pages. Move comparison
   work to the existing external runner and persist small reviewed candidate
   checkpoints, preserving every name, full comparison coverage and human
   rejections. Do not solve CPU pressure by skipping large surname blocks or
   silently truncating candidate lists.
3. Separately measure ordinary user requests and private feed generation. Hundreds
   of simultaneously visible browser sessions are a request-count issue as well
   as a CPU issue: registered-user counts alone do not determine traffic. Avoid
   rebuilding shared data per user and quantify refresh/request budgets before
   claiming that scale. No polling frequency or private-feed behavior changes in
   this batch.
4. Continue the existing restoration order 1 → 2 → 6 → 7 → 3 → 4 → 5 → 8 after
   Batch 2 CPU acceptance. Do not call all historical 503 causes eliminated.

## Production rollout and DDH reservation correction

Code 0c766f99 deployed and pushed. Both sampled feeds retain HTTP 200, 128/154
VEVENTs and their exact baseline UID sets. The temporary CPU Worker and all
CPU261012 file/doctor/job rows are removed; two original preview files and the
accounting ledger remain. Production secrets retain their binding names; API
values are redacted, so they were not copied or replaced.

The production safety stop predates this batch: a DDH request finished at
2026-10-11T05:55:38.469Z with 4,004 actual reads versus 1,072 reserved, and 32
actual writes versus 176 reserved. Its inline account refresh read the receipt
ledger but the ingestion quote allowed only 1,024 reads. Routine SQL aggregation
reduces CPU, not those underlying D1 reads. The corrected DDH quote includes a
65,536-read envelope only when that request refreshed its account grant, covering
the bounded receipt/recovery/CAS reads; normal unused-reservation settlement
returns the surplus. An unchanged poll still does not refresh or download.

The two manual queue workflow checks authenticated successfully but found zero
eligible work; this is not proof of changed-roster delivery because the cost stop
was active. Clear only the identified old stop after deploying this correction,
expire its grant and require normal settled account re-admission. Do not refund
receipts or clear other stop reasons. At the last settled account sample there
were 1,768,605 reads and 46,051 writes; the full diagnostic CLI remained STOP for
billing/sample/query-diagnostic evidence, not a quota-limit breach. No CLI GO or
full Batch 2 acceptance is claimed.

Closeout: reservation correction 2c5632f5 deployed successfully (manual Pages
683228dd; Git-integrated deployments may supersede it with the same code). Only
the exact diagnosed stop was cleared, conditional on its known completed receipt
and absence of another overrun; the grant was expired first. Normal admission
then issued a fresh grant valid to 22:11:18Z, with an empty stop reason. One
unresolved reservation remains held; none was manually refunded. Both sampled
feeds still match their exact baseline UID sets. The final settled analytics
cutoff 21:45:26Z records 1,791,965 reads / 50,882 writes; the subsequent isolated
cleanup measured 3,378 reads / 3,228 writes. The full diagnostic CLI remains STOP,
so this record does not claim diagnostic or full CPU acceptance. Maintenance
verification run 38178709371 uses the deployed commit and no forced view seeding.

At approximately the user's 25%-remaining stopping point, the next work remains
externalizing dense identity comparison without lost coverage, obtaining actual
changed-roster production CPU/publication acceptance, and then proceeding to
Batch 6. Keep Free tier, current quotas, stable feeds and historical grades.
