# Bounded roster restoration batch — 1 October 2026

This checkpoint replaces the pre-reset work list. Production overwrite protection
was released as 098bc2a2; MMC/MCH/DDH contacts were verified before this batch.
User authorised a combined restoration batch and safe automatic continuation.

## Implemented and locally verified

- Deterministic manifests and hash-pinned, ordered chunks for MMC, MCH, DDH and
  VHH; 1,250 weighted facts per chunk, 512 KiB payload, 512 doctors, 25,000
  events, 5,000 issues, and 120-day maximum roster span.
- Inactive staging, transaction receipts, durable event/presence cursors,
  initialization lease, count validation and fenced activation of disjoint terms.
  Current terms remain intact; incomplete terms remain hidden.
- Exact retained-term corrections keep the existing diff path. Large overlapping
  replacements remain guarded. Completion retries repair bookkeeping after an
  activation without deleting live data.
- Shared UTC-day reservations: 10,000 indexed writes and 500,000 reads for
  automated import/publication work across all sites. This leaves headroom under
  the restoration stop thresholds for ordinary traffic. It does not replace
  account-wide analytics admission. Failed work consumes its reservation;
  progress survives and resumes after the allowance resets.
- Corrections reserve old/new daily presence, metadata and SMS continuity costs.
  Local migration-index checks prevent write estimates ignoring added indexes.
- Source-scoped overview queue commits with roster activation. Seven-date
  publication steps retain the previous R2 manifest until successful completion.
  Partial monthly updates preserve untouched dates. Every step reserves an upper
  read bound; unused reads are refunded only with complete D1 billing metadata.
- Immediate publication after queue imports plus a six-hour continuation workflow,
  using the existing global roster concurrency lock and rotating source order.
  DDH checks LastModified before downloading; retained MMC/MCH/VHH jobs resume
  without fetching unrelated sources. Paused/exhausted work skips workbook parsing.
- Optional one-time current-term view seeding uses at most 32 compact active-file
  records per source. It does not rebuild historical terms.
- Cached automatic On shift launch restored for eligible on-duty users. Legacy
  broad builders, history queries, identity fan-out and global warm-up stay closed.
  Preview remains closed. Existing contact flows are unchanged.

## Evidence and rollout

Local checks pass planner/full-term execution (18,000 shifts per source), HTTP
middleware/protocol interruption and retries, receipt rollback, budget deferral,
inactive visibility, current-term preservation, atomic activation, completion
repair, partial-month publication and failed manifest promotion. Contact-cache
checks exercise 36,000 unchanged refreshes with zero D1 use; ingress idempotency,
source isolation, request attribution, client request budgets, rollout guards,
queue-failure safeguards and indexed database-cost fixtures pass. Pages Functions
build succeeds. Synthetic fixtures were used only in local ephemeral databases.

Account-wide admission before release: GO, 27,489 reads / 46 writes. The pending
0032 indexes were already present in Production. Bounded inspection read 694 sync
runs and 34 claims; migrations 0032–0034 applied successfully without backfilling
roster events. New tables are empty on creation.

Safari readback: Sync Monash roster files, Sync Monash Paediatrics roster files,
VHH Active Medical Roster to Production, Sync MMC clinician contacts and Sync DDH
clinician contacts are enabled. Legacy MMC allocations/bootstrap and VHH Preview
flows are disabled.

Release/readback and first maintenance outcome are recorded in the final report
and subsequent checkpoint. Do not infer real next-term completion from local tests.

Rollback: disable ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED,
FACILITY_AUTOMATIC_PUBLICATION_ENABLED and automatic launch, and cancel scheduled
maintenance. Keep progress tables and valid published objects for recovery;
never delete current terms or re-enable historical whole-workbook contact flows.

Remaining service work: compact colleague tools, bounded manual roster mutations,
Creator switching/directory enrichment, doctor discovery, cross-device settings,
and VHH contact support. All-site next-term completion still depends on provider
publication and may span UTC budget windows.
