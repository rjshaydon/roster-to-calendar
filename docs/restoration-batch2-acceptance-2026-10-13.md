# Batch 2 production acceptance — 13 October 2026 (Melbourne)

Status: **not accepted**. No safeguards or clinical data were changed in this pass.

## Passed

- Subscription, external identity audit and identity-operation integration tests
  pass, including full history, current ownership, merge/reversal, durable
  rejection, replay, bounded resume and checksum failure.
- Two real live subscription feeds return HTTP 200 and match fresh read-only
  rendering of current production account/roster inputs exactly, excluding the
  generated timestamp. Staff: 128 events, no IDs added/removed. Owner: 155 events,
  one additional DDH roster event compared with the earlier 154-event capture;
  no prior event IDs removed. Full current fields match, not only counts.
- Initial native feed samples: 6–8 ms CPU, six statements, zero D1 writes,
  complete metadata. These are limited samples, not a universal CPU guarantee.
- The production registration runner yields to a newly queued MMC import.
  The retained source file is present. Successful workflow badges are not counted
  as completed registration or calendar builds.

## Outstanding / failed acceptance

- Registration R2 checkpoint remains active with zero examined in the sampled
  checkpoint. No external production audit job is complete.
- No subscription prime checkpoint exists. Overnight runs explicitly report
  `deferred, built: 0`. The normal grant initially had 19,920 write headroom,
  below the unchanged 20,000 source protection; the new MMC import lowers it
  further. Neither audit nor warm-up may bypass this guard.
- A broader native trace contains a dynamic feed at 11 ms; bounded patch-page
  callbacks at 11–19 ms; contact workbook extraction at 19 ms; and calendar-load
  state requests at 137–145 ms. All sampled responses succeeded, but Free-tier
  CPU acceptance fails. No blanket assertion that historical 503 causes are fixed.
- Calendar-load code reads shared monthly site publications, JSON parses them,
  then filters to authorized doctor markers. This is a candidate for CPU
  attribution and offloading/per-doctor publication, not a measured proof of
  where all 137–145 ms are spent. Subscription ICS caching does not optimize
  this separate JSON app-calendar route.
- Fresh MCH changed-file delivery remains unobserved in this pass; its latest
  completed bounded import remains 80 doctors / 2,364 events.

## Budget evidence

Settled account analytics: about 2,165,309 reads / 58,532 writes in the current
UTC day. The diagnostic sample is not general clearance: billing evidence and a
second sample are missing. The app's normal admission remains authoritative.
D1 resets at 11 am Melbourne during daylight saving, not local midnight.

## Required next work

1. Complete natural MMC delivery and registration/audit when normally admitted.
2. Complete prime/warm-up, verify cached production feeds and changed-input
   invalidation, then sample actual Pages costs/CPU and receipt settlement.
3. Isolate the app-calendar CPU path and move broad roster parsing to bounded
   external work or doctor-specific artifacts while preserving scope, ranges,
   historical grades and current personal settings. Do not trade CPU relief for
   repeated broad D1 reads.
4. Verify remaining patch/extraction CPU excess and material MCH delivery before
   declaring acceptance. Existing automated deferral/resume remains enabled.
