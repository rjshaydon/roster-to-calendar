# Batch 2 production acceptance — 13 October 2026 (Melbourne)

## Deferred until after Batch 2

Investigate the iPhone bottom menu remaining halfway up the screen after keyboard
use or reopening the app. Scrolling sometimes restores it; sometimes force-quitting
is required. Check keyboard/viewport offsets and app resume behavior, preserving
the normal bottom placement. Reported 13 October; intentionally not part of the
current calendar CPU acceptance work.

## Conditional legacy history read follow-up

The optimized authenticated calendar path previously loaded the old account
snapshot before checking the browser's publication revision. For accounts with
claims, it now pins the legacy cache key and built revision, and
loads historical content only when assembling a changed calendar or returning
a fallback during publication. Creator views without claims initially retained their snapshot-based selection
semantics; the indexed selection follow-up below removes that read as well.

The route also no longer builds an identity-filtered preview that is immediately
discarded when a published calendar is available. Historical selection itself
enforces aliases; fallback previews still receive the identity filter before
being returned. This removes redundant preview work in the existing build-mode
path as well as the candidate reader.

Local authenticated-route tests verify no legacy R2 download on unchanged
checks, exact historical/current event preservation, identity isolation and
fallback during source preparation. Additional publication tests cover interrupted
catalogue builds, completed-person receipt reuse, staged generation rebasing,
stale parent/current-pointer rejection and watchdog lease coalescing. These tests
do not establish native CPU acceptance. Production remains in `build` mode.

External preparation run 38350316863 yielded at its time bound after completing
MMC month derivatives; the watchdog automatically started continuation run
38351770569. A successful yielded run is progress, not a completed backfill.

## Small identity index groups

Doctor-source publication now prepares 16 immutable SHA-selected index groups
outside Cloudflare. The current source pointer carries shared term/month metadata
and group pointers. Calendar checks read only the group(s) containing the approved
aliases, instead of parsing every doctor's pointer at that site. This adds no D1
queries and retains the existing number of R2 reads for a single alias at a site.
The full catalogue remains available to external builders and as a migration
fallback. Unchanged group content is reused, and promotion requires all 16 group
objects to exist. Missing groups defer delivery rather than appearing empty.

Tests verify exact event/review parity, unchanged and colleague-only revisions,
one-group reads, missing-group handling and old catalogue compatibility. The
watchdog requests conversion of existing ungrouped generations; the interactive
reader still awaits native CPU acceptance before activation.

## Live read-only comparison and indexed Creator selection

The user explicitly approved a temporary authenticated read-only Worker in their
own Cloudflare account, using private subscription tokens and local private
responses. The probe rejected D1/R2 mutations, read the real account/roster/cache
inputs, and substituted a temporary authentication hash in memory. No production
accounts were edited. Its CPU timings are native full-request samples, but are
not a universal Pages performance guarantee.
The temporary Worker, its secret and its local access-key file were removed after
the comparison; private response/proof files remain restricted to the local user.

Creator aliases can now be obtained through a primary-key `canonical_doctors`
join in the existing snapshot-registry query. This adds at most one indexed row
read, with no extra SQL statement. Missing canonical selection retains the
previous snapshot-based selection. Approved identity expansion still applies.
The legacy cache loader checks stored revision metadata before decoding history;
a missing or concurrently replaced historical cache defers instead of silently
publishing a calendar without those shifts.

Across live comparisons, all 128 standard-user events and 150 Creator events,
including review fields and ordering, matched the existing reader exactly.
Optimized initial requests took about 0.65–0.92 seconds versus about 1.5–2.8
seconds for the existing reader. Unchanged checks took about 0.29–0.45 seconds
and performed zero legacy-cache downloads. Native CPU ranged 13–24 ms for
initial requests, 5–9 ms for standard-user unchanged checks, and 6–11 ms for
Creator unchanged checks. Thus acceptance still fails; `build` mode remains.

Stream-based cache decoding and a date-sort fast path were tested privately and
discarded because native samples did not establish sufficient CPU improvement.
The remaining full-response assembly/decoding should be profiled or moved out of
interactive Worker execution; do not enable the candidate merely because it is
much faster than the existing reader.

External preparation run 38357051324 completed (not yielded) at 06:40 UTC.
Read-only readiness confirmed matching source generations and grouped catalogues
for MMC/MCH/DDH/VHH. Casey has no shared source manifest in this pipeline and is
still part of the later Casey restoration batch.

Targeted doctor, subscription, materialization, authentication containment,
clinician-switching and cold-start checks pass. The broad `test:fixtures` run
reaches an existing source-text assertion at line 922 expecting the former
active-file-selection implementation in the unmodified `derived.js`; record
this separately from calendar runtime parity rather than weakening that test.

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
