# Workflow reliability — 22 October 2026

## Incident

GitHub run 39706801777 reported expected HTTP 200 `repair-grades` postponement
as a failed staff-history run. MMC runs 39706773156 (Term 3) and 39698807970
(Term 4) failed with `D1_ERROR: Network connection lost` during indexed,
read-only `bounded-patch-page` comparisons. Term 4 subsequently imported at
00:33 UTC. These errors did not indicate Free-tier CPU or daily quota exhaustion.

## Repair

- Staff-history runners distinguish explicit deferred/superseded responses from
  unexpected HTTP, network, JSON and processing failures. Only expected outcomes
  finish normally, with an honest workflow summary; genuine failures stay red.
- Dispatch acquires its existing R2 lease before a single bounded preflight.
  Pending imports or unavailable optional capacity postpone dispatch for five
  minutes rather than repeatedly launching blocked runners. Source publication
  callbacks and scheduled maintenance revisit changed inputs. Unchanged history
  still exits without D1 queries in the history API.
- Reused maintenance grants expose their actual expiry and require at least one
  minute remaining. Both affected runners renew before expiry; refreshes retain
  existing analytics, unknown-receipt, concurrency, daily-limit and reserve gates.
- Comparison-page transport retries transient reads at the same cursor at most
  three times, with two/four-second backoff. A read rejected by maintenance can
  request one fresh analytics admission and retry once. Persistent deferral
  retains queued work. Source/baseline fences are checked on every page.
- Stale/invalid inputs and uncertain derived writes are not blindly retried.
  Existing durable staged-write cursors and atomic activation remain unchanged.
  The fix preserves comparison progress within the running process; a persistent
  interruption across separate workflow invocations can still restart inspection.

No thresholds, Free-tier plan, source Flow schedules, clinical interface or email
notification preferences changed. No schema migration or bulk reimport is needed.

## Verification

Passed workflow-reliability tests, including the actual queue executable with
synthetic transport interruptions; maintenance accounting; staff-history API,
cache and dispatch coalescing; MMC/MCH fixture patch parity, atomic rollback and
stale fences; queue failure reporting; roster delivery reliability; cold routes.

The pre-deployment read-only account diagnostic at 15:11 Melbourne returned GO:
632,520 reads, 13,476 writes (settled through 14:56). Daily safeguards remain
4,000,000 reads / 100,000 writes, with protected publication capacity.
