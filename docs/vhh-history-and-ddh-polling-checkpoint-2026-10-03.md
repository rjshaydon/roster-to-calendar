# VHH historical recovery and DDH polling — 3 October 2026

Approved scope: fill VHH's missing roster history and check DDH every five minutes.
The broader reconciliation and browser freshness proposal remains separately
tracked in `vhh-history-and-roster-freshness-plan.md`.

## Implementation

- DDH's dedicated Cloudflare Worker checks FindMyShift LastModified every five
  minutes. It no longer issues an unscoped processor kick. Current-term checking
  continues when the next term enters the existing publication window.
- One provider metadata request covers at most two roster windows. Exact
  successful filename/version matches and known waiting/incomplete windows avoid
  repeat downloads and ordinary unchanged D1/R2 writes. Queued work retries the
  existing guarded dispatch instead of downloading its input again.
- Changed DDH input needs a fresh account grant and bounded ingestion reservation.
  The existing staged import, delivery-order protection and shared publication
  remain in force. Repeated provider failures avoid rewriting the same error.
- VHH recovery explicitly filters a complete historical date window. Live
  extraction still respects hidden sheets. Seniority evidence is now confined
  to the Victorian medical term containing each assignment.

## Recovery evidence

Pinned downloaded workbook SHA-256:
`bcebfcbf599e22cdba2e634cdb0ae9025d6b1fc3a5054a71c000f55f35528043`.
The downloaded and SharePoint workbooks were not changed. Raw data and private
recovery plans are retained outside Git, with a source copy in private R2.

Active VHH provider inputs before recovery covered 27 July–23 August and
21 September–31 January. Two disjoint manual historical contributions filled:

| Shift-start dates | Staff identities | Events | Fact batches |
| --- | ---: | ---: | ---: |
| 4 May–26 July | 53 | 1,202 | 3 |
| 24 August–20 September | 46 | 376 | 1 |

Recovery used the existing bounded manual import functions through an
authenticated Cloudflare D1 API adapter. Each phase used the application meter,
shared account reservation and settlement receipt. The original active provider
set was pinned before preparation; activation kept both original files active.
Both contributions completed without budget deferral. Whole-import measured
cost: 6,059 reads and 26,269 indexed writes, including accounting. Historical
contributions do not pretend to be newer SharePoint revisions or update contacts.

Pre-import two-sample admission was GO: 222,953 account reads, 3,347 writes;
projected ordinary reads approximately 305,403 for the UTC day.

## Validation and live acceptance

VHH fixtures, historical scope, cross-term grade isolation, DDH worker scope,
current/next-term polling, zero-write ingress idempotency, real-workbook bounded
wire protocol, source isolation, emergency quota guards and account reservations
passed. Pages Functions compiled successfully. The broad fixture suite retains
its previously recorded source-text assertion failure in derived completion.

Release PR #13 merged as `b39bd1e8`; canonical Production deployment
`e86d24cd-14ba-4772-afe0-f8718f0ccda8` succeeded. The dedicated DDH Worker was
restored with cron `*/5 * * * *` and its protected watchdog credential. Its
authenticated production check at 10:52 UTC returned unchanged for
3 August–1 November, using the real FindMyShift provider.

The publication harness initially omitted rollout flags and therefore refused
publishing. No production pause was lifted. Fresh canonical Production readback
confirmed emergency pause false, automatic publication true and account budget
true. Using those verified flags, the two historical jobs completed. The
original November term's staff pointer and 19 October visibility date remained
unchanged. The remaining 1 February 2027 maintenance job awaits an actual term
input; the recovery stopped there rather than retrying it.

Read-only verification at 10:58 UTC used the actual published staff, range and
day readers (20 R2 objects, 1,583 D1 reads, zero writes):

| Term | Published staff | Published events | Missing dates |
| --- | ---: | ---: | --- |
| 4 May–2 August | 54 | 1,296 | None |
| 3 August–1 November | 54 | 1,257 | None |

Ten boundary dates passed, including both sides of each historical gap and the
medical term boundary. Both original active file IDs, modification times and
parse revisions were preserved. With the existing future source, VHH now has
continuous retained roster input from 4 May through 31 January.

Final request metering includes FindMyShift first-row reads so scheduled poll
costs and changed-input receipts retain complete billing metadata. Cross-term
recovery windows explicitly fail before any production work. Regression checks
for both changes passed. Safari acceptance could not be repeated because the
Mac was locked; the published-reader verification completed without the browser.

A five-minute check interval alone is not proof of a fifteen-minute end-to-end
delivery guarantee or external calendar refresh. SharePoint flows remain enabled
file-change triggers; their clinician-contact flows are the five-minute schedules.
