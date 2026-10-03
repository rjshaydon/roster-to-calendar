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

## Validation and pending live acceptance

VHH fixtures, historical scope, cross-term grade isolation, DDH worker scope,
current/next-term polling, zero-write ingress idempotency, real-workbook bounded
wire protocol, source isolation, emergency quota guards and account reservations
passed. Pages Functions compiled successfully. The broad fixture suite retains
its previously recorded source-text assertion failure in derived completion.

Publication, canonical deployment and scheduled poll acceptance are recorded
below when complete. A five-minute check interval alone is not proof of a
fifteen-minute end-to-end delivery guarantee or external calendar refresh.
