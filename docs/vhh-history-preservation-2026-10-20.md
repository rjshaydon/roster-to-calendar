# VHH historical rollover correction — 20 October 2026

## Fault and repair

VHH's current workbook hides fortnight sheets as they expire. The live parser
correctly excludes hidden drafts, but an incremental correction also deleted the
old live file's expired prefix. Previous one-off history imports did not cover
the next rollover. Missing input therefore covered 21 September–18 October;
an overnight end date falsely advertised 21 September as covered.

The external processor now selects only hidden, past headers before the visible
window. It skips history already covered by separate active files. Each gap or
retiring prefix is split at medical-term and ownership boundaries, with at most
three windows, one term/92 days each, the existing 180-day publication bound,
1250-fact batches and 32-active-file ceiling retained.

Disjoint gaps use an ordinary bounded supplement. A retiring prefix is staged
invisibly and activated in the SAME D1 transaction as its live patch or full
replacement. Server activation rejects an uncovered prefix. A transaction failure
rolls back both changes, and budget deferral retains the old active calendar.
Archives retain dated roster grades and do not take over whole-term authority.
The original workbook remains retained in R2; archives are reproducible from
that parent's pinned input. No private source exports or test Workers are used.

A source-serialized daily recovery workflow checks for gaps. It downloads the
retained workbook only when a gap exists, runs through account admission, and
uses the normal shared-view and personal-calendar publication drivers. The
ordinary source processor handles future rollover preservation immediately.

Published coverage now includes shift-start bounds computed from the existing
bounded catalogue read, without another D1 query. Historical input known absent
from a committed manifest returns a coverage notice; a missing artifact within
actual coverage still reports preparation failure.

## Evidence

The latest retained workbook, examined in memory only, yields exactly:
21 September–18 October; 43 doctors; 371 events; 6 fact batches; 13 compact
batches. This gap is disjoint from all currently active shift-start ownership.

New synthetic tests cover gap planning, hidden-future exclusion, incomplete
archive rejection, both patch and full replacement rollback/retry, exact active
shift counts, no archive authority takeover, overnight boundaries and absent
versus broken publication. Existing materialisation/HTTP continuation, wire
protocol, real MMC/MCH patches, cached staging reuse, VHH parsing, ingress
idempotency and status-message checks pass. Pages Functions compile. The older
quota test retains an unrelated source-text assertion expecting a literal
`rosterWritesExplicitlyPaused` in advanced repository repair (now guarded by
`advancedRosterMaintenanceEnabled`).

## Account maintenance stop

Initial live admission was STOP. The sole overrun receipt at 02:45:42 UTC read
8,643 rows against 1,072 reserved, with only 37 writes against 176 reserved.
That reservation identifies the changed-input FindMyShift check (128 writes and
1,024 reads plus receipt allowance). It omitted reads already spent checking
term windows. The correction reserves observed earlier reads plus a bounded
changed-input allowance; unused capacity still refunds normally. Unchanged
FindMyShift polling remains zero-write.

Initial analytics: 1,618,258 reads and 17,992 writes, but the ordinary-traffic
forecast left no uncommitted read allowance under the existing 4-million read
policy. No quota/forecast ceilings are raised for this repair. Production
recovery must pass fresh account admission; the daily retry is 00:25 UTC
(11:25 am Melbourne during daylight saving).
