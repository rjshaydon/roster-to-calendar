# Automated roster term protection — 1 October 2026

The complete-save endpoint previously reused the automated source's active-file
pointer for every delivery, including a different term. It now selects an exact
term-range match from at most 32 active files for the facility using compact
indexed coverage rows. A disjoint new term gets its own file; an older-term
correction preserves the latest-term source pointer. Ambiguous, unprepared or
partially overlapping term ranges fail before roster replacement.

The existing explicit 1,250-fact automation ceiling now also applies to first
imports. Large new-term deliveries remain safely rejected pending bounded
staging/chunking; this change does not establish full new-term readiness.

Completion retries recognise the original queued input after its file id has
been remapped. Failure callbacks reject completed runs, and error cleanup only
deletes inactive derived files. An activation followed by failed bookkeeping
retains live data; retry repairs the bookkeeping with an unchanged roster save.

Validation: facility materialisation/endpoint regressions cover current-term
retention, independent next-term files, 14-day visibility, old-term correction,
duplicate callbacks with zero writes, oversized first import, overlapping-term
rejection, indexed target lookup, late failure and interrupted bookkeeping.
Existing ingress idempotency, source isolation and queue failure checks pass.

Release admission at 02:13 UTC: account-wide D1 checker GO; 19,414 rows read,
36 rows written. No contact code, flow, feature flag or database migration changed.
