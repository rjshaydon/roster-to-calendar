# At a glance Phase 0 baseline

Measured locally on 5 September 2026 from production baseline `aa9eed8`. These
are estimates from SQLite query plans, not Cloudflare billing measurements.
They describe the old implementation and are not Phase 0 acceptance failures.

## Focused scale fixture

The deterministic fixture contains 5 active files, 600 doctors and 109,200
events (120 doctors per ED across 182 dates). Parsing correctness remains in
the existing Excel, FindMyShift, contact, queue and roster fixture suites.

| Path | Rows returned | Estimated rows examined | Plan finding |
| --- | ---: | ---: | --- |
| Staff, one ED/term | 120 | At least 120 memberships plus correlated event probes | Scans active files, then performs a correlated file-event lookup and DISTINCT sort |
| Coverage, one ED | 1 | 21,840 | Indexed by source, but walks the ED's complete event history for MIN/MAX |
| On shift, one ED/date | 120 | 120 | Bounded indexed lookup on source and start date |
| Access, one doctor/term | 91 | Up to 182 | Bounded indexed doctor/date lookup; small authentication/account reads remain legitimate |
| Contact discovery | 8 | At most the newest indexed source rows | The lookup is small, but the old ten-second request repeats authentication, access and R2 discovery |

The test intentionally asserts only that bounded paths use their indexes and
records broad existing behaviour. Future phases—not Phase 0—must make repeated
authenticated cache refreshes use zero D1 rows and remove event-history reads
from Staff and coverage.

## Capacity worksheet

Assumptions: 50 simultaneously visible On shift pages, 12 viewing hours per
day, one combined refresh each 60 seconds, hidden pages paused, immediate check
on opening, and 15-minute access validity.

| Work | Daily operations | Intended D1 rows |
| --- | ---: | ---: |
| Visible refresh checks | 36,000 | 0 after valid authorisation |
| Initial page checks | approximately 50 | Small bounded authentication/access work |
| Access renewal, worst case with no cross-tab sharing | 2,400 | 0 roster-event rows; use compact access state |
| Full 24-hour visible worst case | 72,000 refreshes | Exceeds the plan's 50% Worker-request planning target and approaches its 70% stop threshold |

The 12-hour case uses 36% of a 100,000-request daily Worker allowance before
other traffic. Cross-tab coordination, jitter, conditional responses and
backoff are therefore requirements, not optional polish. At 50% observed or
forecast use, optional work is paused and investigated; at 70%, optional D1
reads/builds stop. Actual Cloudflare row and request metrics are reserved for a
separately approved, tightly bounded rollout.

## Future acceptance gates

- Staff and coverage reads do not access `roster_events`.
- An unchanged authenticated refresh performs zero D1 operations.
- On shift remains an indexed ED/date lookup until shared snapshots replace it.
- Identical ingestion rewrites no events, membership, daily presence or cache objects.
- One correction changes only its affected facts and dependent summaries.
