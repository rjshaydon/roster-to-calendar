# Staff-history grade and refresh correction — 20 October 2026

Status: implementation and bounded production grade repair completed on Medium. Final visual acceptance is pending because the Mac locked.

## Read-only findings

- Active compact MMC term contributions and file doctor records say Unknown for Richard Haydon in Terms 1/2 and Patrick Tan in Term 1. Indexed aggregate reads of the matching stored shifts prove SMS (22 and 28 shifts respectively) and Senior Registrar (45 shifts) throughout those respective samples. Original shift grades are intact.
- buildFacilityOverviewFacts uses doctor.seniority || event.seniority, so the nonempty Unknown directory value masks dated shift evidence. The history fragment faithfully copies the resulting incorrect compact fact. Lexical MAX/maxText merging can also prefer Unknown over a known grade; resolve known grades deliberately, never alphabetically.
- Compact unknown contribution counts: MMC Term 1 148, MMC Term 2 143, MCH Term 1 73, DDH Term 3 7. These are contribution rows, not distinct people; not every unknown is necessarily recoverable. Inspect source-term evidence before changing it.
- Patrick has active MMC contributions for Terms 1 and 3, not Term 2, including approved identity aliases. A bounded check also found no Patrick/Pat Tan contribution for Term 2 at another site and no matching active MMC Term 2 file-directory row. Do not fabricate contiguous Terms 1–3; investigate source evidence/other unresolved spelling if needed.
- coalesceStaffHistory partitions on upcoming, so current and future adjacent terms cannot combine. Final sorting is hospital first rather than start date first. Term labels repeat the year.
- readStaffHistoryView refuses the complete published catalogue whenever source/identity watermark differs, even while replacement preparation is running. The client clears its cache on a stale details response and error rendering replaces existing results.
- refreshVisibleRosterView explicitly calls clearStaffHistorySearch before refreshing ED staff, even when history has not semantically changed.
- History access expiry uses a fixed quarter-hour boundary. The browser expiry timer destroys the cache and asks the user to search again instead of renewing access. A newly opened search near a boundary can disappear seconds later.

## Implementation batch

1. Correct grade selection at fact preparation. Use dated same-person, same-source, same-term known shift evidence, honoring explicit dated overrides. Treat blank/Unknown/Not recorded as missing. Do not borrow today's grade or a different term's grade. Detect contradictory known evidence rather than choosing lexical MAX or the last iterated shift. Keep genuine missing evidence explicit.
2. Repair existing affected compact grade facts through bounded, resumable background work under normal account admission and publication priority. Use indexed per-file/person/term evidence, never a full event-table scan or any interactive search query. Update only genuinely changed summaries and their required digests/revisions, leaving original events/source files/accounts unchanged. Estimate and reserve all induced writes (including indexes/triggers/publication); do not claim one write per row. Version prepared grade/fragments inputs so existing cached Unknown fragments cannot be reused. Republish only affected term metadata/history fragments and atomically rebuild the catalogue; unchanged terms and rosters must not be reingested. Keep external preparation outside Worker CPU where practical.
3. Coalesce adjacent/overlapping term records only for the same approved person, source and actual historical grade, regardless of current/future status. Preserve missing terms, gaps and grade changes. Sort resulting ranges by start date then hospital, with deterministic ties. Compact labels: Term 1 – Term 4 2026; across years: Term 1 2026 – Term 3 2027. Show hospital, grade, compact term label and date range. Remove repeated Recorded on roster and Upcoming suffixes from these rows, retaining source alias spelling only where helpful.
4. Separate authorization validity from data freshness. Retain a complete authorized snapshot while source changes trigger preparation. Remove unconditional revision-refresh cache clearing; validate semantic history revision cheaply through existing refresh hooks, without a new polling loop. Assemble replacement name pages and matching histories in a separate generation, then swap them atomically. Never combine slots from different catalogue revisions. No change means no rerender/refetch of existing details. Preserve query, focus, selection and continuation count. A preparing/stale/transient network response must not replace existing valid results with an error or an empty view. Initial loading/error UI remains appropriate when no complete data exists.
5. Renew access quietly before quarter-hour expiry while the view is active, coalescing with normal requests and rechecking on foreground return. Unchanged authorization must not invalidate data. Logout, real expiry without successful renewal, denied access, or account/scope changes still clear private state; late requests cannot restore it. Do not extend authorization beyond what the server grants. Avoid retries on every keystroke; use bounded coalesced refresh and existing source-revision hooks.

## Acceptance and release

- Recovered Richard grades produce DDH SMS Terms 1–4 and MMC SMS Terms 1–4, each one range, correctly ordered.
- Patrick Term 1 grade is recovered; Terms 1 and 3 remain separate unless actual Term 2 roster evidence is recovered. DDH Term 4 sorts last.
- JMS consecutive terms coalesce; grade changes, missing terms, simultaneous sites and year boundaries remain correct. Historical-only names and approved alias search survive.
- An open result survives unchanged source publications, pending background preparation, a new complete catalogue, and a quarter-hour authorization renewal. Unchanged content does not rerender; genuine changes replace a coherent result. Logout/access loss/account switching clear results safely.
- Focused builder/client/renderer integration tests, unchanged-replay checks, and bounded authenticated CPU/D1 checks cover the changed paths. Confirm real examples through native authenticated Safari, using aggregate production queries only; do not export private roster rows to local/test environments.
- Retain current safeguards and Cloudflare Free.

## Release evidence

Runtime commit a0213c95; deployment https://aa1492b5.roster-to-calendar.pages.dev.
Additional regression tests committed as 34668095. Successful repair/preparation
run 39359875438 completed at 02:39 UTC (13:39 AEDT), publishing 718 current
approved/roster identities. Earlier passes yielded to the existing unfinished
maintenance concurrency guard and resumed durable checkpoints; no reservations
were manually refunded or bypassed.

Aggregate verification: all 364 formerly Unknown MMC/MCH contribution summaries
are now known. Richard's MMC Terms 1–4 all say SMS; Patrick's MMC Terms 1 and 3
say Senior Registrar. Patrick still has no Term 2 contribution. Seven DDH Term 3
unknown summaries remain: the repair does not invent a grade when evidence is
missing or contradictory. Original shift/event records were not changed.

Passed staff-history API/builder/client/repair tests; actual browser function
checks for unchanged authorization renewal, retained histories and unchanged
DOM; revision fingerprint isolation; facility materialization/access/snapshot
suites; maintenance-account-budget; cold route startup; syntax and whitespace
checks. No new native CPU maximum is claimed for this refinement. Existing
full-handler evidence predates it; interactive readers now avoid watermark HEAD
checks and stream the same prepared name envelope.

Live Safari before lock confirmed compact current/future grouping and retained
search text/results while repair was in progress. Final corrected visual check
and an unchanged preparation replay remain pending; do not mark these as passed.
The user has been asked to unlock the Mac. Usage reached 15% remaining during
verification, so no further implementation was started.
