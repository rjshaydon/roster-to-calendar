# All EDs staff-history search — 19 October 2026

Status: implemented, merged and accepted on production on 20 October 2026.
See the implementation checkpoint for release and validation evidence.
This is a focused addition before restoration Batch 4, not a renumbering.

## User-visible contract

Keep At a glance → ED staff → All EDs and the existing Find staff input.
With no search text, retain the selected term's department staff lists and the
existing SMS/CMO continuity and JMS term-membership rules. Do not populate those
ordinary lists with everyone who has ever worked at an ED.

With All EDs selected and nonblank search text, search names across all available
published roster history, independently of the selected term. Use the existing
results area, showing each matched person's recorded sites, terms and historical
grades. A short results heading says “Across available rosters” so the selected
term control cannot silently imply that results are limited to that term. No new
search button, panel, modal or navigation is needed. Clearing the search restores
the selected-term view. Single-ED browsing remains unchanged in this first release.

Use the existing site's grouping/cards and name style, with compact history lines
inside each result. Coalesce adjacent terms only for the same person, site and
recorded grade; preserve breaks, simultaneous sites and grade changes. Label
visible future terms “Upcoming”; do not describe future allocations as completed
work. Show the site's full recorded range and term labels where available.

Remove the repeated per-ED “Partial roster coverage” badge: it currently derives
from individual members' first/last applicable dates, which are not proof of an
incomplete hospital roster. Preserve useful messages for genuinely missing or
unreadable source data, once in the appropriate results context. Normal success
needs no status banner. Persisted SMS/CMO membership does not prove leave or work.

## Evidence and identity rules

Search the entire name, including known approved roster aliases, regardless of
capitalisation and usual comma/whitespace formatting. Match partial names without
a D1 LIKE/history scan. Approved identity merges combine aliases across sites;
similar spellings or identical names alone never merge different people. Preserve
unlinked roster identities as separate results with their site/name evidence.
Search must also find historical staff absent from current department lists.

Reuse the existing published staff manifests and identity-history machinery, but
not its Creator-only endpoint. Its current term-history helper includes continuing
SMS membership; that cannot be used unchanged as proof of work in every term.
Use actual per-term roster contributions/evidence. Exclude synthetic continuity,
provider-only directory entries and admin-only designations as work evidence.
Where evidence proves roster listing but not a working shift, say “Recorded on
roster”, not “Worked”. Leave-only records must not imply clinical attendance.
Do not infer missing grades or overwrite historical grades with today's grade.
The result describes available roster records, not a complete employment history.

## Backend and cache design

1. Inspect existing compact identity/name caches and published term staff data.
   Reuse their source/alias keys and versioning; add only the missing historical
   search/history projection, avoiding a duplicate identity system. Historical
   roster_doctors exists, but the normal published signup directory deliberately
   covers current/next terms and must not be broadened for this feature.
2. Build compact per-source/per-term history fragments from already-materialised
   contributions or published staff objects. Each fragment records real roster
   evidence, names/alias references, term and grade; no full shift payloads,
   contacts, account emails or admin audit details. If cached staff data lacks
   sufficient evidence, use a bounded indexed contribution read in background
   preparation, never an interactive event-table scan.
3. Assemble a revisioned search catalogue and history summaries outside Workers
   where assembly could exceed Free-tier CPU, using the established runner.
   Seed existing retained published terms in bounded resumable steps under the
   current account safeguards. Prepare on demand as a job, not while typing.
   Keep incomplete/unreadable terms distinguishable from a true no-match.
4. Attach incremental refresh to successful roster publication, identity
   merge/split/rename and relevant historical grade corrections. Rebuild only
   affected fragments, then atomically publish the coherent catalogue pointer.
   Identity membership and history must come from compatible revisions.
   Unchanged source revisions cause no rebuild or D1 writes. No new polling.
5. Serve a bounded, authenticated compact name catalogue once for this view,
   using conditional revisions/browser caching. Filter names locally while
   typing. Retrieve compact history only for matching results, in bounded pages
   with explicit continuation rather than silently returning the first names.
   Debounce/coalesce requests, cancel stale searches, cache matching histories,
   and preserve input focus/caret and existing results during updates. Do not
   load all term staff objects on each keypress or request each person's entire
   calendar. Normal history reads use published R2 data, with no event SQL.
6. Access is checked server-side before returning catalogue/history. Require
   At a glance enabled and current SMS/CMO all-site eligibility, retaining the
   existing Creator troubleshooting exception. Do not expand Registrar/JMS
   historical or cross-site permissions, or expose identity-edit APIs. Cache by
   authorised scope/revision; clear sensitive view state on logout, account
   switching or access loss. A browser-hidden control is not an access check.

## Acceptance and release as one batch

- Historical-only staff are found from All EDs even when another term is selected.
- Given name, surname, comma formatting and approved alternative spelling find
  the intended person; distinct same-name people remain separate.
- Results retain multiple sites, breaks, past grade transitions and future labels.
- Continuity, leave-only evidence and genuinely missing terms are not presented
  as confirmed clinical work or as complete employment history.
- Existing blank-search staff lists, SMS/CMO persistence, JMS membership,
  single-ED browsing and staff actions remain correct. Historical results must
  not let an existing edit action accidentally change the selected current term.
- Direct endpoint requests from ineligible/disabled users fail before history
  reads; logout/access expiry/account switching cannot reuse authorised caches.
- Typing makes no D1 reads/writes, does not restart searches per keystroke, and
  remains responsive on iPhone. Repeat/unchanged searches reuse cached results;
  background preparation is bounded and resumes without duplicate writes.
- Initial preparation is complete before activation. Deploy version-compatible
  browser assets and server code together. Verify real historical and current
  examples through authenticated Safari, without exporting private roster rows.
- Rollback disables historical-search mode while preserving normal selected-term
  staff browsing and all existing published roster data. Record any unavailable
  historical roster sources rather than manufacturing coverage.

No dependency upgrade, global rebuild, paid service, contact polling change or
clinical roster edit is required. Authority handoff and diagnostic-checker remediation was completed before this
branch. Unrelated legacy test assertions remain under Batch 5/8, outside this feature.
