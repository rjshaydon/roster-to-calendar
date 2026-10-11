# Batch 2 acceptance — 5 October 2026

Status: implemented and committed on `codex/restoration-batch2`; isolated cloud Preview accepted. Production activation awaits the separate approval required by the doctor identity plan.

## Delivered feature set

- Creator review of preferred full names, canonical IDs, source aliases and account links.
- Read-only previews; versioned atomic merge, name correction, ID redirect, alias move and account linking; durable history and exact reversal. Later dependent changes block reversal rather than receiving a partial undo.
- Explicit named confirmation for multiple accounts. Accounts, subscription tokens, source shift rows and historical grades remain separate and unchanged.
- Indexed people search and bounded suggestions, durable rejection, obsolete-suggestion suppression, site-scoped manual audits and resumable progress. A single manual start processes up to 250 identities through short sequential requests and can be stopped after its current batch.
- Formatting-only automatic alias linking, with account conflicts and alias limits enforced. Spelling and nickname variations require Creator review. Existing approved associations are preserved.
- Exact affected-owner snapshot invalidation, retryable publication jobs, preferred names and approved aliases in personal/profile calendars and existing subscription feeds. Reassignment filters old cached shifts; an empty roster returns a valid empty subscription.
- Registration checks published metadata; unchanged metadata performs zero D1 work. Weekly suggestions use the indexed changed-feature queue, defer during imports, and begin Sunday 03:30 Melbourne time, with resumptions until 04:30. Candidate auditing creates no roster imports, snapshot jobs or R2 writes.

## Safeguards

One operation: at most eight existing people, 32 aliases, 16 accounts, 64 redirects; a calendar has at most 16 approved aliases. Transactions have a 400-statement ceiling. Discovery processes at most 25 identities and stops between identities after ten seconds. Each blocking group has at most 24 neighbours. All mutation/publication requests use existing account-wide maintenance admission and reservations in production. Runtime schema creation is absent.

The public identity revision is an opaque wake-up marker. It can wake unaffected open tabs; database snapshot invalidation remains limited to exact affected owners. Phone calendar providers retain their own subscription-fetch schedules.

## Verification

Passed: identity operation suite; bounded identity with 100,001 historical events; restoration mutations; Batch 1 freshness; watchdog; D1 quota; facility access. Identity tests cover injected transaction failures, stale previews, legacy-claim races, every edit type and reversal, reserved IDs, multi-account confirmation, rejected suggestions, indexed discovery, unchanged audit, feed continuity, unchanged historical grades and removal of earlier cached identities.

Cloud acceptance used a separate database (`741ce534-eafe-4aa7-9305-7b7ed5661979`) and bucket (`roster-identity-batch2-preview`), containing synthetic records only. The older Preview database has an incompatible parked identity schema and was left intact. No production migration or identity mutation was performed.

The existing synthetic subscription URL delivered both aliases after merging and removed the moved alias after exact reversal. Internal refresh retries completed. Registration and auditing processed 124 published aliases; the repeated audit examined zero identities. Browser verification confirmed Creator search, merge preview/approval, history, completed publication and reversed operations. Desktop layout was checked and corrected; mobile uses a single-column form layout.

Measured cloud D1 costs below include authentication where applicable; identity request metadata was complete:

| Request | Rows read | Rows written |
| --- | ---: | ---: |
| Merge preview | 28 | 0 |
| Merge commit | 51 | 33 |
| Exact reversal | 63 | 33 |
| Initial registration batch, maximum observed | 83 | 381 |
| Suggestion batch, maximum observed | 1,456 | 39 |
| Unchanged manual audit | 5 | 11 |

Registration batches took 4.8–5.0 seconds in the engine; suggestion batches took 3.8–4.0 seconds. The unchanged audit took 140 ms. Its small writes are run-history bookkeeping, not identity or roster changes. Existing-feed requests used 10–12 D1 statements with no writes; their row metadata was partial, so they are not presented as complete row-cost measurements.

Final settled account-wide safety sample: **GO**, 402,717 reads and 3,885 writes; projection approximately 598,359 reads and 14,049 writes for the UTC day. Analytics has a settlement delay; this sample includes only settled test usage. Additional synthetic test requests remain bounded by the measurements above.

## Production activation proposal

Production preflight found the original compatible identity schema and 48 people. Its migration ledger ends at 0038; the 0039 term policy is already applied (zero mismatched term rows), so record that verified baseline before applying 0040. The new migration adds identity metadata and indexes without roster/event backfill. Existing identity aliases and names are not overwritten by the migration.

After explicit approval: apply only migration 0040 with the existing schema/ledger baseline checked; release this branch; enable identity review and bounded registration; enable the weekly audit in both Pages and the existing watchdog; verify credentials and budgeted requests; initialize the published registry in bounded batches; check subscriptions and usage. Approve real spelling/nickname merges individually through the new interface.

Rollback: disable the three identity flags in Pages and the two watchdog flags, restore the preceding application release, and retain additive schema/history. Reverse any subsequently approved real identity operation through its exact preview before retiring that feature.

The implementation is ready for staged activation; real clinicians and their phone subscriptions have not been merged as a test. Richer historical evidence, upload/date-range audit filters, identity-dispute workflows remain follow-up work, with disputes/account lifecycle scheduled in Batch 6.


## Production rollout checkpoint — 6 October 2026

The user authorized the grouped production activation. PRs 27–31 are merged;
production is `0f57c850`. Migration 0040 is applied after verification of the
0039 baseline, with selective private recovery exports and a Time Travel bookmark.
Identity review, registration and scheduled audit flags are enabled in Pages;
registration and weekly audits are enabled in the existing five-minute watchdog.
The existing credentials are retained. No clinician merge was performed as a test.

Live acceptance found and corrected a Creator selected-calendar regression
(PR 28). Creator calendars and feeds now expand approved aliases from the selected
roster identity rather than replacing that selection with the Creator's own
account links. Existing ordinary and Creator subscription URLs retained their
133 and 154 events respectively, including identical event UIDs in comparisons
against the preceding release. The selected calendar displayed 59 events after
refresh, including the next term.

The account safety gate initially deferred initialization because its runtime
inventory omitted the isolated preview database. PR 30 restored parity with the
four-database inventory. Quotas and admission rules were not relaxed.

Initialization then examined 400 published names and registered 368 new aliases;
the directory contained 399 people. A production CPU/memory limit interrupted
its next request. PR 31 reduces both registration and audit requests to five
identities, retaining durable cursor/lease progress. The complete identity
operations suite passes on the released code, including Creator calendar/feed
coverage (its missing test import was corrected). The earlier compatibility test
log had failed at that import and must not be treated as a passing run.

At this checkpoint, final initialization and initial suggestion audit are still
pending. Safari requires a fresh user sign-in for interface acceptance. The
03:10 UTC automatic identity call authenticated successfully (200, 14 reads,
zero writes) but deferred work while MMC/VHH imports were queued; this does not
yet establish successful automatic registration. The weekly audit is configured
for the existing Sunday window and its scheduling policy is tested; a future
production Sunday run has not yet been observed.

Latest settled account-wide check: GO, 126,120 reads and 1,397 writes. These
figures have an analytics settlement delay. Private rollout diagnostics remain
outside Git. Finish the saved initialization, review the first candidate list,
verify a successful automatic registration checkpoint, then repeat subscription
comparison if aliases changed. Ambiguous matches require human review.


### Resumed live acceptance — 6 October, after user sign-in

Production login/calendar succeeded. The reduced five-identity requests passed
live acceptance and registration completed after examining 468 published roster
names. The directory contains 442 people; 417 aliases were registered during
initialization. The suggestion audit examined 50 names and saved three pending
suggestions, visibly presented with Same person / Different people / Review.
No suggestion was accepted or rejected and no manual identity operation was made.

The next account-wide CLI sample returned STOP: 226,217 reads / 21,150 writes,
with short-interval projected reads about 5.35 million. The settled timeline
shows a maintenance burst at 02:55–03:05 UTC; 03:10 usage fell to 1,406 reads /
zero writes. This is not proof of a continuing runaway, but optional manual work
was stopped at its durable checkpoint pending a safe follow-up sample.

Automatic registration is still unverified: its import-active guard encounters
three old queued runs (MMC 1 October twice; VHH 4 October), rather than fresh
imports. Reconcile these records with their actual processing state before
clearing them or changing the guard. Do not silently ignore them. Initial audit
completion and this automatic acceptance are the remaining activation checks.
Proof screenshot: private `/private/tmp/identity-production-suggestions.png`.


### Full cached-name audit and search presentation — 6 October

PR 32 and production follow-ups (`081daae9`, `a851bfc5`) remove empty-search
People lists/requests, preserve selected people, consult the cached name table
before registration, and continue matching through admitted requests without
requiring repeated Suggestions clicks. Browser verification confirmed blank
search hides People, surname search finds Jay/Jayantha, and clearing it hides
People again. The fixed initialization path creates no new registration run for
unchanged published names (regression test passed).

A five-name matching request repeatedly hit the Cloudflare CPU/memory limit late
in the audit. Matching now checkpoints one name per request; registration retains
five. Live matching requests succeeded at approximately 33–35 ms CPU, 740 reads
and 13 writes including admission/bookkeeping. The audit is confirmed complete:
464 checkpoint-counted names, 14 candidates total, 11 pending and three previously
dismissed. The interrupted request had already marked individual fingerprints
before its run counter committed, so that counter is not the directory size.
Jay/Jayantha is now a visible shortened-name suggestion. No identity operations
were committed during verification. Latest settled CLI sample is GO: 289,410
reads / 22,963 writes, projection about 2.15 million reads. Later verification
traffic remains subject to settlement delay and the per-request gate.

Safari was found on an older immutable production deployment URL, `1240d29c…`;
it was moved to the current canonical `rtc.curiousmind.app`. Existing deployment
URLs retain their old client code. Automated registration acceptance remains
blocked by old queued import records as described above; that is separate from
the now-completed manual cached-name audit.

### Automatic maintenance acceptance — 10 October 2026

Production commit `88139c5` and the Pages/Worker identity flags were verified.
The five-minute Worker cron is deployed; the weekly audit window is Sunday
03:30–04:30 Melbourne time. Added actual scheduler/handler regression coverage
for the window boundaries, paused automation, completed-week suppression and
active-import exclusion. Existing identity operations, bounded linking, editor
and suggestion-cache suites passed.

Reconciled exactly three old queued runs after checking their retained source
files, newer successful imports of the same filename, and inactive staging:
MMC provider versions 394/395 and VHH provider version 1154. They are now
`superseded`, with an explanatory message; no roster, account or identity data
was deleted or merged. The exact SQL and before/after records are private.
The reconciliation reported 14 rows read and 12 rows written including indexes
and import bookkeeping. Production queue readback confirms all three changes.

The 09:40 AEDT scheduled tick then advanced registration run
`e79fcc61-475a-40b3-93a0-1afa7fab6e22` from 100 to 105 examined names. This is
actual automatic progress, not a manually authenticated replay. The account
admission remained open with no stop reason. Full catch-up is not yet complete:
the cached historical-name inspection found 525 names without registered aliases.

Cursor validation also found 209 of those names before the resumed cursor.
Registration now checks the bounded cached names/aliases before marking a pass
complete and rewinds before the earliest missing or changed name. This avoids
publishing an unchanged receipt that permanently skips history added during a
run. Each registration request still processes at most five names; the completion
inspection has a 2,001-row guard and reads no shifts. A growing-directory
regression proves the missing historical name is registered before completion.

The account-wide comparison returned GO at 652,156 reads / 35,017 writes,
projecting approximately 670,000 reads for that UTC quota day. These are settled
analytics, not a claim that the later maintenance requests are already included.
Both existing subscription URLs returned HTTP 200 before and after reconciliation
with exactly the same event UIDs: 133 ordinary-account and 154 Creator events.
No human suggestion decisions were changed. Private live logging was stopped.

Remaining acceptance: observe complete catch-up (including the earlier historical
names) and the real scheduled audit on Sunday 11 October. Today's scheduler and
handler tests are not a substitute for that future production run.

### Completion follow-up — 11 October, 13:40 AEDT

Settled account admission returned GO at 170,920 reads / 4,449 writes. The
automatic registration checkpoint had reached 288 examined names; the bounded
names-cache inspection (3,116 rows read) found 445 missing/changed aliases across
996 cached names. This includes 149 historical Casey names; it does not enable
Casey roster-backed views or read historical shift rows.

No scheduled audit for 11 October exists: the original Sunday 03:30–04:30 window
passed while maintenance was blocked. A delayed Sunday pass may now start later
on that Sunday, and an existing weekly run resumes on subsequent days. R2 holds
its durable run/week checkpoint; completed weeks and idle weekdays do zero D1
work. Auditing waits for registration to finish and retains import priority,
account admission, per-request limits and human review of ambiguous matches.

A manually dispatched continuation workflow processes serial one-name requests
through the same production endpoint, with a 600-checkpoint / fifteen-minute
cap. It stops on admission deferral, active imports, busy leases or uncertain
responses, with no blind retries or safeguard bypass. Automated tests cover
delayed starts, weekday resumption, completed-cache no-ops, historical-registration
priority and continuation limits. Production completion remains to be observed.

### Live bounded continuation result — 11 October, 13:58 AEDT

Production Pages commit da0810d2 succeeded and watchdog version
15b8cf48-8a05-4739-95d3-0650ae316970 is deployed on the unchanged five-minute
cron. GitHub run 38106265471 made 54 serial registration calls, then returned
deferred/busy when the scheduled worker held the lease. This is a successful
safe stop, not completed catch-up. Together with scheduled progress, the missing
name count fell from 445 to 389 across the same 996 cached names. The durable
run had reached 344 examined names at the observed stopping checkpoint.

Completed identity receipt totals for the UTC day were 95,575 reads / 4,027
writes. Two earlier receipts remain unfinished: their request deadlines were
02:17:15Z and 02:52:15Z; recovery bounds are 03:11:15Z and 03:46:15Z respectively.
Their reservations remain conservatively held until settled analytics covers
those bounds. The account grant has no stop reason and retains routine headroom.
These are not proved CPU failures: the interruption/settlement cause remains
unverified, and Batch 2 must not be marked fully accepted on this evidence.

Both existing subscription URLs still returned HTTP 200, with exactly unchanged
UID sets (128 ordinary-user events / 154 Creator events). No suggestion decisions
were merged or dismissed during acceptance.

Remaining: finish the 389-name historical catch-up; observe the automatic due
weekly audit start and complete after registration; verify recovery of the two
unfinished receipts without repeated accumulation; repeat settled account-wide
usage/continuity checks. Codex allowance was down to 10%, so no additional manual
continuation job or feature batch was started. Normal gated scheduled progress
remains enabled.
