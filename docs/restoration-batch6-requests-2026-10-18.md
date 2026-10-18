# Batch 6 requests — owner preview checkpoint

18 October 2026, Melbourne. Security/password migration/passkeys remain as previously deployed. This checkpoint implements the remaining private request workflow, but does **not** declare Batch 6 complete.

## Implemented and tested

- Account **Messages and requests** and Admin → Users → **Requests**, with mutually exclusive admin accordions and one-column list/detail/review. Only the owner's admin preview is enabled in this release; ordinary staff do not gain the new panel yet.
- Preview before sending; private channel per participant; explicit reply audience; open/resolved and unread/waiting filters; 20-item cursor pages; private drafts/caches in account-bound memory, cleared on logout/account change. Escaped plain text, no attachments, email or group disclosure.
- Missing-name requests can receive explicitly reviewed canonical context. Known occupied-name requests carry a server-verified site/alias. Relevant roster identity, terms and historical grades are shown through the existing bounded evidence readers.
- Separate private question to an existing linked holder; private decision notices to displaced holders even when no prior question was needed; reply, resolve, reopen, collapsed history; explicit conversion of identified earlier identity reports. Conversion preserves provenance and never treats the old hard-coded MMC label as verified site/ownership evidence. Parser reports remain separate. The old reporting path is retained for staff while the new feature is gated.
- Reviewed ownership assignment/reassignment and reversal are implemented behind an off flag. Canonical links and scoped legacy claims, exact guards, identity history/job, case outcome, decision messages and unread state share one D1 transaction. Generic corrections are bound to case ID/version as well as identity preview; generic account-link edits cannot bypass reassignment from this workspace.
- Old/new account cache invalidation is in the ownership transaction; existing scoped identity publication is dispatched afterward. Private events, settings, passwords and feed tokens are not merged. Account deletion clears private channels/drafts/free-text case notes while retaining minimal operation references.
- Separate private support admission bucket, 10 writes/account/minute; password and passkey buckets unchanged. Existing account-wide daily admission, protected roster headroom and caps remain in force. A message review prepares/reuses the existing daily grant on demand through the separate admission endpoint, so a quiet roster day does not block messages solely because its grant expired. The write still reserves atomically against the same cap. No continuous polling; an existing foreground session check carries a compact indexed unread summary. Ordinary roster/contact reads gain no preflight.

## Verification and cost

`test:support-requests`, `test:support-ui`, account-browser/session/authentication, identity operations, maintenance-budget, admin-user and cold-route tests pass. Browser integration exercises real SQLite API logic in an isolated 390-pixel fixture with all external traffic blocked. Covers review-before-write, private holder conversations, missing-name attachment, isolated editor, transfer, draft preservation, mobile width and logout privacy. Local tests also cover replay, rollback, concurrent session/claim changes, cross-case preview rejection, deletion and indexed 500-account/case pagination.

Native evidence: `restoration-batch6-native-2026-10-18.json`. Complete middleware and budget accounting included, with synthetic D1/accounts only; no provider/mail/production roster binding. At 500 accounts: summary 3 row reads, list 5, recent detail at most 47, all zero writes. Genuine new support/identity requests measured max 5/6 ms CPU. Four-site/multiple-holder and duplicate identity-request writes still produced samples above the 10 ms target (up to 14/12 ms in the final aggregate); transfers therefore stay off. These measurements do not establish full publication/feed acceptance. The later automatic affected-holder notice and case-bound generic-token refinements have local transaction/browser coverage; their complete boundary CPU test remains part of the gated review.

The scale test caught an unsafe planner choice in the duplicate-open-request lookup and a growing unfinished-receipt check. An explicit partial index on open canonical requests and a partial index on unfinished maintenance receipts made both bounded. The disposable grant safely stopped on the early scan; repeated corrected scale tests passed. No production safeguard was relaxed or cleared.

One broader quota test (`test:d1-quota`) has an existing stale source-text assertion expecting `rosterWritesExplicitlyPaused` in advanced repository repair; HEAD already uses `advancedRosterMaintenanceEnabled`. It fails unchanged on that baseline expectation. Its preceding disabled-automation checks passed. This batch does not change that repair path.

## Release gates and remaining work

Rollout JSON: `requestsEnabled:true`, `requestsAdminOnly:true`, `requestTransfersEnabled:false`. No credentials/extra Pages variable added. Migration 0051 is additive; no roster/history backfill or existing claim changes. Read-only production budget review identified today's two disposable databases missing from the verified inventory. Their exact IDs are now included so actual usage is counted; caps and forecasts are unchanged. Disposable request Worker/database removed after testing.

Before general activation:

1. Owner checks the interface on desktop and iPhone: Requests → Contact admin → review/send a harmless own-account request → reply → resolve/reopen; verify the accordion, draft and mobile behavior. Do not invite staff or convert their reports merely for testing.
2. Focused High review of cross-user privacy, atomic ownership/reversal/legacy claims, current authorization, case-bound previews and former-owner cached/feed access. Review the recent-auth path, interrupted identity-editor acknowledgement recovery and scoped publication/retry behavior. The current Medium implementation is not a substitute for that planned focused review.
3. Finish CPU headroom for rare duplicate/large ownership commits; measure supported boundaries and cold/warm complete paths. Prefer a bounded server-stored or authenticated prepared preview: consume it once with the existing exact SQL guards, so the commit need not rebuild its nine-read ownership snapshot. Bind it to actor/case/version/expiry and keep durable idempotent outcome recovery. This needs review before implementation/activation; it is not permission to omit current authority or stale-state guards. Keep the off flag until this passes; do not remove guards/accounting to reach the target.
4. Synthetic connected publication/feed acceptance for changed ownership and reversal, then controlled general messaging/ownership rollout and ordinary production observation. Verify the requests gate is visible only to intended accounts. Preserve the current password/passkey rollout.

Rollback: disable the three request rollout flags (or restore the previous Pages deployment). Keep additive tables and already accepted messages; do not drop private histories. The small rate-bucket addition is independent of existing login/passkey admission.

## Focused High review follow-up

The send deferral was the protected optional-write headroom, rather than an exhausted D1 limit or an unresolved receipt. Both retired synthetic databases are now classified by exact verified ID/day/settled-cutoff costs for the recurring forecast only. Actual account usage, unknown reservations, daily ceilings and protected roster headroom are unchanged. Aggregate evidence is in `restoration-batch6-safeguard-review-2026-10-18.json`; no general exemption for previews or deleted databases exists.

Review corrections:

- Retain the reviewed identity form, its confirmations and original mutation ID when the response is uncertain. Block navigation to another action until retry acknowledges that exact saved result. Browser integration loses the response **after** a real SQLite commit for both request creation and an identity correction, then verifies exactly one durable result.
- Reversal fences all former holders, including legacy-only claimants, against a later canonical link to another person. A concurrent-link regression test verifies rollback and preserved current ownership.
- Ordinary calendar authority comes from current canonical links/claims, never a cached doctor selection. The real account calendar reader is tested against a ready old cache after transfer, including interrupted publication. Creator troubleshooting selection remains separate.
- Connected synthetic acceptance covers four sites, subscription artifact and dynamic fallback, former-owner revocation, publication failure/retry, reversal and unchanged historical grades/private events. Also covers the 64-claim bounds with eight and sixteen aliases.
- Migration 0052 adds a private, five-minute, actor/case/version-bound ownership preview. Commit reads it once instead of reconstructing the reviewed ownership snapshot. Existing exact SQL guards, current authorization, recent-auth, case/version checks and daily accounting still apply; consumption shares the atomic decision transaction and rolls back on failure. Durable mutation replay works after consumption. Expired previews reject; indexed cleanup removes at most ten expired previews on another explicit review. Preview preparation now has its own daily reservation and existing support admission.

All local request/browser, identity, subscription, account/session, maintenance-budget and restoration mutation checks pass. Owner-only messaging and disabled transfers remain the deployment gates pending focused native CPU measurements and a successful live owner send/reply test. The new disposable CPU database is inventoried **before** test writes so it cannot cause unknown-database maintenance pauses.

Additional review results:

- Live owner-only Safari send → reply → resolve → reopen → final resolve passed through the actual daily safeguard. One clearly labelled owner acceptance case is retained as resolved; no staff were contacted or relinked.
- Decision notices now use one bounded guarded insert/update across separate private channels, avoiding repeated per-channel binding/result work. An already-open identity request is a read-only lookup; its new explanation remains unsent. An uncertain accepted save is recovered by an actor-bound, read-only outcome lookup even without a new write grant. No extra receipt/reservation is created for these no-ops.
- Unread acknowledgements that really advance a cursor now have admission/accounting; unchanged cursors remain zero-write. Deferral never falsely marks the message read. Requests can retry publication from the current operation’s existing identity history.
- The owner preview gives immediate “Sending…” feedback, retains the draft on deferral and allows exact-result recovery. General staff rollout and transfers are still off until acceptance is finished.
- The second native run passed 500-account indexed bounds, real unread writes, 64 claims and eight private channels. Batched commits reached 10 ms, reversals 9 ms; detail/list were at most 6/2 ms in individual aggregate queries. A transfer preview still had an 11 ms sample, caught by querying that group separately. The final response now returns only display evidence; full legacy claim/timestamp rows stay in its private server-side prepared snapshot. Final small-boundary measurement follows without another large directory seed.
- Ordinary messages defer roster-term inspection module initialization. The observed actual UI order (summary/list, then compose) measured support creation 5–10 ms and identity creation up to 6 ms. Earlier direct-to-create/large-fixture setup samples are preserved in evidence, not presented as normal UI timings. Aggregate telemetry event counts are not HTTP request counts.
- The exact retired review database is also classified as one-off only for its observed settled usage on 18 October, under explicit reviewed bounds. Later-arriving analytics buckets still consume actual daily usage. This avoids another forecast distortion after test cleanup; unknown reservations and real daily ceilings are unchanged.

## Focused review completion / owner acceptance gate

Final evidence: `restoration-batch6-focused-review-2026-10-18.json`. The compact prepared-preview response passed a final **27-request** native boundary run: eight participants / 64 claims with eight and sixteen aliases, all HTTP 200. Individual aggregate maxima: transfer preview **8 ms**, reversal preview **9 ms**, transfer commit **7 ms**, reversal commit **9 ms**. This closes the earlier ownership CPU gate for the tested supported bounds; it is measured acceptance, not a guarantee about every future runtime sample. The preceding 500-account tests verify indexed directory independence. Both follow-up disposable Workers/databases are deleted; their exact IDs remain in inventory for delayed account analytics.

Reversal additionally refuses restoration to a former holder with an unrelated legacy claim, pending explicit identity review. A five-minute expiry can be edited into a fresh, current ownership review while retaining the explanation and requiring confirmation again. Browser integration covers expiry, post-commit lost responses, private holder conversations and mobile layout. The current request's publication can be retried through its identity history; general unattended background retries remain in the later resilience batch.

Rollout now permits **owner-only** request decisions: `requestsEnabled:true`, `requestsAdminOnly:true`, `requestTransfersEnabled:true`. Ordinary staff still have no Requests panel. This is the acceptance environment, not general staff activation. No real staff ownership changes or report conversions were performed for testing. Existing password migration/passkeys, calendars and contact syncing remain unchanged.

Remaining before Batch 6 is complete:

1. Owner iPhone acceptance of Requests → review/send/reply/resolve/reopen; no real staff reassignment is needed merely to test.
2. After acceptance, disable `requestsAdminOnly` for the controlled general messaging release, verify the ordinary Account panel and private scope, then observe normal traffic. This is a small rollout/configuration step, suitable for Medium.
3. Keep the existing daily limits and gates in place; a genuine lack of protected headroom must still defer new writes. All actual test usage remains counted. The standalone CLI projection does not apply verified maintenance classifications and is not the runtime grant decision; its independent STOP must not be silently relabelled GO.

## Staff release after iPhone acceptance

The owner reported “Tested seems to work”, satisfying the iPhone acceptance gate. General staff activation now sets `requestsAdminOnly:false`; Requests and reviewed owner decisions remain enabled. Staff use Account → Messages and requests. Ownership previews/commits remain server-enforced owner-only, with recent authentication and transaction guards unchanged.

Focused API/connected calendar tests and the isolated mobile browser suite passed again. Explicit staff attempts to preview or commit an owner decision return 403; private conversation isolation remains covered. No real staff account was impersonated or messaged for release testing.

Pre-release settled analytics through 11:44 UTC on 18 October: account 351,869 reads / 48,569 writes, including disposable test costs; production 296,149 reads / 8,431 writes. The runtime grant has no stop reason and remains valid; read-only verification returned 3,633,639 / 48,782 maximum reads/writes with 151,475 / 19,214 allocated. The independent CLI forecast still reports STOP because it does not classify verified one-off maintenance. No daily cap, reserve, outstanding receipt or actual usage was relaxed.

Batch 6 implementation and acceptance gates are complete for this release. Normal staff use remains the observation stage; genuine daily safeguard deferrals stay active. General unattended background retries remain in the later resilience batch.
