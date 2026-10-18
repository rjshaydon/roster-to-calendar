# Batch 6 — identity requests, admin support and secure accounts

Consolidated plan: 18 October 2026, Melbourne. Replaces the earlier mixed proposal/checkpoint text in this file; prior versions remain in Git.

## Current status and agreed scope

**Security implementation and rollout are complete.** Supabase password signup/migration, opaque cookie sessions, password changes, deletion coordination, sign-out/other-device revocation, migration reporting and optional passkeys are deployed. The user reported that the passkey flow works after the requested device test. This records the user's acceptance, not testing of every Android/Citrix/device combination.

General migration upgrades each existing account on its next successful password login; unmigrated accounts are expected during rollout. Do not treat migration of every dormant user as an unfinished coding requirement. Dated legacy routine-auth compatibility closes on 20 October at 11 am Melbourne; retain the tested fresh sign-in/update path and secure session verification. Normal production observation continues.

Evidence and operational details:
- `authentication-rollout-2026-10-16.md`
- `authentication-security-review-2026-10-17.md`
- `authentication-passkeys-2026-10-18.md`
- `authentication-passkey-connected-2026-10-18.json`

**Remaining Batch 6 scope:** private user/admin support conversations, requests about occupied or incorrect roster identities, the admin case queue, reviewed resolution/ownership reassignment and in-app unread/outcome notifications.

Confirmed choices: keep current ownership during review; communication is with the admin, not clinician-to-clinician chat; use the existing Account/Admin surfaces; keep Cloudflare Free and current D1 safeguards. Mandatory email confirmation is not part of signup or migration. Email recovery/invitations remain separately gated and are not prerequisites of this remaining release. Do not enable email delivery or a paid provider as a side effect. Temporary dual access remains excluded.

## Findings from the 18 October code audit

- `claimSelectedRosterName` still stops at “Conflict notification is still to be added” for another account's claim. Replace this dead end with a durable request flow.
- `reportRosterIdentityIssue` stores a generic admin issue with a hard-coded MMC source and no conversation, participants or case status. It is not the required support/dispute system.
- There are no request/conversation/message persistence tables or complete inbox/reply/resolve APIs/UI yet.
- Batch 2 already provides canonical people, site aliases, term/grade evidence, reviewed identity operations, version guards, history/reversal and publication jobs. Reuse these.
- An `account-link` operation is not a complete contested-ownership transfer: existing canonical account links and legacy site/key claims must both be reconciled. Linking, removing a claim and closing a case in separate requests is not an acceptable transfer transaction.
- Current identity commits build and submit their own D1 batch. Add composable, guarded commit preparation so request resolution and ownership/identity changes can commit together, while preserving existing standalone operation behavior/tests.

## Deliverable 1 — requests and private conversations

### User interface

Add a collapsible **Messages and requests** section in Account. Include **Contact admin** for ordinary app support and **Ask admin to review** wherever a contested roster name currently leads to a dead end. Carry the selected site, roster spelling and canonical identity when known; allow “my name is missing/wrong” reports without requiring an invented roster link.

The user reviews the subject/explanation before submitting. On submission show the durable case and its replies/status. Users can reply, see outcomes and reopen a resolved case when necessary. Plain status labels are Open, Waiting for admin, Waiting for you and Resolved, presented for that user's conversation.

Use one-column list/detail presentation. Close/back returns to the list. A private cached open list/detail reappears immediately when unchanged, without replacing a live message draft. A failed request load stays within that section and never blanks the calendar/On shift.

### Privacy model

A case may involve the requester and an existing linked account, but this does not create a shared conversation between them. Each has a separate private channel with the admin. The admin case view labels the audience for each channel/reply; a user cannot see another party's email, private explanation or messages. Contact an existing holder when admin review needs their input; send relevant participants a decision notice after resolution. A submitted allegation alone does not justify automatic ownership changes or broad peer notifications.

Messages are bounded plain text (initial maximum 4,000 characters, subject 120), escaped on display, with no attachments/rich HTML. Submissions/replies are rate-limited by authenticated account using a separate bucket from password/passkey admission.

## Deliverable 2 — admin queue and review workspace

Add **Requests** under Admin → Users, preserving existing mutually exclusive collapsible sections and the QR card below the sections. Show a count/badge only when nonzero. Keep parser issues in Parser and technical diagnostics in System.

The queue offers open/resolved and waiting/unread filters, bounded pages (initially 20), and one selected case at a time. A case shows its private conversation(s), relevant linked accounts, canonical person/site names and existing site/term/historical-grade evidence. History is collapsed. Actions are Reply, Resolve, Reopen and Review identity; the correction flow returns to the case through Back to request.

Admin inspection keeps the authenticated admin as actor; the viewed user's account is never authorization. A support request need not force an identity operation. Unsubstantiated or withdrawn requests can be closed with an explanation and no roster-link change.

## Deliverable 3 — safe identity/ownership resolution

Supported outcomes:
1. Keep the current link, explaining the decision.
2. Correct spelling/identity or separate a wrong site alias using Batch 2 operations.
3. Reassign the reviewed roster identity/access to the correct existing account.
4. Identify duplicate accounts and review the intended linkage without merging their private events/settings or resetting credentials.

Before a change, show exactly which person, aliases and account access will change. Preserve historical grades, raw rosters, unrelated identities/claims, personal settings, custom events, authentication credentials and existing feed token identities.

For reassignment, include canonical `account_people` links and relevant legacy `account_claims` in one bounded scope. Do not assume one spelling identifies the whole person, one case has only one current holder, or an account-link edit alone removes older access. Over-large/ambiguous scopes stop for a narrower review rather than widening a scan.

**Atomic commit:** current case/version/participant guards, exact old owner and alias/version guards, reviewed identity/link/claim changes, operation history, case outcome, publication job and decision-message/unread bookkeeping must commit in one transaction. Reuse/refactor the existing identity statement builder rather than wrapping several independently committed API calls. Retain all current guards in existing standalone identity operations.

Every submitted mutation has a client idempotency key scoped to the actor/action. A repeated request or lost reply retrieves the same outcome; it must not submit a second case/message/transfer. Stale previews must be re-reviewed. Reversal is also reviewed against intervening changes, linked to the case, and cannot blindly overwrite a newer identity or ownership decision.

Current server authorization must reject the former ownership scope after commit, even while updated artifacts publish asynchronously. Existing identity jobs rebuild only affected account/profile/feed views. Do not serve the old owner's cached scope merely because the replacement is preparing. Already downloaded phone calendar copies follow the calendar client's refresh schedule; server revocation is not a promise to remotely erase an offline client's copy immediately.

## Deliverable 4 — durable storage, unread state and lifecycle

Add additive, indexed tables for cases, private channels/participants, sequenced messages/read cursors, resolution references, mutation idempotency and compact per-account inbox summary/version. Exact schema and query plans are checked before migration. Inbox/message pages are cursor-based, initially 20; no global user/roster scans or unbounded thread downloads.

Case state and each private channel's waiting/read state are distinct, so two parties cannot give one case contradictory global “waiting for you” states. Use conditional monotonic read cursors. Mark read only when a thread is actually opened/advanced, not per render. Update message, unread summary/version and case state atomically with explicit actions; unchanged inputs do no writes.

Use authenticated, account-bound in-memory UI caches first. Clear drafts/private caches on logout, expiry and account change; reject late responses through the existing transport generation/expected-account protections. Durable cross-device history lives on the server, not in browser password storage. Do not introduce permanent free-text message storage in localStorage to make the list appear fast.

Bind authorship, role, participants and recipients on the server. Deny cross-user access and arbitrary participant addition. Current admin authority and recent authentication protect consequential ownership changes. Normal support replies do not demand a password repeatedly.

Account deletion removes that account's private channel/free-text content and participation, clears client state and leaves other parties' private channels protected. Keep the minimal reviewed identity-operation audit under existing rules. A deleted participant or changed owner makes a pending transfer stale and requires review; deletion must not auto-approve it. Define these cleanup writes in the account-deletion integration and test lost responses.

Keep old reports inspectable. Convert only clearly identified legacy identity reports with a stable deduplication reference and preserved provenance; unknown source/canonical person stays unknown. Do not convert old free text into ownership proof or treat the old hard-coded MMC label as reliable. Ordinary parser reports remain in their existing location. No broad roster/history migration is needed.

## Deliverable 5 — notifications without background traffic growth

Use a small unread badge on Account and the admin Requests section only when needed. Counts derive from the compact recipient summary, not conversation recounts on login or every calendar request.

Refresh the summary on sign-in, a genuine return to the app and opening the relevant account/request surface. Prefer adding the summary to an existing coalesced status/context response where its measured cost fits, rather than introducing a separate frequent request. Keep the calendar's first paint independent. If a separate summary call is necessary, perform it asynchronously after paint, coalesce identical calls and apply a short freshness window.

Only fetch lists/messages when their revision changed or the user explicitly opens an uncached page. Preserve draft/focus/scroll and skip unchanged repaint/storage writes. Do not add a periodic message poll, per-keystroke search query, message work to contact fingerprints, or inbox work to every calendar request. Continuously open views receive new content on the agreed open/foreground/sign-in checks; real-time chat delivery is outside scope.

Decision notifications become visible only after the decision is durable. A later publication failure remains a separate retriable job, not a reason to repeat a successful transfer or declare an uncertain transfer successful.

## Implementation and release sequence

Proceed as one coordinated remaining feature release, with independent enable/stop controls for support requests and consequential ownership mutations:

1. **Prove the sensitive boundary early:** implement/test the small guarded ownership-commit path against native D1 using synthetic data. Measure the real supported scope before investing in UI polish. If preparation exceeds Free CPU, use the existing external preparation runner and keep a small server-verified atomic commit; do not split an ownership transfer into unsafe partial commits or weaken guards.
2. **Build the complete request feature:** indexed persistence/APIs, session-bound transport, private user/admin list/detail/reply UI, unread summaries and scoped legacy-report/lifecycle integration.
3. **Connect reviewed resolution:** composable identity/claim transaction, case/version/idempotency guards, exact-scope preview, durable decision messages and affected publication/reversal behavior.
4. **Review, test and release together:** focused higher-effort review of authority/privacy/transfers, one coordinated test/cost pass, additive production migration, deployment with gates off, bounded synthetic live acceptance, enable the accepted features and verify normal calendar/login/On shift behavior.

No per-button approval cycle or code change is needed to repeat already accepted password/passkey migration. Do not reset real users, transfer clinician claims, contact real staff or send emails as artificial tests. Exercise identity changes with isolated synthetic people/claims; use the owner and disposable accounts only for approved live support journeys. Stop before production activation if security or measured D1/CPU gates fail.

## Acceptance: Batch 6 is complete when

- A contested name creates one durable request and changes no ownership while pending; a normal support request also works end to end.
- User and admin can read/reply/resolve/reopen across sessions/devices, with correct audience/privacy, stable drafts, accurate unread state and no irrelevant empty badges.
- A reviewed correction/transfer commits exactly once, survives response loss, rejects stale/concurrent changes, preserves unrelated data/history and can be safely reviewed for reversal.
- Old server access/feed scope is invalidated; the correct replacement views publish through current bounded jobs. An interrupted publication cannot repeat the transfer or leak revoked content.
- Negative tests deny forged actors/recipients, cross-party channels, admin-viewed-account confusion, revoked sessions, CSRF, stale previews and duplicate submissions. Deletion and role changes are covered.
- The real complete app imports/starts and calendar/On shift remain usable when request APIs fail. Changed messages do not trigger unrelated calendar/contact rebuilds.
- Native representative cold/warm interactive requests meet the Free CPU target. Normal list/detail paths are small; ownership transactions retain explicit statement/row budgets and existing account-wide admission. Use indexed query-plan/row-cost tests at representative 100/500-account scale, including a large historical message set. Unchanged render/input/summary checks add no writes, no global scan and no polling loop.
- Record migrations, source/deployment, exact enabled scope, actual costs, synthetic cleanup and rollback. A stop switch for requests must not disable authentication or undo committed identity decisions; schema and pending publication jobs remain recoverable.

## Effort and remaining decisions

**Medium** is appropriate for implementation against this consolidated plan. Use a focused **High** review for the ownership/privacy/transaction boundary before activation. Low suits subsequent isolated wording/layout fixes. There is no need to repeat the completed authentication redesign or keep all routine implementation on High.

No new provider, paid tier, email approval or material product decision is required with these defaults. Privacy is tightened to separate admin/private channels. If live chat, temporary dual access, email recovery/notifications or merging private account data is requested later, treat it as separate scope and obtain the relevant decisions.

Separate remaining batches retain their original order: 7, 3, 4, 5, 8. The displaced iPhone bottom bar, broader diagnostics/retries, Casey ingestion and reporting requirements are not silently absorbed into Batch 6.
