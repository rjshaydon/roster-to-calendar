# Batch 6 — identity requests, admin support and account sessions

Planning date: 15 October 2026, Melbourne.
Status, updated 16 October: product plan approved; Supabase password fallback plus passkey evaluation selected after the failed local KDF feasibility gate. The Sydney Free project and replacement encrypted configuration are verified, with rollout disabled. Two isolated synthetic provider lifecycle runs passed and removed their accounts; one captured native trace used 7 ms CPU. Server-side session/migration coordination and existing state-route cookie support are prepared and locally tested. No production schema migration, staff transfer, application deployment or browser password-storage change has occurred. See the preparation document for the exact remaining activation gates.
Messaging scope confirmed by user: identity requests and admin support, not clinician-to-clinician chat.
Ownership confirmed by user: keep existing ownership unchanged while a dispute is reviewed.
Session clarification: the work environment is a personal Citrix desktop secured by tap-on/tap-off, not a shared browser. The user approved 30-day inactivity expiry and a six-month outer reauthentication deadline, as well as implementation and native password testing. Security is first priority, CPU efficiency an immediate second; weakening password storage to meet Free CPU is not authorized.

## Current evidence and scope

- The 4 October batch plan contains an outline, not an implementation design.
- The original PRD describes disputes, replies, transfers and temporary aliases as intended features. It is not evidence that those journeys were already implemented before lockdown. Its old site/access and Gmail assumptions are superseded by current production behavior.
- `claimSelectedRosterName` currently stops with “Conflict notification is still to be added.” `reportRosterIdentityIssue` records an admin issue, including a hard-coded MMC source, rather than a structured dispute with participants and a reply thread.
- Batch 2 supplies durable people/aliases, reviewed operations, account linking, reversal, historical evidence and affected-calendar publication. Reuse those mechanisms; do not build a second identity system.
- The browser currently keeps a password in session/local storage and includes it in routine requests. Server verification uses salted SHA-256. These are separate issues: secure session cookies remove routine password handling, but do not alone modernize password storage.
- Invitations already use Postmark. Do not introduce Gmail, another paid provider, email verification, password-reset email or external identity hosting without a separate decision.
- Preserve current site/grade access rules, clear automatic signup matching, Creator account switching, shared/private calendar boundaries, historical grades, settings and subscription URLs.

## Proposed user interface

### Ordinary users

Keep the existing Account modal. Add a collapsible **Messages and requests** section and a small unread indicator on Account only when needed. Existing messages appear from local account-scoped cache immediately; refresh on opening, login and foreground return. Do not add another main navigation item or background polling loop.

When an otherwise eligible roster identity is occupied, show **Ask admin to review**. Explain that the request concerns a name already linked elsewhere without revealing the other account's email. The form contains the roster name/site, a short explanation and a preview before submitting. Never change ownership merely because a dispute was submitted.

Users can also choose **Contact admin** for app support. Each request has a single reply thread and a plain status: **Open**, **Waiting for you**, **Waiting for admin**, or **Resolved**. Closing a report and rejecting an identity request have explanatory text rather than conflating rejection with technical failure. Allow an explicit reopen or a new linked request when circumstances change.

### Creator

Add **Requests** as a collapsible section in Admin → Users, following the existing mutually exclusive sections and single-column review layout. Keep parser errors in Parser and system diagnostics in System. Show an open/unread count only when nonzero.

Opening a request hides the list while showing its conversation, relevant accounts/roster aliases and existing site/term/grade evidence. History is collapsed. Actions: **Reply**, **Resolve**, **Review identity**, and an appropriate reviewed correction/transfer. Preserve request context when opening the existing identity review; provide **Back to request**.

A transfer preview states exactly which person, aliases/accounts and calendar access will change. A correction of one wrong site alias must not transfer an entire unrelated person. User-specific custom events/settings are not copied from the previous owner. Do not expose another user's private settings in a support conversation.

## Ownership and dispute rules

1. Resolve requests against the current canonical person, not only a spelling. A different alias must not bypass an existing ownership conflict.
2. Preserve automatic linking of clear, genuinely unoccupied identities. Ambiguous, occupied and duplicate-account cases go to review. A support message is not proof of identity.
3. Bind server-side requester identity from the authenticated session. Clients cannot choose an author, recipient, owner role or approval state.
4. Keep current ownership while a case is open. Notify the relevant existing claimant and Creator inside the app. A participant cannot see other cases or unrelated identities.
5. Resolve using a reviewed Batch 2 operation with actor/reason, before/after versions, idempotency key and case linkage. Account/alias ownership checks, case status and version checks must be in the committing transaction, not only the preview.
6. Concurrent admin edits, simultaneous claims, repeated submission and lost-response retries must not duplicate a case, transfer twice or overwrite a more recent resolution.
7. Invalidate affected app/profile/feed artifacts and contact/access projections using current identity publication rules. Old access must stop at commit even while replacement artifacts are preparing; a pending rebuild is not permission to serve the former owner's private calendar.
8. Notify participants after the durable decision. A message failure cannot undo a committed transfer; an uncertain outcome remains inspectable and resumable. Reversal is reviewed against intervening edits, not a blind snapshot overwrite.

Temporary dual access is outside this batch: the user confirmed unchanged ownership during review. If requested later, design an explicit, expiring grant separate from the canonical alias mapping; do not merge different people or inherit SMS/Director privileges. Its permission and expiry tests become required before enabling it. This replaces the old PRD's ambiguous “duplicate alias” proposal.

## Backend and data design

Prefer the existing D1/R2 stack, without a new real-time service. Add migrated, indexed tables for requests, participants, message sequence/history, resolution references and compact recipient counters. Exact schema/indices are checked with realistic query plans before migration. No runtime DDL or global roster/event scans.

- Requests identify their type, requester, canonical person/site alias where relevant, current state, version and resolution operation. Support cases need no forced roster identity.
- Participants determine message visibility and maintain a monotonic last-read sequence. Mark read when a thread is opened; do not write a receipt per render or per message in a polling loop.
- Messages are plain text with bounded lengths; initial release has no attachments, rich HTML or public user directory. Escape content and rate-limit submissions/replies.
- Inbox and admin lists use cursor pagination, initially 20 records. Conversations use bounded pages. Update unread/open counters atomically with explicit actions; do not count every user's messages on each login.
- Reuse existing login metadata for a compact unread summary where it can remain bounded. Message failures must not block calendar rendering. New replies appear on login, opening Account/Admin and foreground return; live delivery into a continuously open background-free view is not promised in this scope.
- Keep old identity reports inspectable. Convert only clearly identified reports to support/request records; do not guess owners or transfer names from old free text. Parser issues stay where they are.

## Session design and alternatives

**Recommended:** opaque random session tokens in host-only `Secure`, `HttpOnly`, `SameSite=Lax` cookies, with only token hashes stored in D1. Check session expiry/revocation and current account authority through indexed lookups. Prefer joining the existing account lookup; if a separate query is needed, budget one indexed session lookup per authenticated application request. Public roster fingerprints and existing bounded contact-capability reads must not gain routine D1 authentication reads.

This is preferred over self-contained long-lived JWTs: immediate revocation and current account privileges are simpler to enforce. Signed short-lived access cookies can reduce reads, but still need a revocation/role-change strategy; do not silently accept stale privileges to save a row read. A new managed authentication provider is an alternative only if the password feasibility gate cannot meet the current constraints and the user approves it.

Preserve the original authenticated Creator as actor during account inspection. A target-account field never becomes an authorization credential. Migrate state, identity, calendar, comparisons, settings and admin APIs through a common verifier/request helper. Automation bearer tokens and external subscription tokens retain their separate authentication paths.

Cookie authentication adds exact same-origin validation and CSRF defenses for browser mutations. No wildcard credentialed CORS. Rotate session identity after authentication and privilege/risk changes. Invalidate sessions on logout, account deletion and password reset/change as appropriate. Add **Sign out other devices** after reauthentication. Do not expose session tokens to JavaScript, URLs or logs.

Approved UX: retain **Keep me signed in on this device**, default off, with the same option for a personal device and the user's personal Citrix desktop. Remove the blanket 16-hour work-computer rule. Tap-off secures the desktop but does not revoke an application session; explicit app logout still revokes it. Browser close alone is not reliable logout because browsers can restore session cookies. Offer a nonpersistent choice for genuinely shared or temporary browsers.

Distinguish inactivity expiry from a fixed reauthentication deadline. Approved remembered-device policy: expire after 30 days without meaningful use, extending the inactivity deadline as the person uses the app; retain an outer limit of 180 days from the last password/passkey authentication. Thus daily users are not asked to sign in every 30 days. These figures are product choices, not OWASP-prescribed durations. Sensitive account/ownership/admin actions require appropriately recent authentication. Silent session-token rotation does not reset the outer authentication deadline. A no-fixed-deadline alternative permits a stolen, continually used credential to remain valid until revoked and was not chosen.

Implement remembered-device persistence separately from the lifetime of individual rotated session tokens. Define renewal and server-side inactivity enforcement precisely, preserve the whole-shift On shift workflow, and do not let automatic contact/roster polling alone keep a remembered device alive forever. No per-request or per-keystroke activity writes.

Session renewal should be coalesced across requests/tabs and never write last-seen on every request. Choose a tested rotation/grace protocol that handles parallel requests and response loss without unlimited token reuse. Bound expired-session cleanup with an expiry index; never enumerate every account at login.

On a cold restart, reuse one minimal authentication/access envelope before exposing a shared or logged-out user's private cached content; do not fetch/parse rosters as part of authentication. Already validated views remain cache-fast. Measure any cold-start change against the existing baseline rather than promising impossible zero-latency revalidation. Logout/expiry/account switching must not leave the previous person's data visible or allow a late response to repaint it.

## Password feasibility gate

Before choosing the final password implementation, benchmark a maintained password-hashing implementation and realistic parameters in the native Free-tier runtime, using synthetic passwords only. Local Node success is not proof of Workers compatibility or CPU cost. Check registration, sign-in, invitation acceptance, password change and concurrent attempts.

Explain the alternatives before activation: (1) local strong password hashing, with Argon2id preferred and native PBKDF2 a compatibility candidate at appropriate security parameters; (2) a maintained external authentication provider, which moves costly password verification outside Workers but adds a provider, migration, availability and quota/cost dependencies; (3) passkeys, which avoid storing reusable password secrets for passkey authentication but require enrollment, recovery and actual Safari/iPhone/Citrix compatibility checks. A password fallback still needs strong storage. Provider pricing/free quotas and capabilities must be verified before recommending a specific service. A client-generated password hash must not be treated as a shortcut: if accepted directly, it becomes a reusable password-equivalent secret.

Use versioned password records and upgrade existing credentials after successful reauthentication, without bulk password export, forced wholesale resets or password processing in GitHub Actions. Routine cookie-authenticated requests must not repeat password hashing. Apply abuse protection before expensive verification, and ensure rejected guesses cannot cause an unbounded stream of D1 counter writes. Verify any platform rate-limiter binding is available within the existing plan/configuration limits before choosing it.

Do not retain fast SHA-256 for new password storage, reduce security parameters merely to fit Free CPU, write custom cryptography, or add a paid/external service silently. If an appropriate local password verifier is not viable, present concrete alternatives (for example managed authentication or passkeys) and their migration/cost implications before activating new password storage. The secure-cookie migration can be assessed independently, but Batch 6 must not conceal an unresolved password-storage limitation.

## Migration, rollout and rollback

1. Record current startup, authenticated-request and D1 costs. Verify sessions/password feasibility, indices, environment/binding headroom and final user choices before building the complete feature batch.
2. Build backward-compatible schema and modules locally/isolated preview, then integrate the complete request/reply/resolve and session/account journeys. Use additive migrations, not production data rewrites.
3. Perform a short higher-effort security/correctness review before auth activation. No real claim transfer, password reset, email or notification is manufactured against clinicians for testing.
4. Stage the compatible server verifier before the new client. Initially allow a bounded legacy-password-to-cookie exchange for existing signed-in clients. Remove stored browser passwords only after a valid cookie session is established; clear in-memory copies and stop including them in routine bodies. Avoid an authentication-loop failure if exchange fails.
5. Gate compatibility by an explicit migration window. Old cached clients must get a clear update/sign-in path when legacy routine authentication closes. Do not silently keep password-in-every-request forever or automatically fall back to it after cookie revocation.
6. Preserve existing invitation and phone-feed behavior, settings and private export tokens. Existing Postmark invitations are not a reason to send new categories of mail.
7. Verify a real signed-in Safari/iPhone journey with the new assets actually loaded, not just static HTTP success or a warm old page. Add cache versioning and a full startup/import check to prevent a repeat of the missing clock import incident.
8. Observe normal traffic and roll back an independent feature boundary if necessary. Schema remains additive. An emergency rollback must not restore invalid sessions or silently regress to plaintext browser password persistence after credentials have been migrated. Retain a session-compatible rollback release before closing legacy authentication.

## Acceptance

- User: clear match still automatically opens their own calendar; contested match creates one durable request; replies, unread state and resolution persist across devices.
- Admin: correct scoped list/detail, term/grade evidence, reviewed correction/transfer/rejection, preserved reason/history and safe reversal. Duplicate accounts are reviewed rather than automatically merged.
- Authorization: deny cross-user threads, forged authors, stale previews, CSRF, session fixation, expired/revoked sessions, role changes and shared-browser cache leakage. Admin viewing never impersonates the target for authorization.
- Lifecycle: existing and new signup/login, invite acceptance, password change, explicit logout, other-device revocation, browser restart, expiry, offline/resume, failed renewal and parallel tabs. Private subscription URLs remain compatible, and ownership changes invalidate the right feeds.
- Delivery: source rosters, historical grades and personal custom data remain correct. Identity publication failures retain honest state without exposing revoked content.
- Cost: indexed reads only; no work while typing; no new global polling, per-render writes, whole-table count loops, repeated hash work or roster scans. Measure at least representative 100- and 500-user traffic envelopes plus native cold/warm auth and request workloads. Respect all existing D1 admission/receipt protections and Free-plan constraints.
- Startup: exercise the actual imported browser dependencies and a cold signed-in startup, both calendar and On shift, plus the new request view. One inbox failure must not blank the application.

## User choices and effort

- Confirmed: identity requests and admin support.
- Confirmed: keep ownership unchanged during review; no temporary dual-access feature in this batch.
- Clarified: work uses a personal Citrix desktop; do not classify it as a shared browser or impose the earlier 16-hour cap solely because it is at work.
- Confirmed: remembered-device policy, 30 days of inactivity with a 180-day outer reauthentication limit; same persistence choice for personal devices and personal Citrix browsers.
- Technical gate tested: Argon2id at 19 MiB/two passes used 124 and 142 ms of native CPU; PBKDF2 SHA-256/600,000 and SHA-512/220,000 were rejected by the runtime's 100,000-iteration ceiling. See `password-feasibility-2026-10-15.md` and its sanitized native evidence. Do not weaken settings or treat HTTP 200 as CPU acceptance.
- Confirmed architecture direction: remain on Cloudflare Free; Supabase for migrated password authentication/fallback and passkey evaluation. The user created `roster-to-calendar-auth` on Supabase Free in Sydney (`ap-southeast-2`), project `nncmejyfkujpfuivlart`, verified in Safari on 16 October. Do not use the earlier Mumbai project.
- Credential handling: the user replaced/revoked the exposed default key. The replacement is saved as encrypted `ROSTER_AUTH_CONFIG`, disabled. The superseded revoked config was retained under `ROSTER_AUTH_CONFIG_REVOKED` because the dashboard cannot directly overwrite concealed values. Read-only and synthetic account tests passed; the temporary Workers and synthetic accounts were removed. Never emit entire credential-bearing accessibility states, even with a redaction regex.
- Email proof: existing app password possession is not proof of email ownership. Public provider settings show email autoconfirm is off. The migration foundation creates an unconfirmed provider account and refuses credential retirement until verified provider proof is available. Configure and test confirmation/recovery delivery using synthetic accounts before any clinician migration. Do not silently mark legacy emails verified.
- Passkeys: Supabase currently documents native registration/sign-in APIs, but labels them experimental. Keep password fallback independent; verify actual project support, client version and real iPhone/Android/Citrix journeys before enabling an optional passkey prompt. Do not introduce a separate WebAuthn implementation by default.

Use **Medium** for implementation against this plan, with a focused **High** review of authentication, access transitions and migration before production activation. Light/Low suits later wording/layout adjustments, not the account/session migration. Extra High is not required for the routine implementation pass.

The iPhone displaced bottom bar, natural MCH update observation and MMC next-term source availability remain separate tracked follow-ups; do not silently mix provider or mobile navigation changes into this authentication batch.

## Primary references checked

- OWASP session management: https://cheatsheetseries.owasp.org/cheatsheets/Session_Management_Cheat_Sheet.html
- OWASP password storage: https://cheatsheetseries.owasp.org/cheatsheets/Password_Storage_Cheat_Sheet.html
- Cloudflare Web Crypto: https://developers.cloudflare.com/workers/runtime-apis/web-crypto/
- Cloudflare Workers limits: https://developers.cloudflare.com/workers/platform/limits/

These references support the security/platform constraints; session durations and UI defaults are product proposals, not requirements quoted from OWASP.

## Session implementation checkpoint — 16 October

- Prepared migration 0046 creates indexed hashed-token sessions; role/password changes and account deletion revoke access, and a database trigger prevents legacy clients restoring retired weak password hashes. No live migration applied.
- The gated session endpoint supports password authentication for exact provider identities, status, logout, recent reauthentication, signing out other devices and explicit foreground activity renewal. Admission must use a staged `AUTH_LOGIN_LIMITER` binding before D1/provider password work; missing binding fails closed. Production Pages binding availability/configuration is not yet accepted.
- Session cookies are host-only, Secure, HttpOnly, SameSite=Lax. All session POSTs require an exact HTTPS Origin; no credentialed wildcard CORS. Existing state/calendar routes accept verified cookies while retaining Creator as actor and current account/claim checks. Invalid cookies never fall back to body passwords. Public/contact capability paths retain their existing early gates.
- Inactivity renewal is coalesced to at most one successful update per six hours and only called by the explicit activity endpoint. Normal account/status/calendar/contact reads never renew. Tokens remain stable across parallel tabs and lost renewal responses; fresh authentication replaces the token, and security changes revoke it. This refines the earlier unspecified periodic rotation/grace proposal: no silent periodic token rotation is introduced. Authentication/180-day deadlines never slide.
- The existing-account coordinator verifies the old password before reserving a stable UUID. A lost provider response resumes by exact UUID; no email-based password reset/overwrite. Email confirmation plus provider password proof are both required before transactional hash retirement. Delivery readiness and a finite migration deadline are separate closed gates.
- Local suites pass for these behaviors and existing signup/access/startup containment. These are server-side integrations: browser transport/storage removal, complete signup/invite/recovery/password-change/deletion journeys, real email delivery, full native session/D1/browser acceptance and the focused security review remain required before rollout. Messages/requests UI and passkey evaluation remain separate unimplemented Batch 6 scope.

## Connected implementation checkpoint — 16 October

Browser cookie transport, email confirmation/recovery, strong signup/invitations, password change, other-device revocation and durable provider deletion are now implemented together locally. The private admission Worker is staged; the production app/config remains disabled and unmodified. Native synthetic provider lifecycle passed. After focused context optimizations, the final 100-user/200-request test had no failures; 28 sampled traces were 0–9 ms, and ordinary context reads used at most four rows/no writes. These are bounded synthetic results, not full production/device acceptance.

See `supabase-authentication-preparation-2026-10-16.md` (latest checkpoint) and sanitized `supabase-authentication-connected-2026-10-16.json` for exact scope, test coverage, removed disposable resources and remaining gates. Next connected batch: verify actual email allowance/delivery, focused High security review, browser/roster-backed acceptance, explicit canary/rollback and Free-provider inactivity handling. Do not activate global migration in lieu of these steps. Messaging/disputes and passkeys remain later Batch 6 work.
