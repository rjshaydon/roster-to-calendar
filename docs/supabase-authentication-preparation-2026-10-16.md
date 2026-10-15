# Supabase authentication preparation — 16 October 2026

## Verified setup

The user's new `roster-to-calendar-auth` project is healthy, Free, and in Sydney (`ap-southeast-2`). Its approved origin is `https://nncmejyfkujpfuivlart.supabase.co`. The earlier Mumbai project is not an authentication destination. Read-only public Auth settings returned HTTP 200 with email authentication enabled, signups enabled and email autoconfirm disabled. No private API key was revealed, stored or transferred, and no staff/test Auth user was created.

## Prepared code

- `functions/_lib/supabase-auth.js`: server-only provider adapter. Pins the Sydney origin, accepts only modern publishable/secret key formats, blocks local outbound traffic, refuses redirects, bounds provider responses to 64 KiB and times out after ten seconds. Provider error bodies and tokens are not reflected in error messages or logged. Sign-in verifies the exact expected provider UUID and email; migration never overwrites a provider account based only on matching email. New API keys use `apikey`, not a JWT Bearer header.
- `functions/_lib/authentication-migration.js`: reserves a stable external UUID against the current legacy credential revision. A lost response/retry cannot select a different account. Final provider readiness and old-hash retirement run in one D1 transaction after verified provider proof. Concurrent local password changes or account deletion prevent retirement. Repeated completion is idempotent. No role, roster claim, historical grade, feed token or personal calendar setting is changed.
- `migrations/0045_account_authentication_links.sql`: additive indexed mapping table only, no backfill or automatic migration. Not applied to any live database.
- `functions/api/state.js`: rejects an existing account with blank password fields. The former legacy login path would adopt an arbitrary submitted password when those fields were empty; that would be a critical bypass after credential retirement. The new regression test exercises the real handler after a simulated migration. This guard is not deployed yet.

The adapter and migration primitives are intentionally not exposed as API routes. They are not a complete login/session implementation. In particular, no code yet sends a real staff password to Supabase or retires a real legacy hash. Provider secrets are not in source, tests or this document.

## Validation

`npm run test:authentication-foundation` passed with mocked provider responses, actual SQLite schema/migrations and the actual legacy login handler. It covers approved-destination enforcement, local isolation, credential/input bounds, redacted provider/network failures, incomplete responses, exact identity binding, unconfirmed/disabled users, safe account creation, pending migration retries, transaction rollback, competing password changes, idempotent completion, account deletion and the legacy reinitialization bypass. SQLite query plans use indexed account/mapping lookups. No external network or real credentials are used by this suite.

`npm run test:bounded-identity` passed, covering existing signup matching, claims and account workflows. The broader `test:fixtures` stops at its existing source-text assertion at line 922 (“automated corrections must diff against the stable active file identity”); both that assertion and its target `functions/api/automation/derived.js` are identical to HEAD. This is not an authentication failure, but the broad suite is not green and must not be reported as passing. No native CPU cost or complete browser integration acceptance is claimed by these local tests.

## Next integration steps

1. After specific credential-transfer approval, copy the existing secret directly from the authenticated dashboard into encrypted Cloudflare server configuration, including the isolated test environment. Keep rollout disabled. Do not create a new key or grant wider access as part of this step.
2. Verify the private admin and ordinary sign-in API behavior with synthetic accounts only; provider documentation and mocks are not live proof. In particular verify caller-selected UUID handling, modern API key headers, email confirmation state, rate limits, uncertain responses and resumed migration. Do not overwrite pre-existing provider accounts to resolve conflicts.
3. Configure confirmation/recovery delivery and its existing Postmark allowance. Creating an unconfirmed account with an admin API does not itself complete email confirmation. Test the explicit confirmation flow and prove fallback availability under Free's inactivity policy; do not manufacture keepalive traffic to evade that policy.
4. Build cookie sessions and the common API verifier/client transport before enabling any hash retirement. Remove routine password bodies/storage only after session exchange succeeds; migrated accounts must never fall back to the legacy password initializer. Finish signup/invite, password change, account deletion and secure recovery integration, including inactive-account retirement policy. Provider deletion/revocation must be reconciled with local account deletion and session revocation.
5. Supabase's passkey API is currently **experimental** and requires explicit SDK opt-in. Registration requires a confirmed account. Test it as an optional journey with reliable password fallback; verify browser credentials/tokens stay out of persistent JavaScript storage and preserve the approved idle/outer expiry and recent-authentication rules. Test iPhone, Android and the real Citrix environment.
6. Run the planned security review, native full-route CPU/D1 tests, browser cold startup/account switching, and isolated canary rollout. Only then activate migration for real accounts. Batch 6 request/reply/dispute UI remains part of the approved larger plan and is not implemented by these foundations.

## References checked

- [Supabase API keys](https://supabase.com/docs/guides/getting-started/api-keys)
- [Admin account creation](https://supabase.com/docs/reference/javascript/auth-admin-createuser)
- [Password sign-in](https://supabase.com/docs/reference/javascript/auth-signinwithpassword)
- [Migration patterns](https://supabase.com/docs/guides/platform/migrating-to-supabase/auth0)
- [Experimental passkey support and confirmed-account requirement](https://supabase.com/docs/guides/auth/passkeys)
- [Production SMTP requirements](https://supabase.com/docs/guides/auth/auth-smtp)
