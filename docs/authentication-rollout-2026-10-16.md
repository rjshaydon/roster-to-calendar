# Password authentication rollout — 16 October 2026

## Current status

The password/session implementation is local and production activation remains disabled. It follows the user's seamless policy: no confirmation emails; new passwords stored by Supabase; existing accounts migrated after successful existing-password verification; optional passkeys are later work. The supplied password is transmitted over TLS for explicit authentication only, not ordinary calendar/contact requests. Source rosters, account role, claims, historical grades, settings and feed identifiers retain their existing authority.

`npm run test:authentication-journey` now connects the real browser transport to the real auth/state handlers and actual SQLite migrations, with only provider HTTP responses mocked. It tests wrong-password denial, exact-subject migration, hash retirement, cookie/account envelope, roster claim/feed preservation, no-mail signup straight into the envelope, normal reads without provider work or writes, paused migration with successful returning-provider login, logout and replay denial. It is included in `test:account-browser`.

## Rollout controls

These properties live inside the existing encrypted `ROSTER_AUTH_CONFIG`; never write actual credentials to source or chat.

- `enabled` and `sessionsEnabled`: keep enabled once any account has migrated. These are not an emergency migration-pause mechanism.
- `migrationEnabled:true`: explicit admission for existing weak-credential migrations. Missing/false fails closed for migrations, while ready provider accounts can still sign in.
- `migrationPaused:true`: stops further legacy migrations without disabling ready-account authentication or existing valid cookies.
- `migrationEmails`: optional array of at most 100 explicitly selected email addresses. Omit for general rollout; empty means no new migrations. Invalid configuration cannot grant broader migration access. The cohort selector is a rollout control, not authentication.
- `legacyMigrationUntil`: optional finite deadline for a deliberately bounded rollout. If supplied, invalid/expired blocks new legacy migrations. Omit for the user's ongoing migrate-at-next-login policy; no artificial expiry is necessary.
- `signupEnabled`: strong provider signup independently enabled. Disable during the first existing-account canary, then enable for all new users when accepted. No state route can create a fresh weak-hash account while cookie mode is active, including the old missing-Creator bootstrap path.
- `signupMonthlyLimit`: independent bounded creation admission, default 1,000. This is not email capacity.
- `emailDeliveryReady`: remains false for this password rollout. Postmark is not a signup/migration requirement. Optional email recovery/invitations require their separate acceptance before enabling.

During a named cohort rollout, existing non-cohort accounts retain legacy state authentication. Every supplied session cookie, including malformed/expired cookies, stays on the cookie verifier and cannot select password fallback. New account/invitation requests cannot use compatibility mode. A migrated account whose email is outside the cohort still gets the secure upgrade path at password login; blank hashes never accept new passwords or restore weak hashes. An exact current provider identity remains usable when migration is paused or its cohort changes.

## Activation sequence after review

1. Focused High security review completed on 17 October; see `authentication-security-review-2026-10-17.md`. Preserve the credential-version, expected-account and policy-compatibility fixes. Medium is appropriate for the remaining staged acceptance and rollout work.
2. Run the additive schema and compatible app release in an isolated preview with private synthetic fixtures and no live staff emails. Verify real provider signup/migration, cold/returning login, cookie attributes over HTTPS, invalid-cookie refusal, stale tabs, password change, logout and browser restart. Repeat native cost checks for the changed seamless paths; the earlier CPU results predate these changes.
3. Check account/provider limits and Supabase Free pause/resume consequences. Do not add provider keepalive polling. Define how an actual paused provider is resumed; existing valid D1 cookies do not need provider requests but fresh sign-in does.
4. Stage additive migrations 0045–0049 and the compatible app/server/service binding with authentication activation still off. Verify existing production behavior before enabling a named owner canary. Actual encrypted settings must be preserved through a private transfer; never emit an entire credential-bearing browser state.
5. Enable sessions and migration for the selected owner cohort, leaving signup off for this brief canary. Test actual Safari/iPhone use and the user's Citrix environment; verify ready-account returning login also when new migration is paused. Existing non-cohort users remain on their current login.
6. Once accepted, omit the migration cohort, enable strong signup and continue existing-account migration on successful login. Offer passkeys only after their separate implementation and browser/provider acceptance.

## Rollback and freeze

Retain a session-compatible application release containing the cookie verifier and retired-hash guard. To stop adoption, pause migration and disable signup; keep session authentication available for migrated users. Never roll back to an old password initializer, restore retired hashes, or disable the session verifier while users depend on it. A failed provider operation must leave legacy hashes intact until proof/transactional retirement; once retired, failure recovery goes through the provider-compatible login path rather than weak-password reconstruction.

## Browser check performed

Safari exercised the actual app assets against a loopback-only, in-memory synthetic fixture: legacy password login, claimed-account shell, cookie-based page reload, Account controls and logout to the entrance. No real staff data/provider credentials or emails were used. The fixture translated a separate synthetic localhost cookie at its boundary; therefore this proves UI integration, not production HTTPS/Secure-cookie enforcement. There were no actual roster files: an empty roster workspace was expected, and real calendar/phone/device acceptance is still required. New-account creation was exercised by the connected transport/API test, not browser credential-entry automation. The fixture tab was logged out/closed and its server stopped.

The Account current/new password inputs had shared ambiguous labels; they now have separate labels, and another account's secure password controls are hidden from the administrator. The state endpoint explicitly refuses password changes through its legacy profile-update path in modern mode; the dedicated authenticated provider action is required.

## Security review update — 17 October

The review adds migration 0049 for credential versions and pending operations, and account expectations for shared-cookie requests. Existing old-password proof cannot issue a new session after or during a password mutation; old reset/invitation links are invalidated too. A pending legacy reset also closes its cohort compatibility path. The only new-password exception to seamless migration is a provider-incompatible existing password: verify the old app password, then ask for a stronger replacement inline, without email. See the review for native D1 evidence and interruption handling.

Before general activation, explicitly test already-open old clients. The current all-account cookie gate can require them to reload/sign in; non-cohort users are protected during a named canary. If that is unacceptable for a general cutover, implement a bounded legacy-state grace period during the Medium staging batch. Never accept a retired old hash, fall back from a rejected cookie, or put provider password work into ordinary roster/contact handlers.
