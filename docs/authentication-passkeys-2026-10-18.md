# Optional passkeys — 18 October 2026

Status: implementation and isolated tests complete; production feature and Supabase provider remain disabled pending activation confirmation and connected acceptance. The additive production preference migration has been applied. The disposable native test Worker and D1 database were removed after evidence collection.

## Approved behavior

Password sign-in remains available. After explicit password sign-in/signup, users without a passkey are offered optional setup. “Don’t remind me again” is an account-wide durable preference. The dialog explains that a passkey can be added later in Account settings. Account → Passkeys allows later enrollment, removal and changing the reminder preference. A normal session resume does not open the prompt. A failed provider check never masquerades as having no passkey.

The prompt runs after calendar paint without awaiting provider work in login/hydration. Passkey UI code loads only when the feature is enabled and supported. Collapsed Account settings do not list keys until opened. No passkey polling, calendar/contact refreshes, roster scans or routine state writes were added.

## Security and deployment

- Provider remains the approved Sydney Supabase project. WebAuthn verification occurs there; app session cookies stay opaque, Secure, HttpOnly and host-only. Supabase tokens are never returned to JS or saved in browser storage.
- Intended RP ID: `rtc.curiousmind.app`; only allowed origin: `https://rtc.curiousmind.app`; display name: **Roster to Calendar**. Preview/default Pages hosts cannot enroll or use production passkeys.
- Existing exact ready provider subject/email link is required to issue an app session. No provider-only identity can be adopted by email. Pending credential mutations refuse issuance; existing atomic session/version checks apply.
- Supabase currently defaults user verification to preferred. Our browser requests required and the app also checks the signed assertion’s UP/UV bits before accepting it; provider verification authenticates those bits.
- Optional enrollment immediately after password login uses a five-minute AES-GCM encrypted host-only HttpOnly cookie containing provider proof. It is purpose/origin-bound and matched against the exact app session hash and credential version. Logout/password reset invalidates that grant. Later setup asks for the current password if no current grant is available.
- Provider challenges are consumed by Supabase. Local challenge binding, expiry, exact origin, bounded JSON (32 KiB), provider destination pinning, response bounds and private admission service remain in force. Password admission stays at five requests per account/minute; anonymous passkey start/finish share a separate 60 requests per public IP/minute bucket, allowing for a shared hospital network. Address values are hashed before the private limiter and are not logged. This route is capped at eight D1 statements.
- If logout/revocation occurs during provider enrollment, revalidation removes the exact new credential rather than accepting it from the revoked session.
- Removing a key needs a recent authentication, confirmation and membership in the authenticated user’s provider key list. Password fallback remains present for all enrolled app accounts.
- Additive migration `0050_account_passkey_preferences.sql` adds one boolean preference to account profiles. No roster/history/user linkage changes.
- Activation uses `passkeysEnabled:true` inside the existing non-secret `ROSTER_AUTH_ROLLOUT`. It defaults off, adds no environment binding and never changes password migration/session flags.
- Supabase calls passkeys experimental. Keep the password fallback and feature gate; no assurance of native iPhone/Android acceptance is claimed before user testing.

## Validation

`npm run test:passkeys` covers origin/preview/disabled gates, exact actor/session binding, grant tamper/expiry/duplicate cookie/key rotation, preference suppression across sessions and unchanged writes, no roster reads, verified identity/pending operation denial, UP/UV enforcement, logout during registration cleanup, cancellation/logout races and provider-token containment.

Existing authentication foundation, sessions/coordinator/credential-race, account-browser/email/lifecycle/journey, admin-user interactions, clinician switching and client request-budget checks pass. The isolated headless mobile browser fixture uses synthetic data only and checks optional setup/dismissal/re-enable/later setup, no reminder when a key exists, logout/provider-failure suppression, no collapsed-setting fetch and mobile layout. It does not simulate Face ID.

Native Cloudflare probe (disposable empty D1, synthetic provider-proof bytes, no Supabase calls), six requests per stage: session issuance plus encrypted grant max **3 ms CPU**; registration session/link checks plus decrypt/reseal max **5 ms**; status reads max **2 ms**. These are component costs, not a connected end-to-end provider acceptance result. Evidence is in `authentication-passkey-native-2026-10-18.json`.

## Deployment checkpoint

- Commit `b40c609b`, pushed to `main`.
- Pages deployment `05cdfc04`, client `20261018-passkeys`, app passkey flag explicitly false.
- Native admission service version `184721e9-4fb9-4143-8d65-1d451d29eae3`; password limit unchanged, separate anonymous passkey bucket prepared.
- Production migration 0050 applied successfully (the only pending migration).
- Live public options confirm general password migration/session support enabled, passkeys/recovery disabled, no named migration cohort. The new passkey endpoint refuses work while disabled.
- Isolated complete-app mobile entrance smoke test passes with no startup errors; disabled passkeys do not download their lazy UI modules.
- Disposable native Worker and database removed successfully; no synthetic provider account was needed for component tests.

## Remaining activation checks

1. Confirm enabling Supabase passkeys with the exact RP values above, then save provider configuration. Browser-control policy requires action-time confirmation for a new security-sensitive sign-in method; the confirmation question is pending.
2. Activate only the app passkey feature; run one bounded synthetic connected enrollment/sign-in/fallback/deletion journey. Remove its exact synthetic account and key. The connected acceptance script is prepared and refuses to create an account while the app flag is off.
3. Ask the owner to enroll their own passkey and test actual Safari/Home Screen Face ID, optional dismissal and Account setup. Credential creation itself is a user action.
4. Record connected/real-device acceptance. If provider checks fail, turn off only passkeys; preserve password authentication and migrated credentials.
