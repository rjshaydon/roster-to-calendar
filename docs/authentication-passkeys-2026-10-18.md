# Optional passkeys — 18 October 2026

Status: optional passkeys are active in production following explicit user approval. All 20 connected Supabase acceptance checks passed; the exact synthetic account and provider credentials were removed. The owner subsequently reported “That works” after the requested device check; device acceptance is recorded as user-reported. Password sign-in and general migration remain active. The additive reminder preference migration is applied; disposable native test resources are deleted.

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

## Approved activation and connected acceptance

The user approved the exact provider settings, app activation and one disposable synthetic connected test. Supabase is enabled with RP name **Roster to Calendar**, RP ID `rtc.curiousmind.app`, sole origin `https://rtc.curiousmind.app`. Activation commit `32651324`, Pages deployment `c80d3a57`. Live options confirm production passkeys on, general password migration/session support on, no named cohort, email recovery off; the default Pages host still reports passkeys off.

All **20 connected checks passed** on 18 October at 11:37 am Melbourne: no-email password signup/login, encrypted setup grant, reminder suppression/re-enable, setup without repeat password, exact RP, actual Supabase verification of a synthetic P-256 WebAuthn registration and signed assertion, passkey session/workspace, own-key removal, password fallback and full synthetic account/provider cleanup. This is genuine provider verification using a software-generated synthetic key; it is not a Face ID or Android hardware acceptance test. See `authentication-passkey-connected-2026-10-18.json`.

After cleanup: **60 accounts, one migrated, zero synthetic accounts remaining, zero pending credential changes**. General migration is enabled for everyone; remaining accounts upgrade only on a future successful password login. No staff account, roster, claim or feed was altered by acceptance testing. No email was sent.

### Owner device check and acceptance

Reload once to receive the current options. In the owner app use **Admin → Account → Passkeys → Add a passkey**, or sign out and back in with a password to receive the optional setup prompt. The owner must complete the native credential/biometric prompt. After enrollment, sign out and choose **Sign in with a passkey**. Repeat from the iPhone Home Screen app. Password sign-in stays available throughout.

If passkey provider acceptance regresses, set only `passkeysEnabled:false`; preserve password/session verification and migrated credentials. Native CPU evidence above is component evidence; connected acceptance does not independently measure whole-route CPU.

The owner subsequently confirmed that the requested passkey flow works. This closes the owner device acceptance item without claiming independent coverage of every browser/OS/device.
