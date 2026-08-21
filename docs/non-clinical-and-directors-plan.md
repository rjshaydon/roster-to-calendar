# Non-clinical accounts and Director access

## Goal

Allow the Creator to create an account for a person who does not appear on a clinical roster, and independently grant Director access to either a clinical or non-clinical user. The contents and behaviour of the Director view will be specified separately.

## Product decisions

- Keep the existing authentication roles (`creator` and `user`) unchanged. `Non-clinical` is an account classification, not an administrator role.
- Store `nonClinical` and `directorViewEnabled` as independent booleans. A director may still work clinical shifts, and a non-clinical account does not automatically receive Director access.
- Only the Creator can set either value through the Admin interface or API.
- A non-clinical account is valid with no roster claims, no selected roster doctor, and no clinical shifts.
- Do not build or expose the actual Director view until its requirements are defined. This change establishes its persisted entitlement and Creator controls.

## User interface

### Admin → Users → Create user account

Add two unchecked checkboxes below the temporary-password field:

- `Non-clinical`
- `Director view`

Submit both values with `adminCreateUser`. Keep `Create and enter account` as the submit action. For a non-clinical account, entering the new account must show a suitable empty state rather than asking the user to select a roster name or upload a roster.

The two checkboxes remain independent; selecting either one must not silently select the other.

### Admin → Users → Current users

Add a third permission checkbox labelled `Director` to the right of `Who/When?` and `At a glance`, aligned on the same row as 'Who/When?' at desktop widths and wrapping cleanly on narrow screens.

The checkbox reflects the stored Director entitlement and updates it optimistically, using the same rollback and status-message pattern as the existing permission toggles. No Director-view navigation or content is added in this phase.

## Persistence and API

1. Add a D1 migration with two `INTEGER NOT NULL DEFAULT 0` columns on `account_profiles`:
   - `non_clinical`
   - `director_view_enabled`
2. Add matching `ensureColumn` calls so local/new databases and migrated production databases converge on the same schema.
3. Thread both flags through account insert/update, account loading, account-list loading, account response preparation, and user summaries.
4. Extend `adminCreateUser` to accept strict booleans for both flags. Persist them before returning the new user summary.
5. Add a Creator-only action such as `setUserDirectorViewEnabled`, following `setUserInsightsEnabled` and `setUserFacilityOverviewEnabled` for validation, persistence, and response shape.
6. Do not expose a standard-user endpoint for changing either flag. Reject attempts to change the Creator account's effective permissions consistently with the existing feature toggles.

## Non-clinical account behaviour

The current login, account-entry, and hydration paths try to auto-match an account with no claims to roster doctors. Guard those paths with `nonClinical !== true` so a non-clinical user is not accidentally linked to a similarly named clinician:

- account creation (`autoClaimMatchedCanonicalDoctors`)
- login fallback (`autoClaimMatchedRosterNames`)
- Creator impersonation/account entry
- client hydration (`resolveCurrentAccountClaims`)

Keep manual claim editing available to the Creator unless later requirements say otherwise; this makes an incorrect classification recoverable without deleting the account. If a roster claim is deliberately added, the account can display those shifts while retaining its non-clinical classification.

Add a non-clinical empty state that confirms the account is ready and does not instruct the user to claim a roster identity or upload a roster. Calendar/subscription code must tolerate an empty doctor key and an empty event set.

## Client state

- Extend server-user normalization with safe defaults for `nonClinical` and `directorViewEnabled`.
- Carry both flags in login and impersonation responses so account hydration does not infer them from claims.
- Track the Director entitlement in current-view state, but do not use it to reveal a view until that view is defined.
- Reset the new current-view values during logout/login failure alongside the existing feature flags.

## Verification

1. Run `npm run check` and the fixture suite.
2. Verify Creator-only API enforcement for creation and Director toggling.
3. Create and enter each combination:
   - clinical, not Director
   - clinical Director
   - non-clinical, not Director
   - non-clinical Director
4. Confirm a non-clinical account remains claim-free through creation, direct login, refresh, and Creator impersonation.
5. Confirm a normal clinical account still auto-claims the correct roster identity.
6. Toggle `Director` in the Users list, refresh, and confirm persistence and optimistic rollback on a failed request.
7. Check desktop alignment and narrow-screen wrapping of all three permission checkboxes.
8. Confirm existing accounts receive both defaults as false and retain their existing `Who/When?` and `At a glance` settings.

## Deferred until the Director view is defined

- Director navigation and screen placement
- Which hospitals or program scope each director can see
- Whether a director can see identifiable staff, schedules, summaries, or exports
- Hospital Director versus Program Director scope
- Any editing, approval, delegation, or audit capabilities

These requirements may require a scoped relationship (for example, hospital IDs plus a program-wide marker) in addition to the boolean entitlement established here.
