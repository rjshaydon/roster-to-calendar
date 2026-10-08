# Account reliability fixes — 8 October 2026

Production fix: 56f57c2d.

## Evidence

- D1's settled UTC-day sample at approximately 02:00 AEDT: 448,340 reads and 4,496 writes; second-sample budget decision GO. No read blowout was observed.
- Pages invocation metrics for 7 October AEDT: 194 exceeded-resource invocations. Live Safari reload reproduced a login 503. The production log stream recorded `exceededCpu`, 10 ms CPU and no application logs for the failing state requests.
- A static spreadsheet-library import in FindMyShift initialized spreadsheet machinery during unrelated cold requests. It now initializes only during workbook construction; the compiled worker confirms initialization is inside the deferred import.
- Four post-deployment invalid-login checks returned the expected 401 responses, with successful worker outcomes and CPU measurements of 14, 8, 3 and 1 ms. Cloudflare permits occasional CPU allowance rollover; these samples do not guarantee every authenticated operation stays within its limit.

## Changes

- Unlinked clinical fast login returns bounded published name matches and dropdown data. Resolve-account-claims no longer replaces discovered suggestions with empty arrays. Deferred context refresh explicitly renders the name dropdown. Clear full-name matches now link automatically before the calendar is returned. Capitalization, titles and surname-first comma formatting are normalized. Competing spellings, duplicate names at one site, incomplete publications and occupied identities require confirmation. Existing links are preserved; indexed ownership assertions run in the same transaction as the new claims. Repeat logins make no additional link writes.
- Admin Users search filters existing cards locally. It preserves the search input and its caret. Each user's hidden roster-name options are generated only when Edit is opened.
- Permission controls remain in place during saves, with per-user pending protection. A failed save reverts only the affected user's fields. Directory responses begun before a permission change cannot overwrite it.
- Server permission saves update the requested profile fields only, preserving roster claims, subscription tokens and session settings.

## Validation and remaining verification

Passed: bounded identity integration (100,001 historical-event fixture; no event scans for discovery), admin input/permission race tests, creator login containment, client request budget, facility access, request attribution, syntax checks and compiled Pages build.

The broader fixture suite reached its existing roster-processor source-code assertion at line 922 and failed there. Its workbook/parser scenarios before that assertion passed. The suite's October 4 publication-window expectation was updated to the already implemented prior-month policy, and callers now await the deferred workbook builder.

Authenticated production signup/admin checks and a real iPhone check remain outstanding. The pre-deployment Safari reload signed the user out after its resource failure; no account permissions were changed during live diagnosis. The user has been asked to sign back in. All unrelated restoration work and untracked files were preserved.

## Afternoon signup incident

Arnav Mehta's account and two roster claims committed before a signup 503 at 16:56 AEDT. Cloudflare recorded an exceeded-resource failure in that minute; later live requests confirmed `exceededCpu` at 10 ms on both `/api/state` and `/api/roster-revision`. This is distinct from the earlier eager spreadsheet initialization fix.

New-account and first-link fast login now return the authenticated identity envelope before parsing site-month calendar publications. The existing automatic calendar request then delivers the shifts without asking the user to confirm clear name matches. A transient creation 503 gets at most one login-only recovery attempt; database safety/daily limits are not retried. The public revision route uses R2 HEAD/ETags instead of downloading and decompressing manifests, with zero D1 reads.

Tests cover actual signup, automatic multi-site calendar delivery, replay without link writes, committed-creation recovery and database-limit exclusions. Live signup after this deployment still needs confirmation; these fixes do not establish that every application route fits the 10 ms CPU allowance.
