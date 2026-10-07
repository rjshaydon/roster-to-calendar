# Account reliability fixes — 8 October 2026

Production fix: 56f57c2d.

## Evidence

- D1's settled UTC-day sample at approximately 02:00 AEDT: 448,340 reads and 4,496 writes; second-sample budget decision GO. No read blowout was observed.
- Pages invocation metrics for 7 October AEDT: 194 exceeded-resource invocations. Live Safari reload reproduced a login 503. The production log stream recorded `exceededCpu`, 10 ms CPU and no application logs for the failing state requests.
- A static spreadsheet-library import in FindMyShift initialized spreadsheet machinery during unrelated cold requests. It now initializes only during workbook construction; the compiled worker confirms initialization is inside the deferred import.
- Four post-deployment invalid-login checks returned the expected 401 responses, with successful worker outcomes and CPU measurements of 14, 8, 3 and 1 ms. Cloudflare permits occasional CPU allowance rollover; these samples do not guarantee every authenticated operation stays within its limit.

## Changes

- Unlinked clinical fast login returns bounded published name matches and dropdown data. Resolve-account-claims no longer replaces discovered suggestions with empty arrays. Deferred context refresh explicitly renders the name dropdown. Suggestions appear first and a single suggestion is preselected for user confirmation; identities are never silently claimed.
- Admin Users search filters existing cards locally. It preserves the search input and its caret. Each user's hidden roster-name options are generated only when Edit is opened.
- Permission controls remain in place during saves, with per-user pending protection. A failed save reverts only the affected user's fields. Directory responses begun before a permission change cannot overwrite it.
- Server permission saves update the requested profile fields only, preserving roster claims, subscription tokens and session settings.

## Validation and remaining verification

Passed: bounded identity integration (100,001 historical-event fixture; no event scans for discovery), admin input/permission race tests, creator login containment, client request budget, facility access, request attribution, syntax checks and compiled Pages build.

The broader fixture suite reached its existing roster-processor source-code assertion at line 922 and failed there. Its workbook/parser scenarios before that assertion passed. The suite's October 4 publication-window expectation was updated to the already implemented prior-month policy, and callers now await the deferred workbook builder.

Authenticated production signup/admin checks and a real iPhone check remain outstanding. The pre-deployment Safari reload signed the user out after its resource failure; no account permissions were changed during live diagnosis. The user has been asked to sign back in. All unrelated restoration work and untracked files were preserved.
