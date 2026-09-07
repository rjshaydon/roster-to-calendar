# D1 client request audit — 7 September 2026

## Result

The incident's exact historical URL and SQL fingerprint remain unavailable,
but the local investigation established three concrete amplification paths and
closed two of them:

1. A missing `FACILITY_LEGACY_READS_PAUSED` variable selected the legacy path,
   and a missing `FACILITY_SHARED_EMERGENCY_PAUSED` variable did not pause
   builders. Automatic Pages deployments update code without necessarily
   applying `wrangler.toml` variables. Both helpers now default closed for
   missing or malformed values.
2. One Creator login issued 11 authenticated `/api/state` requests in 15
   seconds. Four were post-login calendar revalidations and two were automatic
   console-history writes triggered by ordinary UI messages. The revised local
   build issues six requests: one login, initial calendar load, roster-status
   check, user list, account-context load and one bounded calendar revalidation.
3. Opening By stream repeated a failed metadata request immediately. It now
   attempts that request once per opening.

This explains how a browser lifecycle could amplify D1 work, and the missing
variable explains why those requests were allowed to reach legacy query paths.
It does not prove which historical route accounted for the 8.23 million rows
omitted by Cloudflare's query-fingerprint data.

## Local runtime observations

The test used the isolated local Wrangler database and the deterministic local
Creator account. Cloudflare D1 and R2 were not used.

| Sequence | Before | After | Database-write-capable requests after |
| --- | ---: | ---: | ---: |
| Creator login plus 15 seconds | 11 | 6 | 0 automatic console writes |
| Post-login calendar revalidation | 4 | 1 | 0 |
| By stream metadata after failure | 2 | 1 | 0 |
| Hidden On shift contact refresh | 0 | 0 | 0 |

The fail-closed local At a glance checks returned 503 for On shift, ED staff
and By stream. Each request still performs ordinary account authentication,
but the route does not enter a legacy roster collection query.

## Trigger matrix

| Trigger | Endpoint/action | Current request ceiling | D1 implication |
| --- | --- | ---: | --- |
| Creator login | `login` | 1 | Bounded authentication/login reads |
| Initial calendar hydration | `loadCalendarEvents` | 1 plus one 503 retry | Registry/account reads; retry forbids inline building |
| Creator supporting context | `calendarStoreStatus`, `listUsers`, `loadAccountContext` | 1 each; status callers coalesce | Status is disabled before repository reads; remaining calls must stay bounded |
| Ordinary UI status message | formerly `appendConsoleMessage` | 0 | No longer writes D1 |
| Open On shift while paused | `queryFacilityOverviewOnShift` | 1 per explicit open | Authentication only, then blocked before legacy roster query |
| Open ED staff while paused | `queryFacilityOverviewStaff` | 1 per explicit open | Authentication only, then blocked before legacy roster query |
| Open By stream while paused | `queryFacilityOverviewMetadata` | 1 per explicit open | Authentication only, then blocked before legacy roster query |
| Visible On shift contact refresh | `queryFacilityOverviewContactList` | 1 per 60 seconds, one timer, no overlap | **Future blocker:** current route authenticates through D1 |
| Hidden tab | none for contact/status polling | 0 | No polling D1 |

## Fifty-visible-page capacity

Assumption: 50 simultaneously visible On shift pages, 12 viewing hours and a
60-second refresh interval.

```text
50 × 12 × 60 = 36,000 contact refresh requests/day
```

The acceptance target remains zero D1 rows for an unchanged authenticated
contact refresh. The present `/api/state` contact action cannot meet that target
because it performs password/account verification before loading a shared R2
object. Therefore broad shared-reader activation remains blocked even if the
bootstrap canary succeeds. A later phase must use the planned short-lived
authorised session path whose unchanged refresh reads R2/browser state without
D1, while preserving the 15-minute revocation bound.

## Regression gates

`npm run test:client-request-budget` verifies:

- ordinary UI status messages do not persist console rows;
- login has one background calendar retry and hidden tabs perform none;
- a calendar retry cannot repeat inline building;
- `calendarStoreStatus` callers coalesce and retry at most once;
- By stream makes one metadata request per opening; and
- visible/hidden contact-refresh capacity assumptions remain explicit.

The default-closed rollout tests separately prove missing and malformed pause
variables cannot restore legacy reads or builders.
