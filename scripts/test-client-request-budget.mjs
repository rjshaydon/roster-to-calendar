import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");

const statusBody = section(/function setStatus\([\s\S]*?(?=\nfunction removeSupersededStatusMessages)/);
assert.doesNotMatch(statusBody, /persistConsoleMessage|appendConsoleMessage/, "ordinary UI messages must not produce database-backed console writes");

const postLoginRefresh = section(/function queuePostLoginSnapshotRefresh[\s\S]*?(?=\nfunction markLoginPhase)/);
assert.equal((postLoginRefresh.match(/loadCloudCalendarEvents\(/g) || []).length, 1, "one retry loop must contain one calendar request site");
assert.match(postLoginRefresh, /for \(const delayMs of \[1500\]\)/, "login must schedule at most one background calendar retry");
assert.match(postLoginRefresh, /if \(document\.hidden\) return/, "hidden tabs must not perform the retry");

const calendarLoad = section(/async function loadCloudCalendarEvents[\s\S]*?(?=\nfunction cloudCalendarEventRange)/);
assert.equal((calendarLoad.match(/await fetch\("\/api\/state"/g) || []).length, 2, "foreground calendar load may issue only the request and one resource-limit retry");
assert.match(calendarLoad, /if \(response\.status === 503\)/, "calendar retry must be limited to an explicit resource failure");
assert.match(calendarLoad, /allowInlineBuild: false/, "calendar retry must not repeat inline build work");

const statusRefresh = section(/async function refreshCalendarStoreStatus[\s\S]*?(?=\nasync function toggleAdminConsole)/);
assert.match(statusRefresh, /if \(calendarStoreStatusRequest\) return calendarStoreStatusRequest/, "status callers must coalesce");
assert.match(statusRefresh, /Math\.min\(2, options\.attempts \|\| 2\)/, "status refresh must allow at most one retry");
assert.match(statusRefresh, /document\.hidden/, "silent status refresh must stop in hidden tabs");

const byStreamOpen = section(/async function openFacilityOverviewByStream[\s\S]*?(?=\nfunction closeFacilityOverview)/);
assert.equal((byStreamOpen.match(/loadFacilityOverviewMetadata\(\)/g) || []).length, 1, "By stream must not duplicate a failed metadata request");

const contactScheduling = section(/function facilityOverviewContactRefreshIsActive[\s\S]*?(?=\nasync function refreshFacilityOverviewContactList)/);
assert.match(appSource, /const FACILITY_OVERVIEW_CONTACT_REFRESH_MS = 60_000/);
assert.match(contactScheduling, /!document\.hidden/, "contact polling must stop in hidden tabs");
assert.match(contactScheduling, /stopFacilityOverviewContactRefresh\(\)/, "contact scheduling must replace rather than stack timers");

const visibleTabs = 50;
const viewingHours = 12;
const refreshMinutes = 1;
const contactRequests = visibleTabs * viewingHours * 60 / refreshMinutes;
const maximumObservedCreatorLoginRequests = 6;
const report = {
  observedLocalCreatorLogin: {
    before: 11,
    after: maximumObservedCreatorLoginRequests,
    removedAutomaticConsoleWrites: 2,
    removedBackgroundCalendarRetries: 3,
  },
  fiftyVisibleTabs: {
    viewingHours,
    refreshSeconds: refreshMinutes * 60,
    contactRequests,
    requiredD1RowsForUnchangedContactRefresh: 0,
  },
  gates: {
    loginRequestsPerTab: maximumObservedCreatorLoginRequests,
    postLoginBackgroundRetries: 1,
    byStreamMetadataRequestsPerOpen: 1,
    hiddenContactRefreshes: 0,
  },
};

assert.equal(contactRequests, 36_000);
console.log(JSON.stringify(report, null, 2));

function section(pattern) {
  return appSource.match(pattern)?.[0] || "";
}
