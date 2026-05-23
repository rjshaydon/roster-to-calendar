import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { onRequestPost as handleStatePost } from "../functions/api/state.js";
import { onRequestGet as handleFeedGet } from "../functions/api/feed.js";
import { buildPreviewFromDerivedEvents } from "../functions/_lib/d1-calendar.js";
import { buildRosterView, doctorOptions, parseUploadForm, parserRuleDefaults, previewSummary, setParserExtensions } from "../public/static/roster.js";

function cloneWorkbook(workbook) {
  return XLSX.read(XLSX.write(workbook, { type: "array", bookType: "xlsx" }), { type: "array", cellDates: true });
}

function workbookFile(workbook, name) {
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  return new File([bytes], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function workbookDataUrl(workbook) {
  const bytes = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${Buffer.from(bytes).toString("base64")}`;
}

function withWorkbookCell(workbook, sheetName, cell, value) {
  const copy = cloneWorkbook(workbook);
  copy.Sheets[sheetName][cell] = value;
  return copy;
}

function withWorkbookDate(workbook, sheetName, cell, date) {
  return withWorkbookCell(workbook, sheetName, cell, { t: "d", v: date, w: date.toLocaleDateString("en-AU") });
}

async function assertRejectsMixedTermUpload(label, workbook, filename, expectedParts) {
  const formData = new FormData();
  formData.append("rosterFiles", workbookFile(workbook, filename));
  await assert.rejects(
    () => parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: formData })),
    (error) => {
      assert.ok(expectedParts.every((part) => error.message.includes(part)), `${label}: ${error.message}`);
      return true;
    },
  );
}

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const caseyWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)), {
  cellDates: true,
});
const mchWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const caseyBytes = await readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)));
const mchBytes = await readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)));
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const styleSource = await readFile(new URL("../public/static/styles.css", import.meta.url), "utf8");
const calendarMigrationSource = await readFile(new URL("../migrations/0001_calendar_store.sql", import.meta.url), "utf8");
const insightIndexMigrationSource = await readFile(new URL("../migrations/0005_roster_insight_index.sql", import.meta.url), "utf8");
const d1CalendarSource = await readFile(new URL("../functions/_lib/d1-calendar.js", import.meta.url), "utf8");
assert.match(
  appSource,
  /function pasteCopiedEvent[\s\S]*openCustomEventModal\(previewEventToCustomEvent\(shifted\), targetDate, \{ draft: true \}\);/,
  "pasting should open a custom-event draft instead of persisting immediately",
);
assert.doesNotMatch(
  appSource.match(/function openCustomEventModal[\s\S]*?function closeCustomEventModal/)?.[0] || "",
  /renderInlineWhoInsight/,
  "custom-event modal should not request roster coworker insights",
);
assert.match(
  appSource.match(/function openReviewModal[\s\S]*?function closeReviewModal/)?.[0] || "",
  /canUseRosterInsights\(\) && !isLeaveEvent\(event\)[\s\S]*renderInlineWhoInsight/,
  "leave review modals should not request roster coworker insights",
);
assert.match(appSource, /let cloudStateSaveQueue = Promise\.resolve\(\);/, "cloud saves should be serialized");
assert.match(
  appSource.match(/async function enterUserAccount[\s\S]*?async function enterDoctorProfileView/)?.[0] || "",
  /queueBackgroundCloudStateSave\(capturePendingCloudStateSave\(\) \|\| creatorCalendarSavePayload\(\) \|\| snapshotCloudSavePayload\(\), \{ delayMs: 1500 \}\)/,
  "switching from the creator account should queue the creator doctor save without blocking entry",
);
assert.match(
  appSource.match(/async function enterUserAccount[\s\S]*?async function enterDoctorProfileView/)?.[0] || "",
  /accountSwitchStartedAt[\s\S]*renderCachedCalendarSnapshotForContext\(targetContext[\s\S]*validateClaimedAccountCalendarInBackground\(targetContext/,
  "switched-account entry should render a cached target snapshot before background validation",
);
assert.match(appSource, /function markAccountSwitchPhase/, "account switching should expose debug timings separately from login timings");
assert.match(appSource, /function renderCachedCalendarSnapshotForContext/, "calendar switching should have an explicit-context snapshot renderer");
assert.match(appSource, /function validateDoctorProfileCalendarInBackground/, "doctor-profile switching should validate cached snapshots in the background");
assert.match(
  appSource.match(/async function hydrateAuthenticatedWorkspace[\s\S]*?function markLoginPhase/)?.[0] || "",
  /currentUserEmail === OWNER_EMAIL[\s\S]*forceCreatorDoctorSession\(\)[\s\S]*loadCloudCalendarEvents/,
  "creator hydration should normalize the creator doctor before calendar events load",
);
assert.match(
  stateSource.match(/if \(action === "adminLoadUser"\)[\s\S]*?if \(action === "claimRosterName"\)/)?.[0] || "",
  /prepareLightweightAccountResponse[\s\S]*snapshot: null[\s\S]*issueConfig: null/,
  "admin account switching should use the lightweight account response and defer calendar snapshots",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "adminLoadUser"\)[\s\S]*?if \(action === "claimRosterName"\)/)?.[0] || "",
  /prepareAccountResponse/,
  "admin account switching should not run full account enrichment before calendar load",
);
assert.match(
  appSource.match(/async function returnToCreatorAccount[\s\S]*?async function clearLocalWorkspace/)?.[0] || "",
  /forceCreatorDoctorSession\(\);[\s\S]*renderCachedCalendarSnapshotForContext\(targetContext[\s\S]*validateClaimedAccountCalendarInBackground/,
  "returning to the creator should render the cached creator calendar before calendar validation",
);
assert.match(
  appSource.match(/function currentAccount[\s\S]*?function canUseRosterInsights/)?.[0] || "",
  /viewedAccountEmail\(\)[\s\S]*function viewedAccountEmail[\s\S]*adminViewingEmail \|\| currentUserEmail[\s\S]*function isOwnerAccount\(\) \{\s*return isViewingCreatorAccount\(\);/,
  "switched-user account surfaces should be driven by viewed identity rather than creator authentication",
);
assert.match(
  appSource.match(/async function deleteAccount[\s\S]*?function deleteLocalAccountData/)?.[0] || "",
  /creatorDeletingSwitchedUser[\s\S]*if \(creatorDeletingSwitchedUser\)[\s\S]*returnToCreatorCalendar[\s\S]*if \(deletingViewedAccount\)[\s\S]*setActiveCalendarContext\(\"claimed-account\", \{ email: \"\" \}\)/,
  "deleting a switched user should return to creator, while user self-delete should clear into logged-out context",
);
assert.doesNotMatch(
  appSource.match(/function deleteLocalAccountData[\s\S]*?function clearDeletedAccountClaims/)?.[0] || "",
  /accountState\.currentEmail = OWNER_EMAIL/,
  "deleting the current local user should not silently select the creator account",
);
assert.match(
  appSource.match(/function snapshotCloudSavePayload[\s\S]*?function forceCreatorDoctorSession/)?.[0] || "",
  /accountEmail: viewedAccountEmail\(\)[\s\S]*requestEmail: adminViewingEmail \? authenticatedAccountEmail\(\) : viewedAccountEmail\(\)[\s\S]*targetEmail: adminViewingEmail \? viewedAccountEmail\(\) : \"\"/,
  "switched-user settings should save to the viewed account using creator authentication only for the request",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "login"\)[\s\S]*?const account = await verifyD1Account/)?.[0] || "",
  /autoClaimMatchedRosterNames|prepareAccountResponse|loadRepositoryIndex|buildIssueConfig/,
  "login should stay lightweight and avoid broad account hydration",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function prepareLightweightAccountResponse[\s\S]*?export async function prepareAccountResponse/)?.[0] || "",
  /queryRosterFiles\(options\.db\)/,
  "lightweight login responses should not load full roster-file metadata",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function prepareLightweightAccountResponse[\s\S]*?export async function prepareAccountResponse/)?.[0] || "",
  /includeImportRefs === false[\s\S]*imports: \[\]/,
  "lightweight login responses should be able to skip import refs entirely",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /removedImportIds\.length[\s\S]*deleteDerivedRosterFile[\s\S]*deleteRawRosterFile/,
  "ordinary saves must only delete roster database rows when explicit removed import ids are supplied",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /staleFileIds|canReconcileToFullSet|queryRosterFileRanges/,
  "ordinary saves must not infer roster deletions from partial account import snapshots",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "adminCreateUser"\)[\s\S]*?if \(action === "resolveAccountClaims"\)/)?.[0] || "",
  /autoClaimMatchedRosterNames|prepareAccountResponse|loadRepositoryIndex|buildIssueConfig/,
  "admin user creation should stay lightweight and avoid broad account hydration",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "loadCalendarEvents"\)[\s\S]*?if \(action === "loadInsightImports"\)/)?.[0] || "",
  /loadRepositoryIndex/,
  "calendar event loads should be D1-first and avoid repository-index hydration",
);
assert.match(
  stateSource.match(/if \(action === "loadDoctorProfile"\)[\s\S]*?if \(action === "saveDoctorProfile"\)/)?.[0] || "",
  /cachedRevision[\s\S]*snapshotCurrent[\s\S]*calendarRevision/,
  "doctor profile loads should support cached-revision validation without rebuilding the snapshot",
);
assert.match(
  stateSource,
  /async function queryDoctorProfileCalendarRevision[\s\S]*queryCalendarRevision/,
  "doctor profile cache validation should use a lightweight calendar revision",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /loadRepositoryIndex/,
  "ordinary cloud saves should not load the repository index",
);
assert.match(
  stateSource.match(/if \(action === "claimRosterName"\)[\s\S]*?if \(action === "listUsers"\)/)?.[0] || "",
  /loadSqlDoctorCandidates[\s\S]*findDoctorClaimCandidate/,
  "manual roster claims should resolve from SQL doctor candidates",
);
assert.match(
  stateSource.match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /replaceAccountCustomEvents[\s\S]*stripRelationalCustomEventsFromSession/,
  "account saves should keep custom-event truth in D1 rows rather than session JSON",
);
assert.match(
  stateSource.match(/if \(action === "resolveAccountClaims"\)[\s\S]*?if \(action === "adminLoadUser"\)/)?.[0] || "",
  /resolvedClaims[\s\S]*includeAvailableDoctors:[\s\S]*&& !resolvedClaims\.length/,
  "claim resolution should only load full doctor candidates while an account still has no claims",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "listUsers"\)[\s\S]*?if \(action === "calendarStoreStatus"\)/)?.[0] || "",
  /cleanupResolvedAdminIssues/,
  "creator user-list loading should not synchronously clean every account",
);
assert.match(
  stateSource.match(/if \(action === "listUsers"\)[\s\S]*?if \(action === "calendarStoreStatus"\)/)?.[0] || "",
  /globalParserExtensions[\s\S]*listD1Users\(context\.env\.ROSTER_DB, \{ globalParserExtensions \}\)/,
  "creator user-list loading should reuse one parser-rule load for stale issue filtering",
);
assert.match(
  stateSource.match(/async function userSummaryFromRecord[\s\S]*?function insightsEnabledForRecord/)?.[0] || "",
  /filterResolvedAdminIssuesForSummary/,
  "user summaries should hide resolved stale parser issues without foreground cleanup",
);
assert.match(
  appSource.match(/async function hydrateAuthenticatedWorkspace[\s\S]*?function markLoginPhase/)?.[0] || "",
  /adminTargetEmail && adminTargetEmail !== OWNER_EMAIL && !currentRosterClaims\.length[\s\S]*resolveCurrentAccountClaims\(adminTargetEmail\)[\s\S]*!adminTargetEmail && currentUserEmail !== OWNER_EMAIL && !currentRosterClaims\.length[\s\S]*resolveCurrentAccountClaims\(\)[\s\S]*loadCloudCalendarEvents[\s\S]*void loadServerUsers\(\)/,
  "claim resolution should be skipped for already-claimed account entry while still resolving unclaimed users before calendar load",
);
assert.match(
  stateSource.match(/async function buildDerivedAccountSnapshot[\s\S]*?function rawRosterObjectKey/)?.[0] || "",
  /snapshotFileRefs[\s\S]*d1RepositoryImportRefsForClaims\(db, claims\)[\s\S]*fileRefs: snapshotFileRefs/,
  "claimed-account snapshots should derive source file refs from D1 claims when lightweight account state omits imports",
);
assert.match(
  appSource.match(/async function claimSelectedRosterName[\s\S]*?async function updatePreview/)?.[0] || "",
  /applyCloudStateData\(data\)[\s\S]*loadCloudCalendarEvents\(\{ adminTargetEmail: adminViewingEmail \? viewedAccountEmail\(\) : "" \}\)[\s\S]*if \(loadedCalendar\)[\s\S]*bootstrapImports\(\)/,
  "manual roster claims should load a D1 calendar snapshot before bootstrapping repository refs",
);
assert.match(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /cloudAvailable && selectedFiles\.length && selectedFiles\.some\(\(entry\) => !entry\.file\)[\s\S]*loadCloudCalendarEvents\(\{ adminTargetEmail: adminViewingEmail \? viewedAccountEmail\(\) : "" \}\)[\s\S]*No D1 calendar events were found/,
  "cloud repository refs should use D1 calendar loading instead of falling through to local browser parsing",
);
assert.match(
  appSource.match(/function renderFilesMarkup[\s\S]*?function renderFilesList/)?.[0] || "",
  /statusOnlyEntries[\s\S]*const displayFiles = selectedFiles\.length \? selectedFiles : statusOnlyEntries/,
  "Admin Files should render D1 roster status files when local selected files are empty",
);
assert.match(
  appSource.match(/async function replaceActiveRostersWithCurrentUploads[\s\S]*?async function reparseRosterFile/)?.[0] || "",
  /retainedRosterEntriesFromStatus[\s\S]*Rebuild requires at least one retained roster file/,
  "rebuild-all should use retained roster sources and refuse an empty retained roster set",
);
assert.match(
  appSource.match(/async function reparseRosterFile[\s\S]*?async function refreshCalendarStoreStatus/)?.[0] || "",
  /calendarStoreStatus\?\.files[\s\S]*ensureRosterEntrySource\(entry\)/,
  "per-file refresh should be able to hydrate status-only files from retained R2 source",
);
assert.match(
  d1CalendarSource.match(/async function runTransactionalBatch[\s\S]*?function chunkRows/)?.[0] || "",
  /db\.batch\(statements\)/,
  "D1 roster imports should keep file replacement statements transactional",
);
assert.match(
  d1CalendarSource.match(/function bulkInsertEventStatements[\s\S]*?function bulkInsertIssueStatements/)?.[0] || "",
  /chunkRows\(rows, 1\)/,
  "D1 roster event inserts should stay under SQL variable limits",
);
assert.match(
  d1CalendarSource.match(/function bulkInsertIssueStatements[\s\S]*?async function runTransactionalBatch/)?.[0] || "",
  /chunkRows\(rows, 1\)/,
  "D1 roster issue inserts should stay under SQL variable limits",
);
assert.match(
  stateSource.match(/async function calendarStoreStatus[\s\S]*?function summarizeExpectedRosterFiles/)?.[0] || "",
  /queryRawRosterFiles[\s\S]*retainedOnlyFiles[\s\S]*retainedSourceOnly/,
  "calendar status should include retained R2 source pointers without derived rows",
);
assert.match(
  appSource.match(/async function loginWithEmail[\s\S]*?async function restoreCloudState/)?.[0] || "",
  /restoreCloudState\(\{ \.\.\.options, deferHydration: true[\s\S]*renderLoginState\(\);\s*closeLoginModal\(\);[\s\S]*hydrateAuthenticatedWorkspace\(\{ \.\.\.options, includeBootstrap: true \}/,
  "successful login should reveal the shell before background workspace hydration completes",
);
assert.doesNotMatch(
  stateSource,
  /ADMIN_ISSUE_DISMISS_PREFIX|ADMIN_ISSUE_IGNORE_PREFIX|PARSER_EXTENSION_RULES_KEY|PARSER_RULE_SUGGESTIONS_KEY|loadParserExtensionRules\(|saveParserExtensionRules\(/,
  "D1-only state routes should not retain dead KV-era helper scaffolding",
);
assert.match(
  appSource.match(/function buildResolvedPreviewEvents[\s\S]*?function latestPreviewEventsByIdentity/)?.[0] || "",
  /activeCustomEventIds[\s\S]*!activeCustomEventIds\.has\(event\.id\)[\s\S]*previewCustomEventIds[\s\S]*customEventsMaterialized === true && previewCustomEventIds\.has\(event\.id\)/,
  "D1-loaded previews should merge newly added local custom events and drop materialized custom events removed from active state",
);
assert.match(
  appSource.match(/function openReviewModal[\s\S]*?function closeReviewModal/)?.[0] || "",
  /isCustomPreviewEvent\(previewEvent\)[\s\S]*openCustomEventModal/,
  "materialized custom events should open in the custom-event editor before generic review handling",
);
assert.match(
  appSource.match(/function deletePreviewEvent[\s\S]*?function resetImportedEvent/)?.[0] || "",
  /isCustomPreviewEvent\(event\)[\s\S]*ensureEditableCustomEvent\(event\)[\s\S]*removeCustomEventForActiveCalendar/,
  "materialized custom events should be rehydrated before deletion",
);
assert.match(
  appSource.match(/function renderWorkspaceFromSnapshot[\s\S]*?async function ensureSelectedFilesLoaded/)?.[0] || "",
  /applySessionState[\s\S]*reconcileMaterializedPreviewCustomEvents\(\)/,
  "rendering a D1 snapshot should reconcile materialized custom events into editable state",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /serverStatusComplete[\s\S]*serverSyncedCount[\s\S]*roster file\$\{serverSyncedCount === 1 \? "" : "s"\} synced/,
  "healthy server roster status should drive the System-card synced count",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /Sync issue detected: \$\{serverSyncedCount\}\/\$\{serverExpectedCount\}/,
  "System-card sync issues should compare server counts rather than empty local file handles",
);
assert.match(appSource, /data-admin-user-seniority-filter/, "admin users should expose a seniority filter");
assert.doesNotMatch(
  appSource.match(/async function deleteAccount[\s\S]*?function deleteLocalAccountData/)?.[0] || "",
  /if \(creatorCanDelete\) await loadServerUsers\(\);\s*closeAccountsModal\(\);/,
  "deleting another account should not close the Admin modal",
);
assert.match(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /async function calendarStoreStatus[\s\S]*resolveSelectedRosterFileDoctorRows\(db, selectedDoctorKey\)/,
  "calendar status should resolve only the selected doctor instead of rebuilding every canonical option",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function buildDerivedDoctorProfileSnapshot[\s\S]*?function matchDoctorClaims/)?.[0] || "",
  /queryRosterFileDoctorsForKeys\(db, doctorKeysForOption\(profile\)\)/,
  "doctor profile load should resolve only the requested doctor instead of rebuilding every canonical option",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function buildDerivedDoctorProfileSnapshot[\s\S]*?function matchDoctorClaims/)?.[0] || "",
  /if \(!doctorDiagnostics\.length\)[\s\S]*queryRosterFileDoctors\(db\)[\s\S]*resolveCanonicalDoctorOptionForKey/,
  "doctor profile load should reserve full canonical rebuilds for the rare targeted-lookup miss path",
);
assert.match(appSource, /data-replace-active-rosters/, "creator UI should expose a roster recovery action");
assert.match(appSource, /<strong>Roster database<\/strong>/, "system card should use plain roster-database language");
assert.match(appSource, /source file\$\{retainedSourceTotal === 1 \? \"\" : \"s\"\} retained/, "system card should report retained raw source coverage");
assert.match(appSource, />Check status<\/button>/, "system card should expose a non-mutating status check");
assert.match(appSource, />Rebuild from roster files<\/button>/, "system card should expose an explicit roster rebuild action");
assert.doesNotMatch(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /loadRepositoryIndex|loadSnapshotRecord|persistSnapshotRecord|storeSnapshotForAccount|storeSnapshotForDoctorProfile/,
  "D1 serving code should not retain repository-index or snapshot serving helpers",
);
assert.doesNotMatch(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /calendarStoreStatus = \{ unavailable: true \}/,
  "failed status checks should not erase the last valid D1 status",
);
assert.match(
  appSource.match(/async function saveSelectedRosterFilesToD1[\s\S]*?function emptyRosterPersistenceSummary/)?.[0] || "",
  /entriesToSave = options\.force === true[\s\S]*entries\.filter\(\(entry\) => !persistedIds\.has\(entry\.id\) \|\| failedIds\.has\(entry\.id\)\)/,
  "ordinary roster sync should process only missing or failed files",
);
assert.match(appSource, /function rosterSyncLabel/, "file cards should expose live roster sync labels");
assert.match(appSource, /Missing \/ unresolved shift codes/, "system card should expose unresolved shift-code review");
assert.match(appSource, /data-add-manual-shift-code/, "missing shift-code heading should expose a manual Add action");
assert.match(appSource, /parserRuleSeniorityOption/, "shift-code editor should expose multi-seniority selection");
assert.match(appSource, /function selectedParserRuleSeniorities/, "shift-code editor should save selected seniorities as a batch");
assert.match(appSource, /function openManualParserRuleModal/, "shift-code editor should support manual rule creation without an existing issue");
assert.match(appSource, /parserRuleExistsForIssue/, "system shift-code review should hide only issues resolved by active rules");
assert.match(appSource, /matchingParserRuleGroup/, "saved shift-code rules should reopen with equivalent seniorities selected");
assert.match(appSource, /<span class="who-team-role">/, "Who role labels should render separately from doctor buttons");
assert.doesNotMatch(appSource, /data-who-role-shift-code|handleWhoRoleRuleClick/, "Who role labels should not be clickable shift-code controls");
assert.match(appSource, /isWhoNightIcShift[\s\S]*nightIcRank/, "Night IC shifts should rank above other night staff in Who lists");
assert.match(appSource, /function whoDisplayTeamLabel[\s\S]*Night main team[\s\S]*Night Hub[\s\S]*Night SSU/, "Night teams should use explicit main, hub, and SSU labels");
assert.match(appSource, /night main team", "night hub", "night ssu"/, "Night Hub should sort between Night main team and Night SSU");
assert.match(appSource, /Night Hub/, "Hub night shift-code rules should preview as Night Hub");
assert.match(appSource, /refreshActiveWhoInsightSurfaces/, "saving shift-code rules should refresh active Who insight panels");
assert.match(appSource, /function synthesizeIncompleteShiftCodeIssues/, "derived code-only shift titles should synthesize unresolved shift-code issues");
assert.match(appSource, /parserRuleIgnore/, "shift-code editor should expose persistent ignore mode");
assert.match(appSource, /data-ignore-shift-code/, "missing shift-code queue should open the shared ignore rule flow");
assert.match(appSource, /Ignored shift codes/, "ignored shift codes should remain editable in hospital rule sections");
assert.match(appSource, /parserRuleSeniorityAll/, "shift-code seniority picker should expose an All option");
assert.match(appSource, /function normalizeParserRuleSenioritySelection/, "shift-code seniority picker should keep All and Unknown selections consistent");
assert.match(appSource, /const key = `\$\{source\}\|\$\{code\}`/, "unresolved shift-code grouping should be by hospital and code");
assert.match(appSource, /formatShiftCodeSeniorities/, "grouped unresolved shift-code rows should summarize detected seniorities");
assert.match(styleSource, /#parserRuleForm[\s\S]*overflow-y: auto/, "shift-code editor form should scroll vertically when it exceeds available height");
assert.match(appSource, /normalizeDdhParserRuleCodeText/, "DDH shift-code issues should use parser-equivalent label codes");
assert.match(appSource, /seniority !== "Unknown"[\s\S]*some\(\(rule\) => rule\.code === code\)/, "Unknown-seniority shift-code issues should resolve by source/code");
assert.match(appSource, /isKnownResolvedShiftCodeValue/, "derived warnings should not be synthesized for built-in recognised shift labels");
assert.match(appSource, /function shouldShowPreviewIssue/, "Warnings panel should centralize resolved shift-code issue filtering");
assert.match(
  appSource.match(/function shouldShowPreviewIssue[\s\S]*?function pruneResolvedLatestPreviewIssues/)?.[0] || "",
  /isKnownResolvedShiftCodeValue[\s\S]*isShiftCodeResolvedByActiveRules/,
  "Warnings panel should hide review-derived shift-code warnings resolved by built-ins or active rules",
);
assert.match(
  appSource.match(/function buildClientPreviewData[\s\S]*?function synthesizeIncompleteShiftCodeIssues/)?.[0] || "",
  /issues\.filter\(shouldShowPreviewIssue\)/,
  "active preview data should use the shared Warnings issue filter",
);
assert.match(
  appSource.match(/function openParserRuleModalFromPreviewIssue[\s\S]*?async function saveParserRuleFromModal/)?.[0] || "",
  /shouldShowPreviewIssue[\s\S]*That parser warning has already been resolved/,
  "shift-code warning editor should refresh resolved stale cards instead of opening missing issues",
);
assert.match(
  appSource.match(/function rebuildClientPreview[\s\S]*?function buildClientPreviewData/)?.[0] || "",
  /pruneResolvedLatestPreviewIssues\(\)/,
  "preview rendering should prune resolved warnings once review context is current",
);
assert.match(
  calendarMigrationSource,
  /CREATE TABLE IF NOT EXISTS roster_issues/,
  "calendar migration should persist parser diagnostics beside indexed roster events",
);
assert.match(
  d1CalendarSource,
  /INSERT INTO roster_issues/,
  "derived roster saves should persist import-time parser diagnostics",
);
assert.match(
  d1CalendarSource,
  /export async function queryDoctorIssues/,
  "calendar loads should read stored parser diagnostics from D1",
);
assert.match(
  appSource.match(/function synthesizeIncompleteShiftCodeIssues[\s\S]*?function incompleteShiftCodeIssueForReviewItem/)?.[0] || "",
  /if \(baseData\?\.derivedFromD1\) return \[\];/,
  "D1 calendar loads should not synthesize parser warnings from already-indexed events",
);
assert.match(
  appSource.match(/async function reportPreviewIssues[\s\S]*?async function reportAccountError/)?.[0] || "",
  /if \(latestPreview\?\.derivedFromD1\) return;/,
  "D1 calendar loads should not report parser warnings during normal rendering",
);
assert.doesNotMatch(
  appSource.match(/function applyIssueConfig[\s\S]*?function sanitizeParserExtensions/)?.[0] || "",
  /pruneResolvedLatestPreviewIssues\(\)/,
  "parser config application should stay lightweight during login and account switching",
);
assert.match(appSource, /if \(parsedRosterSources\)[\s\S]*await updatePreview\(\)[\s\S]*else if \(latestPreview\)/, "saving parser rules should refresh the visible preview before trying to reparse cloud file refs");
assert.match(
  appSource.match(/function renderParserRulesCard[\s\S]*?function collectUnknownShiftIssues/)?.[0] || "",
  /parserRuleSuggestions\.length \? `[\s\S]*<strong>User suggestions<\/strong>/,
  "empty user-suggestion sections should be omitted",
);
assert.match(appSource, /function exportHospitalOptions/, "one-off exports should expose hospital options");
assert.match(appSource, /Recognised hospitals &amp; default locations/, "account modal should expose recognised hospital locations");
assert.match(appSource, /data-account-location-key/, "account modal locations should bind to shared settings keys");
assert.match(appSource, /ACCOUNT_HOSPITAL_LOCATION_ORDER = \["mmc", "ddh", "mch", "casey"\]/, "account modal should keep hospital locations in the expected vertical order");
assert.match(d1CalendarSource, /CREATE TABLE IF NOT EXISTS account_hospital_locations/, "D1 should store account hospital locations relationally");
assert.match(d1CalendarSource, /function applyAccountHospitalLocations/, "SQL-first roster reads should apply account hospital defaults");
assert.match(
  appSource.match(/function buildLocationOptionMarkup[\s\S]*?function detectLocationPreset/)?.[0] || "",
  /locationOptionSourceTypes\(source\)/,
  "location options should consider the roster event source, not only detected imports",
);
assert.match(
  appSource.match(/async function buildDerivedCalendarFilePayload[\s\S]*?function assertDerivedCalendarFilePayload/)?.[0] || "",
  /includeLocations: true/,
  "D1 roster materialisation should retain canonical onsite locations for rebuilds",
);
assert.match(appSource, /skipStatus: true/, "rebuild saves should avoid full aggregate status checks after every file");
assert.match(appSource, /mergeLightweightRosterStatus/, "rebuild saves should merge lightweight per-file status");
assert.match(appSource, /function matchesExportHospitals/, "one-off exports should support hospital filtering");
assert.match(appSource, /function canCopySubscriptionUrl/, "subscription URL availability should be separate from one-off exports");
assert.match(
  appSource.match(/async function handleExportAction[\s\S]*?function downloadIcs/)?.[0] || "",
  /await navigator\.clipboard\.writeText\(url\)[\s\S]*saveCloudState\(snapshot\)\.catch/,
  "subscription URLs should be copied before async feed persistence runs",
);
assert.match(
  appSource.match(/async function handleExportAction[\s\S]*?function downloadIcs/)?.[0] || "",
  /subscriptionUrl\("webcal", exportConfig\.mode === "range" \? "range" : "full"\)/,
  "Apple Calendar action should open the subscription URL via webcal",
);
assert.match(
  appSource.match(/function renderExportModal[\s\S]*?async function handleExportAction/)?.[0] || "",
  /data-export-action="apple">Open in Apple Calendar/,
  "Apple Calendar one-off import should not be disabled by subscription availability",
);
assert.doesNotMatch(
  appSource.match(/async function reportAccountError[\s\S]*?async function updateAccountDetails/)?.[0] || "",
  /currentUserRole === "creator" && !adminViewingEmail/,
  "creator-owned unresolved shift codes should be reportable into the admin issue flow",
);
assert.doesNotMatch(
  appSource.match(/async function buildDerivedCalendarFilePayload[\s\S]*?function assertDerivedCalendarFilePayload/)?.[0] || "",
  /rawFile:/,
  "derived-file saves should not carry retained raw file bytes",
);
assert.match(
  d1CalendarSource,
  /queryRosterFileDoctorsForKeys/,
  "D1 helpers should expose a narrow selected-doctor file lookup",
);
assert.match(appSource, /function rosterSyncSummary/, "system card should expose aggregate roster sync progress");
assert.doesNotMatch(appSource, /function scheduleFailedRosterRetry/, "failed roster syncs should not schedule automatic retry storms");
assert.match(appSource, /data-reparse-import/, "file cards should expose a visible reparse action");
assert.match(appSource, /Reparse produced 0 events/, "zero-event reparses should remain visibly failed");
assert.match(appSource, /formatted\.replace\(\/,\\s0\(\\d:\\d\{2\}\\s\?pm\)\$\/i, \", \$1\"\)/, "PM timestamps should drop their leading zero");
assert.match(
  appSource.match(/accountsBody\.addEventListener\(\"click\"[\s\S]*?\n\}\);/)?.[0] || "",
  /data-reparse-import[\s\S]*reparseRosterFile/,
  "admin file reparse buttons should invoke the single-file reparse path",
);
assert.doesNotMatch(
  appSource.match(/async function renderWhoInsight[\s\S]*?async function renderWhenInsight/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "who insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/async function renderInlineWhoInsight[\s\S]*?function renderInlineWhoGroups/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "inline who insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "when insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/function isRosterShiftEvent[\s\S]*?function chooseNextOverlapDate/)?.[0] || "",
  /clinical support|\\bcso\?\\b/i,
  "clinical support shifts should count as rostered work for insight lookups",
);
assert.doesNotMatch(
  appSource.match(/function renderWhenInsightResult[\s\S]*?function comparisonDoctorOptions/)?.[0] || "",
  /doctor\.key === selectedKey/,
  "when insight rendering should not reference an out-of-scope selectedKey",
);
assert.match(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /fetchRosterOverlapDoctors[\s\S]*const options = prioritizeDoctorOptions[\s\S]*doctorKeys: \[selectedKey\][\s\S]*renderWhenInsightResult\(\{ options, selectedComparison/,
  "general when insights should keep all compact overlap doctors in the dropdown while loading one selected doctor's events",
);
assert.match(
  calendarMigrationSource,
  /CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events \(source_type, start_date, end_date\);/,
  "calendar migration should include the source/date range index used by insight lookups",
);
assert.match(
  insightIndexMigrationSource,
  /CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events \(source_type, start_date, end_date\);/,
  "existing databases should receive the insight query index through a forward migration",
);
assert.match(
  d1CalendarSource,
  /SELECT DISTINCT[\s\S]*od\.event_json,[\s\S]*od\.doctor_key,[\s\S]*od\.display_name,[\s\S]*od\.source_type[\s\S]*FROM roster_daily_overlaps/,
  "coworker lookup should select from roster_daily_overlaps with DISTINCT deduplication",
);
const caseyFormData = new FormData();
caseyFormData.append("rosterFiles", new File([caseyBytes], "Casey_Term_2_2026_DRAFT.xlsm", { type: "application/vnd.ms-excel.sheet.macroEnabled.12" }));
const parsedCaseyUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: caseyFormData }));
assert.equal(parsedCaseyUpload.sources.casey.length, 1);
const mchFormData = new FormData();
mchFormData.append("rosterFiles", new File([mchBytes], "Paeds_Term_2_2026.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
const parsedMchUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mchFormData }));
assert.equal(parsedMchUpload.sources.mch.length, 1);

const mmcUpload = new FormData();
mmcUpload.append("rosterFiles", workbookFile(mmcWorkbook, "AdultTerm1.2026.xlsx"));
const parsedMmcUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mmcUpload }));
assert.equal(parsedMmcUpload.sources.mmc.length, 1);

await assertRejectsMixedTermUpload(
  "MMC date typo should identify the worksheet and cell",
  withWorkbookDate(mmcWorkbook, "Week 1", "H4", new Date("2025-02-17T00:00:00")),
  "AdultTerm1.2026 typo.xlsx",
  ["AdultTerm1.2026 typo.xlsx has dates from multiple terms", "Week 1 cell H4", "2025-02-17", "Term 1 2025"],
);

await assertRejectsMixedTermUpload(
  "Casey date typo should identify the worksheet and cell",
  withWorkbookCell(caseyWorkbook, "May 4", "B2", { t: "s", v: "01-Feb", w: "01-Feb" }),
  "Casey_Term_2_2026 typo.xlsm",
  ["Casey_Term_2_2026 typo.xlsm has dates from multiple terms", "May 4 cell B2", "2026-02-01", "Term 4 2025"],
);

await assertRejectsMixedTermUpload(
  "MCH date typo should identify the worksheet and cell",
  withWorkbookDate(mchWorkbook, "Week 1", "F19", new Date("2025-05-04T00:00:00")),
  "Paeds_Term_2_2026 typo.xlsx",
  ["Paeds_Term_2_2026 typo.xlsx has dates from multiple terms", "Week 1 cell F19", "2025-05-04", "Term 1 2025"],
);

const ddhSheetName = ddhWorkbook.SheetNames[0];
await assertRejectsMixedTermUpload(
  "DDH date typo should identify the worksheet and cell",
  withWorkbookCell(ddhWorkbook, ddhSheetName, "B1", { t: "s", v: "Mon. Feb. 3, 2025", w: "Mon. Feb. 3, 2025" }),
  "Dandenong typo.xlsx",
  ["Dandenong typo.xlsx has dates from multiple terms", `${ddhSheetName} cell B1`, "2025-02-03", "Term 1 2025"],
);

const mixedUpload = new FormData();
mixedUpload.append("rosterFiles", workbookFile(mmcWorkbook, "AdultTerm1.2026.xlsx"));
mixedUpload.append("rosterFiles", workbookFile(withWorkbookDate(mmcWorkbook, "Week 1", "H4", new Date("2025-02-17T00:00:00")), "OnlyThisFileIsBad.xlsx"));
await assert.rejects(
  () => parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mixedUpload })),
  (error) => error.message.includes("OnlyThisFileIsBad.xlsx has dates from multiple terms") && !error.message.includes("AdultTerm1.2026.xlsx has dates from multiple terms"),
);

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook, caseyWorkbook);
const defaultRules = parserRuleDefaults();
const mmcRules = defaultRules.mmc || [];
const mchRules = defaultRules.mch || [];
const ddhRules = defaultRules.ddh || [];
const hasMmcRule = (seniority, code) => mmcRules.some((rule) => rule.seniority === seniority && rule.code === code);
const hasMchRule = (seniority, code, base = "") => mchRules.some((rule) => rule.seniority === seniority && rule.code === code && (!base || rule.base === base));
assert.ok(hasMmcRule("SMS", "AGC"));
assert.ok(hasMmcRule("CMO", "AGC"));
assert.ok(hasMmcRule("SMS", "CS"));
assert.ok(hasMmcRule("CMO", "CSO"));
assert.equal(hasMmcRule("Senior Registrar", "AGC"), false);
assert.equal(hasMmcRule("HMO", "AGC"), false);
assert.equal(hasMmcRule("Senior Registrar", "CS"), false);
assert.equal(hasMmcRule("HMO", "CSO"), false);
assert.equal(hasMmcRule("SMS", "ACR"), false);
assert.equal(hasMmcRule("SMS", "ARR"), false);
assert.equal(hasMmcRule("SMS", "ASSR"), false);
assert.ok(hasMmcRule("Senior Registrar", "SWA"));
assert.ok(hasMmcRule("Transitional/Intermediate Registrar", "SWP"));
assert.ok(hasMmcRule("Junior Registrar", "AHJ"));
assert.ok(hasMmcRule("HMO", "PHJ"));
assert.ok(hasMmcRule("Intern", "NSSJ"));
for (const seniority of ["Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "Intern"]) {
  assert.ok(hasMmcRule(seniority, "ASSJ"), `default MMC rules should include ASSJ for ${seniority}`);
  assert.ok(hasMmcRule(seniority, "PSSJ"), `default MMC rules should include PSSJ for ${seniority}`);
}
assert.ok(hasMchRule("SMS", "CS", "CS"));
assert.ok(hasMchRule("CMO", "OCS", "CS Office"));
assert.ok(hasMchRule("HMO", "PHNW", "PHNW"));
assert.ok(ddhRules.some((rule) => rule.code === "ROVER AM" && rule.base === "Rover" && rule.period === "AM"));
assert.ok(ddhRules.some((rule) => rule.code === "ROVER PM" && rule.base === "Rover" && rule.period === "PM"));
const nssjRule = mmcRules.find((rule) => rule.seniority === "HMO" && rule.code === "NSSJ");
assert.equal(nssjRule.startTime, "23:00");

const unmappedTimedMmcWorkbook = XLSX.utils.book_new();
const unmappedTimedMmcSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "0800-1730 ASSJ"],
]);
for (let index = 0; index < 7; index += 1) {
  unmappedTimedMmcSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(unmappedTimedMmcWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(unmappedTimedMmcWorkbook, unmappedTimedMmcSheet, "Week 1");
const unmappedTimedMmcView = buildRosterView([{ id: "unmapped-assj", workbook: unmappedTimedMmcWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(
  unmappedTimedMmcView.events.some((event) => event.rawValue === "0800-1730 ASSJ" && event.title === "MMC: SSU AM" && event.start.includes("08:00:00") && event.end.includes("17:30:00")),
  "default ASSJ rules should resolve Senior Registrar SSU AM while preserving explicit roster time",
);
assert.equal(unmappedTimedMmcView.issues.some((issue) => issue.rawValue === "0800-1730 ASSJ"), false, "default ASSJ rules should not enter the unresolved shift-code workflow");
assert.equal(nssjRule.endTime, "08:30");
for (const sourceRules of [defaultRules.ddh, defaultRules.casey]) {
  assert.equal(sourceRules.some((rule) => rule.seniority === "Senior Registrar" && rule.base === "CS"), false);
  assert.equal(sourceRules.some((rule) => rule.seniority === "HMO" && rule.base === "CS"), false);
  assert.ok(sourceRules.some((rule) => rule.seniority === "SMS" && rule.base === "CS"));
  assert.ok(sourceRules.some((rule) => rule.seniority === "CMO" && rule.base === "CS"));
}
assert.ok(doctors.length > 100);
const richard = doctors.find((doctor) => doctor.displayName === "Richard HAYDON");
assert.ok(richard);
assert.deepEqual(richard.sourceTypes, ["mmc", "ddh"]);
assert.ok(doctors.find((doctor) => doctor.displayName === "Brianna Dawn MURPHY"));
assert.ok(doctors.find((doctor) => doctor.displayName === "Patrick TAN"));
assert.equal(doctors.find((doctor) => doctor.displayName === "Aarushi Pathania"), undefined);
assert.equal(doctors.find((doctor) => doctor.displayName === "HMO MUST BE"), undefined);

const michaelMerged = doctorOptions(mmcWorkbook, [], [], mchWorkbook).filter((doctor) => doctor.displayName.toUpperCase().includes("MICHAEL COMAN"));
assert.equal(michaelMerged.length, 1);
assert.deepEqual(michaelMerged[0].sourceTypes, ["mmc", "mch"]);
assert.ok(michaelMerged[0].aliases.some((alias) => alias.sourceType === "mmc" && alias.key === "MICHAEL COMAN"));
assert.ok(michaelMerged[0].aliases.some((alias) => alias.sourceType === "mch" && alias.key === "DR MICHAEL COMAN"));
const michaelMergedView = buildRosterView(mmcWorkbook, [], michaelMerged[0].key, undefined, {}, {}, michaelMerged[0].aliases, [], mchWorkbook);
assert.ok(michaelMergedView.events.some((event) => event.source === "MMC"));
assert.ok(michaelMergedView.events.some((event) => event.source === "MCH"));
assert.ok(michaelMergedView.events.some((event) => event.title === "Conference Leave"));

const markDouglas = doctors.find((doctor) => doctor.displayName === "Mark DOUGLAS");
assert.ok(markDouglas);
const markView = buildRosterView(mmcWorkbook, [], markDouglas.key);
assert.ok(markView.events.some((event) => event.title === "MMC: AM shift"));
assert.ok(markView.events.some((event) => event.title === "MMC: PM shift"));

const deslinAraullo = doctors.find((doctor) => doctor.displayName === "Deslin ARAULLO");
assert.ok(deslinAraullo);
const deslinView = buildRosterView(mmcWorkbook, [], deslinAraullo.key);
assert.ok(deslinView.events.some((event) => event.title === "MMC: Hub PM"));
assert.ok(deslinView.events.some((event) => event.title === "MMC: Swing AM"));

const caseyDoctors = doctorOptions([], [], caseyWorkbook);
assert.ok(caseyDoctors.length > 130);
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Andrew DYALL"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Dennis CHUNG"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Rizwana SADAF"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Victor Ki Chung LI"));
assert.equal(caseyDoctors.find((doctor) => doctor.displayName === "Rostered staff"), undefined);

const patrickTan = doctorOptions(mmcWorkbook, [], caseyWorkbook).find((doctor) => doctor.displayName === "Patrick TAN");
assert.ok(patrickTan);
assert.deepEqual(patrickTan.sourceTypes, ["mmc", "casey"]);
const patrickTanView = buildRosterView(mmcWorkbook, [], patrickTan.key, undefined, {}, {}, [], caseyWorkbook);
const patrickCaseyEvents = patrickTanView.events.filter((event) => event.source === "Casey");
assert.ok(patrickCaseyEvents.length > 40);
assert.equal(patrickCaseyEvents.some((event) => event.start.startsWith("2025")), false);
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: MIC PM"));
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: AM shift" && event.rawValue === "Orient 0800-1730" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));
const patrickMergedLeave = patrickCaseyEvents.find((event) => event.title === "Annual Leave" && event.start === "2026-07-27");
assert.ok(patrickMergedLeave);
assert.equal(patrickMergedLeave.end, "2026-08-03");
assert.equal(patrickMergedLeave.rawValue, "Annual Leave");

const suzanFoxCasey = caseyDoctors.find((doctor) => doctor.displayName === "Suzan FOX");
assert.ok(suzanFoxCasey);
const suzanFoxCaseyView = buildRosterView([], [], suzanFoxCasey.key, undefined, {}, {}, suzanFoxCasey.aliases, caseyWorkbook);
const suzanMergedLeave = suzanFoxCaseyView.events.find((event) => event.title === "Annual Leave" && event.start === "2026-07-27");
assert.ok(suzanMergedLeave);
assert.equal(suzanMergedLeave.end, "2026-08-03");
assert.equal(suzanMergedLeave.rawValue, "Annual Leave");

const derivedLeavePreview = buildPreviewFromDerivedEvents([
  {
    id: "leave-a",
    title: "MMC: Annual Leave",
    rawValue: "Annual Leave",
    source: "MMC",
    sources: ["MMC"],
    start: "2026-07-27",
    end: "2026-08-03",
    allDay: true,
  },
  {
    id: "leave-b",
    title: "Casey: Conference Leave",
    rawValue: "Conference Leave",
    source: "Casey",
    sources: ["Casey"],
    start: "2026-07-27",
    end: "2026-08-03",
    allDay: true,
  },
  {
    id: "shift-a",
    title: "MMC: AM",
    rawValue: "AM",
    source: "MMC",
    start: "2026-07-28T08:00:00",
    end: "2026-07-28T17:00:00",
    allDay: false,
  },
]);
const derivedLeaveEvents = derivedLeavePreview.events.filter((event) => /leave/i.test(event.title));
assert.equal(derivedLeaveEvents.length, 1, "overlapping leave from multiple sources should render once");
assert.deepEqual(derivedLeaveEvents[0].sources, ["MMC", "Casey"]);
assert.equal(derivedLeavePreview.events.some((event) => event.id === "shift-a"), true, "non-leave shifts must remain visible");
const adjacentMixedLeavePreview = buildPreviewFromDerivedEvents([
  { id: "conference-week", title: "Conference Leave", rawValue: "Conference Leave", source: "MMC", start: "2026-02-02", end: "2026-02-09", allDay: true },
  { id: "annual-week", title: "Annual Leave", rawValue: "Annual Leave", source: "MMC", start: "2026-02-09", end: "2026-02-16", allDay: true },
]);
assert.deepEqual(
  adjacentMixedLeavePreview.events.map((event) => [event.title, event.start, event.end]),
  [["Conference Leave", "2026-02-02", "2026-02-09"], ["Annual Leave", "2026-02-09", "2026-02-16"]],
  "adjacent different leave types should not be merged",
);

const andrewDyallCasey = caseyDoctors.find((doctor) => doctor.displayName === "Andrew DYALL");
const andrewCaseyView = buildRosterView([], [], andrewDyallCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: TL AM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: UFD PM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: MIC AM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: PAEDS PM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: CS" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));

const bashirCasey = caseyDoctors.find((doctor) => doctor.displayName === "Bashir GONDAL");
const bashirCaseyView = buildRosterView([], [], bashirCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(bashirCaseyView.events.some((event) => event.title === "Casey: SSU AM"));

const dennisCasey = caseyDoctors.find((doctor) => doctor.displayName === "Dennis CHUNG");
const dennisCaseyView = buildRosterView([], [], dennisCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(dennisCaseyView.events.some((event) => event.title === "Casey: Night shift" && event.start.includes("23:00:00") && event.end.startsWith("2026-05-06")));
assert.equal(dennisCaseyView.events.filter((event) => event.title === "Casey: Night shift" && event.start.startsWith("2026-05-05") && event.end.startsWith("2026-05-06")).length, 1);

const jasonAwCasey = caseyDoctors.find((doctor) => doctor.displayName === "Jason AW");
const jasonAwCaseyView = buildRosterView([], [], jasonAwCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(jasonAwCaseyView.events.some((event) => event.title === "Annual Leave"));

const mustafaCasey = caseyDoctors.find((doctor) => doctor.displayName === "Mustafa Al ASAAD");
const mustafaCaseyView = buildRosterView([], [], mustafaCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(mustafaCaseyView.events.some((event) => event.title === "Conference Leave"));

const mchDoctors = doctorOptions([], [], [], mchWorkbook);
assert.ok(mchDoctors.length >= 60);
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Adam WEST"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Mark LIM"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Firas HAMDAN"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Peter AHN"));
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "ONCALL 0000-0800"), undefined);
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "requested off"), undefined);

const adamWestMch = mchDoctors.find((doctor) => doctor.displayName === "Adam WEST");
const adamWestMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730 CS"));
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: PM shift" && event.rawValue === "1430-0000" && event.end.startsWith("2026-05-09")));

const bobSeithMch = mchDoctors.find((doctor) => doctor.displayName === "Bob SEITH");
const bobSeithMchView = buildRosterView([], [], bobSeithMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: DEMT" && event.rawValue === "0800-1730 DEMT"));
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730CS"));

const andrewHardyMch = mchDoctors.find((doctor) => doctor.displayName === "Andrew HARDY");
const andrewHardyMchView = buildRosterView([], [], andrewHardyMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(andrewHardyMchView.events.some((event) => event.title === "MCH: CS Office" && event.rawValue === "0800-1730 OCS"));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Exam Leave" && event.rawValue === "ME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.start === "2026-06-08" && event.end === "2026-06-15"));

const adamWestMchWeek6 = adamWestMchView.events.filter((event) => event.rawValue === "PHNW 0800-1730");
assert.ok(adamWestMchWeek6.some((event) => event.title === "MCH: PHNW"));
const noSpacePhnwMchWorkbook = withWorkbookCell(mchWorkbook, "Week 6", "F21", { t: "s", v: "0800-1730PHNW", w: "0800-1730PHNW" });
const noSpacePhnwMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], noSpacePhnwMchWorkbook);
assert.ok(noSpacePhnwMchView.events.some((event) => event.title === "MCH: PHNW" && event.rawValue === "0800-1730PHNW"));

const markLimMch = mchDoctors.find((doctor) => doctor.displayName === "Mark LIM");
const markLimMchView = buildRosterView([], [], markLimMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(markLimMchView.events.some((event) => event.title === "MCH: Night shift" && event.rawValue === "2300-0830" && event.end.startsWith("2026-05-09")));
assert.equal(markLimMchView.events.filter((event) => event.title === "MCH: Night shift" && event.rawValue === "2300-0830" && event.start.startsWith("2026-05-08") && event.end.startsWith("2026-05-09")).length, 1);
assert.ok(markLimMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "C/L" && event.allDay));

const firasMch = mchDoctors.find((doctor) => doctor.displayName === "Firas HAMDAN");
const firasMchView = buildRosterView([], [], firasMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(firasMchView.events.some((event) => event.title === "Annual Leave" && event.rawValue === "AL 0.5" && event.allDay));

const marianPanlilioMch = mchDoctors.find((doctor) => doctor.displayName === "Marian PANLILIO");
const marianPanlilioMchView = buildRosterView([], [], marianPanlilioMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(marianPanlilioMchView.events.some((event) => event.title === "Sick Leave" && event.rawValue.trim() === "S/L PM" && event.allDay));

const houshmandMch = mchDoctors.find((doctor) => doctor.displayName === "Houshmand REFAEI");
const houshmandMchView = buildRosterView([], [], houshmandMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.equal(houshmandMchView.events.some((event) => String(event.rawValue || "").includes("EDO")), false);

const overlappingConferenceWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(overlappingConferenceWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "20-Jul", "21-Jul", "22-Jul", "23-Jul", "24-Jul", "25-Jul", "26-Jul"],
  ["Dr Michael Coman", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L"],
  ["Daily LEAVE", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave"],
]), "Week 1");
const michaelComan = doctorOptions([], [], overlappingConferenceWorkbook, mchWorkbook).find((doctor) => doctor.displayName === "Michael COMAN");
assert.ok(michaelComan);
const michaelComanView = buildRosterView([], [], michaelComan.key, undefined, {}, {}, michaelComan.aliases, overlappingConferenceWorkbook, mchWorkbook);
const michaelConferenceEvents = michaelComanView.events.filter((event) => event.title === "Conference Leave" && event.start === "2026-07-20");
assert.equal(michaelConferenceEvents.length, 1);
assert.equal(michaelConferenceEvents[0].end, "2026-07-27");
assert.equal(michaelConferenceEvents[0].rawValue, "C/L / CME/L");
assert.ok(michaelComanView.reviewItems.some((item) => item.id === michaelConferenceEvents[0].id));

const dailyLeave = doctorOptions([], [], overlappingConferenceWorkbook).find((doctor) => doctor.displayName === "Daily LEAVE");
assert.ok(dailyLeave);
const dailyLeaveView = buildRosterView([], [], dailyLeave.key, undefined, {}, {}, [], overlappingConferenceWorkbook);
const dailyAnnualLeave = dailyLeaveView.events.filter((event) => event.title === "Annual Leave");
assert.equal(dailyAnnualLeave.length, 1);
assert.equal(dailyAnnualLeave[0].start, "2026-07-20");
assert.equal(dailyAnnualLeave[0].end, "2026-07-27");
assert.ok(dailyLeaveView.reviewItems.some((item) => item.id === dailyAnnualLeave[0].id));

const michaelAnnualWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(michaelAnnualWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "27-Jul", "28-Jul", "29-Jul", "30-Jul", "31-Jul", "1-Aug", "2-Aug"],
  ["Michael COMAN", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave"],
]), "Week 1");
const michaelAnnualOption = doctorOptions([], [], michaelAnnualWorkbook, mchWorkbook).find((doctor) => doctor.displayName === "Michael COMAN");
assert.ok(michaelAnnualOption);
assert.deepEqual(michaelAnnualOption.sourceTypes, ["casey", "mch"]);
const michaelAnnualView = buildRosterView([], [], michaelAnnualOption.key, undefined, {}, {}, michaelAnnualOption.aliases, michaelAnnualWorkbook, mchWorkbook);
const michaelAnnualEvents = michaelAnnualView.events.filter((event) => /leave/i.test(event.title));
assert.equal(michaelAnnualEvents.length, 2);
assert.ok(michaelAnnualEvents.some((event) => event.title === "Conference Leave" && event.start === "2026-07-20" && event.end === "2026-07-27"));
assert.ok(michaelAnnualEvents.some((event) => event.title === "Annual Leave" && event.start === "2026-07-27" && event.end === "2026-08-03"));
assert.ok(michaelAnnualView.events.some((event) => event.source === "MCH"));

const mergedAnnualWorkbook = XLSX.utils.book_new();
const mergedAnnualSheet = XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "27-Jul", "28-Jul", "29-Jul", "30-Jul", "31-Jul", "1-Aug", "2-Aug"],
  ["Merged LEAVE", "Annual Leave", "", "", "", "", "", ""],
]);
mergedAnnualSheet["!merges"] = [{ s: { r: 2, c: 1 }, e: { r: 2, c: 7 } }];
XLSX.utils.book_append_sheet(mergedAnnualWorkbook, mergedAnnualSheet, "Week 1");
const mergedAnnualDoctor = doctorOptions([], [], mergedAnnualWorkbook).find((doctor) => doctor.displayName === "Merged LEAVE");
assert.ok(mergedAnnualDoctor);
const mergedAnnualView = buildRosterView([], [], mergedAnnualDoctor.key, undefined, {}, {}, [], mergedAnnualWorkbook);
const mergedAnnualEvents = mergedAnnualView.events.filter((event) => event.title === "Annual Leave");
assert.equal(mergedAnnualEvents.length, 1);
assert.equal(mergedAnnualEvents[0].start, "2026-07-27");
assert.equal(mergedAnnualEvents[0].end, "2026-08-03");
assert.equal(mergedAnnualEvents[0].rawValue, "Annual Leave");

const annualSynonymWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(annualSynonymWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "3-Aug", "4-Aug", "5-Aug", "6-Aug", "7-Aug", "8-Aug", "9-Aug"],
  ["Paeds Annual", "Paeds AL", "", "", "", "", "", ""],
  ["Casey Annual", "Casey AL", "", "", "", "", "", ""],
  ["Paeds Sick", "Paeds S/L", "", "", "", "", "", ""],
]), "Week 1");
const paedsAnnual = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Paeds ANNUAL");
const caseyAnnual = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Casey ANNUAL");
const paedsSick = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Paeds SICK");
assert.ok(buildRosterView([], [], paedsAnnual.key, undefined, {}, {}, paedsAnnual.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Annual Leave" && event.rawValue === "Paeds AL"));
assert.ok(buildRosterView([], [], caseyAnnual.key, undefined, {}, {}, caseyAnnual.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Annual Leave" && event.rawValue === "Casey AL"));
assert.ok(buildRosterView([], [], paedsSick.key, undefined, {}, {}, paedsSick.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Sick Leave" && event.rawValue === "Paeds S/L"));

const view = buildRosterView(mmcWorkbook, ddhWorkbook, richard.key);
const summary = previewSummary(view.events);

const aftabMmc = doctors.find((doctor) => doctor.displayName === "Aftab SAMDANI");
assert.ok(aftabMmc);
const aftabMmcView = buildRosterView(mmcWorkbook, [], aftabMmc.key);
assert.ok(aftabMmcView.events.some((event) => event.title === "Conference Leave" && event.rawValue.toUpperCase() === "CME LEAVE" && event.allDay));

assert.equal(view.events.length, 38);
assert.equal(summary.date_range, "2026-02-02 to 2026-05-02");
assert.ok(view.events.some((event) => event.title === "Conference Leave" && event.rawValue.includes("Dandenong CL")));
assert.ok(view.reviewItems.length >= view.events.length);
assert.ok(view.events.some((event) => event.rawValue.includes("Annual leave")));
assert.ok(view.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(view.events.some((event) => event.title === "Sick Leave"));

const ddhFullWorkbook = XLSX.utils.book_new();
const ddhFullSheet = XLSX.utils.aoa_to_sheet([
  ["", "Mon. Feb. 02, 2026", "Tue. Feb. 03, 2026", "Wed. Feb. 04, 2026", "Thu. Feb. 05, 2026", "Fri. Feb. 06, 2026", "Sat. Feb. 07, 2026", "Sun. Feb. 08, 2026"],
  ["Richard Haydon", "", "", "", "", "", "", ""],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["Jim BARTON", "AVAO AM", "", "Orange PM (on-call)", "AVAO PM", "Clinical Support", "", ""],
  ["", "07:30-17:00", "", "15:00-00:00", "14:30-00:00", "", "", ""],
  ["Caroline BOLT", "Orange PM (on-call)", "", "AVAO AM", "", "Orange AM IC", "", ""],
  ["", "15:00-00:00", "", "07:30-17:00", "", "08:00-18:00", "", ""],
  ["Di FLOOD", "CS AM", "SSU SMS", "Clinical Support", "", "HITH PM", "", ""],
  ["", "", "07:30-17:30", "", "", "", "", ""],
]);
XLSX.utils.book_append_sheet(ddhFullWorkbook, ddhFullSheet, "Sheet1");
const ddhFullDoctors = doctorOptions([], ddhFullWorkbook);
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Jim BARTON"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Caroline BOLT"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Di FLOOD"));
assert.equal(ddhFullDoctors.find((doctor) => doctor.displayName === "SENIOR MEDICAL STAFF"), undefined);

const jim = ddhFullDoctors.find((doctor) => doctor.displayName === "Jim BARTON");
const jimView = buildRosterView([], ddhFullWorkbook, jim.key);
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO AM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: CS"));

const diFlood = ddhFullDoctors.find((doctor) => doctor.displayName === "Di FLOOD");
const diFloodView = buildRosterView([], ddhFullWorkbook, diFlood.key);
assert.ok(diFloodView.events.some((event) => event.title === "DDH: CS AM"));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: SSU" && event.start.includes("07:30:00")));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: HITH PM"));

const mmcPdfBytes = await readFile(fileURLToPath(new URL("../fixtures/AdultMMCTerm2.2026.Ver1.pdf", import.meta.url)));
const formData = new FormData();
formData.append("rosterFiles", new File([mmcPdfBytes], "AdultMMCTerm2.2026.Ver1.pdf", { type: "application/pdf" }));
const parsedPdf = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: formData }));
const pdfDoctors = doctorOptions(parsedPdf.sources.mmc, parsedPdf.sources.ddh);
assert.ok(pdfDoctors.length > 50);
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Richard HAYDON"));
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Abi THANIKASALAM"));
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Titus HACKMAN"));
const pdfRichard = pdfDoctors.find((doctor) => doctor.displayName === "Richard HAYDON");
const pdfView = buildRosterView(parsedPdf.sources.mmc, parsedPdf.sources.ddh, pdfRichard.key);
assert.ok(pdfView.events.some((event) => event.title === "MMC: SSU PM"));
assert.ok(pdfView.events.some((event) => event.rawValue === "0800-1730" && event.title === "MMC: AM shift"));

class MemoryStore {
  constructor() {
    this.records = new Map();
    this.deletedKeys = [];
    this.accountListCalls = 0;
    this.accountGetCalls = 0;
    this.d1 = new MemoryD1();
    this.r2 = new MemoryR2();
  }

  async get(key, type) {
    if (String(key || "").startsWith("account:")) this.accountGetCalls += 1;
    const value = this.records.get(key);
    if (value === undefined) return null;
    return type === "json" ? JSON.parse(value) : value;
  }

  async put(key, value) {
    this.records.set(key, String(value));
  }

  async delete(key) {
    this.deletedKeys.push(key);
    this.records.delete(key);
  }

  async list(options = {}) {
    const prefix = options.prefix || "";
    if (prefix === "account:") this.accountListCalls += 1;
    return {
      keys: [...this.records.keys()].filter((name) => name.startsWith(prefix)).map((name) => ({ name })),
    };
  }

  resetMetrics() {
    this.accountListCalls = 0;
    this.accountGetCalls = 0;
  }
}

class MemoryR2 {
  constructor() {
    this.objects = new Map();
  }

  async put(key, value, options = {}) {
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.objects.set(key, {
      bytes,
      httpMetadata: options.httpMetadata || {},
      customMetadata: options.customMetadata || {},
    });
  }

  async get(key) {
    const item = this.objects.get(key);
    if (!item) return null;
    return {
      httpMetadata: item.httpMetadata,
      customMetadata: item.customMetadata,
      arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength),
    };
  }

  async head(key) {
    const item = this.objects.get(key);
    if (!item) return null;
    return { customMetadata: item.customMetadata };
  }

  async delete(keys) {
    const list = Array.isArray(keys) ? keys : [keys];
    for (const k of list) this.objects.delete(k);
  }

  async list({ prefix = "" } = {}) {
    const matching = [...this.objects.keys()].filter((k) => k.startsWith(prefix)).sort();
    return { objects: matching.map((key) => ({ key })), truncated: false };
  }
}

class MemoryD1 {
  constructor() {
    this.files = new Map();
    this.doctors = new Map();
    this.fileDoctors = new Map();
    this.events = new Map();
    this.issues = new Map();
    this.rawFiles = new Map();
    this.accountProfiles = new Map();
    this.accountClaims = new Map();
    this.accountStates = new Map();
    this.accountHospitalLocations = new Map();
    this.canonicalDoctors = new Map();
    this.customEvents = new Map();
    this.subscriptionTokens = new Map();
    this.parserRules = new Map();
    this.parserRuleSuggestions = new Map();
    this.doctorProfiles = new Map();
    this.dailyOverlaps = new Map();
    this.consoleMessages = [];
    this.nextConsoleMessageId = 1;
    this.failNextEventInsert = false;
  }

  prepare(sql) {
    return new MemoryD1Statement(this, sql);
  }

  async batch(statements) {
    const snapshots = Object.fromEntries([
      "files",
      "doctors",
      "fileDoctors",
      "events",
      "issues",
      "rawFiles",
      "accountProfiles",
      "accountClaims",
      "accountStates",
      "accountHospitalLocations",
      "canonicalDoctors",
      "customEvents",
      "subscriptionTokens",
      "parserRules",
      "parserRuleSuggestions",
      "doctorProfiles",
      "dailyOverlaps",
    ].map((key) => [key, new Map(this[key])]));
    try {
      const results = [];
      for (const statement of statements) results.push(await statement.run());
      return results;
    } catch (error) {
      for (const [key, value] of Object.entries(snapshots)) this[key] = value;
      throw error;
    }
  }
}

class MemoryD1Statement {
  constructor(db, sql) {
    this.db = db;
    this.sql = String(sql || "").replace(/\s+/g, " ").trim();
    this.args = [];
  }

  bind(...args) {
    this.args = args;
    return this;
  }

  async run() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("CREATE ")) return { success: true };
    if (sql.startsWith("ALTER TABLE")) return { success: true };
    if (sql.startsWith("INSERT INTO roster_files")) {
      this.db.files.set(args[0], {
        id: args[0],
        name: args[1],
        source_type: args[2],
        active: args[3],
        size: args[4],
        last_modified: args[5],
        added_at: args[6],
        uploaded_at: args[7],
        uploaded_by: args[8],
        parsed_at: args[9],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO raw_roster_files")) {
      this.db.rawFiles.set(args[0], {
        file_id: args[0],
        name: args[1],
        source_type: args[2],
        size: args[3],
        last_modified: args[4],
        object_key: args[5],
        type: args[6],
        data_url: args[7],
        uploaded_at: args[8],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_file_doctors")) {
      for (const key of [...this.db.fileDoctors.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.fileDoctors.delete(key);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_events")) {
      for (const [key, value] of [...this.db.events.entries()]) if (value.file_id === args[0]) this.db.events.delete(key);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_issues")) {
      for (const [key, value] of [...this.db.issues.entries()]) if (value.file_id === args[0]) this.db.issues.delete(key);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_files")) {
      this.db.files.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM raw_roster_files")) {
      this.db.rawFiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_doctors")) {
      for (let index = 0; index < args.length; index += 4) {
        this.db.doctors.set(`${args[index]}|${args[index + 1]}`, {
          source_type: args[index],
          doctor_key: args[index + 1],
          display_name: args[index + 2],
          updated_at: args[index + 3],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_file_doctors")) {
      for (let index = 0; index < args.length; index += 4) {
        this.db.fileDoctors.set(`${args[index]}|${args[index + 1]}|${args[index + 2]}`, {
          file_id: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_events")) {
      if (this.db.failNextEventInsert) {
        this.db.failNextEventInsert = false;
        throw new Error("Injected event insert failure.");
      }
      for (let index = 0; index < args.length; index += 16) {
        this.db.events.set(args[index], {
          id: args[index],
          file_id: args[index + 1],
          source_type: args[index + 2],
          doctor_key: args[index + 3],
          display_name: args[index + 4],
          start_date: args[index + 5],
          end_date: args[index + 6],
          start_ts: args[index + 7],
          end_ts: args[index + 8],
          title: args[index + 9],
          raw_value: args[index + 10],
          seniority: args[index + 11],
          location: args[index + 12],
          all_day: args[index + 13],
          time_label: args[index + 14],
          event_json: args[index + 15],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_issues")) {
      for (let index = 0; index < args.length; index += 14) {
        this.db.issues.set(args[index], {
          id: args[index],
          file_id: args[index + 1],
          source_type: args[index + 2],
          doctor_key: args[index + 3],
          display_name: args[index + 4],
          start_date: args[index + 5],
          raw_value: args[index + 6],
          seniority: args[index + 7],
          status: args[index + 8],
          message: args[index + 9],
          resolution_type: args[index + 10],
          suggested_title: args[index + 11],
          time_label: args[index + 12],
          issue_json: args[index + 13],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_profiles")) {
      const previous = this.db.accountProfiles.get(args[0]) || {};
      this.db.accountProfiles.set(args[0], {
        ...previous,
        email: args[0],
        real_name: args[1],
        role: args[2],
        insights_enabled: args[3],
        subscription_token: args[4],
        password_salt: args[5] || previous.password_salt || "",
        password_hash: args[6] || previous.password_hash || "",
        admin_issues_json: args[7] || "[]",
        local_parser_extensions_json: args[8] || "[]",
        created_at: args[9] || previous.created_at || "",
        updated_at: args[10] || args[5],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_claims")) {
      for (const key of [...this.db.accountClaims.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.accountClaims.delete(key);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM subscription_tokens WHERE email")) {
      for (const [token, row] of [...this.db.subscriptionTokens.entries()]) {
        if (row.email === args[0]) this.db.subscriptionTokens.delete(token);
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO subscription_tokens")) {
      this.db.subscriptionTokens.set(args[0], {
        token: args[0],
        email: args[1],
        created_at: args[2],
        updated_at: args[3],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO parser_rules")) {
      this.db.parserRules.set(args[0], {
        id: args[0],
        scope: "global",
        source_type: args[1],
        seniority: args[2],
        code: args[3],
        title: args[4],
        rule_json: args[5],
        updated_at: args[6],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM parser_rules")) {
      for (const [key, rule] of [...this.db.parserRules.entries()]) {
        if (rule.scope === "global" && rule.source_type === args[0] && rule.seniority === args[1] && rule.code === args[2]) {
          this.db.parserRules.delete(key);
        }
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM parser_rule_suggestions")) {
      this.db.parserRuleSuggestions.clear();
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO parser_rule_suggestions")) {
      this.db.parserRuleSuggestions.set(args[0], {
        id: args[0],
        email: args[1],
        status: "pending",
        suggestion_json: args[2],
        created_at: args[3],
        updated_at: args[4],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_claims")) {
      for (let index = 0; index < args.length; index += 6) {
        this.db.accountClaims.set(`${args[index]}|${args[index + 1]}|${args[index + 2]}`, {
          email: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
          matched_at: args[index + 4],
          updated_at: args[index + 5],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_states")) {
      this.db.accountStates.set(args[0], {
        email: args[0],
        session_json: args[1],
        updated_at: args[2],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_states")) {
      this.db.accountStates.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_hospital_locations")) {
      for (const key of [...this.db.accountHospitalLocations.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.accountHospitalLocations.delete(key);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_hospital_locations")) {
      for (let index = 0; index < args.length; index += 4) {
        this.db.accountHospitalLocations.set(`${args[index]}|${args[index + 1]}`, {
          email: args[index],
          source_type: args[index + 1],
          location: args[index + 2],
          updated_at: args[index + 3],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM canonical_doctors")) {
      this.db.canonicalDoctors.clear();
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO canonical_doctors")) {
      this.db.canonicalDoctors.set(args[0], {
        canonical_key: args[0],
        display_name: args[1],
        source_type: args[2],
        source_types_json: args[3],
        aliases_json: args[4],
        has_events: args[5],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM custom_events")) {
      for (const key of [...this.db.customEvents.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.customEvents.delete(key);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO custom_events")) {
      this.db.customEvents.set(`${args[0]}|${args[1]}`, {
        owner_email: args[0],
        id: args[1],
        title: args[2],
        start_date: args[3],
        end_date: args[4],
        all_day: args[5],
        start_time: args[6],
        end_time: args[7],
        location: args[8],
        include: args[9],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_profiles")) {
      this.db.accountProfiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO doctor_profiles")) {
      this.db.doctorProfiles.set(args[0], {
        profile_id: args[0],
        doctor_key: args[1],
        display_name: args[2],
        source_types_json: args[3],
        state_json: args[4],
        created_at: args[5],
        updated_at: args[6],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO console_messages")) {
      this.db.consoleMessages.push({
        id: this.db.nextConsoleMessageId++,
        actor_email: args[0],
        message: args[1],
        is_error: args[2],
        created_at: args[3],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM console_messages")) {
      this.db.consoleMessages = [...this.db.consoleMessages]
        .sort((left, right) => String(right.created_at).localeCompare(String(left.created_at)) || right.id - left.id)
        .slice(0, 50);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM doctor_profiles")) {
      this.db.doctorProfiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("UPDATE roster_files SET active")) {
      const file = this.db.files.get(args[1]);
      if (file) file.active = args[0];
      return { success: true };
    }
    if (sql.includes("CREATE TABLE IF NOT EXISTS roster_daily_overlaps")) {
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_daily_overlaps")) {
      for (let index = 0; index < args.length; index += 8) {
        const id = args[index + 4];
        this.db.dailyOverlaps.set(id, {
          date: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
          event_id: id,
          event_json: args[index + 5],
          start_ts: args[index + 6],
          end_ts: args[index + 7],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_daily_overlaps")) {
      const pattern = args[0];
      const prefix = pattern.replace(/%/g, "");
      for (const key of [...this.db.dailyOverlaps.keys()]) {
        if (key.startsWith(prefix)) this.db.dailyOverlaps.delete(key);
      }
      return { success: true };
    }

    throw new Error(`Unsupported MemoryD1 run SQL: ${sql}`);
  }

  async all() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("PRAGMA table_info(account_profiles)")) {
      return {
        results: [
          "email",
          "real_name",
          "role",
          "insights_enabled",
          "subscription_token",
          "password_salt",
          "password_hash",
          "admin_issues_json",
          "local_parser_extensions_json",
          "created_at",
          "updated_at",
        ].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(raw_roster_files)")) {
      return {
        results: ["file_id", "name", "source_type", "size", "last_modified", "object_key", "type", "data_url", "uploaded_at"].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(roster_issues)")) {
      return {
        results: [
          "id",
          "file_id",
          "source_type",
          "doctor_key",
          "display_name",
          "start_date",
          "raw_value",
          "seniority",
          "status",
          "message",
          "resolution_type",
          "suggested_title",
          "time_label",
          "issue_json",
        ].map((name) => ({ name })),
      };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("file_name") && sql.includes("event_count")) {
      const requestedKeys = sql.includes("roster_file_doctors.doctor_key IN") ? new Set(args) : null;
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .filter((doctor) => !requestedKeys || requestedKeys.has(doctor.doctor_key))
          .map((doctor) => {
            const file = this.db.files.get(doctor.file_id);
            return {
              file_id: doctor.file_id,
              file_name: file?.name || "",
              file_source_type: file?.source_type || doctor.source_type,
              active: file?.active ?? 0,
              source_type: doctor.source_type,
              doctor_key: doctor.doctor_key,
              display_name: doctor.display_name,
              event_count: [...this.db.events.values()].filter((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key).length,
            };
          })
          .sort((left, right) => String(left.file_id).localeCompare(String(right.file_id)) || String(left.display_name).localeCompare(String(right.display_name))),
      };
    }
    if (sql.includes("FROM canonical_doctors")) {
      return {
        results: [...this.db.canonicalDoctors.values()]
          .filter((doctor) => !sql.includes("WHERE has_events = 1") || doctor.has_events === 1)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM custom_events")) {
      return {
        results: [...this.db.customEvents.values()]
          .filter((event) => event.owner_email === args[0])
          .sort((left, right) => left.start_date.localeCompare(right.start_date) || left.start_time.localeCompare(right.start_time) || left.title.localeCompare(right.title) || left.id.localeCompare(right.id)),
      };
    }
    if (sql.includes("FROM account_hospital_locations")) {
      return {
        results: [...this.db.accountHospitalLocations.values()]
          .filter((row) => row.email === args[0])
          .sort((left, right) => left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("GROUP BY file_id, doctor_key")) {
      const pairs = new Set();
      for (let index = 0; index < args.length; index += 2) pairs.add(`${args[index]}:${args[index + 1]}`);
      const grouped = new Map();
      for (const event of this.db.events.values()) {
        const key = `${event.file_id}:${event.doctor_key}`;
        if (!pairs.has(key)) continue;
        grouped.set(key, (grouped.get(key) || 0) + 1);
      }
      return {
        results: [...grouped.entries()].map(([key, count]) => {
          const [file_id, doctor_key] = key.split(":");
          return { file_id, doctor_key, count };
        }),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("GROUP BY file_id")) {
      const ids = new Set(args);
      const grouped = new Map();
      for (const event of this.db.events.values()) {
        if (!ids.has(event.file_id)) continue;
        grouped.set(event.file_id, (grouped.get(event.file_id) || 0) + 1);
      }
      return { results: [...grouped.entries()].map(([file_id, count]) => ({ file_id, count })) };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("GROUP BY file_id")) {
      const ids = new Set(args);
      const grouped = new Map();
      for (const doctor of this.db.fileDoctors.values()) {
        if (!ids.has(doctor.file_id)) continue;
        grouped.set(doctor.file_id, (grouped.get(doctor.file_id) || 0) + 1);
      }
      return { results: [...grouped.entries()].map(([file_id, count]) => ({ file_id, count })) };
    }
    if (sql.includes("FROM roster_events") && sql.includes("(roster_events.file_id = ? AND roster_events.doctor_key = ?)")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const pairs = new Set();
      for (let index = 0; index < args.length - 2; index += 2) {
        pairs.add(`${args[index]}:${args[index + 1]}`);
      }
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => pairs.has(`${event.file_id}:${event.doctor_key}`))
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .sort((left, right) => left.start_ts.localeCompare(right.start_ts))
          .map((event) => ({ event_json: event.event_json })),
      };
    }
    if (sql.includes("FROM roster_issues") && sql.includes("(roster_issues.file_id = ? AND roster_issues.doctor_key = ?)")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const pairs = new Set();
      for (let index = 0; index < args.length - 2; index += 2) {
        pairs.add(`${args[index]}:${args[index + 1]}`);
      }
      return {
        results: [...this.db.issues.values()]
          .filter((issue) => this.db.files.get(issue.file_id)?.active === 1)
          .filter((issue) => pairs.has(`${issue.file_id}:${issue.doctor_key}`))
          .filter((issue) => issue.start_date <= end && issue.start_date >= start)
          .sort((left, right) => left.start_date.localeCompare(right.start_date))
          .map((issue) => ({ issue_json: issue.issue_json })),
      };
    }
    if (sql.includes("FROM roster_daily_overlaps AS mine")) {
      const overlapKeyCount = (sql.match(/mine\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1;
      const overlapKeys = new Set(args.slice(0, overlapKeyCount));
      const start = args[overlapKeyCount];
      const end = args[overlapKeyCount + 1];
      const sourceOffset = overlapKeyCount + 4;
      const hasSourceFilter = sql.includes("od.source_type IN");
      const hasDoctorFilter = sql.includes("od.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("od.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/od\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/od\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/od\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(sourceOffset, sourceOffset + sourceCount));
      const doctorKeys = new Set(args.slice(sourceOffset + sourceCount, sourceOffset + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(sourceOffset + sourceCount + doctorCount, sourceOffset + sourceCount + doctorCount + excludedDoctorCount));
      const myEvents = [...this.db.events.values()]
        .filter((event) => this.db.files.get(event.file_id)?.active === 1)
        .filter((event) => overlapKeys.has(event.doctor_key))
        .filter((event) => event.start_date <= end && event.end_date >= start);
      const results = [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => myEvents.some((mine) => mine.source_type === event.source_type && event.start_date <= mine.end_date && event.end_date >= mine.start_date))
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .filter((event, index, events) => events.findIndex((item) => item.id === event.id) === index);
      if (!sql.includes("event_json")) {
        return {
          results: results
            .map((event) => ({ doctor_key: event.doctor_key, display_name: event.display_name, source_type: event.source_type }))
            .filter((event, index, eventsList) => eventsList.findIndex((item) => item.doctor_key === event.doctor_key && item.source_type === event.source_type) === index),
        };
      }
      return { results };
    }
    if (sql.includes("FROM roster_daily_overlaps") && !sql.includes("AS mine")) {
      const start = args[0];
      const end = args[1];
      const hasSourceFilter = sql.includes("od.source_type IN");
      const hasDoctorFilter = sql.includes("od.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("od.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/od\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/od\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/od\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(2, 2 + sourceCount));
      const doctorKeys = new Set(args.slice(2 + sourceCount, 2 + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(2 + sourceCount + doctorCount, 2 + sourceCount + doctorCount + excludedDoctorCount));
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .filter((event, index, events) => events.findIndex((item) => item.id === event.id) === index),
      };
    }
    if (sql.includes("FROM roster_events AS mine")) {
      const overlapKeyCount = (sql.match(/mine\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1;
      const overlapKeys = new Set(args.slice(0, overlapKeyCount));
      const end = args[overlapKeyCount];
      const start = args[overlapKeyCount + 1];
      const sourceOffset = overlapKeyCount + 4;
      const hasSourceFilter = sql.includes("other_events.source_type IN");
      const hasDoctorFilter = sql.includes("other_events.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("other_events.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/other_events\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/other_events\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/other_events\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(sourceOffset, sourceOffset + sourceCount));
      const doctorKeys = new Set(args.slice(sourceOffset + sourceCount, sourceOffset + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(sourceOffset + sourceCount + doctorCount, sourceOffset + sourceCount + doctorCount + excludedDoctorCount));
      const myEvents = [...this.db.events.values()]
        .filter((event) => this.db.files.get(event.file_id)?.active === 1)
        .filter((event) => overlapKeys.has(event.doctor_key))
        .filter((event) => event.start_date <= end && event.end_date >= start);
      const results = [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => myEvents.some((mine) => mine.source_type === event.source_type && event.start_date <= mine.end_date && event.end_date >= mine.start_date))
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .filter((event, index, events) => events.findIndex((item) => item.id === event.id) === index)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.start_ts.localeCompare(right.start_ts));
      if (!sql.includes("event_json")) {
        return {
          results: results
            .map((event) => ({ doctor_key: event.doctor_key, display_name: event.display_name, source_type: event.source_type }))
            .filter((event, index, events) => events.findIndex((item) => item.doctor_key === event.doctor_key && item.source_type === event.source_type) === index),
        };
      }
      return { results };
    }
    if (sql.includes("SELECT DISTINCT roster_events.seniority AS seniority")) {
      const keys = new Set(args);
      return {
        results: [...new Set([...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => keys.has(event.doctor_key))
          .map((event) => event.seniority)
          .filter(Boolean))]
          .sort()
          .map((seniority) => ({ seniority })),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("doctor_key IN")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const keys = new Set(args.slice(0, -2));
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => keys.has(event.doctor_key))
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .sort((left, right) => left.start_ts.localeCompare(right.start_ts))
          .map((event) => ({ event_json: event.event_json })),
      };
    }
    if (sql.includes("FROM roster_issues") && sql.includes("doctor_key IN")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const keys = new Set(args.slice(0, -2));
      return {
        results: [...this.db.issues.values()]
          .filter((issue) => this.db.files.get(issue.file_id)?.active === 1)
          .filter((issue) => keys.has(issue.doctor_key))
          .filter((issue) => issue.start_date <= end && issue.start_date >= start)
          .sort((left, right) => left.start_date.localeCompare(right.start_date))
          .map((issue) => ({ issue_json: issue.issue_json })),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("display_name")) {
      const end = args[0];
      const start = args[1];
      const hasSourceFilter = sql.includes("roster_events.source_type IN");
      const hasDoctorFilter = sql.includes("roster_events.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("roster_events.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/roster_events\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/roster_events\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/roster_events\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(2, 2 + sourceCount));
      const doctorKeys = new Set(args.slice(2 + sourceCount, 2 + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(2 + sourceCount + doctorCount, 2 + sourceCount + doctorCount + excludedDoctorCount));
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.start_ts.localeCompare(right.start_ts)),
      };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("roster_file_doctors.file_id AS file_id")) {
      const hasKeyFilter = sql.includes("roster_file_doctors.doctor_key IN");
      const keys = new Set(args);
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .filter((doctor) => !hasKeyFilter || keys.has(doctor.doctor_key))
          .map((doctor) => {
            const file = this.db.files.get(doctor.file_id);
            return {
              file_id: doctor.file_id,
              file_name: file?.name || "",
              file_source_type: file?.source_type || doctor.source_type,
              active: file?.active ?? 0,
              source_type: doctor.source_type,
              doctor_key: doctor.doctor_key,
              display_name: doctor.display_name,
              event_count: [...this.db.events.values()].filter((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key).length,
            };
          })
          .sort((left, right) => String(left.file_name).localeCompare(String(right.file_name)) || left.display_name.localeCompare(right.display_name)),
      };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("DISTINCT")) {
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .map((doctor) => ({
            source_type: doctor.source_type,
            doctor_key: doctor.doctor_key,
            display_name: doctor.display_name,
          }))
          .filter((doctor, index, doctors) => doctors.findIndex((item) => item.source_type === doctor.source_type && item.doctor_key === doctor.doctor_key) === index)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM roster_files") && sql.includes("doctor_count") && sql.includes("event_count")) {
      const includeInactive = !sql.includes("WHERE roster_files.active = 1");
      return {
        results: [...this.db.files.values()]
          .filter((file) => includeInactive || file.active === 1)
          .sort((left, right) => String(left.added_at || "").localeCompare(String(right.added_at || "")) || left.name.localeCompare(right.name))
          .map((file) => ({
            id: file.id,
            name: file.name,
            source_type: file.source_type,
            active: file.active,
            size: file.size,
            last_modified: file.last_modified,
            added_at: file.added_at,
            uploaded_at: file.uploaded_at,
            uploaded_by: file.uploaded_by,
            doctor_count: [...this.db.fileDoctors.values()].filter((doctor) => doctor.file_id === file.id).length,
            event_count: [...this.db.events.values()].filter((event) => event.file_id === file.id).length,
          })),
      };
    }
    if (sql.includes("FROM raw_roster_files") && !sql.includes("WHERE file_id = ?")) {
      return {
        results: [...this.db.rawFiles.values()].sort((left, right) => String(left.uploaded_at || "").localeCompare(String(right.uploaded_at || "")) || String(left.file_id || "").localeCompare(String(right.file_id || ""))),
      };
    }
    if (sql.includes("MIN(roster_events.start_date)") && sql.includes("FROM roster_files")) {
      const includeInactive = !sql.includes("WHERE roster_files.active = 1");
      return {
        results: [...this.db.files.values()]
          .filter((file) => includeInactive || file.active === 1)
          .map((file) => {
            const events = [...this.db.events.values()].filter((event) => event.file_id === file.id);
            const starts = events.map((event) => event.start_date).filter(Boolean).sort();
            const ends = events.map((event) => event.end_date).filter(Boolean).sort();
            return {
              id: file.id,
              name: file.name,
              source_type: file.source_type,
              active: file.active,
              last_modified: file.last_modified,
              added_at: file.added_at,
              uploaded_at: file.uploaded_at,
              start_date: starts[0] || "",
              coverage_end_date: starts[starts.length - 1] || "",
              end_date: ends[ends.length - 1] || "",
              event_count: events.length,
            };
          })
          .sort((left, right) => String(left.source_type).localeCompare(String(right.source_type)) || String(left.start_date).localeCompare(String(right.start_date))),
      };
    }
    if (sql.includes("FROM roster_files") && sql.includes("INNER JOIN roster_file_doctors") && sql.includes("roster_file_doctors.doctor_key IN")) {
      const keys = new Set(args);
      const seen = new Set();
      const results = [];
      for (const doctor of [...this.db.fileDoctors.values()]) {
        const file = this.db.files.get(doctor.file_id);
        if (!file || file.active !== 1 || !keys.has(doctor.doctor_key) || seen.has(file.id)) continue;
        seen.add(file.id);
        results.push({
          id: file.id,
          name: file.name,
          source_type: file.source_type,
          active: file.active,
          size: file.size,
          last_modified: file.last_modified,
          added_at: file.added_at,
          uploaded_at: file.uploaded_at,
          uploaded_by: file.uploaded_by,
        });
      }
      return {
        results: results.sort((left, right) => String(left.added_at || "").localeCompare(String(right.added_at || "")) || left.name.localeCompare(right.name)),
      };
    }
    if (sql.startsWith("SELECT file_id, source_type, doctor_key, display_name FROM roster_file_doctors")) {
      return {
        results: [...this.db.fileDoctors.values()]
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM account_profiles") && sql.includes("LEFT JOIN account_claims")) {
      const results = [];
      const tokenFilter = sql.includes("WHERE account_profiles.subscription_token = ?");
      const emailFilter = sql.includes("WHERE account_profiles.email = ?");
      for (const profile of [...this.db.accountProfiles.values()].sort((left, right) => left.email.localeCompare(right.email))) {
        if (tokenFilter && profile.subscription_token !== args[0]) continue;
        if (emailFilter && profile.email !== args[0]) continue;
        const state = this.db.accountStates.get(profile.email) || null;
        const claims = [...this.db.accountClaims.values()]
          .filter((claim) => claim.email === profile.email)
          .sort((left, right) => left.source_type.localeCompare(right.source_type) || left.display_name.localeCompare(right.display_name));
        if (!claims.length) {
          results.push({
            email: profile.email,
            real_name: profile.real_name,
            role: profile.role,
            insights_enabled: profile.insights_enabled,
            subscription_token: profile.subscription_token,
            password_salt: profile.password_salt,
            password_hash: profile.password_hash,
            admin_issues_json: profile.admin_issues_json,
            local_parser_extensions_json: profile.local_parser_extensions_json,
            created_at: profile.created_at,
            updated_at: profile.updated_at,
            source_type: null,
            doctor_key: null,
            display_name: null,
            matched_at: null,
            session_json: state?.session_json || null,
          });
          continue;
        }
        for (const claim of claims) {
          results.push({
            email: profile.email,
            real_name: profile.real_name,
            role: profile.role,
            insights_enabled: profile.insights_enabled,
            subscription_token: profile.subscription_token,
            password_salt: profile.password_salt,
            password_hash: profile.password_hash,
            admin_issues_json: profile.admin_issues_json,
            local_parser_extensions_json: profile.local_parser_extensions_json,
            created_at: profile.created_at,
            updated_at: profile.updated_at,
            source_type: claim.source_type,
            doctor_key: claim.doctor_key,
            display_name: claim.display_name,
            matched_at: claim.matched_at,
            session_json: state?.session_json || null,
          });
        }
      }
      return { results };
    }
    if (sql.startsWith("SELECT rule_json FROM parser_rules")) {
      return {
        results: [...this.db.parserRules.values()]
          .filter((rule) => rule.scope === "global")
          .map((rule) => ({ rule_json: rule.rule_json })),
      };
    }
    if (sql.startsWith("SELECT suggestion_json FROM parser_rule_suggestions")) {
      return {
        results: [...this.db.parserRuleSuggestions.values()]
          .filter((suggestion) => suggestion.status === "pending")
          .sort((left, right) => right.updated_at.localeCompare(left.updated_at))
          .map((suggestion) => ({ suggestion_json: suggestion.suggestion_json })),
      };
    }
    if (sql.startsWith("SELECT * FROM doctor_profiles ORDER BY")) {
      return {
        results: [...this.db.doctorProfiles.values()]
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.profile_id.localeCompare(right.profile_id)),
      };
    }
    if (sql.includes("FROM console_messages")) {
      return {
        results: [...this.db.consoleMessages]
          .sort((left, right) => String(right.created_at).localeCompare(String(left.created_at)) || right.id - left.id)
          .slice(0, Number(args[0] || 50)),
      };
    }
    throw new Error(`Unsupported MemoryD1 all SQL: ${sql}`);
  }

  async first() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ? AND doctor_key = ?")) {
      return {
        count: [...this.db.events.values()].filter((event) => event.file_id === args[0] && event.doctor_key === args[1]).length,
      };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_events WHERE file_id")) {
      return {
        count: [...this.db.events.values()].filter((event) => event.file_id === args[0]).length,
      };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_file_doctors WHERE file_id")) {
      return {
        count: [...this.db.fileDoctors.values()].filter((doctor) => doctor.file_id === args[0]).length,
      };
    }
    if (sql.startsWith("SELECT session_json FROM account_states WHERE email")) {
      return this.db.accountStates.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT email FROM subscription_tokens WHERE token")) {
      return this.db.subscriptionTokens.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_profiles WHERE subscription_token")) {
      return { count: [...this.db.accountProfiles.values()].filter((profile) => profile.subscription_token).length };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_profiles")) {
      return { count: this.db.accountProfiles.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_claims")) {
      return { count: this.db.accountClaims.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_states")) {
      return { count: this.db.accountStates.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM doctor_profiles")) {
      return { count: this.db.doctorProfiles.size };
    }
    if (sql.startsWith("SELECT * FROM doctor_profiles WHERE profile_id")) {
      return this.db.doctorProfiles.get(args[0]) || null;
    }
    if (sql.includes("FROM raw_roster_files") && sql.includes("WHERE file_id = ?")) {
      return this.db.rawFiles.get(args[0]) || null;
    }
    throw new Error(`Unsupported MemoryD1 first SQL: ${sql}`);
  }
}

function repositoryFile(id, overrides = {}) {
  return {
    repoId: id,
    id,
    name: `${id}.xlsx`,
    sourceType: "mmc",
    active: true,
    size: 12,
    lastModified: 1,
    doctors: [{
      key: "TITUS HACKMAN",
      displayName: "Titus HACKMAN",
      sourceType: "mmc",
    }],
    ...overrides,
  };
}

async function seedRepository(store, files) {
  await store.put("repository:index", JSON.stringify({ version: 1, files }));
  for (const file of files) {
    await store.put(`repository:file:${file.id}`, JSON.stringify({
      ...file,
      dataUrl: `data:application/octet-stream;base64,${Buffer.from(file.id).toString("base64")}`,
    }));
  }
}

function seedD1Repository(db, files) {
  db.canonicalDoctors.clear();
  for (const file of files) {
    db.files.set(file.id, {
      id: file.id,
      name: file.name || `${file.id}.xlsx`,
      source_type: file.sourceType || "mmc",
      active: file.active === false ? 0 : 1,
      size: file.size || 0,
      last_modified: file.lastModified || 0,
      added_at: file.addedAt || "",
      uploaded_at: file.uploadedAt || "",
      uploaded_by: file.uploadedBy || "",
      parsed_at: file.parsedAt || "",
    });
    for (const doctor of file.doctors || []) {
      db.fileDoctors.set(`${file.id}|${doctor.sourceType || file.sourceType || "mmc"}|${doctor.key}`, {
        file_id: file.id,
        source_type: doctor.sourceType || file.sourceType || "mmc",
        doctor_key: doctor.key,
        display_name: doctor.displayName || doctor.key,
      });
    }
  }
}

async function seedUser(store, email, password, realName = "Titus Hackman", db = null) {
  await postState(store, {
    action: "login",
    email,
    password,
    mode: "create",
    realName,
  }, db);
}

async function postState(store, payload, db = null) {
  const { response, body } = await postStateRaw(store, payload, db);
  assert.equal(response.ok, true, body.error || "state request failed");
  return body;
}

async function postStateRaw(store, payload, db = null) {
  const rosterDb = db || store?.d1 || new MemoryD1();
  const response = await handleStatePost({
    request: new Request("http://fixture.test/api/state", {
      method: "POST",
      body: JSON.stringify(payload),
      headers: { "content-type": "application/json" },
    }),
    env: { ROSTER_DB: rosterDb, ROSTER_FILES: store?.r2 || new MemoryR2() },
    waitUntil: () => {},
  });
  const body = await response.json();
  return { response, body };
}

function memoryD1AccountRecord(db, email) {
  const profile = db.accountProfiles.get(email);
  if (!profile) return null;
  return {
    email,
    realName: profile.real_name,
    role: profile.role,
    adminIssues: JSON.parse(profile.admin_issues_json || "[]"),
    localParserExtensions: JSON.parse(profile.local_parser_extensions_json || "[]"),
    claims: [...db.accountClaims.values()].filter((claim) => claim.email === email).map((claim) => ({
      sourceType: claim.source_type,
      key: claim.doctor_key,
      displayName: claim.display_name,
    })),
  };
}

const stateStore = new MemoryStore();
stateStore.d1 = new MemoryD1();
const creatorPassword = "fixture-password";
const wranglerConfig = await readFile(new URL("../wrangler.toml", import.meta.url), "utf8");
assert.ok(/binding\s*=\s*"ROSTER_DB"/.test(wranglerConfig), "wrangler.toml must bind D1 ROSTER_DB");
assert.ok(!/ROSTER_STORE/.test(wranglerConfig), "wrangler.toml must not reintroduce KV ROSTER_STORE");
await postState(stateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await seedRepository(stateStore, [repositoryFile("fixture-roster", {
  name: "AdultMMCTerm2.2026.Ver1.pdf",
  sourceType: "mmc",
})]);

const creatorImports = await postStateRaw(stateStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorImports.response.status, 410);
const emptyD1Status = await postState(stateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(emptyD1Status.total, 0, "KV repository metadata must not be used when D1 roster_files is empty");

const d1StateStore = new MemoryStore();
const d1Store = new MemoryD1();
const d1Doctor = doctorOptions(parsedMmcUpload.sources.mmc, [], [], [])[0];
assert.ok(d1Doctor?.key, "fixture should expose at least one MMC doctor");
await postState(d1StateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
const d1Doctors = doctorOptions(parsedMmcUpload.sources.mmc, [], [], []);
const d1EventsByDoctor = Object.fromEntries(d1Doctors.map((doctor) => [
  doctor.key,
  buildRosterView(parsedMmcUpload.sources.mmc, [], doctor.key, {}, {}, {}, [], [], []).events,
]));
await postState(d1StateStore, {
  action: "uploadRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "d1-mmc",
    name: "AdultTerm1.2026.xlsx",
    size: 123,
    lastModified: 1,
    addedAt: "2026-01-01T00:00:00.000Z",
    sourceType: "mmc",
  },
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  dataUrl: workbookDataUrl(mmcWorkbook),
}, d1Store);
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "d1-mmc",
    name: "AdultTerm1.2026.xlsx",
    size: 123,
    lastModified: 1,
    addedAt: "2026-01-01T00:00:00.000Z",
    sourceType: "mmc",
  },
  doctors: d1Doctors,
  eventsByDoctor: d1EventsByDoctor,
}, d1Store);
const retainedRaw = await postState(d1StateStore, {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "d1-mmc",
}, d1Store);
assert.match(retainedRaw.dataUrl, /^data:/, "retained raw files should be fetchable for browser-side reparsing");
const legacyRawStore = new MemoryStore();
const legacyRawDb = new MemoryD1();
await postState(legacyRawStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, legacyRawDb);
legacyRawDb.rawFiles.set("legacy-raw", {
  file_id: "legacy-raw",
  object_key: "",
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  data_url: workbookDataUrl(mmcWorkbook),
  uploaded_at: "2026-01-01T00:00:00.000Z",
});
const migratedRaw = await postState(legacyRawStore, {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "legacy-raw",
}, legacyRawDb);
assert.match(migratedRaw.dataUrl, /^data:/, "legacy raw files should remain fetchable during lazy migration");
assert.equal(legacyRawDb.rawFiles.get("legacy-raw")?.object_key, "rosters/legacy-raw", "legacy raw files should be promoted to R2 when fetched");
assert.equal(legacyRawDb.rawFiles.get("legacy-raw")?.data_url, "", "lazy migration should clear the inline D1 payload after promotion");
legacyRawDb.rawFiles.set("retained-only:1:1", {
  file_id: "retained-only:1:1",
  name: "Dandenong retained only.xlsx",
  source_type: "ddh",
  size: 1,
  last_modified: 1,
  object_key: "rosters/retained-only:1:1",
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  data_url: "",
  uploaded_at: "2026-01-01T00:01:00.000Z",
});
const retainedOnlyStatus = await postState(legacyRawStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, legacyRawDb);
const retainedOnlyFile = retainedOnlyStatus.files.find((file) => file.id === "retained-only:1:1");
assert.equal(retainedOnlyFile?.retainedSourceOnly, true, "calendar status should list retained R2 source files without derived rows");
assert.equal(retainedOnlyFile?.status, "retained");
const reparsedD1Status = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: d1Doctor.key,
}, d1Store);
assert.equal(reparsedD1Status.files.find((file) => file.id === "d1-mmc")?.rawSourceAvailable, true, "durably stored raw files should be visible in status");

const missingRawDb = new MemoryD1();
await postState(new MemoryStore(), {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, missingRawDb);
await postState(new MemoryStore(), {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "missing-raw", name: "missing-raw.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "missing-raw-shift", source: "MMC", title: "MMC shift", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "MMC shift" }],
  },
}, missingRawDb);
const missingRawReparse = await postStateRaw(new MemoryStore(), {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "missing-raw",
}, missingRawDb);
assert.equal(missingRawReparse.response.status, 404, "files without retained raw bytes should fail source fetch explicitly");
await postState(d1StateStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{
      id: "d1-mmc",
      name: "AdultTerm1.2026.xlsx",
      size: 123,
      lastModified: 1,
      addedAt: "2026-01-01T00:00:00.000Z",
      sourceType: "mmc",
    }],
    session: {
      doctorKey: d1Doctor.key,
      settings: {},
    },
  },
}, d1Store);
assert.ok(d1Store.events.size > 0, "D1 should contain derived roster events after creator save");
const transactionalDb = new MemoryD1();
const transactionalStore = new MemoryStore();
await postState(transactionalStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, transactionalDb);
await postState(transactionalStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "safe-refresh", name: "safe-refresh.xlsx", sourceType: "mmc" },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "original", source: "MMC", title: "Original", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Original" }],
  },
}, transactionalDb);
transactionalDb.failNextEventInsert = true;
const failedRefresh = await postStateRaw(transactionalStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "safe-refresh", name: "safe-refresh.xlsx", sourceType: "mmc" },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "replacement", source: "MMC", title: "Replacement", allDay: true, start: "2026-02-04", end: "2026-02-04", rawValue: "Replacement" }],
  },
}, transactionalDb);
assert.equal(failedRefresh.response.ok, false, "injected replacement failure should surface");
assert.deepEqual(
  [...transactionalDb.events.values()].map((event) => event.title),
  ["Original"],
  "failed derived-file refresh should preserve the previous events transactionally",
);
const d1CreatorLogin = await postState(d1StateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1CreatorLogin.snapshot, null, "login should not build or return a full calendar snapshot");
assert.equal(d1CreatorLogin.state.session.doctorKey, d1Doctor.key, "creator login should keep selected doctor metadata");
assert.equal(d1CreatorLogin.subscription.enabled, true, "creator account should expose subscription URL capability");
const d1CreatorCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1CreatorCalendar.snapshot?.preview?.derivedFromD1, true);
assert.ok(d1CreatorCalendar.snapshot.preview.events.length > 0);
assert.ok(d1CreatorCalendar.snapshot.detectedSources.mmc.length > 0, "creator D1 snapshots should retain detected roster sources");
const d1CreatorFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1CreatorLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1CreatorFeedResponse.ok, true, "creator subscription feed should resolve from the creator-selected doctor");
const d1CreatedUser = await postState(d1StateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
  targetRealName: d1Doctor.displayName,
  targetPassword: "d1-password",
}, d1Store);
assert.ok(d1CreatedUser.user.claims.length > 0, "admin-created account should immediately claim exact roster matches");
const d1DirectLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1DirectLogin.snapshot, null, "claimed login should not build or return a full calendar snapshot");
assert.equal(d1DirectLogin.state.session.doctorKey, d1Doctor.key, "claimed login should default to the claimed doctor");
const d1DirectCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(d1DirectCalendar.snapshot.detectedSources.mmc.length > 0, "claimed D1 snapshots should retain linked roster sources");
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "ddh-rover-diagnostics",
    name: "Dandenong 2026.xlsx",
    sourceType: "ddh",
    active: true,
  },
  doctors: [{ key: d1Doctor.key, displayName: d1Doctor.displayName, sourceType: "ddh" }],
  eventsByDoctor: {
    [d1Doctor.key]: [{
      id: "ddh-rover-event",
      source: "DDH",
      seniority: "Unknown",
      title: "DDH: Rover AM",
      allDay: false,
      start: "2026-05-14T08:00:00",
      end: "2026-05-14T18:00:00",
      rawValue: "Rover AM",
      timeLabel: "08:00-18:00",
    }],
  },
  issuesByDoctor: {
    [d1Doctor.key]: [
      {
        id: "DDH::Unknown::2026-05-14::Rover AM",
        source: "DDH",
        seniority: "Unknown",
        startDay: "2026-05-14",
        rawValue: "Rover AM",
        status: "unknown",
        message: "DDH shift code not recognised; using explicit roster time.",
        resolutionType: "shift_code",
        suggestedTitle: "DDH: Rover AM",
        timeLabel: "08:00-18:00",
      },
      {
        id: "DDH::Unknown::2026-05-15::Mystery AM",
        source: "DDH",
        seniority: "Unknown",
        startDay: "2026-05-15",
        rawValue: "Mystery AM",
        status: "unknown",
        message: "DDH shift code not recognised; using explicit roster time.",
        resolutionType: "shift_code",
        suggestedTitle: "DDH: Mystery AM",
        timeLabel: "08:00-18:00",
      },
    ],
  },
}, d1Store);
assert.equal(
  [...d1Store.issues.values()].filter((issue) => issue.file_id === "ddh-rover-diagnostics").length,
  2,
  "derived roster saves should persist parser diagnostics transactionally",
);
const d1StoredIssueCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(
  d1StoredIssueCalendar.snapshot.preview.events.some((event) => event.title === "DDH: Rover AM"),
  "DDH Rover shifts should render from indexed events",
);
assert.equal(
  d1StoredIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Rover AM"),
  false,
  "stale DDH Rover diagnostics should be hidden on first D1 calendar load",
);
assert.equal(
  d1StoredIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Mystery AM"),
  true,
  "calendar loads should show unresolved diagnostics stored during import",
);
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "ddh-rover-diagnostics",
    name: "Dandenong 2026.xlsx",
    sourceType: "ddh",
    active: true,
  },
  doctors: [{ key: d1Doctor.key, displayName: d1Doctor.displayName, sourceType: "ddh" }],
  eventsByDoctor: {
    [d1Doctor.key]: [{
      id: "ddh-rover-event",
      source: "DDH",
      seniority: "Unknown",
      title: "DDH: Rover AM",
      allDay: false,
      start: "2026-05-14T08:00:00",
      end: "2026-05-14T18:00:00",
      rawValue: "Rover AM",
      timeLabel: "08:00-18:00",
    }],
  },
  issuesByDoctor: {},
}, d1Store);
assert.equal(
  [...d1Store.issues.values()].some((issue) => issue.file_id === "ddh-rover-diagnostics"),
  false,
  "reparsed roster files should replace stale diagnostics instead of accumulating them",
);
const d1ClearedIssueCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(
  d1ClearedIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Mystery AM"),
  false,
  "cleared import diagnostics should not reappear on later D1 calendar loads",
);
for (const [key, event] of [...d1Store.events.entries()]) if (event.file_id === "ddh-rover-diagnostics") d1Store.events.delete(key);
for (const [key, issue] of [...d1Store.issues.entries()]) if (issue.file_id === "ddh-rover-diagnostics") d1Store.issues.delete(key);
for (const [key, doctor] of [...d1Store.fileDoctors.entries()]) if (doctor.file_id === "ddh-rover-diagnostics") d1Store.fileDoctors.delete(key);
d1Store.files.delete("ddh-rover-diagnostics");
await seedUser(d1StateStore, "admin-enter-match@example.com", "admin-enter-password", d1Doctor.displayName, d1Store);
for (const key of [...d1Store.accountClaims.keys()]) {
  if (key.startsWith("admin-enter-match@example.com|")) d1Store.accountClaims.delete(key);
}
const adminEnteredResolution = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "admin-enter-match@example.com",
}, d1Store);
assert.ok(
  adminEnteredResolution.claims.some((claim) => claim.key === d1Doctor.key && claim.sourceType === "mmc"),
  "creator-entered matching users should resolve and persist D1 roster claims",
);
const adminEnteredCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "admin-enter-match@example.com",
}, d1Store);
assert.equal(
  adminEnteredCalendar.snapshot?.preview?.derivedFromD1,
  true,
  "creator-entered matching users should load calendars from D1 after claim resolution",
);
for (const key of [...d1Store.accountClaims.keys()]) {
  if (key.startsWith("admin-enter-match@example.com|")) d1Store.accountClaims.delete(key);
}
const d1OnsiteMmcEvent = d1DirectCalendar.snapshot.preview.events.find((event) => event.source === "MMC" && !/\\b(CS|leave|conference|PHNW)\\b/i.test(`${event.title} ${event.rawValue}`));
assert.ok(d1OnsiteMmcEvent?.location, "D1 account calendar load should apply SQL-backed hospital defaults to onsite shifts");
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    version: 1,
    imports: d1DirectLogin.state.imports,
    session: {
      doctorKey: d1Doctor.key,
      settings: {
        defaultLocationMmc: "User One MMC Location",
      },
    },
  },
}, d1Store);
for (const row of d1Store.events.values()) {
  const event = JSON.parse(row.event_json);
  if (event.source !== "MMC" || /\b(CS|leave|conference|PHNW)\b/i.test(`${event.title} ${event.rawValue}`)) continue;
  event.location = "";
  row.location = "";
  row.event_json = JSON.stringify(event);
  break;
}
const d1CustomLocationCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(
  d1CustomLocationCalendar.snapshot.preview.events.some((event) => event.source === "MMC" && event.location === "User One MMC Location"),
  "account SQL hospital location should override shared roster-event location",
);
const d1UserFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1UserFeedResponse.ok, true, "claimed user subscription feed should resolve");
assert.match(await d1UserFeedResponse.text(), /LOCATION:User One MMC Location/, "subscription feed should apply account hospital defaults");
const d1OverrideEvent = d1DirectCalendar.snapshot.preview.events[0];
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    version: 1,
    imports: d1DirectLogin.state.imports,
    session: {
      doctorKey: d1Doctor.key,
      exportRange: { startDate: "2026-02-01", endDate: "2026-02-28", allFuture: false },
      settings: {},
      overrides: {
        [d1OverrideEvent.id]: {
          title: "D1 Edited Shift",
        },
      },
      customEvents: [{
        id: "d1-custom-event",
        ownerEmail: "d1-user@example.com",
        title: "D1 Custom Event stale",
        startDate: "2026-02-12",
        endDate: "2026-02-12",
        allDay: true,
        include: true,
      }, {
        id: "d1-custom-event",
        ownerEmail: "d1-user@example.com",
        title: "D1 Custom Event",
        startDate: "2026-02-13",
        endDate: "2026-02-13",
        allDay: true,
        include: true,
      }],
    },
  },
}, d1Store);
const d1SessionLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1SessionLogin.state.session.exportRange.startDate, "2026-02-01", "D1 account session should load without KV state");
assert.deepEqual(
  d1SessionLogin.state.session.customEvents.map((event) => [event.id, event.title, event.startDate]),
  [["d1-custom-event", "D1 Custom Event", "2026-02-13"]],
  "D1 session save should collapse duplicate custom event ids with the latest value winning",
);
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    ...d1SessionLogin.state,
    session: {
      ...d1SessionLogin.state.session,
      customEvents: [
        ...d1SessionLogin.state.session.customEvents,
        { id: "different-id-same-event", ownerEmail: "d1-user@example.com", title: "D1 Custom Event", startDate: "2026-02-13", endDate: "2026-02-13", allDay: true, include: true },
      ],
    },
  },
}, d1Store);
const d1IdentityDedupedLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1IdentityDedupedLogin.state.session.customEvents.length, 1, "matching logical custom events should collapse even when ids differ");
const d1SessionCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(d1SessionCalendar.snapshot.preview.events.some((event) => event.title === "D1 Edited Shift"), "D1 calendar load should apply session overrides");
assert.ok(d1SessionCalendar.snapshot.preview.events.some((event) => event.title === "D1 Custom Event"), "D1 calendar load should include session custom events");
assert.equal(d1SessionCalendar.snapshot.preview.customEventsMaterialized, true, "D1 calendar snapshots should declare custom events already materialized");
assert.equal(
  d1SessionCalendar.snapshot.preview.events.filter((event) => event.title === "D1 Custom Event").length,
  1,
  "D1 calendar snapshots should carry each logical custom event once",
);
const d1AdminLoad = await postState(d1StateStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
}, d1Store);
assert.equal(d1AdminLoad.snapshot, null, "admin user switch should not build or return a full calendar snapshot");
const d1AdminCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
}, d1Store);
assert.deepEqual(
  d1AdminCalendar.snapshot.preview.events.map((event) => event.id),
  d1SessionCalendar.snapshot.preview.events.map((event) => event.id),
  "direct and creator-switched user loads should use the same D1 calendar events",
);

const recreateDb = new MemoryD1();
const recreateStore = new MemoryStore();
seedD1Repository(recreateDb, [{
  id: "recreate-roster",
  name: "recreate.xlsx",
  sourceType: "mmc",
  active: true,
  doctors: [{ key: "TITUS HACKMAN", displayName: "Titus Hackman", sourceType: "mmc" }],
  eventsByDoctor: {
    "TITUS HACKMAN": [{ id: "recreate-shift", source: "MMC", title: "Roster shift", seniority: "Intern", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Roster shift" }],
  },
}]);
recreateDb.events.set("recreate-event", {
  id: "recreate-event",
  file_id: "recreate-roster",
  source_type: "mmc",
  doctor_key: "TITUS HACKMAN",
  display_name: "Titus Hackman",
  start_date: "2026-02-03",
  end_date: "2026-02-03",
  start_ts: "2026-02-03",
  end_ts: "2026-02-03",
  title: "Roster shift",
  raw_value: "Roster shift",
  seniority: "Intern",
  location: "",
  all_day: 1,
  time_label: "",
  event_json: JSON.stringify({ id: "recreate-shift", source: "MMC", title: "Roster shift", seniority: "Intern", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Roster shift" }),
});
await postState(recreateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, recreateDb);
const firstMatchedCreate = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "recreated@example.com",
  targetRealName: "Titus Hackman",
  targetPassword: "recreated-password",
}, recreateDb);
assert.equal(firstMatchedCreate.user.claims.length, 1, "first matched create should auto-claim the roster doctor");
assert.deepEqual(firstMatchedCreate.user.seniorities, ["Intern"], "user summaries should include roster-derived seniorities");
await postState(recreateStore, {
  action: "deleteAccount",
  email: "recreated@example.com",
  password: "recreated-password",
}, recreateDb);
const recreatedMatchedUser = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "recreated@example.com",
  targetRealName: "Titus Hackman",
  targetPassword: "recreated-password-2",
}, recreateDb);
assert.equal(recreatedMatchedUser.user.claims.length, 1, "recreating a self-deleted matched user should auto-claim cleanly");
const madeUpUser = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "made-up@example.com",
  targetRealName: "Made Up",
  targetPassword: "made-up-password",
}, recreateDb);
assert.equal(madeUpUser.user.claims.length, 0, "creating a no-roster user should remain lightweight and unclaimed");
const d1Insights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
}, d1Store);
assert.ok(Array.isArray(d1Insights.coworkers));
const d1InsightsExcludingSelectedDoctor = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(
  d1InsightsExcludingSelectedDoctor.coworkers.every((row) => row.doctorKey !== d1Doctor.key),
  "coworker lookup should exclude the selected doctor in SQL",
);
const d1OverlapInsights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(Array.isArray(d1OverlapInsights.coworkers), "overlap coworker lookup should return SQL-derived rows");
assert.ok(
  d1OverlapInsights.coworkers.every((row) => row.doctorKey !== d1Doctor.key),
  "overlap coworker lookup should exclude the selected doctor in SQL",
);
const d1OverlapDoctors = await postState(d1StateStore, {
  action: "queryRosterOverlapDoctors",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(Array.isArray(d1OverlapDoctors.doctors), "overlap doctor lookup should return compact doctor rows");
const d1RepositoryFile = [...d1Store.files.keys()][0];
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: d1RepositoryFile,
    name: "AdultTerm1.2026.xlsx",
    sourceType: "mmc",
    active: true,
  },
  doctors: [...d1Store.fileDoctors.values()].map((doctor) => ({
    key: doctor.doctor_key,
    displayName: doctor.display_name,
    sourceType: doctor.source_type,
  })),
  eventsByDoctor: Object.fromEntries(
    [...d1Store.fileDoctors.values()].map((doctor) => [
      doctor.doctor_key,
      [...d1Store.events.values()]
        .filter((event) => event.doctor_key === doctor.doctor_key)
        .map((event) => JSON.parse(event.event_json)),
    ]),
  ),
}, d1Store);
const d1Status = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  expectedFileIds: ["d1-mmc", "missing-d1-mmc"],
}, d1Store);
assert.equal(d1Status.total, 1);
assert.equal(d1Status.populated, 1);
assert.equal(d1Status.remaining, 0);
assert.equal(d1Status.expectedFiles.expectedCount, 2);
assert.equal(d1Status.expectedFiles.persistedCount, 1);
assert.equal(d1Status.expectedFiles.populatedCount, 1);
assert.deepEqual(d1Status.expectedFiles.persistedFileIds, ["d1-mmc"]);
assert.deepEqual(d1Status.expectedFiles.populatedFileIds, ["d1-mmc"]);
assert.deepEqual(d1Status.expectedFiles.missingFileIds, ["missing-d1-mmc"]);
assert.ok(d1Status.accounts.profiles >= 2);
assert.ok(d1Status.accounts.claims >= 1);
assert.ok(d1Status.accounts.states >= 2);
assert.ok(d1Status.accounts.subscriptionTokens >= 2);
await d1StateStore.delete("repository:index");
const d1OnlyRepositoryStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1OnlyRepositoryStatus.total, 1, "D1 roster_files should supply calendar status without KV repository index");
assert.equal(d1OnlyRepositoryStatus.populated, 1);
const d1UserList = await postState(d1StateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.ok(d1UserList.availableDoctors.some((doctor) => doctor.key === d1Doctor.key));
assert.equal(
  d1UserList.availableDoctors.find((doctor) => doctor.key === d1Doctor.key)?.claimedBy,
  "d1-user@example.com",
  "D1 doctor directory should include claimed account metadata for the Creator switcher",
);
const d1NoKvIndexLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvIndexLogin.snapshot, null, "D1 account login should stay lightweight without KV repository index");
const d1NoKvIndexCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvIndexCalendar.snapshot?.preview?.derivedFromD1, true, "D1 account calendar load should work without KV repository index");
assert.ok(d1NoKvIndexCalendar.snapshot.fileRefs.some((ref) => ref.id === d1RepositoryFile), "D1 claimed-account calendar snapshots should include source file refs for the Account modal");
const d1NoKvIndexEnrichment = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.deepEqual(d1NoKvIndexEnrichment.availableDoctors, [], "claimed-account enrichment should not reload the full doctor directory");
await seedUser(d1StateStore, "d1-unmatched@example.com", "d1-unmatched-password", "Unmatched Person", d1Store);
const d1UnmatchedEnrichment = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "d1-unmatched@example.com",
  password: "d1-unmatched-password",
}, d1Store);
assert.ok(d1UnmatchedEnrichment.availableDoctors.some((doctor) => doctor.key === d1Doctor.key), "D1 doctor directory should load only when claim resolution still leaves the account unclaimed");
const d1ClaimResolution = await postState(d1StateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: d1Doctor.key,
    displayName: d1Doctor.displayName,
    sourceTypes: ["mmc"],
  },
}, d1Store);
assert.equal(d1ClaimResolution.mode, "claimed-account");
assert.equal(d1ClaimResolution.email, "d1-user@example.com");
const d1DoctorProfile = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
}, d1Store);
assert.equal(d1DoctorProfile.snapshot?.preview?.derivedFromD1, true);
assert.equal(d1DoctorProfile.snapshotStale, false);
assert.ok(d1DoctorProfile.snapshot.preview.events.length > 0);
assert.ok(d1DoctorProfile.snapshot.fileRefs.some((ref) => ref.id === d1RepositoryFile), "D1 doctor profile should derive file refs without KV repository index");
const d1RepositoryDoctors = [...d1Store.fileDoctors.values()].map((doctor) => ({
  key: doctor.doctor_key,
  displayName: doctor.display_name,
  sourceType: doctor.source_type,
}));
const d1RepositoryEventsByDoctor = Object.fromEntries(
  [...d1Store.fileDoctors.values()].map((doctor) => [
    doctor.doctor_key,
    [...d1Store.events.values()]
      .filter((event) => event.doctor_key === doctor.doctor_key)
      .map((event) => JSON.parse(event.event_json)),
  ]),
);
await postState(d1StateStore, {
  action: "resetDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: d1RepositoryFile,
}, d1Store);
const d1InactiveRepositoryStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1InactiveRepositoryStatus.total, 0, "inactive D1 roster files should not drive active calendar status");
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: d1RepositoryFile,
    name: "AdultTerm1.2026.xlsx",
    sourceType: "mmc",
    active: true,
  },
  doctors: d1RepositoryDoctors,
  eventsByDoctor: d1RepositoryEventsByDoctor,
}, d1Store);
const d1ProfileOverrideEvent = d1DoctorProfile.snapshot.preview.events[0];
await postState(d1StateStore, {
  action: "saveDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
  state: {
    version: 1,
    imports: [],
    session: {
      doctorKey: d1Doctor.key,
      hadPreview: true,
      settings: {},
      overrides: {
        [d1ProfileOverrideEvent.id]: { title: "D1 Profile Edited Shift" },
      },
      customEvents: [{
        id: "d1-profile-custom-event",
        title: "D1 Profile Custom Event",
        startDate: "2026-02-13",
        endDate: "2026-02-13",
        allDay: true,
        include: true,
      }],
    },
  },
}, d1Store);
await d1StateStore.delete(`doctor-profile:${d1Doctor.key}::mmc`);
await d1StateStore.delete(`snapshot:doctor-profile:${d1Doctor.key}::mmc`);
const d1OnlyDoctorProfile = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
}, d1Store);
assert.equal(d1OnlyDoctorProfile.snapshot?.preview?.derivedFromD1, true);
assert.equal(d1OnlyDoctorProfile.snapshot?.preview?.customEventsMaterialized, true);
assert.ok(d1OnlyDoctorProfile.snapshot.preview.events.some((event) => event.title === "D1 Profile Edited Shift"), "D1 doctor profile should apply stored overrides without KV profile state");
assert.ok(d1OnlyDoctorProfile.snapshot.preview.events.some((event) => event.title === "D1 Profile Custom Event"), "D1 doctor profile should include stored custom events without KV profile state");
const d1ProfileStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.ok(d1ProfileStatus.accounts.doctorProfiles >= 1);
const leaveMergeStore = new MemoryStore();
const leaveMergeDb = new MemoryD1();
const leaveDoctor = { key: "LEAVE DOCTOR", displayName: "Leave Doctor", sourceType: "mmc" };
await seedRepository(leaveMergeStore, [
  repositoryFile("leave-mmc", { sourceType: "mmc", doctors: [leaveDoctor] }),
  repositoryFile("leave-ddh", { sourceType: "ddh", doctors: [{ ...leaveDoctor, sourceType: "ddh" }] }),
]);
await postState(leaveMergeStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, leaveMergeDb);
for (const [fileId, sourceType] of [["leave-mmc", "mmc"], ["leave-ddh", "ddh"]]) {
  await postState(leaveMergeStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ ...leaveDoctor, sourceType }],
    eventsByDoctor: {
      [leaveDoctor.key]: [{
        id: `${fileId}-leave`,
        source: sourceType.toUpperCase() === "DDH" ? "DDH" : "MMC",
        title: "Annual Leave",
        allDay: true,
        start: "2026-04-06",
        end: "2026-04-13",
        rawValue: "Annual Leave",
        monthKey: "2026-04",
      }],
    },
  }, leaveMergeDb);
}
await seedUser(leaveMergeStore, "leave@example.com", "leave-password", "Leave Doctor", leaveMergeDb);
await postState(leaveMergeStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "leave@example.com",
  claims: [{ sourceType: "mmc", key: leaveDoctor.key }, { sourceType: "ddh", key: leaveDoctor.key }],
}, leaveMergeDb);
const leaveLogin = await postState(leaveMergeStore, {
  action: "login",
  email: "leave@example.com",
  password: "leave-password",
}, leaveMergeDb);
assert.equal(leaveLogin.snapshot, null, "login should stay lightweight for duplicate leave accounts");
const leaveCalendar = await postState(leaveMergeStore, {
  action: "loadCalendarEvents",
  email: "leave@example.com",
  password: "leave-password",
}, leaveMergeDb);
const mergedLeave = leaveCalendar.snapshot.preview.events.filter((event) => event.title === "Annual Leave");
assert.equal(mergedLeave.length, 1);
assert.deepEqual(mergedLeave[0].sources.sort(), ["DDH", "MMC"]);
const aliasClaimsStore = new MemoryStore();
const aliasClaimsDb = new MemoryD1();
await postState(aliasClaimsStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, aliasClaimsDb);
for (const [fileId, sourceType, key, title] of [
  ["alias-mmc", "mmc", "ALIAS DOCTOR", "MMC Alias Shift"],
  ["alias-ddh", "ddh", "DR ALIAS DOCTOR", "DDH Alias Shift"],
]) {
  await postState(aliasClaimsStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName: "Alias Doctor", sourceType }],
    eventsByDoctor: {
      [key]: [{
        id: `${fileId}-shift`,
        source: sourceType.toUpperCase(),
        title,
        allDay: true,
        start: "2026-05-18",
        end: "2026-05-19",
        rawValue: title,
        monthKey: "2026-05",
      }],
    },
  }, aliasClaimsDb);
}
await seedUser(aliasClaimsStore, "alias@example.com", "alias-password", "Alias Doctor", aliasClaimsDb);
await postState(aliasClaimsStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "alias@example.com",
  claims: [
    { sourceType: "mmc", key: "ALIAS DOCTOR" },
    { sourceType: "ddh", key: "DR ALIAS DOCTOR" },
  ],
}, aliasClaimsDb);
const aliasCalendar = await postState(aliasClaimsStore, {
  action: "loadCalendarEvents",
  email: "alias@example.com",
  password: "alias-password",
  doctorKey: "ALIAS DOCTOR",
}, aliasClaimsDb);
assert.deepEqual(
  aliasCalendar.snapshot.preview.events.map((event) => event.title).sort(),
  ["DDH Alias Shift", "MMC Alias Shift"],
  "calendar load should include all selected doctor alias keys across hospitals",
);
const typoAliasStore = new MemoryStore();
const typoAliasDb = new MemoryD1();
await postState(typoAliasStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, typoAliasDb);
await postState(typoAliasStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "aeshan-mch", name: "aeshan-mch.xlsx", sourceType: "mch", active: true },
  doctors: [{ key: "AESHAN KULARATNE", displayName: "Aeshan KULARATNE", sourceType: "mch" }],
  eventsByDoctor: { "AESHAN KULARATNE": [{ id: "aeshan-mch-shift", source: "MCH", title: "MCH Shift", allDay: false, start: "2026-02-03T08:00:00", end: "2026-02-03T17:00:00", rawValue: "MCH Shift" }] },
}, typoAliasDb);
seedD1Repository(typoAliasDb, [
  repositoryFile("aeshan-ddh", { sourceType: "ddh", doctors: [{ key: "AESHAN KULURATNE", displayName: "Aeshan KULURATNE", sourceType: "ddh" }] }),
  repositoryFile("zero-only", { sourceType: "ddh", doctors: [{ key: "ZERO PERSON", displayName: "Zero PERSON", sourceType: "ddh" }] }),
]);
const typoDoctors = await postState(typoAliasStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, typoAliasDb);
const aeshanOption = typoDoctors.availableDoctors.find((doctor) => doctor.displayName === "Aeshan KULARATNE");
assert.ok(aeshanOption, "event-backed typo variants should keep the event-backed display name");
assert.deepEqual(aeshanOption.aliases.map((alias) => alias.key).sort(), ["AESHAN KULARATNE", "AESHAN KULURATNE"]);
assert.equal(typoDoctors.availableDoctors.some((doctor) => doctor.key === "ZERO PERSON"), false, "zero-event standalone identities should be hidden from the picker");
const typoProfile = await postState(typoAliasStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "AESHAN KULARATNE::ddh+mch",
  doctorKey: "AESHAN KULARATNE",
  displayName: "Aeshan KULARATNE",
  sourceTypes: ["ddh", "mch"],
}, typoAliasDb);
assert.deepEqual(typoProfile.snapshot.preview.events.map((event) => event.title), ["MCH Shift"]);
assert.equal(typoProfile.snapshot.profileCoverage.zeroEventAliases.length, 1);
assert.deepEqual(typoProfile.snapshot.profileCoverage.absentSources, []);
const conflictingAliasDb = new MemoryD1();
const conflictingAliasStore = new MemoryStore();
await postState(conflictingAliasStore, { action: "login", email: "rhaydon@gmail.com", password: creatorPassword }, conflictingAliasDb);
for (const [fileId, sourceType, key, displayName] of [
  ["aeshan-a", "mmc", "AESHAN KULARATNE", "Aeshan KULARATNE"],
  ["aeshan-b", "mch", "AESHAN KULURATNE", "Aeshan KULURATNE"],
]) {
  await postState(conflictingAliasStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName, sourceType }],
    eventsByDoctor: { [key]: [{ id: `${fileId}-shift`, source: sourceType.toUpperCase(), title: `${displayName} Shift`, allDay: false, start: "2026-05-05T08:00:00", end: "2026-05-05T17:00:00", rawValue: "Shift" }] },
  }, conflictingAliasDb);
}
const conflictingDoctors = await postState(conflictingAliasStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, conflictingAliasDb);
assert.equal(conflictingDoctors.availableDoctors.filter((doctor) => doctor.key === "AESHAN KULARATNE" || doctor.key === "AESHAN KULURATNE").length, 2, "overlapping working shifts should block typo merges");
const fourRosterStore = new MemoryStore();
const fourRosterDb = new MemoryD1();
await postState(fourRosterStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, fourRosterDb);
for (const [fileId, sourceType, key, displayName, title, start] of [
  ["creator-mmc-t1", "mmc", "RICHARD HAYDON", "Richard HAYDON", "MMC Term 1", "2026-02-03"],
  ["creator-ddh-t1", "ddh", "RICHARD HAYDON", "Richard Haydon", "DDH Term 1", "2026-02-04"],
  ["creator-mmc-t2", "mmc", "RICHARD HAYDON", "Richard HAYDON", "MMC Term 2", "2026-05-05"],
  ["creator-ddh-t2", "ddh", "HAYDON RICHARD", "HAYDON, Richard", "DDH Term 2", "2026-05-06"],
]) {
  await postState(fourRosterStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    selectedDoctorKey: "RICHARD HAYDON",
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName, sourceType }],
    eventsByDoctor: {
      [key]: [{
        id: `${fileId}-shift`,
        source: sourceType.toUpperCase(),
        title,
        allDay: true,
        start,
        end: start,
        rawValue: title,
        monthKey: start.slice(0, 7),
      }],
    },
  }, fourRosterDb);
}
const fourRosterStatus = await postState(fourRosterStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
}, fourRosterDb);
assert.equal(fourRosterStatus.total, 4, "all four active roster files should be reported");
assert.equal(fourRosterStatus.files.filter((file) => file.eventCount > 0).length, 4, "each roster file should have D1 event rows");
assert.equal(fourRosterStatus.files.filter((file) => file.selectedDoctorEventCount > 0).length, 4, "each roster file should have selected creator event rows");
const fourRosterCalendar = await postState(fourRosterStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
}, fourRosterDb);
assert.deepEqual(
  fourRosterCalendar.snapshot.preview.events.map((event) => event.title).sort(),
  ["DDH Term 1", "DDH Term 2", "MMC Term 1", "MMC Term 2"],
  "creator calendar load should include both hospitals across both terms",
);
assert.equal(fourRosterCalendar.diagnostics.cacheEngine, "r2", "default calendar loads should report r2 cache engine");
assert.equal(fourRosterCalendar.diagnostics.cacheHit, false, "default calendar loads should report cache miss on first load");
assert.equal(fourRosterCalendar.diagnostics.selectedDoctorKey, undefined, "default calendar loads should avoid file/doctor diagnostics");
const fourRosterDiagnosticCalendar = await postState(fourRosterStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
  diagnostics: true,
}, fourRosterDb);
assert.equal(fourRosterDiagnosticCalendar.diagnostics.queryMode, "file-doctor-pairs");
assert.equal(fourRosterDiagnosticCalendar.diagnostics.selectedDoctorFiles.length, 4, "diagnostics should include each resolved file/doctor pair when requested");
const fourRosterExpectedStatus = await postState(fourRosterStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
  expectedFileIds: ["creator-mmc-t1", "creator-ddh-t1", "creator-mmc-t2", "creator-ddh-t2"],
}, fourRosterDb);
assert.equal(fourRosterExpectedStatus.expectedFiles.expectedCount, 4);
assert.equal(fourRosterExpectedStatus.expectedFiles.persistedCount, 4);
assert.equal(fourRosterExpectedStatus.expectedFiles.activeCount, 4);
assert.deepEqual(fourRosterExpectedStatus.expectedFiles.missingFileIds, []);
for (let index = 1; index <= 55; index += 1) {
  await postState(fourRosterStore, {
    action: "appendConsoleMessage",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    message: `Console message ${index}`,
    isError: index === 55,
  }, fourRosterDb);
}
const consoleHistory = await postState(fourRosterStore, {
  action: "consoleMessages",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, fourRosterDb);
assert.equal(consoleHistory.messages.length, 50, "console history should keep only the latest 50 messages");
assert.equal(consoleHistory.messages[0].message, "Console message 55");
assert.equal(consoleHistory.messages[0].isError, true);
assert.equal(consoleHistory.messages.at(-1).message, "Console message 6");

const partialUploadStore = new MemoryStore();
const partialUploadDb = new MemoryD1();
await postState(partialUploadStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, partialUploadDb);
await postState(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  expectedFileIds: ["partial-mmc", "partial-ddh"],
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{
      id: "partial-mmc-shift",
      source: "MMC",
      title: "Persisted MMC Shift",
      allDay: true,
      start: "2026-02-03",
      end: "2026-02-03",
      rawValue: "Persisted MMC Shift",
      monthKey: "2026-02",
    }],
  },
}, partialUploadDb);
await postState(partialUploadStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [
      { repoId: "partial-mmc", id: "partial-mmc", sourceType: "mmc", name: "partial-mmc.xlsx" },
      { repoId: "partial-ddh", id: "partial-ddh", sourceType: "ddh", name: "partial-ddh.xlsx" },
    ],
    session: { doctorKey: "RICHARD HAYDON", settings: {} },
  },
}, partialUploadDb);
const partialUploadStatus = await postState(partialUploadStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
  expectedFileIds: ["partial-mmc", "partial-ddh"],
}, partialUploadDb);
assert.equal(partialUploadStatus.expectedFiles.persistedCount, 1, "partial upload status should report only persisted D1 files");
assert.equal(partialUploadStatus.expectedFiles.populatedCount, 1, "partial upload status should count only populated roster files as synced");
assert.deepEqual(partialUploadStatus.expectedFiles.missingFileIds, ["partial-ddh"]);
const partialUploadCalendar = await postState(partialUploadStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
}, partialUploadDb);
assert.deepEqual(
  partialUploadCalendar.snapshot.preview.events.map((event) => event.title),
  ["Persisted MMC Shift"],
  "post-login calendar rebuild should only use D1-persisted roster rows",
);
const invalidReplacement = await postStateRaw(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: { "RICHARD HAYDON": [] },
}, partialUploadDb);
assert.equal(invalidReplacement.response.status, 422, "empty derived uploads should be rejected before replacing D1 rows");
assert.equal(
  [...partialUploadDb.events.values()].filter((event) => event.file_id === "partial-mmc").length,
  1,
  "rejected derived uploads must preserve previously indexed D1 events",
);
const sparseReplacement = await postStateRaw(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [
    { key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" },
    { key: "SECOND DOCTOR", displayName: "Second Doctor", sourceType: "mmc" },
  ],
  eventsByDoctor: {
    "RICHARD HAYDON": [{
      id: "partial-mmc-shift",
      source: "MMC",
      title: "Persisted MMC Shift",
      allDay: true,
      start: "2026-02-03",
      end: "2026-02-03",
      rawValue: "Persisted MMC Shift",
      monthKey: "2026-02",
    }],
  },
}, partialUploadDb);
assert.equal(sparseReplacement.response.status, 422, "suspiciously sparse derived uploads should be rejected before replacing D1 rows");
const sharedUploadStore = new MemoryStore();
const sharedUploadDb = new MemoryD1();
await postState(sharedUploadStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, sharedUploadDb);
await seedUser(sharedUploadStore, "shared-user@example.com", "shared-password", "Shared User", sharedUploadDb);
const sharedUserSave = await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  file: { id: "shared-ddh", name: "Shared_DDH_04-05-2026_to_02-08-2026.xlsx", sourceType: "ddh", active: true, lastModified: 20 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "ddh" }],
  eventsByDoctor: {
    "SHARED USER": [{
      id: "shared-ddh-shift",
      source: "DDH",
      title: "Shared DDH Shift",
      allDay: true,
      start: "2026-05-06",
      end: "2026-05-06",
      rawValue: "Shared DDH Shift",
      monthKey: "2026-05",
    }],
  },
}, sharedUploadDb);
assert.equal(sharedUserSave.ok, true, "non-creator users should be able to add roster files to D1");
assert.equal(sharedUploadDb.files.get("shared-ddh")?.uploaded_by, "shared-user@example.com");
const sharedUserReset = await postStateRaw(sharedUploadStore, {
  action: "resetDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  fileId: "shared-ddh",
}, sharedUploadDb);
assert.equal(sharedUserReset.response.status, 403, "non-creator users must not remove D1 roster files");
for (const [fileId, name, lastModified, title] of [
  ["supersede-old", "MMC_Term2_2026_old.xlsx", 10, "Old MMC Shift"],
  ["supersede-new", "MMC_Term2_2026_new.xlsx", 30, "New MMC Shift"],
]) {
  await postState(sharedUploadStore, {
    action: "saveDerivedCalendarFile",
    email: "shared-user@example.com",
    password: "shared-password",
    file: { id: fileId, name, sourceType: "mmc", active: true, lastModified },
    doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mmc" }],
    eventsByDoctor: {
      "SHARED USER": [{
        id: `${fileId}-shift`,
        source: "MMC",
        title,
        allDay: true,
        start: "2026-05-07",
        end: "2026-05-07",
        rawValue: title,
        monthKey: "2026-05",
      }],
    },
  }, sharedUploadDb);
}
assert.equal(sharedUploadDb.files.get("supersede-old")?.active, 0, "older overlapping same-source roster should be deactivated");
assert.equal(sharedUploadDb.files.get("supersede-new")?.active, 1, "latest overlapping same-source roster should remain active");
for (const [fileId, name, lastModified, start, end] of [
  ["adjacent-term-1", "MMC_Term1_2026.xlsx", 10, "2026-05-03", "2026-05-04"],
  ["adjacent-term-2", "MMC_Term2_2026.xlsx", 30, "2026-05-04", "2026-05-04"],
]) {
  await postState(sharedUploadStore, {
    action: "saveDerivedCalendarFile",
    email: "shared-user@example.com",
    password: "shared-password",
    file: { id: fileId, name, sourceType: "mmc", active: true, lastModified },
    doctors: [{ key: "ADJACENT DOCTOR", displayName: "Adjacent Doctor", sourceType: "mmc" }],
    eventsByDoctor: {
      "ADJACENT DOCTOR": [{
        id: `${fileId}-shift`,
        source: "MMC",
        title: fileId,
        allDay: false,
        start,
        end,
        rawValue: fileId,
        monthKey: start.slice(0, 7),
      }, ...(fileId === "adjacent-term-1" ? [{
        id: `${fileId}-early-shift`,
        source: "MMC",
        title: `${fileId}-early`,
        allDay: false,
        start: "2026-02-02",
        end: "2026-02-02",
        rawValue: `${fileId}-early`,
        monthKey: "2026-02",
      }] : [])],
    },
  }, sharedUploadDb);
}
assert.equal(sharedUploadDb.files.get("adjacent-term-1")?.active, 1, "adjacent terms should remain active when the earlier roster only ends on the next term boundary");
assert.equal(sharedUploadDb.files.get("adjacent-term-2")?.active, 1, "adjacent next-term roster should remain active");
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  file: { id: "ambiguous-left", name: "Ambiguous_Left.xlsx", sourceType: "mch", active: true, lastModified: 0 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mch" }],
  eventsByDoctor: {
    "SHARED USER": [{ id: "ambiguous-left-shift", source: "MCH", title: "Ambiguous Left", allDay: true, start: "2026-05-08", end: "2026-05-08", rawValue: "Ambiguous Left", monthKey: "2026-05" }],
  },
}, sharedUploadDb);
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  file: { id: "ambiguous-right", name: "Ambiguous_Right.xlsx", sourceType: "mch", active: true, lastModified: 0 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mch" }],
  eventsByDoctor: {
    "SHARED USER": [{ id: "ambiguous-right-shift", source: "MCH", title: "Ambiguous Right", allDay: true, start: "2026-05-08", end: "2026-05-08", rawValue: "Ambiguous Right", monthKey: "2026-05" }],
  },
}, sharedUploadDb);
assert.ok(memoryD1AccountRecord(sharedUploadDb, "rhaydon@gmail.com").adminIssues.some((issue) => issue.message.includes("Could not determine the latest MCH roster")), "ambiguous supersession should create a creator admin issue");
const conferenceLeaveStore = new MemoryStore();
const conferenceLeaveDb = new MemoryD1();
const conferenceDoctor = { key: "CONFERENCE DOCTOR", displayName: "Conference Doctor", sourceType: "mmc" };
await seedRepository(conferenceLeaveStore, [
  repositoryFile("conference-mmc", { sourceType: "mmc", doctors: [conferenceDoctor] }),
  repositoryFile("conference-mch", { sourceType: "mch", doctors: [{ ...conferenceDoctor, sourceType: "mch" }] }),
]);
await postState(conferenceLeaveStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, conferenceLeaveDb);
await postState(conferenceLeaveStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "conference-mmc", name: "conference-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ ...conferenceDoctor, sourceType: "mmc" }],
  eventsByDoctor: {
    [conferenceDoctor.key]: [{
      id: "conference-mmc-leave",
      source: "MMC",
      title: "Conference Leave",
      allDay: true,
      start: "2026-05-04",
      end: "2026-05-11",
      rawValue: "Conference Leave",
      monthKey: "2026-05",
    }, {
      id: "conference-mmc-separate-leave",
      source: "MMC",
      title: "Conference Leave",
      allDay: true,
      start: "2026-05-18",
      end: "2026-05-25",
      rawValue: "Conference Leave",
      monthKey: "2026-05",
    }],
  },
}, conferenceLeaveDb);
await postState(conferenceLeaveStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "conference-mch", name: "conference-mch.xlsx", sourceType: "mch", active: true },
  doctors: [{ ...conferenceDoctor, sourceType: "mch" }],
  eventsByDoctor: {
    [conferenceDoctor.key]: [{
      id: "conference-mch-leave",
      source: "MCH",
      title: "CME Leave",
      allDay: true,
      start: "2026-05-04",
      end: "2026-05-11",
      rawValue: "CME/L",
      monthKey: "2026-05",
    }],
  },
}, conferenceLeaveDb);
const conferenceProfile = await postState(conferenceLeaveStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${conferenceDoctor.key}::mmc::mch`,
  doctorKey: conferenceDoctor.key,
  displayName: conferenceDoctor.displayName,
  sourceTypes: ["mmc", "mch"],
}, conferenceLeaveDb);
const conferenceLeaves = conferenceProfile.snapshot.preview.events.filter((event) => event.title === "Conference Leave");
assert.equal(conferenceLeaves.length, 2, "overlapping conference/CME leave should merge, separate weeks should remain separate");
const overlappingConferenceLeave = conferenceLeaves.find((event) => event.start === "2026-05-04");
assert.ok(overlappingConferenceLeave);
assert.equal(overlappingConferenceLeave.end, "2026-05-11");
assert.deepEqual(overlappingConferenceLeave.sources.sort(), ["MCH", "MMC"]);
assert.equal(overlappingConferenceLeave.rawValue, "Conference Leave / CME/L");
const d1FeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1FeedResponse.ok, true);
const d1FeedText = await d1FeedResponse.text();
assert.ok(d1FeedText.includes("BEGIN:VCALENDAR"));
assert.ok(d1FeedText.includes("BEGIN:VEVENT"));
assert.ok(d1FeedText.includes("D1 Edited Shift"), "D1 subscription feed should apply D1 session overrides");
assert.ok(d1FeedText.includes("D1 Custom Event"), "D1 subscription feed should include D1 custom events");
await d1StateStore.delete(`subscription:token:${d1DirectLogin.subscription.token}`);
await d1StateStore.delete("snapshot:account:d1-user@example.com");
const d1OnlyFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}&view=range`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1OnlyFeedResponse.ok, true);
const d1OnlyFeedText = await d1OnlyFeedResponse.text();
assert.ok(d1OnlyFeedText.includes("D1 Custom Event"), "D1 feed should resolve account and session without KV token index or snapshot");
const d1NoKvLogin = await postState(null, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvLogin.snapshot, null, "D1-only login should not depend on KV snapshots");
const d1NoKvCalendar = await postState(null, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvCalendar.snapshot?.preview?.derivedFromD1, true, "separate calendar load should work without KV when D1 has account and roster data");
const d1NoKvFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1NoKvFeedResponse.ok, true, "subscription feed should resolve from D1 without KV");

const michaelStateStore = new MemoryStore();
michaelStateStore.d1 = new MemoryD1();
await postState(michaelStateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(michaelStateStore.d1, [
  repositoryFile("michael-mmc", {
    name: "michael-mmc.xlsx",
    sourceType: "mmc",
    doctors: [{ key: "MICHAEL COMAN", displayName: "Michael COMAN", sourceType: "mmc" }],
  }),
  repositoryFile("michael-mch", {
    name: "michael-mch.xlsx",
    sourceType: "mch",
    doctors: [{ key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman", sourceType: "mch" }],
  }),
]);
await seedUser(michaelStateStore, "michael@example.com", "michael-password", "Michael COMAN");
await postState(michaelStateStore, {
  action: "claimRosterName",
  email: "michael@example.com",
  password: "michael-password",
  claim: { sourceType: "mmc", key: "MICHAEL COMAN" },
});
const michaelDirectLogin = await postState(michaelStateStore, {
  action: "login",
  email: "michael@example.com",
  password: "michael-password",
});
const michaelEnrichedLogin = await postState(michaelStateStore, {
  action: "resolveAccountClaims",
  email: "michael@example.com",
  password: "michael-password",
});
assert.deepEqual(michaelEnrichedLogin.state.imports.map((item) => item.repoId).sort(), ["michael-mch", "michael-mmc"]);
assert.equal(michaelEnrichedLogin.claims.some((claim) => claim.sourceType === "mch" && claim.key === "DR MICHAEL COMAN"), true);
assert.equal(michaelEnrichedLogin.suggestedClaims.some((claim) => claim.sourceType === "mch" && claim.key === "DR MICHAEL COMAN"), false);
await postState(michaelStateStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
  claims: [
    { sourceType: "mmc", key: "MICHAEL COMAN" },
    { sourceType: "mch", key: "DR MICHAEL COMAN" },
  ],
});
const michaelAdminLoad = await postState(michaelStateStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
});
assert.deepEqual(michaelAdminLoad.state.imports.map((item) => item.repoId).sort(), [], "admin account loads should not hydrate roster file refs during login");
const michaelPrimaryResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "MICHAEL COMAN",
    displayName: "Michael COMAN",
    sourceTypes: ["mmc", "mch"],
    aliases: [
      { sourceType: "mmc", key: "MICHAEL COMAN", displayName: "Michael COMAN" },
      { sourceType: "mch", key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman" },
    ],
  },
});
assert.equal(michaelPrimaryResolution.mode, "claimed-account");
assert.equal(michaelPrimaryResolution.email, "michael@example.com");
const michaelAliasResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "DR MICHAEL COMAN",
    displayName: "Dr Michael Coman",
    sourceTypes: ["mch"],
  },
});
assert.equal(michaelAliasResolution.mode, "claimed-account");
assert.equal(michaelAliasResolution.email, "michael@example.com");
await postState(michaelStateStore, {
  action: "deleteAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
});
const michaelDeletedResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "MICHAEL COMAN",
    displayName: "Michael COMAN",
    sourceTypes: ["mmc", "mch"],
    aliases: [
      { sourceType: "mmc", key: "MICHAEL COMAN", displayName: "Michael COMAN" },
      { sourceType: "mch", key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman" },
    ],
  },
});
assert.equal(michaelDeletedResolution.mode, "doctor-profile");
assert.equal(michaelDeletedResolution.email, "");
await seedUser(michaelStateStore, "michael@example.com", "michael-password-2", "Michael COMAN");
const michaelRecreatedResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "DR MICHAEL COMAN",
    displayName: "Dr Michael Coman",
    sourceTypes: ["mch"],
  },
});
assert.equal(michaelRecreatedResolution.mode, "claimed-account");
assert.equal(michaelRecreatedResolution.email, "michael@example.com");

const identityStore = new MemoryStore();
identityStore.d1 = new MemoryD1();
await postState(identityStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(identityStore.d1, [
  repositoryFile("identity-ddh", {
    sourceType: "ddh",
    doctors: [
      { key: "AARON BADWAL", displayName: "Aaron BADWAL", sourceType: "ddh" },
      { key: "ANDREA LIM", displayName: "Andrea LIM", sourceType: "ddh" },
      { key: "ABI THANIKASALAM", displayName: "Abi THANIKASALAM", sourceType: "ddh" },
    ],
  }),
  repositoryFile("identity-mch", {
    sourceType: "mch",
    doctors: [{ key: "DR ANDREA LIM", displayName: "Dr Andrea LIM", sourceType: "mch" }],
  }),
]);
const abiAutoClaim = await postState(identityStore, {
  action: "login",
  email: "abi@example.com",
  password: "abi-password",
  mode: "create",
  realName: "Abirama Thanikasalam",
});
assert.deepEqual(abiAutoClaim.claims.map((claim) => `${claim.sourceType}:${claim.key}`), ["ddh:ABI THANIKASALAM"]);
const andreaLogin = await postState(identityStore, {
  action: "login",
  email: "andrea@example.com",
  password: "andrea-password",
  mode: "create",
  realName: "Andrea LIM",
});
const andreaEnrichedLogin = await postState(identityStore, {
  action: "resolveAccountClaims",
  email: "andrea@example.com",
  password: "andrea-password",
});
assert.deepEqual(andreaLogin.claims.map((claim) => `${claim.sourceType}:${claim.key}`).sort(), ["ddh:ANDREA LIM", "mch:DR ANDREA LIM"]);
assert.deepEqual(andreaEnrichedLogin.state.imports.map((item) => item.repoId).sort(), ["identity-ddh", "identity-mch"]);
assert.deepEqual(andreaEnrichedLogin.suggestedClaims, []);
const barryLogin = await postState(identityStore, {
  action: "login",
  email: "barry@example.com",
  password: "barry-password",
  mode: "create",
  realName: "Barry Cunningham",
});
assert.deepEqual(barryLogin.claims, []);
assert.deepEqual(barryLogin.suggestedClaims, []);
assert.deepEqual(barryLogin.state.imports, []);
await postState(identityStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "barry@example.com",
  claims: [{ sourceType: "ddh", key: "AARON BADWAL" }],
});
const aaronAfterBadBarryClaim = await postState(identityStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: { sourceType: "ddh", key: "AARON BADWAL", displayName: "Aaron BADWAL" },
});
assert.equal(aaronAfterBadBarryClaim.mode, "claimed-account");
const usersAfterBadBarryClaim = await postState(identityStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.deepEqual(usersAfterBadBarryClaim.users.find((user) => user.email === "barry@example.com")?.claims.map((claim) => `${claim.sourceType}:${claim.key}`) || [], ["ddh:AARON BADWAL"]);
const barryAfterBadClaim = await postState(identityStore, {
  action: "login",
  email: "barry@example.com",
  password: "barry-password",
});
const barryAfterBadClaimEnriched = await postState(identityStore, {
  action: "resolveAccountClaims",
  email: "barry@example.com",
  password: "barry-password",
});
assert.deepEqual(barryAfterBadClaim.claims.map((claim) => `${claim.sourceType}:${claim.key}`), ["ddh:AARON BADWAL"]);
assert.deepEqual(barryAfterBadClaimEnriched.state.imports.map((item) => item.repoId), ["identity-ddh"]);
await postState(identityStore, {
  action: "claimRosterName",
  email: "barry@example.com",
  password: "barry-password",
  claim: { sourceType: "ddh", key: "AARON BADWAL" },
});
assert.ok(
  memoryD1AccountRecord(identityStore.d1, "barry@example.com").adminIssues.some((issue) => issue.rawValue.includes("Manual roster claim review")),
  "mismatched manual claims should create a Creator review issue",
);
await postState(identityStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "andrea@example.com",
  claims: [
    { sourceType: "ddh", key: "ANDREA LIM" },
    { sourceType: "mch", key: "DR ANDREA LIM" },
  ],
});
const andreaAssigned = await postState(identityStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "andrea@example.com",
});
assert.deepEqual(andreaAssigned.claims.map((claim) => `${claim.sourceType}:${claim.key}`).sort(), ["ddh:ANDREA LIM", "mch:DR ANDREA LIM"]);
assert.deepEqual(andreaAssigned.state.imports.map((item) => item.repoId).sort(), [], "admin account loads should not hydrate roster file refs during login");
await postState(identityStore, {
  action: "removeRosterClaim",
  email: "andrea@example.com",
  password: "andrea-password",
  claim: { sourceType: "ddh", key: "ANDREA LIM" },
});
await postState(identityStore, {
  action: "reportRosterIdentityIssue",
  email: "andrea@example.com",
  password: "andrea-password",
  message: "Wrong roster name.",
});
const andreaProfileAfterReport = identityStore.d1.accountProfiles.get("andrea@example.com");
const andreaRecordAfterReport = {
  claims: [...identityStore.d1.accountClaims.values()].filter((claim) => claim.email === "andrea@example.com").map((claim) => ({
    sourceType: claim.source_type,
    key: claim.doctor_key,
  })),
  adminIssues: JSON.parse(andreaProfileAfterReport.admin_issues_json || "[]"),
};
assert.equal((andreaRecordAfterReport.claims || []).some((claim) => claim.sourceType === "ddh"), false);
assert.ok(andreaRecordAfterReport.adminIssues.length >= 1);

const manyDoctorsStore = new MemoryStore();
manyDoctorsStore.d1 = new MemoryD1();
await postState(manyDoctorsStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(manyDoctorsStore.d1, [
  repositoryFile("many-doctors", {
    doctors: Array.from({ length: 90 }, (_, index) => ({
      key: `DOCTOR ${index}`,
      displayName: `Doctor ${index}`,
      sourceType: "mmc",
    })),
  }),
]);
await seedUser(manyDoctorsStore, "doctor-1@example.com", "doctor-password", "Doctor 1");
await seedUser(manyDoctorsStore, "doctor-2@example.com", "doctor-password", "Doctor 2");
manyDoctorsStore.resetMetrics();
const manyDoctorsLogin = await postState(manyDoctorsStore, {
  action: "login",
  email: "new-doctor@example.com",
  password: "new-password",
  mode: "create",
  realName: "New Doctor",
});
const manyDoctorsEnrichment = await postState(manyDoctorsStore, {
  action: "resolveAccountClaims",
  email: "new-doctor@example.com",
  password: "new-password",
});
assert.equal(manyDoctorsEnrichment.availableDoctors.length, 90);
assert.ok(manyDoctorsStore.accountListCalls <= 2, "available doctor claimed status should avoid repeated account scans");

const profileImports = await postStateRaw(stateStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(profileImports.response.status, 410);

await postState(stateStore, {
  action: "saveDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
  state: { version: 1, imports: [], session: { hadPreview: true } },
  snapshot: {
    preview: {
      count: 1,
      date_range: "2026-07-27 to 2026-07-28",
      events: [{ id: "old-leave", source: "Casey", title: "Annual Leave", allDay: true, start: "2026-07-27", end: "2026-07-28", rawValue: "Annual Leave" }],
      review: [],
      issues: [],
      conflicts: [],
    },
    session: { hadPreview: true },
    doctorOptions: [{ key: "TITUS HACKMAN", displayName: "Titus HACKMAN", sourceTypes: ["mmc"] }],
    detectedSources: { mmc: ["fixture-roster"] },
    fileRefs: [{ repoId: "fixture-roster", id: "fixture-roster", sourceType: "mmc", name: "AdultMMCTerm2.2026.Ver1.pdf" }],
  },
});
const d1ProfileReload = await postState(stateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(d1ProfileReload.profile.profileId, "TITUS HACKMAN::mmc");

await seedUser(stateStore, "patrick@example.com", "patrick-password", "Patrick TAN");
await seedUser(stateStore, "senior@example.com", "senior-password", "Senior Registrar");
const initialUsers = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(initialUsers.users.find((user) => user.email === "patrick@example.com")?.insightsEnabled, false);
const enabledInsights = await postState(stateStore, {
  action: "setUserInsightsEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "patrick@example.com",
  insightsEnabled: true,
});
assert.equal(enabledInsights.user.insightsEnabled, true);
const patrickLogin = await postState(stateStore, {
  action: "login",
  email: "patrick@example.com",
  password: "patrick-password",
});
assert.equal(patrickLogin.insightsEnabled, true);
const seniorLogin = await postState(stateStore, {
  action: "login",
  email: "senior@example.com",
  password: "senior-password",
});
assert.equal(seniorLogin.insightsEnabled, false);
stateStore.d1.accountClaims.set("patrick@example.com|mmc|PATRICK TAN", {
  email: "patrick@example.com",
  source_type: "mmc",
  doctor_key: "PATRICK TAN",
  display_name: "Patrick TAN",
  matched_at: "2026-05-01T00:00:00.000Z",
  updated_at: "2026-05-01T00:00:00.000Z",
});
const writerCodeIssue = {
  id: "writer-code",
  source: "MMC",
  seniority: "Senior Registrar",
  startDay: "2026-05-02",
  date: "2026-05-02",
  rawValue: "WRITER",
  code: "WRITER",
  status: "unknown",
  message: "MMC shift code not recognised.",
  resolutionType: "shift_code",
  fingerprint: "MMC::Senior Registrar::WRITER",
};
await postState(stateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "writer-code-file", name: "Writer Code.xlsx", sourceType: "mmc", size: 1, lastModified: 1 },
  doctors: [{ key: "PATRICK TAN", displayName: "Patrick TAN", sourceType: "mmc" }],
  eventsByDoctor: {
    "PATRICK TAN": [{
      id: "writer-code-event",
      source: "MMC",
      seniority: "Senior Registrar",
      title: "MMC: WRITER",
      allDay: true,
      start: "2026-05-02",
      end: "2026-05-03",
      rawValue: "WRITER",
    }],
  },
  issuesByDoctor: { "PATRICK TAN": [writerCodeIssue] },
  skipStatus: true,
});
assert.ok(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.some((issue) => issue.code === "WRITER"), "D1 ingestion should promote unresolved shift-code diagnostics into Creator Errors");
const patrickWriterCalendar = await postState(stateStore, {
  action: "loadCalendarEvents",
  email: "patrick@example.com",
  password: "patrick-password",
  doctorKey: "PATRICK TAN",
});
assert.ok(patrickWriterCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "WRITER"), "D1-derived user calendar should still expose the warning panel issue");
await postState(stateStore, {
  action: "reportUserError",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "senior@example.com",
  errorId: writerCodeIssue.fingerprint,
  message: writerCodeIssue.message,
  issue: writerCodeIssue,
});
await postState(stateStore, {
  action: "saveLocalParserExtensionRule",
  email: "patrick@example.com",
  password: "patrick-password",
  fingerprint: writerCodeIssue.fingerprint,
  rawValue: "WRITER",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "WRITER",
    kind: "ignore",
    ignore: true,
    includeAsShift: false,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.some((issue) => issue.code === "WRITER"), false, "local ignored shift code should clear that user's warning evidence");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.some((issue) => issue.code === "WRITER"), true, "local ignored shift code should not clear other users");
await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "WRITER",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "WRITER",
    kind: "ignore",
    ignore: true,
    includeAsShift: false,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.some((issue) => issue.code === "WRITER"), false, "creator ignored shift code should clear matching warnings globally");
const n1Issue = {
  source: "MMC",
  seniority: "Senior Registrar",
  date: "2026-05-01",
  rawValue: "N1",
  message: "MMC shift code not recognised.",
  fingerprint: "MMC::Senior Registrar::N1",
};
await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
await postState(stateStore, {
  action: "reportUserError",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "senior@example.com",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 1);
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 1);
await postState(stateStore, {
  action: "saveLocalParserExtensionRule",
  email: "patrick@example.com",
  password: "patrick-password",
  fingerprint: n1Issue.fingerprint,
  rawValue: "N1",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "N1",
    kind: "shift",
    base: "SR IC",
    period: "NIGHT",
    suffix: "",
    allDay: false,
    startTime: "23:00",
    endTime: "09:00",
    includeAsShift: true,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "local parser rule should clear matching user warning evidence");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 1, "local parser rule should not clear other users");
const creatorSuggestionView = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorSuggestionView.issueConfig.parserRuleSuggestions.length, 1, "user shift-code resolutions must be visible to the creator");
await postState(stateStore, {
  action: "decideParserRuleSuggestion",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  suggestionId: creatorSuggestionView.issueConfig.parserRuleSuggestions[0].id,
  decision: "reject",
});
const parserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: n1Issue.fingerprint,
  source: "MMC",
  rawValue: "N1",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "N1",
    kind: "shift",
    base: "SR IC",
    period: "NIGHT",
    suffix: "",
    allDay: false,
    startTime: "23:00",
    endTime: "09:00",
    includeAsShift: true,
  },
});
assert.ok(parserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "N1"));
setParserExtensions(parserSave.parserExtensions);
const srN1Workbook = XLSX.utils.book_new();
const srN1Sheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "2300-0900 N1"],
]);
for (let index = 0; index < 7; index += 1) {
  srN1Sheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srN1Workbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srN1Workbook, srN1Sheet, "Week 1");
const srN1View = buildRosterView([{ id: "sr-n1", workbook: srN1Workbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srN1View.events.some((event) => event.rawValue === "2300-0900 N1" && event.title === "MMC: SR IC Night" && event.start.includes("23:00:00") && event.end.includes("09:00:00")), "Senior Registrar N1 explicit-time rules must render with the saved rule title");
assert.equal(srN1View.issues.some((issue) => issue.rawValue === "2300-0900 N1"), false);
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "global parser rule should keep direct-user warning evidence cleared");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 0, "global parser rule should clear matching switch-user warning evidence");
const staleReport = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "MMC::Senior Registrar::2200-0830 N1",
  message: n1Issue.message,
  issue: {
    ...n1Issue,
    rawValue: "2200-0830 N1",
    fingerprint: "MMC::Senior Registrar::2200-0830 N1",
  },
});
assert.equal(staleReport.ignored, true, "resolved global shift-code warnings must not be requeued from stale user previews");
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "resolved global shift-code warning evidence must remain cleared");
const ssuBatchSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "ASSJ",
  rules: ["HMO", "Intern"].map((seniority) => ({
    source: "MMC",
    seniority,
    code: "ASSJ",
    kind: "shift",
    base: "SSU",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "07:30",
    endTime: "17:30",
    includeAsShift: true,
  })),
});
assert.ok(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "HMO" && rule.code === "ASSJ"), "batch parser save should add HMO ASSJ");
assert.ok(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "Intern" && rule.code === "ASSJ"), "batch parser save should add Intern ASSJ");
assert.equal(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "ASSJ"), false, "batch parser save should not add unselected seniorities");
const pssjBatchSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "PSSJ",
  rules: ["HMO", "Intern"].map((seniority) => ({
    source: "MMC",
    seniority,
    code: "PSSJ",
    kind: "shift",
    base: "SSU",
    period: "PM",
    suffix: "",
    allDay: false,
    startTime: "14:30",
    endTime: "00:00",
    includeAsShift: true,
  })),
});
setParserExtensions(pssjBatchSave.parserExtensions);
const hmoSsuWorkbook = XLSX.utils.book_new();
const hmoSsuSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "HMO"],
  ["", "", "", "Patrick TAN", "", "0800-1730 ASSJ", "1430-0000 PSSJ"],
]);
for (let index = 0; index < 7; index += 1) {
  hmoSsuSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(hmoSsuWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(hmoSsuWorkbook, hmoSsuSheet, "Week 1");
const hmoSsuView = buildRosterView([{ id: "hmo-ssu", workbook: hmoSsuWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(hmoSsuView.events.some((event) => event.rawValue === "0800-1730 ASSJ" && event.title === "MMC: SSU AM" && event.start.includes("08:00:00") && event.end.includes("17:30:00")), "HMO ASSJ should resolve to SSU AM while preserving explicit roster time");
assert.ok(hmoSsuView.events.some((event) => event.rawValue === "1430-0000 PSSJ" && event.title === "MMC: SSU PM" && event.start.includes("14:30:00") && event.end.includes("00:00:00")), "HMO PSSJ should resolve to SSU PM");
assert.equal(hmoSsuView.issues.some((issue) => issue.rawValue.includes("SSJ")), false, "selected seniorities should no longer surface ASSJ/PSSJ as unresolved");
const deletedAssj = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "ASSJ",
  replacementTargets: [
    { source: "MMC", seniority: "HMO", code: "ASSJ" },
    { source: "MMC", seniority: "Intern", code: "ASSJ" },
  ],
  rules: [{
    source: "MMC",
    seniority: "HMO",
    code: "ASSJ",
    kind: "shift",
    base: "SSU",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "07:30",
    endTime: "17:30",
    includeAsShift: true,
  }],
});
assert.ok(deletedAssj.parserExtensions.mmc.some((rule) => rule.seniority === "HMO" && rule.code === "ASSJ"), "replacement save should keep selected matching seniority");
assert.equal(deletedAssj.parserExtensions.mmc.some((rule) => rule.seniority === "Intern" && rule.code === "ASSJ"), false, "replacement save should delete deselected matching seniority");
await postState(stateStore, {
  action: "deleteParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  rule: { source: "MMC", seniority: "HMO", code: "ASSJ" },
});
const hmoAssjReappears = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "MMC::HMO::0800-1730 ASSJ",
  message: "MMC shift code not recognised.",
  issue: {
    source: "MMC",
    seniority: "HMO",
    date: "2026-05-01",
    rawValue: "0800-1730 ASSJ",
    message: "MMC shift code not recognised.",
    fingerprint: "MMC::HMO::0800-1730 ASSJ",
  },
});
assert.equal(hmoAssjReappears.ignored, undefined, "deleted shift-code disambiguations should allow unresolved reports to reappear");
const knownDdhClinicalSupport = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::Clinical Support",
  message: "DDH shift code not recognised; using explicit roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-03-02",
    rawValue: "Clinical Support",
    message: "DDH shift code not recognised; using explicit roster time.",
    fingerprint: "DDH::Unknown::Clinical Support",
  },
});
assert.equal(knownDdhClinicalSupport.ignored, true, "known DDH Clinical Support mappings should not enter unresolved shift-code queues");
const knownDdhSsuSms = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::SSU SMS",
  message: "DDH shift code not recognised; using explicit roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-03-06",
    rawValue: "SSU SMS",
    message: "DDH shift code not recognised; using explicit roster time.",
    fingerprint: "DDH::Unknown::SSU SMS",
  },
});
assert.equal(knownDdhSsuSms.ignored, true, "known DDH SSU SMS mappings should not enter unresolved shift-code queues");
await seedUser(stateStore, "michael@example.com", "michael-password", "Michael Coman");
const michaelProfile = stateStore.d1.accountProfiles.get("michael@example.com");
michaelProfile.admin_issues_json = JSON.stringify([
  {
    id: "MCH::SMS::OCS",
    source: "MCH",
    seniority: "SMS",
    date: "2026-03-25",
    rawValue: "0800-1730 OCS",
    code: "OCS",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: OCS",
    fingerprint: "MCH::SMS::OCS",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MCH::SMS::PHNW",
    source: "MCH",
    seniority: "SMS",
    date: "2026-03-09",
    rawValue: "0800-1730PHNW",
    code: "PHNW",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: PHNW",
    fingerprint: "MCH::SMS::PHNW",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MCH::SMS::AM",
    source: "MCH",
    seniority: "SMS",
    date: "2026-02-10",
    rawValue: "0800-1730",
    code: "AM",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: AM",
    fingerprint: "MCH::SMS::AM",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MMC::CMO::PM",
    source: "MMC",
    seniority: "CMO",
    date: "2026-02-09",
    rawValue: "1430-0000",
    code: "PM",
    timeLabel: "14:30-00:00",
    suggestedTitle: "MMC: PM",
    fingerprint: "MMC::CMO::PM",
    message: "MMC shift code not recognised; using explicit roster time.",
  },
]);
const cleanedUserList = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(cleanedUserList.users.find((user) => user.email === "michael@example.com")?.adminIssues.length, 0, "creator user list should hide stale resolved shift-code issues");
assert.ok(memoryD1AccountRecord(stateStore.d1, "michael@example.com").adminIssues.length > 0, "creator user list should not persist a full stale-issue cleanup during load");
await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "DDH",
  rawValue: "Rover AM",
  rules: ["SMS", "CMO"].map((seniority) => ({
    source: "DDH",
    seniority,
    code: "ROVER AM",
    kind: "shift",
    base: "Rover",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "08:00",
    endTime: "18:00",
    includeAsShift: true,
  })),
});
const staleUnknownRover = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::Rover AM",
  message: "DDH shift label not recognised; using roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-05-14",
    rawValue: "Rover AM",
    message: "DDH shift label not recognised; using roster time.",
    fingerprint: "DDH::Unknown::Rover AM",
  },
});
assert.equal(staleUnknownRover.ignored, true, "stale Unknown-seniority DDH warnings should resolve once a source/code rule exists");
const accParserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: "MMC::Senior Registrar::ACC",
  source: "MMC",
  rawValue: "ACC",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "ACC",
    kind: "shift",
    base: "Clinic",
    period: "AM",
    suffix: "Charge",
    allDay: false,
    startTime: "08:00",
    endTime: "17:30",
    includeAsShift: true,
  },
});
assert.ok(accParserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "ACC"), "Senior Registrar charge/consultant-style codes must persist as explicit rules");
setParserExtensions(accParserSave.parserExtensions);
const srAccWorkbook = XLSX.utils.book_new();
const srAccSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "ACC"],
]);
for (let index = 0; index < 7; index += 1) {
  srAccSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srAccWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srAccWorkbook, srAccSheet, "Week 1");
const srAccView = buildRosterView([{ id: "sr-acc", workbook: srAccWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srAccView.events.some((event) => event.rawValue === "ACC" && event.title === "MMC: Clinic AM Charge"), "Senior Registrar explicit ACC rules must render in user calendars");
assert.equal(srAccView.issues.some((issue) => issue.rawValue === "ACC"), false);
setParserExtensions({});

const deletionStore = new MemoryStore();
deletionStore.d1 = new MemoryD1();
await postState(deletionStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await seedRepository(deletionStore, [
  repositoryFile("keep-roster"),
  repositoryFile("missing-from-save", { name: "missing-from-save.xlsx" }),
  repositoryFile("remove-roster", { name: "remove-roster.xlsx" }),
]);
seedD1Repository(deletionStore.d1, [
  repositoryFile("keep-roster"),
  repositoryFile("missing-from-save", { name: "missing-from-save.xlsx" }),
  repositoryFile("remove-roster", { name: "remove-roster.xlsx" }),
]);

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
});
assert.equal(deletionStore.d1.files.has("missing-from-save"), true, "ordinary creator save must not delete omitted D1 roster files");
let deletionIndex = await deletionStore.get("repository:index", "json");
assert.ok(deletionIndex.files.some((file) => file.id === "missing-from-save"), "ordinary creator save must keep omitted files in the repository index");

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.ok(await deletionStore.get("repository:file:remove-roster", "json"), "D1-only removal must not mutate legacy KV files");
assert.equal(deletionStore.d1.files.has("remove-roster"), false, "creator removal should delete the D1 roster file");
deletionIndex = await deletionStore.get("repository:index", "json");
assert.equal(deletionIndex.files.some((file) => file.id === "remove-roster"), true, "D1-only removal must not mutate legacy KV index");
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));
await postState(deletionStore, {
  action: "replaceActiveRosterFiles",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  keepFileIds: ["keep-roster"],
}, deletionStore.d1);
assert.equal(deletionStore.d1.files.has("keep-roster"), true, "recovery should retain the requested current roster");
assert.equal(deletionStore.d1.files.has("missing-from-save"), false, "recovery should remove active rosters outside the current upload set");

const emptyReplacement = await postStateRaw(deletionStore, {
  action: "replaceActiveRosterFiles",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  keepFileIds: [],
}, deletionStore.d1);
assert.equal(emptyReplacement.response.status, 400, "rebuild-all must not accept an empty retained roster set");

const failedReplacementStore = new MemoryStore();
failedReplacementStore.d1 = new MemoryD1();
await postState(failedReplacementStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(failedReplacementStore.d1, [
  repositoryFile("last-known-good", { name: "last-known-good.xlsx" }),
]);
await postState(failedReplacementStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "new-not-persisted", id: "new-not-persisted", sourceType: "mmc", name: "new-not-persisted.xlsx" }],
    session: {},
  },
});
assert.equal(
  failedReplacementStore.d1.files.has("last-known-good"),
  true,
  "creator save should keep the last-known-good D1 roster set when the replacement files are not yet persisted",
);

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));

await seedUser(deletionStore, "claimed-doctor@example.com", "claimed-password");
await postState(deletionStore, {
  action: "claimRosterName",
  email: "claimed-doctor@example.com",
  password: "claimed-password",
  claim: { sourceType: "mmc", key: "TITUS HACKMAN" },
});
const observerBeforeDelete = await postState(deletionStore, {
  action: "login",
  email: "observer@example.com",
  password: "observer-password",
  mode: "create",
  realName: "Observer Person",
});
const observerBeforeDeleteEnriched = await postState(deletionStore, {
  action: "resolveAccountClaims",
  email: "observer@example.com",
  password: "observer-password",
});
assert.equal(
  observerBeforeDeleteEnriched.availableDoctors.find((doctor) => doctor.key === "TITUS HACKMAN")?.claimedBy,
  "claimed-doctor@example.com",
  "repository doctor should be marked claimed before deleting the linked account",
);
await postState(deletionStore, {
  action: "deleteAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "claimed-doctor@example.com",
});
assert.equal(await deletionStore.get("account:claimed-doctor@example.com", "json"), null, "deleteAccount must remove the claimed account record");
const observerAfterDelete = await postState(deletionStore, {
  action: "login",
  email: "observer@example.com",
  password: "observer-password",
});
assert.equal(
  observerAfterDelete.availableDoctors.find((doctor) => doctor.key === "TITUS HACKMAN")?.claimedBy || "",
  "",
  "repository doctor should become unclaimed after deleting the linked account",
);

await seedUser(deletionStore, "user@example.com", "user-password");
await postState(deletionStore, {
  action: "save",
  email: "user@example.com",
  password: "user-password",
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["keep-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"), "standard users must not delete repository files");

const beforeLoadDeleteCount = deletionStore.deletedKeys.length;
const deletionCreatorImports = await postStateRaw(deletionStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(deletionStore.deletedKeys.length, beforeLoadDeleteCount, "loading imports must not delete repository records");
assert.equal(deletionCreatorImports.response.status, 410);

const deletionProfileImports = await postStateRaw(deletionStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(deletionProfileImports.response.status, 410);

const deletionUserImports = await postStateRaw(deletionStore, {
  action: "loadImports",
  email: "user@example.com",
  password: "user-password",
});
assert.equal(deletionUserImports.response.status, 410);

console.log("Fixture smoke test passed.");
