import {
  applyEventOverrides,
  buildRosterView,
  customEventsToEvents,
  defaultSettings as rosterDefaultSettings,
  doctorOptions as rosterDoctorOptions,
  exportIcs,
  parseUploadForm,
  parserRuleDefaults,
  parserRuleSeniorities,
  previewSummary,
  serializeConflict,
  serializeEvent,
  serializeReviewItem,
  setParserExtensions,
  sourceNames,
} from "./roster.js";
import * as XLSX from "xlsx";

const form = document.querySelector("#roster-form");
const appShell = document.querySelector("#appShell");
const entrancePage = document.querySelector("#entrancePage");
const entranceStatus = document.querySelector("#entranceStatus");
const loginTabButton = document.querySelector("#loginTabButton");
const createTabButton = document.querySelector("#createTabButton");
const entrancePanels = [...document.querySelectorAll("[data-entrance-panel]")];
const fileInput = document.querySelector("#rosterFiles");
const addRosterFilesButton = document.querySelector("#addRosterFilesButton");
const rosterDropOverlay = document.querySelector("#rosterDropOverlay");
const filesButton = document.querySelector("#filesButton");
const accountsButton = document.querySelector("#accountsButton");
const filesModal = document.querySelector("#filesModal");
const filesList = document.querySelector("#filesList");
const filesCloseButton = document.querySelector("#filesCloseButton");
const accountsModal = document.querySelector("#accountsModal");
const accountsCloseButton = document.querySelector("#accountsCloseButton");
const accountsBody = document.querySelector("#accountsBody");
const accountsModalTitle = document.querySelector("#accountsModalTitle");
const accountsModalSubtitle = document.querySelector("#accountsModalSubtitle");
const parserRuleModal = document.querySelector("#parserRuleModal");
const parserRuleModalTitle = document.querySelector("#parserRuleModalTitle");
const parserRuleCloseButton = document.querySelector("#parserRuleCloseButton");
const parserRuleForm = document.querySelector("#parserRuleForm");
const parserRuleIssueId = document.querySelector("#parserRuleIssueId");
const parserRuleSource = document.querySelector("#parserRuleSource");
const parserRuleRawValue = document.querySelector("#parserRuleRawValue");
const parserRuleOriginalCode = document.querySelector("#parserRuleOriginalCode");
const parserRuleOriginalSeniority = document.querySelector("#parserRuleOriginalSeniority");
const parserRuleCode = document.querySelector("#parserRuleCode");
const parserRuleSeniority = document.querySelector("#parserRuleSeniority");
const parserRuleBase = document.querySelector("#parserRuleBase");
const parserRulePeriod = document.querySelector("#parserRulePeriod");
const parserRuleSuffix = document.querySelector("#parserRuleSuffix");
const parserRuleAllDay = document.querySelector("#parserRuleAllDay");
const parserRuleIncludeAsShift = document.querySelector("#parserRuleIncludeAsShift");
const parserRuleIgnore = document.querySelector("#parserRuleIgnore");
const parserRuleTimeFields = document.querySelector("#parserRuleTimeFields");
const parserRuleStartTime = document.querySelector("#parserRuleStartTime");
const parserRuleEndTime = document.querySelector("#parserRuleEndTime");
const parserRuleLocation = document.querySelector("#parserRuleLocation");
const parserRulePreview = document.querySelector("#parserRulePreview");
const insightsModal = document.querySelector("#insightsModal");
const insightsCloseButton = document.querySelector("#insightsCloseButton");
const insightsModalTitle = document.querySelector("#insightsModalTitle");
const insightsModalSubtitle = document.querySelector("#insightsModalSubtitle");
const insightsModalBody = document.querySelector("#insightsModalBody");
const loginBar = document.querySelector("#loginBar");
const loginIdentity = document.querySelector("#loginIdentity");
const logoutButton = document.querySelector("#logoutButton");
const backToCreatorButton = document.querySelector("#backToCreatorButton");
const mobileAccountAccessButton = document.querySelector("#mobileAccountAccessButton");
const loadingScreen = document.querySelector("#loadingScreen");
const loginForm = document.querySelector("#loginForm");
const loginEmail = document.querySelector("#loginEmail");
const loginPassword = document.querySelector("#loginPassword");
const stayLoggedIn = document.querySelector("#stayLoggedIn");
const createAccountForm = document.querySelector("#createAccountForm");
const createRealName = document.querySelector("#createRealName");
const createEmail = document.querySelector("#createEmail");
const createPassword = document.querySelector("#createPassword");
const currentDayPreview = document.querySelector("#currentDayPreview");
const exportButton = document.querySelector("#exportButton");
const mobileTodayButton = document.querySelector("#mobileTodayButton");
const mobileExportButton = document.querySelector("#mobileExportButton");
const mobileSettingsButton = document.querySelector("#mobileSettingsButton");
const mobileAccountButton = document.querySelector("#mobileAccountButton");
const mobileAccountButtonLabel = document.querySelector("#mobileAccountButtonLabel");
const exportModal = document.querySelector("#exportModal");
const exportCloseButton = document.querySelector("#exportCloseButton");
const exportModalBody = document.querySelector("#exportModalBody");
const doctorSection = document.querySelector("#doctorSection");
const doctorSelect = document.querySelector("#doctorSelect");
const doctorName = document.querySelector("#doctorName");
const controlBar = document.querySelector("#controlBar");
const claimSection = document.querySelector("#claimSection");
const claimDoctorSelect = document.querySelector("#claimDoctorSelect");
const claimDoctorButton = document.querySelector("#claimDoctorButton");
const settingsToggle = document.querySelector("#settingsToggle");
const settingsPanel = document.querySelector("#settingsPanel");
const settingsCloseButton = document.querySelector("#settingsCloseButton");
const mobileSettingsControls = document.querySelector("#mobileSettingsControls");
const mobileDoctorSelect = document.querySelector("#mobileDoctorSelect");
const mobileDateFrom = document.querySelector("#mobileDateFrom");
const mobileDateTo = document.querySelector("#mobileDateTo");
const mobileHospitalFilter = document.querySelector("#mobileHospitalFilter");
const mobileLogoutButton = document.querySelector("#mobileLogoutButton");
const previewSection = document.querySelector("#previewSection");
const preview = document.querySelector("#preview");
const issuesPanel = document.querySelector("#issuesPanel");
const issuesList = document.querySelector("#issuesList");
const conflictsPanel = document.querySelector("#conflictsPanel");
const conflictsList = document.querySelector("#conflictsList");
const status = document.querySelector("#status");
const mobileActionBar = document.querySelector("#mobileActionBar");
const reviewModal = document.querySelector("#reviewModal");
const reviewModalBody = document.querySelector("#reviewModalBody");
const reviewCloseButton = document.querySelector("#reviewCloseButton");
const customEventModal = document.querySelector("#customEventModal");
const customEventForm = document.querySelector("#customEventForm");
const customEventCloseButton = document.querySelector("#customEventCloseButton");
const customEventId = document.querySelector("#customEventId");
const customEventTitle = document.querySelector("#customEventTitle");
const customEventStartDate = document.querySelector("#customEventStartDate");
const customEventEndDate = document.querySelector("#customEventEndDate");
const customEventAllDay = document.querySelector("#customEventAllDay");
const customEventTimeFields = document.querySelector("#customEventTimeFields");
const customEventStartTime = document.querySelector("#customEventStartTime");
const customEventEndTime = document.querySelector("#customEventEndTime");
const customEventLocationMode = document.querySelector("#customEventLocationMode");
const customEventCustomLocationField = document.querySelector("#customEventCustomLocationField");
const customEventCustomLocation = document.querySelector("#customEventCustomLocation");
const customEventDeleteButton = document.querySelector("#customEventDeleteButton");
const customEventWhoPanel = document.querySelector("#customEventWhoPanel");
const contextMenu = document.querySelector("#contextMenu");
const switchOverlay = document.querySelector("#switchOverlay");
const switchOverlayTitle = document.querySelector("#switchOverlayTitle");
const switchOverlayMessage = document.querySelector("#switchOverlayMessage");

const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const OWNER_EMAIL = "rhaydon@gmail.com";
const OWNER_DOCTOR_KEY = "RICHARD HAYDON";
const DEFAULT_MMC_LOCATION = "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
const DEFAULT_DDH_LOCATION = "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";
const DEFAULT_CASEY_LOCATION = "Casey Hospital, 62-70 Kangan Drive, Berwick VIC 3806, Australia";
const DEFAULT_MCH_LOCATION = "Monash Children's Hospital, 246 Clayton Road, Clayton VIC 3168, Australia";
const SHIFT_COLOUR_DEFAULTS = {
  day: "#0b8f6a",
  evening: "#c96d14",
  night: "#6152d9",
  cs: "#0f8297",
  leave: "#2d79d6",
  custom: "#c48a12",
  phnw: "#5d6c73",
};
const ACCOUNT_STATE_KEY = "roster-account-state";
const SESSION_STATE_KEY = "roster-session-state-v1";
const ACCOUNT_WORKSPACES_KEY = "roster-account-workspaces-v1";
const CALENDAR_SNAPSHOT_CACHE_KEY = "roster-calendar-snapshot-cache-v1";
const CURRENT_EMAIL_KEY = "roster-current-email";
const CURRENT_PASSWORD_KEY = "roster-current-password";
const PERSISTENT_PASSWORD_KEY = "roster-persistent-password";
const SETTINGS_FIELDS = [
  "showSourcePrefix",
  "showAmPm",
  "showTimes",
  "showRawValues",
  "showNormalizedTitles",
  "currentDayBorderColor",
  "currentDayBorderOpacity",
  "currentDayBackgroundColor",
  "currentDayBackgroundOpacity",
  "currentDayFillStyle",
  "currentDayGradientDirection",
  "currentDayBorderWidth",
  "weekendShadeEnabled",
  "weekendShadeColor",
  "weekendShadeOpacity",
  "shiftColorDay",
  "shiftColorEvening",
  "shiftColorNight",
  "shiftColorCs",
  "shiftColorLeave",
  "shiftColorCustom",
  "shiftColorPhnw",
  "includeLocations",
  "includeAnnualLeave",
  "includeConferenceLeave",
  "includePublicHoliday",
  "includeSickLeave",
  "defaultLocationMmc",
  "defaultLocationDdh",
  "defaultLocationCasey",
  "defaultLocationMch",
  "hospitalFilter",
  "dateFrom",
  "dateTo",
];
const PREVIEW_STYLE_FIELDS = [
  "currentDayBorderColor",
  "currentDayBorderOpacity",
  "currentDayBorderWidth",
  "currentDayBackgroundColor",
  "currentDayBackgroundOpacity",
  "currentDayFillStyle",
  "currentDayGradientDirection",
  "weekendShadeEnabled",
  "weekendShadeColor",
  "weekendShadeOpacity",
];
const PREVIEW_DISPLAY_FIELDS = ["showTimes", "showRawValues", "showNormalizedTitles"];
const STATUS_MESSAGE_LIMIT = 5;
const STATUS_MESSAGE_LIFETIME_MS = 5000;
const STATUS_MESSAGE_FADE_MS = 240;
const STATUS_SUPERSEDED_MESSAGES = new Map([
  ["Calendar loaded.", ["Loading calendar...", "Refreshing calendar..."]],
  ["Calendar refreshed.", ["Refreshing calendar..."]],
  ["Calendar file ready.", ["Building calendar file..."]],
  ["Subscription URL copied.", ["Saving subscription feed..."]],
]);

let doctorOptions = [];
let detectedSources = {};
let selectedFiles = [];
let parsedRosterSources = null;
let doctorRoleIndex = null;
let parsedImportDoctors = new Map();
let settings = defaultSettings();
let previewStyleDraft = null;
let previewDisplayDraft = null;
let overrides = {};
let latestPreview = null;
let reviewIndex = new Map();
let customEvents = [];
let currentPreviewEvents = new Map();
let availablePreviewHospitals = [];
let dragEventId = null;
let copiedEvent = null;
let previewGesture = null;
let pendingPreviewGesture = null;
let suppressPreviewClickUntil = 0;
let mobileScrollLockY = 0;
let mobileScrollLocked = false;
let openReviewId = "";
let conflictSelections = {};
let accountState = loadAccountState();
let restoredSessionState = null;
let statusMessageId = 0;
let currentUserEmail = loadCurrentUserEmail();
let currentUserPassword = sessionStorage.getItem(CURRENT_PASSWORD_KEY) || localStorage.getItem(PERSISTENT_PASSWORD_KEY) || "";
let currentUserRole = currentUserEmail === OWNER_EMAIL ? "creator" : "user";
let authUserEmail = currentUserEmail;
let authUserPassword = currentUserPassword;
let adminViewingEmail = "";
let activeDoctorProfile = null;
let activeCalendarContext = initialCalendarContext();
let cloudAvailable = false;
let cloudSaveTimer = 0;
let pendingCloudSaveSnapshot = null;
let cloudStateSaveQueue = Promise.resolve();
let serverUsers = [];
let currentRosterClaims = [];
let currentSuggestedClaims = [];
let latestNameMatches = [];
let availableRosterDoctors = [];
let currentSubscription = null;
let currentInsightsEnabled = currentUserRole === "creator";
let creatorCalendarSourceFileRefs = [];
let insightsState = null;
let doctorAnalysisCacheKey = "";
let doctorAnalysisCache = new Map();
let insightDoctorOptionsCache = [];
let insightDoctorRoleCache = new Map();
let undoHistory = [];
let redoHistory = [];
let applyingHistory = false;
let lastHistorySignature = "";
let pendingExportMode = "full";
let pendingExportRange = defaultExportRangeState();
let pendingExportHospitals = [];
let currentAdminTab = "system";
let adminUserSeniorityFilter = "";
let calendarStoreStatus = null;
let calendarStoreStatusError = "";
let rosterSyncStates = new Map();
let rosterSyncRefreshTimer = 0;
let lastRosterPersistence = null;
let adminConsoleOpen = false;
let adminConsoleLoading = false;
let adminConsoleMessages = [];
let reportedIssueFingerprints = new Set();
let currentSnapshot = null;
let currentSnapshotStale = false;
let currentSnapshotBuiltAt = "";
let currentCalendarRevision = "";
let lastCacheDiagnostics = null;
let snapshotRefreshPromise = null;
let pendingPreviewSnapToToday = false;
let insightWarmupTimer = 0;
let insightWarmupPromise = null;
let visibleInsightWarmCache = new Map();
let visibleInsightWarmKey = "";
let parserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let globalParserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let localParserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let parserRuleSuggestions = [];
let parserRuleSaveContext = { mode: "global", suggestionId: "", targetEmail: "" };
let dismissedIssueFingerprints = new Set();
let ignoredIssueFingerprints = new Set();
let lastLoginTimings = null;
let lastAccountSwitchTimings = null;

const settingsInputs = Object.fromEntries(
  SETTINGS_FIELDS.map((id) => [id, document.querySelector(`#${id}`)]),
);

const overlayObserver = new MutationObserver(syncOverlayState);
[
  settingsPanel,
  exportModal,
  filesModal,
  accountsModal,
  insightsModal,
  parserRuleModal,
  reviewModal,
  customEventModal,
].filter(Boolean).forEach((surface) => {
  overlayObserver.observe(surface, { attributes: true, attributeFilter: ["class", "aria-hidden"] });
});

forceConsoleSkin();
applyShiftColours(settings);
applyCurrentDayHighlight(settings);
applyWeekendShade(settings);
syncOverlayState();
setEntranceTab("login");

["gesturestart", "gesturechange", "gestureend"].forEach((eventName) => {
  document.addEventListener(eventName, (event) => {
    if (isMobileLayout()) event.preventDefault();
  }, { passive: false });
});

addRosterFilesButton.addEventListener("click", (event) => {
  event.preventDefault();
  fileInput.click();
});

fileInput.addEventListener("change", async () => {
  const accepted = validateIncomingFiles([...fileInput.files]);
  if (!accepted.length) {
    fileInput.value = "";
    return;
  }
  if (!await validateFreshRosterUploads(accepted)) {
    fileInput.value = "";
    return;
  }
  await mergeFiles(accepted);
  fileInput.value = "";
  await analyzeFiles();
});

let rosterDragDepth = 0;

for (const eventName of ["dragenter", "dragover"]) {
  window.addEventListener(eventName, (event) => {
    if (!hasFileDrag(event.dataTransfer)) return;
    event.preventDefault();
    if (eventName === "dragenter") rosterDragDepth += 1;
    syncRosterDragState(event.dataTransfer);
  });
}

window.addEventListener("dragleave", (event) => {
  if (!hasFileDrag(event.dataTransfer)) return;
  event.preventDefault();
  rosterDragDepth = Math.max(0, rosterDragDepth - 1);
  if (rosterDragDepth === 0) clearRosterDragState();
});

window.addEventListener("dragend", clearRosterDragState);

window.addEventListener("drop", async (event) => {
  if (!hasFileDrag(event.dataTransfer)) return;
  event.preventDefault();
  clearRosterDragState();
  const accepted = validateIncomingFiles([...event.dataTransfer.files]);
  if (!accepted.length) return;
  if (!await validateFreshRosterUploads(accepted)) return;
  await mergeFiles(accepted);
  await analyzeFiles();
});

filesButton?.addEventListener("click", openFilesModal);
filesCloseButton?.addEventListener("click", closeFilesModal);
filesModal?.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-files]")) closeFilesModal();
});
exportButton.addEventListener("click", openExportModal);
mobileExportButton.addEventListener("click", openExportModal);
mobileTodayButton?.addEventListener("click", () => {
  snapPreviewToCurrentMonth();
});
exportCloseButton.addEventListener("click", closeExportModal);
exportModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-export]")) closeExportModal();
});
exportModalBody.addEventListener("click", async (event) => {
  const modeButton = event.target.closest("[data-export-mode]");
  if (modeButton) {
    setPendingExportMode(modeButton.dataset.exportMode || "full");
    renderExportModal();
    return;
  }
  const futureToggle = event.target.closest("[data-export-all-future]");
  if (futureToggle) {
    pendingExportRange.allFuture = futureToggle.checked;
    if (pendingExportRange.allFuture) pendingExportRange.endDate = "";
    renderExportModal();
    return;
  }
  const hospitalButton = event.target.closest("[data-export-hospital]");
  if (hospitalButton) {
    const hospital = String(hospitalButton.dataset.exportHospital || "").toUpperCase();
    pendingExportHospitals = pendingExportHospitals.includes(hospital)
      ? pendingExportHospitals.filter((item) => item !== hospital)
      : [...pendingExportHospitals, hospital];
    renderExportModal();
    return;
  }
  const actionButton = event.target.closest("[data-export-action]");
  if (!actionButton) return;
  await handleExportAction(actionButton.dataset.exportAction || "");
});
filesList.addEventListener("click", async (event) => {
  const addButton = event.target.closest("[data-open-file-picker]");
  if (addButton) {
    fileInput.click();
    return;
  }
  const removeButton = event.target.closest("[data-remove-import]");
  if (!removeButton) return;
  if (!canRemoveImports()) return;
  await removeStoredImport(removeButton.dataset.removeImport);
});
accountsButton.addEventListener("click", async () => {
  await openAccountsSurface({ defaultAdminTab: "system" });
});
mobileAccountButton?.addEventListener("click", async () => {
  await openAccountsSurface({ defaultAdminTab: "system" });
});
mobileAccountAccessButton?.addEventListener("click", async () => {
  await openAccountsSurface({ defaultAdminTab: "system" });
});
accountsCloseButton.addEventListener("click", closeAccountsModal);
accountsModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-accounts]")) closeAccountsModal();
});
insightsCloseButton.addEventListener("click", closeInsightsModal);
insightsModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-insights]")) closeInsightsModal();
});
accountsBody.addEventListener("submit", (event) => {
  event.preventDefault();
  const createForm = event.target.closest("[data-create-account-form]");
  if (createForm) {
    createAccountFromOwner(createForm);
    return;
  }
  const formElement = event.target.closest("[data-account-form]");
  if (!formElement) return;
  const email = formElement.querySelector("[data-account-email]")?.value.trim() || "";
  const realName = formElement.querySelector("[data-account-real-name]")?.value.trim() || "";
  const password = formElement.querySelector("[data-account-password]")?.value || "";
  if (!email) return;
  void updateAccountDetails(email, { password, realName });
});
accountsBody.addEventListener("change", (event) => {
  const accountLocationInput = event.target.closest("[data-account-location-key]");
  if (accountLocationInput) {
    updateDefaultLocationSetting(accountLocationInput.dataset.accountLocationKey || "", accountLocationInput.value);
    return;
  }
  const seniorityFilter = event.target.closest("[data-admin-user-seniority-filter]");
  if (seniorityFilter) {
    adminUserSeniorityFilter = String(seniorityFilter.value || "");
    renderAccountsModal();
    return;
  }
  const insightsToggle = event.target.closest("[data-toggle-user-insights]");
  if (!insightsToggle) return;
  void setUserInsightsEnabled(insightsToggle.dataset.toggleUserInsights || "", insightsToggle.checked);
});
accountsBody.addEventListener("click", (event) => {
  const adminTab = event.target.closest("[data-admin-tab]");
  if (adminTab) {
    currentAdminTab = adminTab.dataset.adminTab || "errors";
    renderAccountsModal();
    return;
  }
  const clearAdminErrorsButton = event.target.closest("[data-clear-admin-errors]");
  if (clearAdminErrorsButton) {
    clearAdminErrors(clearAdminErrorsButton.dataset.clearAdminErrors || "", clearAdminErrorsButton.dataset.errorId || "");
    return;
  }
  const ignoreAdminErrorButton = event.target.closest("[data-ignore-admin-error]");
  if (ignoreAdminErrorButton) {
    ignoreAdminErrorForever(ignoreAdminErrorButton.dataset.ignoreAdminError || "", ignoreAdminErrorButton.dataset.errorId || "");
    return;
  }
  const refreshCalendarStoreButton = event.target.closest("[data-refresh-calendar-store]");
  if (refreshCalendarStoreButton) {
    refreshCalendarStoreStatus();
    return;
  }
  const replaceActiveRostersButton = event.target.closest("[data-replace-active-rosters]");
  if (replaceActiveRostersButton) {
    void replaceActiveRostersWithCurrentUploads();
    return;
  }
  const viewConsoleButton = event.target.closest("[data-view-console]");
  if (viewConsoleButton) {
    toggleAdminConsole();
    return;
  }
  const addShiftCodeButton = event.target.closest("[data-add-shift-code]");
  if (addShiftCodeButton) {
    openParserRuleModal(
      addShiftCodeButton.dataset.addShiftCode || "",
      addShiftCodeButton.dataset.errorId || "",
      splitShiftCodeSeniorities(addShiftCodeButton.dataset.shiftCodeSeniorities || ""),
    );
    return;
  }
  const ignoreShiftCodeRuleButton = event.target.closest("[data-ignore-shift-code]");
  if (ignoreShiftCodeRuleButton) {
    openParserRuleModal(
      ignoreShiftCodeRuleButton.dataset.ignoreShiftCode || "",
      ignoreShiftCodeRuleButton.dataset.errorId || "",
      splitShiftCodeSeniorities(ignoreShiftCodeRuleButton.dataset.shiftCodeSeniorities || ""),
      { ignore: true },
    );
    return;
  }
  if (event.target.closest("[data-add-manual-shift-code]")) {
    openManualParserRuleModal();
    return;
  }
  const suggestionButton = event.target.closest("[data-parser-suggestion-action]");
  if (suggestionButton) {
    handleParserSuggestionAction(
      suggestionButton.dataset.parserSuggestionAction || "",
      suggestionButton.dataset.suggestionId || "",
    );
    return;
  }
  const removeImportButton = event.target.closest("[data-remove-import]");
  if (removeImportButton) {
    if (!canRemoveImports()) return;
    void removeStoredImport(removeImportButton.dataset.removeImport);
    return;
  }
  const reparseImportButton = event.target.closest("[data-reparse-import]");
  if (reparseImportButton) {
    if (!canRemoveImports()) return;
    void reparseRosterFile(reparseImportButton.dataset.reparseImport);
    return;
  }
  const openFilePickerButton = event.target.closest("[data-open-file-picker]");
  if (openFilePickerButton) {
    fileInput.click();
    return;
  }
  const editShiftCodeButton = event.target.closest("[data-edit-parser-rule]");
  if (editShiftCodeButton) {
    openParserRuleModalFromRule(
      editShiftCodeButton.dataset.editParserSource || "",
      editShiftCodeButton.dataset.editParserSeniority || "",
      editShiftCodeButton.dataset.editParserRule || "",
    );
    return;
  }
  const deleteShiftCodeButton = event.target.closest("[data-delete-parser-rule]");
  if (deleteShiftCodeButton) {
    deleteParserRule(
      deleteShiftCodeButton.dataset.deleteParserSource || "",
      deleteShiftCodeButton.dataset.deleteParserSeniority || "",
      deleteShiftCodeButton.dataset.deleteParserRule || "",
    );
    return;
  }
  const deleteButton = event.target.closest("[data-delete-account]");
  if (deleteButton) {
    deleteAccount(deleteButton.dataset.deleteAccount);
    return;
  }
  const confirmClaimButton = event.target.closest("[data-confirm-suggested-claim]");
  if (confirmClaimButton) {
    confirmSuggestedClaim(Number(confirmClaimButton.dataset.confirmSuggestedClaim));
    return;
  }
  const rejectClaimButton = event.target.closest("[data-reject-suggested-claim]");
  if (rejectClaimButton) {
    rejectSuggestedClaim(Number(rejectClaimButton.dataset.rejectSuggestedClaim));
    return;
  }
  const removeClaimButton = event.target.closest("[data-remove-roster-claim]");
  if (removeClaimButton) {
    removeRosterClaim(removeClaimButton.dataset.removeRosterClaim || "", removeClaimButton.dataset.claimEmail || "");
    return;
  }
  const reportIdentityButton = event.target.closest("[data-report-roster-identity]");
  if (reportIdentityButton) {
    reportRosterIdentityIssue();
    return;
  }
  const adminAddClaimButton = event.target.closest("[data-admin-add-claim]");
  if (adminAddClaimButton) {
    addAdminRosterClaim(adminAddClaimButton.dataset.adminAddClaim || "");
    return;
  }
  const editClaimsButton = event.target.closest("[data-edit-roster-claims]");
  if (editClaimsButton) {
    toggleAdminRosterClaimControls(editClaimsButton.dataset.editRosterClaims || "");
    return;
  }
  const enterButton = event.target.closest("[data-enter-account]");
  if (enterButton) {
    enterUserAccount(enterButton.dataset.enterAccount);
    return;
  }
  const addButton = event.target.closest("[data-add-account]");
  if (addButton) {
    addLocalAccount();
    return;
  }
  const removeButton = event.target.closest("[data-remove-account]");
  if (removeButton) {
    removeLocalAccount(removeButton.dataset.removeAccount);
  }
});
loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setEntranceTab("login");
  const email = normalizeEmail(loginEmail.value);
  const password = loginPassword.value;
  if (!email || !password) return;
  await loginWithEmail(email, password, { mode: "login", stayLoggedIn: Boolean(stayLoggedIn?.checked) });
});
createAccountForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setEntranceTab("create");
  const realName = createRealName.value.trim();
  const email = normalizeEmail(createEmail.value);
  const password = createPassword.value;
  if (!realName || !email || !password) {
    setEntranceStatus("Real name, email address, and password are required.", true);
    return;
  }
  await loginWithEmail(email, password, { mode: "create", realName });
});
loginTabButton?.addEventListener("click", () => setEntranceTab("login"));
createTabButton?.addEventListener("click", () => setEntranceTab("create"));
logoutButton.addEventListener("click", () => {
  logoutCurrentUser();
});
backToCreatorButton.addEventListener("click", () => {
  void returnToCreatorCalendar();
});
doctorSelect.addEventListener("change", async () => {
  await switchDoctorSelection(doctorSelect.value, { resetRange: true });
});

claimDoctorSelect.addEventListener("change", () => {
  claimDoctorButton.disabled = !claimDoctorSelect.value;
});

claimDoctorButton.addEventListener("click", () => {
  claimSelectedRosterName();
});

settingsToggle.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  toggleSettingsPanel();
});
settingsCloseButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeSettingsPanel();
});
mobileSettingsButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  toggleSettingsPanel();
});
mobileLogoutButton?.addEventListener("click", () => {
  logoutCurrentUser();
});
mobileDoctorSelect?.addEventListener("change", async () => {
  if (mobileDoctorSelect.disabled) return;
  await switchDoctorSelection(mobileDoctorSelect.value, { resetRange: true });
});
mobileDateFrom?.addEventListener("change", () => {
  applyPreviewRangeChange("from", mobileDateFrom.value);
  syncMobileSettingsControls();
});
mobileDateTo?.addEventListener("change", () => {
  applyPreviewRangeChange("to", mobileDateTo.value);
  syncMobileSettingsControls();
});
mobileHospitalFilter?.addEventListener("change", () => {
  settings.hospitalFilter = mobileHospitalFilter.value;
  if (settingsInputs.hospitalFilter) settingsInputs.hospitalFilter.value = settings.hospitalFilter;
  saveCurrentSessionState();
  updatePreview();
});

for (const [key, input] of Object.entries(settingsInputs)) {
  if (isPreviewStyleField(key) && input.type === "color") {
    input.addEventListener("input", () => {
      previewStyleDraft = readPreviewStyleInputs(previewStyleDraft || settings);
      applyCurrentDayHighlight(previewStyleDraft, { previewOnly: true });
      applyWeekendShade(previewStyleDraft, { previewOnly: true });
      syncPreviewStyleControls(previewStyleDraft);
    });
  }
  input.addEventListener("change", () => {
    if (isPreviewStyleField(key)) {
      previewStyleDraft = readPreviewStyleInputs(previewStyleDraft || settings);
      applyCurrentDayHighlight(previewStyleDraft, { previewOnly: true });
      applyWeekendShade(previewStyleDraft, { previewOnly: true });
      syncPreviewStyleControls(previewStyleDraft);
      setStatus("Preview styling updated.");
      return;
    }
    if (isPreviewDisplayField(key)) {
      previewDisplayDraft = readPreviewDisplayInputs(previewDisplayDraft || settings);
      updatePreviewDisplayExample(previewDisplayDraft);
      setStatus("Preview display updated in the sample. Save or close settings to apply it.");
      return;
    }
    settings[key] = input.type === "checkbox" ? input.checked : input.value;
    if (["showTimes", "showRawValues", "showNormalizedTitles"].includes(key)) {
      updatePreviewDisplayExample();
    }
    saveCurrentSessionState();
    if (latestPreview && (key === "dateFrom" || key === "dateTo" || key === "hospitalFilter")) {
      rebuildClientPreview();
      setStatus(key === "hospitalFilter" ? "Hospital filter updated." : "Preview range updated.");
      return;
    }
    if (latestPreview && [
      "showSourcePrefix",
      "showAmPm",
      "includeAnnualLeave",
      "includeConferenceLeave",
      "includePublicHoliday",
      "includeSickLeave",
    ].includes(key)) {
      updatePreview();
      return;
    }
    if (latestPreview && ["showTimes", "showRawValues", "showNormalizedTitles"].includes(key)) {
      rebuildClientPreview();
      setStatus("Preview display updated.");
      return;
    }
    if (["includeLocations", "defaultLocationMmc", "defaultLocationDdh", "defaultLocationCasey", "defaultLocationMch"].includes(key)) {
      setStatus(
        key === "includeLocations"
          ? "Location export setting updated."
          : "Default export locations updated.",
      );
      return;
    }
    if (key.startsWith("shiftColor")) {
      applyShiftColours(settings);
      if (latestPreview) rebuildClientPreview();
      setStatus("Shift colours updated.");
      return;
    }
    if (key.startsWith("currentDay")) {
      applyCurrentDayHighlight(settings);
      syncPreviewStyleControls();
      if (latestPreview) rebuildClientPreview();
      setStatus("Current day highlight updated.");
      return;
    }
    if (key.startsWith("weekendShade")) {
      applyWeekendShade(settings);
      syncPreviewStyleControls();
      if (latestPreview) rebuildClientPreview();
      setStatus("Weekend shading updated.");
      return;
    }
    setStatus("Settings updated.");
  });
}

settingsPanel?.addEventListener("click", (event) => {
  const stepButton = event.target.closest("[data-opacity-step]");
  if (stepButton) {
    const stepper = stepButton.closest("[data-opacity-stepper]");
    const field = stepper?.dataset.opacityStepper;
    const input = field ? settingsInputs[field] : null;
    if (input) {
      const next = Math.max(0, Math.min(100, Number(input.value || 0) + Number(stepButton.dataset.opacityStep || 0)));
      input.value = String(next);
      input.dispatchEvent(new Event("change", { bubbles: true }));
    }
    return;
  }
  const fillModeButton = event.target.closest("[data-fill-mode]");
  if (fillModeButton && settingsInputs.currentDayFillStyle) {
    const mode = fillModeButton.dataset.fillMode === "gradient" ? "gradient" : "solid";
    settingsInputs.currentDayFillStyle.value = mode;
    if (mode === "solid" && settingsInputs.currentDayGradientDirection) {
      settingsInputs.currentDayGradientDirection.value = "";
    }
    if (mode === "gradient" && settingsInputs.currentDayGradientDirection && !settingsInputs.currentDayGradientDirection.value) {
      settingsInputs.currentDayGradientDirection.value = "90deg";
    }
    settingsInputs.currentDayFillStyle.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  const directionButton = event.target.closest("[data-gradient-direction]");
  if (directionButton && settingsInputs.currentDayFillStyle && settingsInputs.currentDayGradientDirection) {
    const direction = directionButton.dataset.gradientDirection || "90deg";
    const isActive = settingsInputs.currentDayFillStyle.value === "gradient"
      && settingsInputs.currentDayGradientDirection.value === direction;
    settingsInputs.currentDayFillStyle.value = isActive ? "solid" : "gradient";
    settingsInputs.currentDayGradientDirection.value = isActive ? "" : direction;
    settingsInputs.currentDayFillStyle.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  if (event.target.closest("[data-preview-style-reset]")) {
    const defaults = defaultSettings();
    previewStyleDraft = pickPreviewStyleSettings(defaults);
    writePreviewStyleInputs(previewStyleDraft);
    applyCurrentDayHighlight(previewStyleDraft, { previewOnly: true });
    applyWeekendShade(previewStyleDraft, { previewOnly: true });
    syncPreviewStyleControls(previewStyleDraft);
    setStatus("Preview styling reset in the sample. Save to apply it.");
    return;
  }
  if (event.target.closest("[data-preview-style-cancel]")) {
    previewStyleDraft = null;
    previewDisplayDraft = null;
    renderSettings();
    closeSettingsPanel({ commit: false });
    setStatus("Preview changes cancelled.");
    return;
  }
  if (event.target.closest("[data-preview-style-save]")) {
    closeSettingsPanel();
    setStatus("Preview settings saved.");
  }
});

reviewModalBody.addEventListener("input", (event) => {
  const titleInput = event.target.closest("[data-override-title]");
  if (titleInput) {
    const id = titleInput.dataset.overrideTitle;
    syncImportedOverride(id, { title: titleInput.value });
    rebuildClientPreview();
    setStatus("Mapping override updated.");
    return;
  }

  const customLocationInput = event.target.closest("[data-override-custom-location]");
  if (!customLocationInput) return;
  const id = customLocationInput.dataset.overrideCustomLocation;
  syncImportedOverride(id, { location: customLocationInput.value.trim() });
  rebuildClientPreview();
  setStatus("Event details updated.");
});

reviewModalBody.addEventListener("change", (event) => {
  const includeInput = event.target.closest("[data-override-include]");
  if (includeInput) {
    const id = includeInput.dataset.overrideInclude;
    syncImportedOverride(id, { include: includeInput.checked });
    rebuildClientPreview();
    setStatus("Inclusion override updated.");
    return;
  }

  const startDateInput = event.target.closest("[data-override-start-date]");
  const endDateInput = event.target.closest("[data-override-end-date]");
  const allDayInput = event.target.closest("[data-override-all-day]");
  const startTimeInput = event.target.closest("[data-override-start-time]");
  const endTimeInput = event.target.closest("[data-override-end-time]");
  const locationModeInput = event.target.closest("[data-override-location-mode]");
  const target = startDateInput || endDateInput || allDayInput || startTimeInput || endTimeInput || locationModeInput;
  if (!target) return;
  const id = (
    startDateInput?.dataset.overrideStartDate ||
    endDateInput?.dataset.overrideEndDate ||
    allDayInput?.dataset.overrideAllDay ||
    startTimeInput?.dataset.overrideStartTime ||
    endTimeInput?.dataset.overrideEndTime ||
    locationModeInput?.dataset.overrideLocationMode
  );
  applyImportedEventFormState(id);
  rebuildClientPreview();
  setStatus("Event details updated.");
});

reviewModalBody.addEventListener("click", (event) => {
  const resetButton = event.target.closest("[data-override-reset]");
  if (resetButton) {
    resetImportedEvent(resetButton.dataset.overrideReset);
    return;
  }
  const whenDoctorButton = event.target.closest("[data-inline-when-doctor]");
  if (whenDoctorButton) {
    event.preventDefault();
    if (!canUseRosterInsights()) return;
    const panel = whenDoctorButton.closest(".event-inline-insight");
    void renderInlineWhenInsight(panel, whenDoctorButton.dataset.inlineWhenDoctor || "");
    return;
  }
  const dateButton = event.target.closest("[data-inline-who-on-date]");
  if (dateButton) {
    event.preventDefault();
    if (!canUseRosterInsights()) return;
    closeReviewModal();
    void openWhoInsightForDate(dateButton.dataset.inlineWhoOnDate || "");
    return;
  }
  const backButton = event.target.closest("[data-inline-back-who]");
  if (backButton) {
    event.preventDefault();
    if (!canUseRosterInsights()) return;
    const panel = backButton.closest(".event-inline-insight");
    void renderInlineWhoInsight(panel, backButton.dataset.inlineBackWho || "", { source: backButton.dataset.inlineBackSource || "" });
  }
});

preview.addEventListener("click", (event) => {
  if (Date.now() < suppressPreviewClickUntil) return;
  closeContextMenu();
  const logoutTrigger = event.target.closest("[data-preview-logout]");
  if (logoutTrigger) {
    logoutCurrentUser();
    return;
  }
  const backTrigger = event.target.closest("[data-preview-back-to-creator]");
  if (backTrigger) {
    void returnToCreatorCalendar();
    return;
  }
  const rangeTrigger = event.target.closest("[data-range-trigger]");
  if (rangeTrigger) {
    openPreviewRangePicker(rangeTrigger.dataset.rangeTrigger);
    return;
  }
  const todayTrigger = event.target.closest("[data-range-today]");
  if (todayTrigger) {
    snapPreviewToCurrentMonth();
    return;
  }
  const whoTrigger = event.target.closest("[data-insight-who]");
  if (whoTrigger) {
    if (!canUseRosterInsights()) return;
    void openWhoInsight(whoTrigger.dataset.insightWho, whoTrigger.dataset.insightWhoEnd);
    return;
  }
  const whenTrigger = event.target.closest("[data-insight-when]");
  if (whenTrigger) {
    if (!canUseRosterInsights()) return;
    void openWhenInsight(whenTrigger.dataset.insightWhen, whenTrigger.dataset.insightWhenEnd);
    return;
  }
  const chip = event.target.closest("[data-review-id]");
  if (chip) {
    openReviewModal(chip.dataset.reviewId, chip.dataset.reviewDate || "");
    return;
  }
  const cell = event.target.closest("[data-add-date]");
  if (!cell) return;
  openCustomEventModal(null, cell.dataset.addDate);
});
preview.addEventListener("pointerdown", (event) => {
  const chip = event.target.closest("[data-review-id]");
  if (!chip || event.button !== 0) return;
  if (isMobileLayout()) {
    startPendingPreviewGesture(event, chip);
    return;
  }
  startPreviewGesture(event, chip);
});
preview.addEventListener("change", (event) => {
  const doctorPicker = event.target.closest("[data-preview-doctor-select]");
  if (doctorPicker) {
    void switchDoctorSelection(doctorPicker.value, { resetRange: true });
    return;
  }

  const rangeInput = event.target.closest("[data-range-input]");
  if (rangeInput) {
    applyPreviewRangeChange(rangeInput.dataset.rangeInput, rangeInput.value);
    return;
  }

  const termStartSelect = event.target.closest("[data-preview-term-start]");
  if (termStartSelect) {
    applyPreviewTermStart(termStartSelect.value);
  }
});
preview.addEventListener("change", (event) => {
  const hospitalSelect = event.target.closest("[data-preview-hospital-filter]");
  if (!hospitalSelect) return;
  settings.hospitalFilter = hospitalSelect.value;
  if (settingsInputs.hospitalFilter) settingsInputs.hospitalFilter.value = settings.hospitalFilter;
  saveCurrentSessionState();
  updatePreview();
});
insightsModalBody.addEventListener("change", (event) => {
  const whoDateInput = event.target.closest("[data-insights-who-date]");
  if (whoDateInput && insightsState?.mode === "who") {
    insightsState.date = whoDateInput.value;
    void renderInsightsModal();
    return;
  }
  const whenDoctorSelect = event.target.closest("[data-insights-when-doctor]");
  if (whenDoctorSelect && insightsState?.mode === "when") {
    insightsState.comparisonDoctorKey = whenDoctorSelect.value;
    void renderInsightsModal();
    return;
  }
  const whenFromInput = event.target.closest("[data-insights-when-from]");
  if (whenFromInput && insightsState?.mode === "when") {
    insightsState.fromDate = whenFromInput.value || formatDateKey(new Date());
    void renderInsightsModal();
    return;
  }
  const whenHospitalToggle = event.target.closest("[data-insights-when-hospital]");
  if (whenHospitalToggle && insightsState?.mode === "when") {
    const hospital = String(whenHospitalToggle.value || "").toUpperCase();
    const selected = new Set(insightsState.hospitalFilters || []);
    if (whenHospitalToggle.checked) {
      selected.add(hospital);
    } else {
      selected.delete(hospital);
    }
    insightsState.hospitalFilters = [...selected];
    void renderInsightsModal();
  }
});
exportModalBody.addEventListener("change", (event) => {
  const rangeInput = event.target.closest("[data-export-range-input]");
  if (!rangeInput) return;
  if (rangeInput.dataset.exportRangeInput === "start") {
    pendingExportRange.startDate = rangeInput.value;
  }
  if (rangeInput.dataset.exportRangeInput === "end") {
    pendingExportRange.endDate = rangeInput.value;
  }
});
insightsModalBody.addEventListener("click", async (event) => {
  const dateButton = event.target.closest("[data-insights-who-on-date]");
  if (dateButton) {
    event.preventDefault();
    await openWhoInsightForDate(dateButton.dataset.insightsWhoOnDate || "");
    return;
  }
  const doctorButton = event.target.closest("[data-insights-when-doctor-key]");
  if (doctorButton) {
    event.preventDefault();
    await openWhenInsightForDoctor(doctorButton.dataset.insightsWhenDoctorKey || "");
    return;
  }
  const profileButton = event.target.closest("[data-insights-doctor-key]");
  if (!profileButton) return;
  event.preventDefault();
  await openDoctorProfileFromInsight(profileButton.dataset.insightsDoctorKey);
});
issuesList.addEventListener("click", (event) => {
  const reviewButton = event.target.closest("[data-open-review]");
  if (reviewButton) {
    event.preventDefault();
    event.stopPropagation();
    openReviewModal(reviewButton.dataset.openReview || "");
    return;
  }
  const quickRuleButton = event.target.closest("[data-preview-add-shift-code]");
  if (quickRuleButton) {
    event.preventDefault();
    event.stopPropagation();
    openParserRuleModalFromPreviewIssue(quickRuleButton.dataset.previewAddShiftCode || "");
    return;
  }
  const card = event.target.closest("[data-review-id]");
  if (!card) return;
  const issue = (latestPreview?.issues || []).find((item) => item.id === card.dataset.reviewId);
  if (isShiftCodeIssue(issue)) {
    openParserRuleModalFromPreviewIssue(card.dataset.reviewId);
    return;
  }
  // Warnings should be fix-first. Non-parser warnings currently have no bespoke
  // resolver, so leave the event editor behind an explicit action instead.
  card.querySelector("[data-open-review]")?.focus();
});
conflictsList.addEventListener("change", async (event) => {
  const select = event.target.closest("[data-conflict-key]");
  if (!select) return;
  conflictSelections[select.dataset.conflictKey] = select.value;
  saveConflictSelections();
  saveCurrentSessionState();
  await updatePreview();
});
conflictsList.addEventListener("click", (event) => {
  if (event.target.closest("select")) return;
  event.target.closest(".issue-card")?.querySelector("[data-conflict-key]")?.focus();
});
preview.addEventListener("contextmenu", (event) => {
  const chip = event.target.closest("[data-review-id]");
  const cell = event.target.closest("[data-add-date]");
  if (!chip && !cell) return;
  event.preventDefault();
  const items = [];
  if (chip) {
    const previewEvent = currentPreviewEvents.get(chip.dataset.reviewId);
    items.push({ label: "Copy Event", action: () => copyPreviewEvent(chip.dataset.reviewId) });
    items.push({ label: "Delete Event", action: () => deletePreviewEvent(chip.dataset.reviewId) });
    if (previewEvent?.source !== "Custom" && hasImportedOverride(chip.dataset.reviewId)) {
      items.push({ label: "Reset Event", action: () => resetImportedEvent(chip.dataset.reviewId) });
    }
  } else if (cell) {
    items.push({ label: "Add Event", action: () => openCustomEventModal(null, cell.dataset.addDate) });
    if (copiedEvent) {
      items.push({ label: "Paste Event", action: () => pasteCopiedEvent(cell.dataset.addDate) });
    }
  }
  if (items.length) {
    openContextMenu(event.clientX, event.clientY, items);
  }
});
preview.addEventListener("dragstart", (event) => {
  const chip = event.target.closest("[data-review-id]");
  if (!chip) return;
  dragEventId = chip.dataset.reviewId;
  chip.classList.add("is-dragging");
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("text/plain", dragEventId);
});
preview.addEventListener("dragend", () => {
  dragEventId = null;
  preview.querySelectorAll(".preview-chip.is-dragging").forEach((chip) => chip.classList.remove("is-dragging"));
  preview.querySelectorAll(".preview-cell.is-drop-target").forEach((cell) => cell.classList.remove("is-drop-target"));
});
preview.addEventListener("dragover", (event) => {
  const cell = event.target.closest("[data-add-date]");
  if (!cell || !dragEventId) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
  preview.querySelectorAll(".preview-cell.is-drop-target").forEach((node) => {
    if (node !== cell) node.classList.remove("is-drop-target");
  });
  cell.classList.add("is-drop-target");
});
preview.addEventListener("dragleave", (event) => {
  const cell = event.target.closest("[data-add-date]");
  if (!cell) return;
  if (cell.contains(event.relatedTarget)) return;
  cell.classList.remove("is-drop-target");
});
preview.addEventListener("drop", (event) => {
  const cell = event.target.closest("[data-add-date]");
  if (!cell || !dragEventId) return;
  event.preventDefault();
  cell.classList.remove("is-drop-target");
  movePreviewEvent(dragEventId, cell.dataset.addDate);
});
reviewCloseButton.addEventListener("click", closeReviewModal);
reviewModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-review]")) {
    closeReviewModal();
  }
});
document.addEventListener("keydown", (event) => {
  if (handleHistoryShortcut(event)) return;
  if (event.key === "Escape") {
    closeReviewModal();
    closeCustomEventModal();
    closeParserRuleModal();
    closeContextMenu();
    closeFilesModal();
    closeExportModal();
    closeAccountsModal();
    closeInsightsModal();
    closeSettingsPanel();
  }
});
document.addEventListener("click", (event) => {
  if (!event.target.closest("#contextMenu")) {
    closeContextMenu();
  }
});
document.addEventListener("pointerdown", (event) => {
  if (
    !settingsPanel.classList.contains("hidden")
    && !event.target.closest("#settingsPanel")
    && !event.target.closest("#settingsToggle")
    && !event.target.closest("#mobileSettingsButton")
  ) {
    event.preventDefault();
    event.stopPropagation();
    closeSettingsPanel();
  }
}, true);
window.addEventListener("resize", () => {
  syncMobileChrome();
});
window.visualViewport?.addEventListener("resize", () => {
  syncMobileChrome();
});
window.visualViewport?.addEventListener("scroll", () => {
  syncMobileChrome();
});
document.addEventListener("pointermove", (event) => {
  if (pendingPreviewGesture && event.pointerId === pendingPreviewGesture.pointerId) {
    updatePendingPreviewGesture(event);
    return;
  }
  if (!previewGesture || event.pointerId !== previewGesture.pointerId) return;
  updatePreviewGesture(event);
});
document.addEventListener("pointerup", (event) => {
  if (pendingPreviewGesture && event.pointerId === pendingPreviewGesture.pointerId) {
    cancelPendingPreviewGesture();
    return;
  }
  if (!previewGesture || event.pointerId !== previewGesture.pointerId) return;
  finishPreviewGesture(event);
});
document.addEventListener("pointercancel", (event) => {
  if (pendingPreviewGesture && event.pointerId === pendingPreviewGesture.pointerId) {
    cancelPendingPreviewGesture();
    return;
  }
  if (!previewGesture || event.pointerId !== previewGesture.pointerId) return;
  cancelPreviewGesture();
});
customEventAllDay.addEventListener("change", () => {
  customEventTimeFields.classList.toggle("hidden", customEventAllDay.checked);
});
customEventLocationMode.addEventListener("change", () => {
  customEventCustomLocationField.classList.toggle("hidden", customEventLocationMode.value !== "custom");
});
customEventCloseButton.addEventListener("click", closeCustomEventModal);
customEventModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-custom-event]")) {
    closeCustomEventModal();
  }
});
parserRuleCloseButton?.addEventListener("click", closeParserRuleModal);
parserRuleModal?.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-parser-rule]")) {
    closeParserRuleModal();
  }
});
parserRuleAllDay?.addEventListener("change", () => {
  parserRuleTimeFields?.classList.toggle("hidden", parserRuleAllDay.checked);
  renderParserRulePreview();
});
parserRuleIgnore?.addEventListener("change", () => {
  syncParserRuleIgnoreControls();
  renderParserRulePreview();
});
parserRuleForm?.addEventListener("input", () => renderParserRulePreview());
parserRuleForm?.addEventListener("change", (event) => {
  if (event.target?.matches?.('input[name="parserRuleSeniorityOption"], input[name="parserRuleSeniorityAll"]')) {
    normalizeParserRuleSenioritySelection(event.target);
  }
  if (event.target === parserRuleIgnore) syncParserRuleIgnoreControls();
  renderParserRulePreview();
});
parserRuleForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await saveParserRuleFromModal();
});
customEventDeleteButton.addEventListener("click", () => {
  const id = customEventId.value;
  if (!id) return;
  removeCustomEventForActiveCalendar(id);
  closeCustomEventModal();
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Manual event removed.");
});
customEventForm.addEventListener("click", (event) => {
  const whenDoctorButton = event.target.closest("[data-inline-when-doctor]");
  if (whenDoctorButton) {
    event.preventDefault();
    void renderInlineWhenInsight(customEventWhoPanel, whenDoctorButton.dataset.inlineWhenDoctor || "");
    return;
  }
  const dateButton = event.target.closest("[data-inline-who-on-date]");
  if (dateButton) {
    event.preventDefault();
    closeCustomEventModal();
    void openWhoInsightForDate(dateButton.dataset.inlineWhoOnDate || "");
    return;
  }
  const backButton = event.target.closest("[data-inline-back-who]");
  if (backButton) {
    event.preventDefault();
    void renderInlineWhoInsight(customEventWhoPanel, backButton.dataset.inlineBackWho || "", { source: backButton.dataset.inlineBackSource || "" });
  }
});
customEventForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const entry = readCustomEventForm();
  if (!entry) return;
  const ownerEmail = activeCalendarEmail();
  const index = customEvents.findIndex((item) => item.id === entry.id && normalizeEmail(item.ownerEmail) === ownerEmail);
  if (index >= 0) {
    customEvents[index] = entry;
    setStatus("Manual event updated.");
  } else {
    customEvents.push(entry);
    setStatus("Manual event added.");
  }
  closeCustomEventModal();
  rebuildClientPreview();
  saveCurrentSessionState();
});

function defaultSettings() {
  return {
    showSourcePrefix: true,
    showAmPm: true,
    showTimes: true,
    showRawValues: false,
    showNormalizedTitles: true,
    currentDayBorderColor: "#c44949",
    currentDayBorderOpacity: "20",
    currentDayBorderWidth: "2",
    currentDayBackgroundColor: "#c44949",
    currentDayBackgroundOpacity: "20",
    currentDayFillStyle: "solid",
    currentDayGradientDirection: "",
    weekendShadeEnabled: true,
    weekendShadeColor: "#e5e7eb",
    weekendShadeOpacity: "30",
    shiftColorDay: SHIFT_COLOUR_DEFAULTS.day,
    shiftColorEvening: SHIFT_COLOUR_DEFAULTS.evening,
    shiftColorNight: SHIFT_COLOUR_DEFAULTS.night,
    shiftColorCs: SHIFT_COLOUR_DEFAULTS.cs,
    shiftColorLeave: SHIFT_COLOUR_DEFAULTS.leave,
    shiftColorCustom: SHIFT_COLOUR_DEFAULTS.custom,
    shiftColorPhnw: SHIFT_COLOUR_DEFAULTS.phnw,
    includeLocations: true,
    includeAnnualLeave: true,
    includeConferenceLeave: true,
    includePublicHoliday: true,
    includeSickLeave: true,
    defaultLocationMmc: DEFAULT_MMC_LOCATION,
    defaultLocationDdh: DEFAULT_DDH_LOCATION,
    defaultLocationCasey: DEFAULT_CASEY_LOCATION,
    defaultLocationMch: DEFAULT_MCH_LOCATION,
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
  };
}

async function mergeFiles(files) {
  let persistenceFailed = false;
  lastRosterPersistence = null;
  for (const file of files) {
    const id = fileFingerprint(file);
    selectedFiles = selectedFiles.filter((entry) => entry.id !== id);
    const entry = {
      id,
      file,
      name: file.name,
      size: file.size,
      lastModified: file.lastModified,
      addedAt: new Date().toISOString(),
      sourceType: "pending",
    };
    selectedFiles.push(entry);
    try {
      await saveStoredImport(entry);
    } catch {
      persistenceFailed = true;
    }
  }
  selectedFiles.sort((left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || left.name.localeCompare(right.name));
  renderFilesList();
  saveCurrentWorkspace();
  if (persistenceFailed) {
    setStatus("Import added, but browser storage was unavailable so it will not persist after reload.", true);
  }
}

function validateIncomingFiles(files) {
  if (files.some((file) => !file.name.match(/\.(xlsx|xlsm|xltx|xltm|pdf)$/i))) {
    setStatus("Only Excel or PDF roster files in .xlsx, .xlsm, .xltx, .xltm, or .pdf format are supported.", true);
    return [];
  }
  return files;
}

function hasFileDrag(dataTransfer) {
  return Boolean(dataTransfer?.types && [...dataTransfer.types].includes("Files"));
}

function syncRosterDragState(dataTransfer) {
  const supported = isSupportedRosterDrag(dataTransfer);
  document.body.classList.toggle("is-roster-dragging", supported);
  rosterDropOverlay.classList.toggle("hidden", !supported);
  rosterDropOverlay.setAttribute("aria-hidden", supported ? "false" : "true");
}

function clearRosterDragState() {
  rosterDragDepth = 0;
  document.body.classList.remove("is-roster-dragging");
  rosterDropOverlay.classList.add("hidden");
  rosterDropOverlay.setAttribute("aria-hidden", "true");
}

function isSupportedRosterDrag(dataTransfer) {
  const files = [...(dataTransfer?.files || [])];
  if (files.length) return files.every(isSupportedRosterFile);

  const fileItems = [...(dataTransfer?.items || [])].filter((item) => item.kind === "file");
  if (!fileItems.length) return false;
  const itemFiles = fileItems.map((item) => item.getAsFile?.()).filter(Boolean);
  if (itemFiles.length === fileItems.length) return itemFiles.every(isSupportedRosterFile);
  return fileItems.every((item) => isSupportedRosterMimeType(item.type));
}

function isSupportedRosterFile(file) {
  return Boolean(file?.name?.match(/\.(xlsx|xlsm|xltx|xltm|pdf)$/i));
}

function isSupportedRosterMimeType(type) {
  return [
    "application/pdf",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel.sheet.macroenabled.12",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    "application/vnd.ms-excel.template.macroenabled.12",
  ].includes(String(type || "").toLowerCase());
}

async function validateFreshRosterUploads(files) {
  try {
    const formData = new FormData();
    for (const file of files) formData.append("rosterFiles", file);
    await parseUploadForm(new Request(`${window.location.origin}/browser-roster-validate`, {
      method: "POST",
      body: formData,
    }));
    return true;
  } catch (error) {
    setStatus(error.message || "Could not validate roster dates.", true);
    return false;
  }
}

async function analyzeFiles(options = {}) {
  if (!selectedFiles.length) {
    setStatus("Add a roster file to begin.");
    return;
  }
  await ensureSelectedFilesLoaded();
  if (!options.preserveVisiblePreview) {
    clearPreviewData();
    doctorOptions = [];
    detectedSources = {};
  }
  parsedRosterSources = null;
  doctorRoleIndex = null;
  parsedImportDoctors = new Map();
  clearDoctorAnalysisCache();
  controlBar.classList.remove("hidden");
  mobileActionBar.classList.remove("hidden");
  setStatus("Detecting roster sources and consultants...");
  try {
    const data = await analyzeFilesInBrowser();
    doctorOptions = doctorOptionsForCurrentAccount(data.doctors || []);
    detectedSources = summarizeDetectedSources(data.imports || []);
    selectedFiles = selectedFiles.map((entry) => {
      const serverEntry = (data.imports || []).find((item) => item.id === entry.id);
      return serverEntry ? { ...entry, sourceType: serverEntry.sourceType } : entry;
    });
    const workspaceSession = loadCurrentSessionState();
    restoredSessionState = cloudAvailable && restoredSessionState
      ? restoredSessionState
      : workspaceSession;
    applySessionState(restoredSessionState, { inheritedSettings: data.settings || {} });
    renderSettings();
    renderFilesList();
    renderDoctorState();
    saveCurrentWorkspace();
    scheduleCloudStateSave();
    if (selectedDoctor()) {
      await updatePreview({ resetRange: options.resetRange !== false });
      return;
    }
  } catch (error) {
    if (!options.preserveVisiblePreview) {
      doctorOptions = [];
      detectedSources = {};
      clearPreviewData();
      renderFilesList();
      syncActionState();
    }
    setStatus(error.message, true);
  }
}

async function analyzeFilesInBrowser() {
  const parsed = await parseCurrentRosterForm(null);
  parsedRosterSources = parsed.sources;
  parsedImportDoctors = doctorsByImportId(parsed.sources);
  const imports = sourceImports(parsed.sources);
  return {
    sources: sourceNames(parsed.sources),
    imports,
    doctors: rosterDoctorOptions(parsed.sources.mmc, parsed.sources.ddh, parsed.sources.casey, parsed.sources.mch),
    settings: rosterDefaultSettings(),
  };
}

async function parseCurrentRosterForm(doctor = null) {
  if (!await ensureSelectedFilesLoaded()) {
    throw new Error("Roster files need to be re-uploaded so they can be parsed into D1.");
  }
  return await parseRosterEntries(selectedFiles, doctor);
}

async function parseRosterEntries(entries, doctor = null) {
  const normalizedEntries = Array.isArray(entries) ? entries : [];
  if (!normalizedEntries.length) {
    throw new Error("Add a roster file to begin.");
  }
  if (normalizedEntries.some((entry) => !entry.file)) {
    throw new Error("Roster files are not loaded yet.");
  }
  return await parseUploadForm(new Request(`${window.location.origin}/browser-roster-parse`, {
    method: "POST",
    body: createFormDataForEntries(normalizedEntries, doctor),
  }));
}

async function parseRosterEntriesLenient(entries, doctor = null) {
  try {
    return await parseRosterEntries(entries, doctor);
  } catch (error) {
    if (!isUnsupportedRosterError(error)) throw error;
    const sources = { mmc: [], ddh: [], casey: [], mch: [] };
    let parsedAny = false;
    for (const entry of entries || []) {
      try {
        const parsed = await parseRosterEntries([entry], doctor);
        sources.mmc.push(...(parsed.sources?.mmc || []));
        sources.ddh.push(...(parsed.sources?.ddh || []));
        sources.casey.push(...(parsed.sources?.casey || []));
        sources.mch.push(...(parsed.sources?.mch || []));
        parsedAny = true;
      } catch (entryError) {
        if (!isUnsupportedRosterError(entryError)) throw entryError;
      }
    }
    if (!parsedAny) throw error;
    return { sources };
  }
}

function isUnsupportedRosterError(error) {
  return /is not a supported MMC workbook, MMC PDF, Dandenong Hospital FindMyShift export, Casey roster, or MCH roster|is not a supported MMC workbook, MMC PDF, Dandenong Hospital FindMyShift export, or Casey roster|is not a supported MMC workbook, MMC PDF, or Dandenong Hospital FindMyShift export/i.test(String(error?.message || error || ""));
}

function sourceImports(sources) {
  return [
    ...(sources.mmc || []).map((entry) => sourceImportMeta(entry, "mmc")),
    ...(sources.ddh || []).map((entry) => sourceImportMeta(entry, "ddh")),
    ...(sources.casey || []).map((entry) => sourceImportMeta(entry, "casey")),
    ...(sources.mch || []).map((entry) => sourceImportMeta(entry, "mch")),
  ];
}

function sourceImportMeta(entry, sourceType) {
  return {
    id: entry.id,
    name: entry.file.name,
    sourceType,
    addedAt: entry.addedAt || "",
    size: entry.file.size,
    lastModified: entry.file.lastModified,
  };
}

function doctorsByImportId(sources) {
  const result = new Map();
  for (const entry of sources.mmc || []) {
    result.set(entry.id, rosterDoctorOptions([entry], [], [], []).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "mmc",
    })));
  }
  for (const entry of sources.ddh || []) {
    result.set(entry.id, rosterDoctorOptions([], [entry], [], []).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "ddh",
    })));
  }
  for (const entry of sources.casey || []) {
    result.set(entry.id, rosterDoctorOptions([], [], [entry], []).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "casey",
    })));
  }
  for (const entry of sources.mch || []) {
    result.set(entry.id, rosterDoctorOptions([], [], [], [entry]).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "mch",
    })));
  }
  return result;
}

function renderSettings() {
  previewStyleDraft = null;
  previewDisplayDraft = null;
  for (const [key, input] of Object.entries(settingsInputs)) {
    if (!input) continue;
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else if (input.type === "color") {
      input.value = isHexColour(settings[key]) ? settings[key] : defaultColourForField(key);
    } else {
      input.value = settings[key] || "";
    }
  }
  applyShiftColours(settings);
  applyCurrentDayHighlight(settings);
  applyWeekendShade(settings);
  syncPreviewStyleControls();
  updatePreviewDisplayExample();
  syncMobileSettingsControls();
}

function isPreviewStyleField(field) {
  return PREVIEW_STYLE_FIELDS.includes(field);
}

function isPreviewDisplayField(field) {
  return PREVIEW_DISPLAY_FIELDS.includes(field);
}

function pickPreviewDisplaySettings(source = settings) {
  return Object.fromEntries(PREVIEW_DISPLAY_FIELDS.map((field) => [field, Boolean(source[field])]));
}

function readPreviewDisplayInputs(base = settings) {
  const next = { ...pickPreviewDisplaySettings(base) };
  for (const field of PREVIEW_DISPLAY_FIELDS) {
    const input = settingsInputs[field];
    if (input) next[field] = Boolean(input.checked);
  }
  return next;
}

function pickPreviewStyleSettings(source = settings) {
  return Object.fromEntries(PREVIEW_STYLE_FIELDS.map((field) => [field, source[field]]));
}

function readPreviewStyleInputs(base = settings) {
  const next = { ...pickPreviewStyleSettings(base) };
  for (const field of PREVIEW_STYLE_FIELDS) {
    const input = settingsInputs[field];
    if (!input) continue;
    next[field] = input.type === "checkbox" ? input.checked : input.value;
  }
  return next;
}

function writePreviewStyleInputs(source = settings) {
  for (const field of PREVIEW_STYLE_FIELDS) {
    const input = settingsInputs[field];
    if (!input) continue;
    if (input.type === "checkbox") {
      input.checked = Boolean(source[field]);
    } else if (input.type === "color") {
      input.value = isHexColour(source[field]) ? source[field] : defaultColourForField(field);
    } else {
      input.value = source[field] || "";
    }
  }
}

function setEntranceTab(tab) {
  const active = tab === "create" ? "create" : "login";
  loginTabButton?.classList.toggle("is-active", active === "login");
  createTabButton?.classList.toggle("is-active", active === "create");
  loginTabButton?.setAttribute("aria-selected", active === "login" ? "true" : "false");
  createTabButton?.setAttribute("aria-selected", active === "create" ? "true" : "false");
  entrancePanels.forEach((panel) => {
    panel.classList.toggle("hidden", panel.dataset.entrancePanel !== active);
  });
}

function isMobileLayout() {
  return window.matchMedia("(max-width: 760px)").matches;
}

function syncMobileViewportInsets() {
  const root = document.documentElement;
  const viewport = window.visualViewport;
  const margin = 12;
  const dockClearance = 65;
  const topClearance = -10;
  const modalTopClearance = 58;
  const modalBottomClearance = 92;
  if (!viewport) {
    root.style.setProperty("--mobile-dock-left", `${margin}px`);
    root.style.setProperty("--mobile-dock-width", `calc(100vw - ${margin * 2}px)`);
    root.style.setProperty("--mobile-dock-top", `calc(100vh - ${dockClearance}px)`);
    root.style.setProperty("--mobile-top-anchor", `${topClearance}px`);
    root.style.setProperty("--mobile-modal-left", `${margin}px`);
    root.style.setProperty("--mobile-modal-width", `calc(100vw - ${margin * 2}px)`);
    root.style.setProperty("--mobile-modal-top", `${modalTopClearance}px`);
    root.style.setProperty("--mobile-modal-bottom", `calc(100vh - ${modalBottomClearance}px)`);
    return;
  }
  root.style.setProperty("--mobile-dock-left", `${Math.round(viewport.offsetLeft + margin)}px`);
  root.style.setProperty("--mobile-dock-width", `${Math.round(Math.max(0, viewport.width - margin * 2))}px`);
  root.style.setProperty("--mobile-dock-top", `${Math.round(viewport.offsetTop + viewport.height - dockClearance)}px`);
  root.style.setProperty("--mobile-top-anchor", `${Math.round(viewport.offsetTop + topClearance)}px`);
  root.style.setProperty("--mobile-modal-left", `${Math.round(viewport.offsetLeft + margin)}px`);
  root.style.setProperty("--mobile-modal-width", `${Math.round(Math.max(0, viewport.width - margin * 2))}px`);
  root.style.setProperty("--mobile-modal-top", `${Math.round(viewport.offsetTop + modalTopClearance)}px`);
  root.style.setProperty("--mobile-modal-bottom", `${Math.round(viewport.offsetTop + viewport.height - modalBottomClearance)}px`);
}

function hasCalendarPreview() {
  return Boolean(latestPreview && selectedDoctor());
}

function toggleSettingsPanel() {
  if (settingsPanel.classList.contains("hidden")) {
    renderSettings();
    settingsPanel.classList.remove("hidden");
    syncMobileSettingsControls();
    syncOverlayState();
    return;
  }
  closeSettingsPanel();
}

function closeSettingsPanel(options = {}) {
  const shouldCommit = options.commit !== false;
  if (shouldCommit) {
    commitSettingsDrafts();
  }
  settingsPanel.classList.add("hidden");
  syncOverlayState();
}

function commitSettingsDrafts() {
  let previewNeedsRebuild = false;
  let changed = false;
  if (previewDisplayDraft) {
    const nextDisplay = readPreviewDisplayInputs(previewDisplayDraft);
    previewNeedsRebuild = ["showTimes", "showRawValues", "showNormalizedTitles"]
      .some((field) => settings[field] !== nextDisplay[field]);
    Object.assign(settings, nextDisplay);
    previewDisplayDraft = null;
    changed = changed || previewNeedsRebuild;
  }
  if (previewStyleDraft) {
    Object.assign(settings, previewStyleDraft);
    previewStyleDraft = null;
    applyCurrentDayHighlight(settings);
    applyWeekendShade(settings);
    syncPreviewStyleControls(settings);
    previewNeedsRebuild = true;
    changed = true;
  }
  if (!changed) return;
  saveCurrentSessionState();
  updatePreviewDisplayExample(settings);
  if (latestPreview && previewNeedsRebuild) rebuildClientPreview();
}

function setMobileBodyScrollLock(locked) {
  if (locked && !isMobileLayout()) {
    if (mobileScrollLocked) setMobileBodyScrollLock(false);
    return;
  }
  if (locked && !mobileScrollLocked) {
    mobileScrollLockY = window.scrollY || document.documentElement.scrollTop || 0;
    document.body.style.position = "fixed";
    document.body.style.top = `-${mobileScrollLockY}px`;
    document.body.style.left = "0";
    document.body.style.right = "0";
    document.body.style.width = "100%";
    mobileScrollLocked = true;
    return;
  }
  if (!locked && mobileScrollLocked) {
    const restoreY = mobileScrollLockY;
    document.body.style.position = "";
    document.body.style.top = "";
    document.body.style.left = "";
    document.body.style.right = "";
    document.body.style.width = "";
    mobileScrollLocked = false;
    mobileScrollLockY = 0;
    window.scrollTo(0, restoreY);
  }
}

function syncOverlayState() {
  const active = [
    settingsPanel,
    exportModal,
    filesModal,
    accountsModal,
    insightsModal,
    parserRuleModal,
    reviewModal,
    customEventModal,
  ].some((surface) => surface && !surface.classList.contains("hidden"));
  document.body.classList.toggle("has-active-popup", active);
  setMobileBodyScrollLock(active);
}

async function openAccountsSurface(options = {}) {
  closeSettingsPanel();
  if (isViewingCreatorAccount()) {
    if (options.defaultAdminTab) currentAdminTab = options.defaultAdminTab;
  }
  renderAccountsModal();
  accountsModal.classList.remove("hidden");
  accountsModal.setAttribute("aria-hidden", "false");
  if (isViewingCreatorAccount()) {
    void loadServerUsers().then(() => {
      if (!accountsModal.classList.contains("hidden")) renderAccountsModal();
    });
    void refreshCalendarStoreStatus({ silent: true });
  }
}

function syncMobileAccountButtons() {
  const label = isViewingCreatorAccount() ? "Admin" : "Account";
  if (mobileAccountButtonLabel) mobileAccountButtonLabel.textContent = label;
  if (mobileAccountButton) mobileAccountButton.setAttribute("aria-label", label);
  if (mobileAccountAccessButton) {
    mobileAccountAccessButton.textContent = label;
    mobileAccountAccessButton.setAttribute("aria-label", label);
  }
}

function syncMobileSettingsControls() {
  const mobile = isMobileLayout();
  mobileSettingsControls?.classList.toggle("hidden", !mobile);
  if (!mobile) return;
  if (mobileDoctorSelect) {
    const pickerOptions = doctorPickerOptions();
    const selected = selectedDoctor();
    if (pickerOptions.length > 1) {
      mobileDoctorSelect.innerHTML = pickerOptions.map((doctor) => `
        <option value="${escapeHtml(doctor.key)}" ${doctor.key === selected?.key ? "selected" : ""}>
          ${escapeHtml(doctor.displayName)}
        </option>
      `).join("");
      mobileDoctorSelect.disabled = false;
    } else {
      mobileDoctorSelect.innerHTML = `<option value="${escapeHtml(selected?.key || "")}">${escapeHtml(selected?.displayName || currentAccount().realName || "Selected doctor")}</option>`;
      mobileDoctorSelect.disabled = true;
    }
  }
  if (mobileDateFrom) mobileDateFrom.value = settings.dateFrom || "";
  if (mobileDateTo) mobileDateTo.value = settings.dateTo || "";
  if (mobileHospitalFilter) {
    const hospitals = latestPreview?.hospitals || [];
    mobileHospitalFilter.innerHTML = `
      <option value="all" ${settings.hospitalFilter === "all" ? "selected" : ""}>All hospitals</option>
      ${hospitals.map((code) => {
        const value = code.toLowerCase();
        return `<option value="${value}" ${settings.hospitalFilter === value ? "selected" : ""}>${escapeHtml(code)}</option>`;
      }).join("")}
    `;
  }
}

function syncMobileChrome() {
  const loggedIn = Boolean(currentUserEmail && currentUserPassword);
  const mobile = isMobileLayout();
  syncMobileViewportInsets();
  const showBar = loggedIn && mobile && hasCalendarPreview();
  mobileActionBar.classList.toggle("hidden", !showBar);
  if (mobileAccountAccessButton) {
    mobileAccountAccessButton.classList.toggle("hidden", !(loggedIn && mobile && !hasCalendarPreview()));
  }
  if (!showBar) closeSettingsPanel();
  syncMobileAccountButtons();
  syncMobileSettingsControls();
}

function renderFilesMarkup({ canRemove = false, heading = "", description = "", canAdd = false } = {}) {
  const hasUsableStatus = Boolean(calendarStoreStatus && calendarStoreStatus.unavailable !== true && !calendarStoreStatusError);
  const statusFiles = new Map((calendarStoreStatus?.files || []).map((file) => [file.id, file]));
  const persistedSelectedFileIds = new Set(calendarStoreStatus?.expectedFiles?.persistedFileIds || []);
  const statusOnlyEntries = hasUsableStatus
    ? (calendarStoreStatus?.files || [])
      .filter((file) => file?.id)
      .map((file) => ({
        id: file.id,
        repoId: file.id,
        name: file.name,
        sourceType: file.sourceType,
        addedAt: "",
        fromRosterDatabase: true,
      }))
    : [];
  const displayFiles = selectedFiles.length ? selectedFiles : statusOnlyEntries;
  if (!displayFiles.length) {
    const emptyMessage = canRemove
      ? "Add rosters and they will stay here until removed."
      : "No files are currently linked to this calendar.";
    return `
      <article class="review-card">
        ${heading ? `<div class="review-top"><div><strong>${escapeHtml(heading)}</strong>${description ? `<span>${escapeHtml(description)}</span>` : ""}</div>${canAdd ? `<button type="button" class="button button-secondary" data-open-file-picker>Add files</button>` : ""}</div>` : ""}
        <article class="issue-card"><strong>No files imported yet.</strong><p>${escapeHtml(emptyMessage)}</p></article>
      </article>
    `;
  }
  return `
    <article class="review-card">
      ${heading ? `<div class="review-top"><div><strong>${escapeHtml(heading)}</strong>${description ? `<span>${escapeHtml(description)}</span>` : ""}</div>${canAdd ? `<button type="button" class="button button-secondary" data-open-file-picker>Add files</button>` : ""}</div>` : ""}
      <div class="file-summary">
        ${displayFiles.map((entry) => `
          <article class="file-pill" data-file-id="${entry.id}">
            <span>${escapeHtml(String(entry.sourceType || "").toUpperCase())}${entry.addedAt ? ` · Imported ${escapeHtml(formatTimestamp(entry.addedAt))}` : " · Roster database"}</span>
            <strong>${escapeHtml(entry.name)}</strong>
            ${rosterSyncLabel(entry) || (statusFiles.has(entry.id)
              ? statusFiles.get(entry.id)?.retainedSourceOnly
                ? `<span>Retained in R2 · not yet synced to D1</span>`
                : `<span>${Number(statusFiles.get(entry.id)?.eventCount || 0)} events · ${Number(statusFiles.get(entry.id)?.selectedDoctorEventCount || 0)} for selected doctor</span>`
              : persistedSelectedFileIds.has(entry.id) ? `<span>Saved in D1 · inactive</span>`
              : entry.file && hasUsableStatus ? `<span>Not yet confirmed in D1</span>`
              : entry.file ? `<span>Roster database status not checked</span>` : "")}
            ${statusFiles.has(entry.id) && statusFiles.get(entry.id)?.rawSourceAvailable !== true
              ? `<span>Source file not retained · re-upload once to enable reparse</span>`
              : ""}
            ${canRemove ? `<button type="button" class="file-remove file-remove-visible" aria-label="Remove file" title="Remove file" data-remove-import="${entry.id}">🗑</button>` : ""}
            ${canRemove && (!entry.fromRosterDatabase || statusFiles.get(entry.id)?.rawSourceAvailable === true) ? `<button type="button" class="file-reparse file-reparse-visible" aria-label="Reparse roster file" title="Reparse roster file" data-reparse-import="${entry.id}">↻</button>` : ""}
          </article>
        `).join("")}
      </div>
    </article>
  `;
}

function renderFilesList() {
  if (!filesList) return;
  filesList.innerHTML = renderFilesMarkup({ canRemove: canRemoveImports(), canAdd: true });
}

function renderFileSurfaces() {
  renderFilesList();
  if (accountsModal && !accountsModal.classList.contains("hidden")) {
    renderAccountsModal();
  }
}

function renderDoctorState() {
  const pickerOptions = doctorPickerOptions();
  doctorSelect.innerHTML = "";
  doctorName.textContent = "";
  doctorName.classList.add("hidden");
  doctorSelect.classList.add("hidden");
  doctorSection.classList.add("hidden");
  syncControlBarVisibility();
  closeSettingsPanel();

  if (!pickerOptions.length) {
    const message = canUseDoctorPicker()
      ? "No consultant names could be matched from the uploaded roster files."
      : "No roster entries are currently linked to your account name.";
    setStatus(message, true);
    renderClaimSection();
    syncActionState();
    return;
  }

  claimSection.classList.add("hidden");
  doctorSection.classList.remove("hidden");

  if (pickerOptions.length === 1 && !canUseCreatorDoctorSwitcher()) {
    doctorName.textContent = pickerOptions[0].displayName;
    doctorName.classList.remove("hidden");
    setStatus("Loading calendar...");
  } else {
    for (const doctor of pickerOptions) {
      const option = document.createElement("option");
      option.value = doctor.key;
      option.textContent = doctor.displayName;
      option.dataset.displayName = doctor.displayName;
      option.dataset.sourceTypes = normalizedDoctorSourceTypes(doctor).join(",");
      option.dataset.accountEmail = doctor.accountEmail || "";
      doctorSelect.append(option);
    }
    const preferredDoctorKey = activeDoctorProfile?.doctorKey || preferredDoctorKeyForCurrentAccount();
    if (preferredDoctorKey && pickerOptions.some((doctor) => doctor.key === preferredDoctorKey)) {
      doctorSelect.value = preferredDoctorKey;
    } else if (restoredSessionState?.doctorKey && pickerOptions.some((doctor) => doctor.key === restoredSessionState.doctorKey)) {
      doctorSelect.value = restoredSessionState.doctorKey;
    }
    doctorSelect.classList.remove("hidden");
    setStatus(preferredDoctorKey ? "Loading calendar..." : "Choose a doctor to load the calendar.");
  }

  syncActionState();
  syncMobileChrome();
}

function renderClaimSection() {
  if (!claimSection) return;
  const shouldShow = !canUseDoctorPicker() && !doctorOptions.length && availableRosterDoctors.length;
  claimSection.classList.toggle("hidden", !shouldShow);
  if (!shouldShow) return;

  const unclaimed = [];
  const claimed = [];
  availableRosterDoctors.forEach((doctor, index) => {
    const item = {
      index,
      claimed: Boolean(doctor.claimedBy),
      label: `${doctor.displayName} (${doctor.sourceType.toUpperCase()})${doctor.claimedBy ? " - already claimed" : ""}`,
    };
    (item.claimed ? claimed : unclaimed).push(item);
  });
  unclaimed.sort((left, right) => left.label.localeCompare(right.label));
  claimed.sort((left, right) => left.label.localeCompare(right.label));
  claimDoctorSelect.innerHTML = `
    <option value="">My name is not listed</option>
    ${unclaimed.length ? `<optgroup label="Unclaimed names">${unclaimed.map((item) => `<option value="${item.index}">${escapeHtml(item.label)}</option>`).join("")}</optgroup>` : ""}
    ${claimed.length ? `<optgroup label="Already claimed">${claimed.map((item) => `<option value="${item.index}" class="claimed-option">${escapeHtml(item.label)}</option>`).join("")}</optgroup>` : ""}
  `;
  claimDoctorButton.disabled = true;
}

async function claimSelectedRosterName(candidateOverride = null) {
  const index = Number(claimDoctorSelect.value);
  const candidate = candidateOverride || (Number.isInteger(index) ? availableRosterDoctors[index] : null);
  if (!candidate) {
    setStatus("If your name is not listed, upload the first roster file for your hospital.", true);
    return;
  }
  if (candidate.claimedBy && normalizeEmail(candidate.claimedBy) !== currentUserEmail) {
    setStatus(`${candidate.displayName} is already claimed. Conflict notification is still to be added.`, true);
    return;
  }

  setStatus("Linking roster name...");
  try {
    const requestEmail = adminViewingEmail ? authUserEmail : currentUserEmail;
    const requestPassword = adminViewingEmail ? authUserPassword : currentUserPassword;
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "claimRosterName",
        email: requestEmail,
        password: requestPassword,
        targetEmail: adminViewingEmail ? currentUserEmail : "",
        claim: candidate,
      }),
    });
    const data = await readJsonResponse(response, "Could not link roster name.");
    await applyCloudStateData(data);
    const loadedCalendar = await loadCloudCalendarEvents({ adminTargetEmail: adminViewingEmail ? viewedAccountEmail() : "" });
    if (isCreatorAuthenticated()) await loadServerUsers();
    if (loadedCalendar) {
      await bootstrapImports();
    } else if (cloudAvailable && selectedFiles.some((entry) => !entry.file)) {
      renderFilesList();
      renderClaimSection();
      syncActionState();
      setStatus(`Linked ${candidate.displayName} (${candidate.sourceType.toUpperCase()}), but no D1 calendar events were found for this roster name.`, true);
      return;
    } else {
      await bootstrapImports();
    }
    renderLoginState();
    setStatus(`Linked ${candidate.displayName} (${candidate.sourceType.toUpperCase()}).`);
  } catch (error) {
    setStatus(error.message || "Could not link roster name.", true);
  }
}

async function updatePreview(options = {}) {
  const doctor = selectedDoctor();
  if (!doctor) {
    setStatus("Choose a doctor before loading the calendar.", true);
    return;
  }
  setStatus("Loading calendar...");
  try {
    if (options.resetRange) {
      settings.dateFrom = "";
      settings.dateTo = "";
    }
    const data = await buildBrowserPreviewData(doctor);
    latestPreview = data;
    if (shouldApplyDefaultPreviewRange(options, data.events || [])) {
      const range = deriveDefaultPreviewRange(data.events || []);
      settings.dateFrom = range.start;
      settings.dateTo = range.end;
      pendingPreviewSnapToToday = true;
      renderSettings();
    }
    indexReviewItems(data.review || []);
    rebuildClientPreview();
    scheduleInsightWarmup();
    cacheCurrentSnapshot(buildActiveSessionState());
    saveCurrentSessionState();
    setStatus(hasUnconfirmedLocalRosterFiles()
      ? "Calendar preview loaded locally. Saving roster files to D1..."
      : "Calendar loaded.");
  } catch (error) {
    clearPreviewData();
    setStatus(error.message, true);
  }
}

async function buildBrowserPreviewData(doctor) {
  if (!parsedRosterSources) {
    const parsed = await parseCurrentRosterForm(doctor);
    parsedRosterSources = parsed.sources;
    parsedImportDoctors = doctorsByImportId(parsed.sources);
  }
  if (!doctor?.key) {
    throw new Error("A doctor selection is required.");
  }
  const validDoctors = new Set(rosterDoctorOptions(parsedRosterSources.mmc, parsedRosterSources.ddh, parsedRosterSources.casey, parsedRosterSources.mch).map((item) => item.key));
  const requestedKeys = new Set([doctor.key, ...(doctor.aliases || []).map((alias) => alias.key)].filter(Boolean));
  if (![...requestedKeys].some((key) => validDoctors.has(key))) {
    throw new Error("The selected doctor was not found in the uploaded roster files.");
  }
  const baseSettings = {
    ...settings,
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
  };
  const view = buildRosterView(
    parsedRosterSources.mmc,
    parsedRosterSources.ddh,
    doctor.key,
    baseSettings,
    overrides,
    conflictSelections,
    doctor.aliases || [],
    parsedRosterSources.casey,
    parsedRosterSources.mch,
  );
  const events = view.events;
  return {
    ...previewSummary(events),
    events: events.map(serializeEvent),
    review: view.reviewItems.map(serializeReviewItem),
    issues: view.issues,
    conflicts: view.conflicts.map(serializeConflict),
    imports: view.imports,
    sources: sourceNames(parsedRosterSources),
    lastParsed: new Date().toISOString(),
  };
}

function rebuildClientPreview() {
  if (!latestPreview) return;
  const doctor = selectedDoctor();
  if (!doctor) return;
  pruneResolvedLatestPreviewIssues();
  const view = buildClientPreviewData(latestPreview);
  renderConflicts(view.conflicts || []);
  renderPreviewGrid(doctor, view);
  renderIssues(view.issues || []);
  void reportPreviewIssues(view.issues || []);
  void reportPreviewConflicts(view.conflicts || []);
  saveCurrentSessionState();
}

function buildClientPreviewData(baseData) {
  const range = deriveDefaultPreviewRange(baseData.events || []);
  const events = buildFilteredPreviewEvents(baseData, settings, range);
  const deletedItems = [];
  const synthesizedIssues = synthesizeIncompleteShiftCodeIssues(baseData);
  const issueKeys = new Set();
  const issues = [
    ...(baseData.issues || []),
    ...synthesizedIssues,
  ].filter((issue) => {
    const key = issueFingerprint(issue?.source, issue?.code || issue?.rawValue, issue?.seniority);
    if (!key) return true;
    if (issueKeys.has(key)) return false;
    issueKeys.add(key);
    return true;
  });
  const hospitals = availableHospitalsForPreview(baseData.events || []);
  if (settings.hospitalFilter === "all" || hospitals.length > availablePreviewHospitals.length) {
    availablePreviewHospitals = hospitals;
  }

  const previewStart = boundedPreviewStart(settings.dateFrom, range.start);
  const previewEnd = boundedPreviewEnd(settings.dateTo, range.end);
  const terms = availablePreviewTerms(baseData.events || []);
  return {
    ...baseData,
    events,
    count: events.length,
    date_range: formatPreviewRange(previewStart, previewEnd) || (events.length ? summarizeEvents(events) : "No events found"),
    previewStart,
    previewEnd,
    terms,
    hospitals: availablePreviewHospitals,
    lastImport: latestImportTimestamp(),
    issues: [
      ...issues.filter(shouldShowPreviewIssue),
      ...deletedItems,
    ],
  };
}

function synthesizeIncompleteShiftCodeIssues(baseData) {
  if (baseData?.derivedFromD1) return [];
  if (!baseData?.review?.some((item) => item.status === "cached" || item.status === "derived")) return [];
  const eventsById = new Map((baseData.events || []).map((event) => [event.id, event]));
  return (baseData.review || [])
    .map((item) => incompleteShiftCodeIssueForReviewItem(item, eventsById.get(item.id)))
    .filter(Boolean);
}

function incompleteShiftCodeIssueForReviewItem(item, event = null) {
  const source = sanitizeIssueSource(item?.source || event?.source);
  const seniority = sanitizeRuleSeniority(item?.seniority || event?.seniority);
  const rawValue = String(item?.rawValue || event?.rawValue || "").trim();
  const normalizedTitle = String(item?.normalizedTitle || item?.suggestedTitle || event?.title || "").trim();
  const code = incompleteShiftCodeFromTitle(source, normalizedTitle) || parserRuleCodeFromRawValue(source, rawValue);
  if (!source || !seniority || !code || !rawValue) return null;
  if (!looksLikeIncompleteShiftCodeTitle(source, normalizedTitle, code)) return null;
  if (isKnownResolvedShiftCodeValue(source, rawValue, normalizedTitle)) return null;
  if (isShiftCodeResolvedByActiveRules({ source, seniority, code, rawValue: code })) return null;
  const id = issueFingerprint(source, code, seniority);
  return {
    id,
    source,
    seniority,
    code,
    startDay: item?.startDay || event?.start?.slice(0, 10) || "",
    rawValue,
    status: "unknown",
    message: `${source} shift code not recognised; using explicit roster time.`,
    resolutionType: "shift_code",
    suggestedTitle: normalizedTitle,
    timeLabel: item?.timeLabel || event?.timeLabel || "",
  };
}

function incompleteShiftCodeFromTitle(source, title) {
  const normalizedSource = sanitizeIssueSource(source);
  const text = String(title || "").trim();
  const prefix = normalizedSource ? `${normalizedSource}:` : "";
  const core = prefix && text.toUpperCase().startsWith(prefix.toUpperCase())
    ? text.slice(prefix.length).trim()
    : text;
  if (normalizedSource === "DDH" && /^[A-Z0-9/]+(?:\s+(?:AM|PM|NIGHT))?$/i.test(core)) {
    return normalizeDdhParserRuleCodeText(core);
  }
  return /^[A-Z0-9/]{2,8}$/.test(core) ? core.toUpperCase() : "";
}

function looksLikeIncompleteShiftCodeTitle(source, title, code) {
  const titleCode = incompleteShiftCodeFromTitle(source, title);
  return Boolean(titleCode && titleCode === String(code || "").trim().toUpperCase());
}

function isSuppressedIssue(issue) {
  const fingerprint = issueFingerprint(issue?.source, issue?.code || issue?.rawValue, issue?.seniority);
  if (!fingerprint) return false;
  return dismissedIssueFingerprints.has(fingerprint) || ignoredIssueFingerprints.has(fingerprint);
}

function buildFilteredPreviewEvents(baseData, filterSettings, defaultRange = deriveDefaultPreviewRange(baseData.events || [])) {
  const events = buildResolvedPreviewEvents(baseData);
  const previewStart = boundedPreviewStart(filterSettings.dateFrom, defaultRange.start);
  const previewEnd = boundedPreviewEnd(filterSettings.dateTo, defaultRange.end);
  const visibleEvents = filterEventsByPreviewRange(events, previewStart, previewEnd)
    .filter((event) => matchesPreviewHospitalFilter(event, filterSettings.hospitalFilter));
  visibleEvents.sort(comparePreviewEvents);
  return visibleEvents;
}

function buildResolvedPreviewEvents(baseData) {
  const activeCustomEventIds = new Set(customEventsForActiveCalendar().map((event) => event.id));
  const baseEvents = new Map(
    (baseData.events || [])
      .filter((event) => !(
        baseData.customEventsMaterialized === true
        && isCustomPreviewEvent(event)
        && !activeCustomEventIds.has(event.id)
      ))
      .map((event) => [event.id, { ...event }]),
  );
  const events = [];
  for (const item of reviewIndex.values()) {
    const event = baseEvents.get(item.id);
    if (!event) continue;
    const override = overrides[item.id] || {};
    const include = typeof override.include === "boolean" ? override.include : item.include;
    if (!include) continue;
    events.push(buildEventOverridePatch(event, item, override));
  }
  const previewCustomEventIds = new Set(
    (baseData.events || [])
      .filter(isCustomPreviewEvent)
      .map((event) => String(event.id || ""))
      .filter(Boolean),
  );
  for (const event of customEventsForActiveCalendar()) {
    if (baseData.customEventsMaterialized === true && previewCustomEventIds.has(event.id)) continue;
    events.push(customEventToPreviewEvent(event));
  }
  const dedupedEvents = latestPreviewEventsByIdentity(events);
  dedupedEvents.sort(comparePreviewEvents);
  return dedupedEvents;
}

function latestPreviewEventsByIdentity(events) {
  const byIdentity = new Map();
  for (const event of events || []) {
    const key = previewEventIdentity(event);
    byIdentity.delete(key);
    byIdentity.set(key, event);
  }
  return [...byIdentity.values()];
}

function previewEventIdentity(event) {
  if (String(event?.source || "").toLowerCase() !== "custom") {
    return `event:${String(event?.id || "")}`;
  }
  return [
    "custom",
    normalizeEmail(event.ownerEmail || activeCalendarEmail()),
    String(event.title || ""),
    String(event.start || ""),
    String(event.end || ""),
    event.allDay === true ? "all-day" : String(event.timeLabel || ""),
    String(event.location || ""),
  ].join("|");
}

function availableHospitalsForPreview(events) {
  const codes = new Set();
  for (const event of events || []) {
    if (event.source) codes.add(displaySourceCode(event.source));
    const titlePrefix = String(event.title || "").match(/^(MMC|DDH|Casey|MCH):/i)?.[1];
    if (titlePrefix) codes.add(displaySourceCode(titlePrefix));
  }
  return [...codes]
    .filter((code) => code === "MMC" || code === "DDH" || code === "Casey" || code === "MCH")
    .sort();
}

function availablePreviewTerms(events) {
  const eventRange = deriveRangeBounds(events || []);
  if (!eventRange.start || !eventRange.end) return [];
  let term = australianTermForDate(parseDateOnly(eventRange.start));
  const lastTerm = australianTermForDate(parseDateOnly(eventRange.end));
  const terms = [];
  while (term.year < lastTerm.year || (term.year === lastTerm.year && term.termNumber <= lastTerm.termNumber)) {
    terms.push({
      label: formatAustralianTermLabel(term),
      value: formatDateKey(term.start),
      year: term.year,
      termNumber: term.termNumber,
    });
    term = nextAustralianTerm(term);
  }
  return terms;
}

function selectedPreviewTermValue(terms, previewStart) {
  if (!terms.length || !previewStart) return "";
  const start = parseDateOnly(previewStart);
  const selectedTerm = australianTermForDate(start);
  const selectedValue = formatDateKey(selectedTerm.start);
  if (terms.some((term) => term.value === selectedValue)) return selectedValue;
  return terms[0]?.value || "";
}

function displaySourceCode(value) {
  const source = String(value || "").trim().toUpperCase();
  if (source === "CASEY") return "Casey";
  if (source === "MMC" || source === "DDH" || source === "MCH") return source;
  return String(value || "").trim();
}

async function buildBrowserIcs(doctor) {
  const events = await buildBrowserExportEvents(doctor, "full");
  if (!events.length) {
    throw new Error("No calendar events were found for the selected doctor.");
  }
  return exportIcs(events, doctor.displayName);
}

async function ensureBasePreviewData(doctor) {
  if (!doctor) throw new Error("Choose a doctor before exporting.");
  if (!parsedRosterSources) {
    const data = await analyzeFilesInBrowser();
    doctorOptions = doctorOptionsForCurrentAccount(data.doctors || []);
  }
  const selected = selectedDoctor();
  if (!latestPreview || selected?.key !== doctor.key) {
    const data = await buildBrowserPreviewData(doctor);
    latestPreview = data;
    indexReviewItems(data.review || []);
  }
  return latestPreview;
}

function normalizeExportRangeState(value = {}) {
  return {
    startDate: String(value.startDate || "").trim(),
    endDate: String(value.endDate || "").trim(),
    allFuture: value.allFuture !== false,
  };
}

function defaultExportRangeState() {
  const doctor = selectedDoctor();
  const rangeSource = latestPreview?.events?.length
    ? deriveRangeBounds(latestPreview.events)
    : doctor && latestPreview
      ? deriveRangeBounds(latestPreview.events || [])
      : { start: "", end: "" };
  const today = formatDateKey(new Date());
  const startDate = rangeSource.start && rangeSource.start > today ? rangeSource.start : today;
  return {
    startDate,
    endDate: rangeSource.end || "",
    allFuture: true,
  };
}

function exportConfigForMode(mode = "full", rangeState = pendingExportRange) {
  if (mode !== "range") return { mode: "full" };
  const normalizedRange = normalizeExportRangeState(rangeState);
  return {
    mode: "range",
    startDate: normalizedRange.startDate,
    endDate: normalizedRange.allFuture ? "" : normalizedRange.endDate,
    allFuture: normalizedRange.allFuture !== false,
  };
}

function buildExportEventsFromBase(baseData, exportConfig = { mode: "full" }) {
  const selectedHospitals = normalizeExportHospitals(exportConfig?.hospitals);
  const events = buildResolvedPreviewEvents(baseData)
    .filter((event) => matchesExportHospitals(event, selectedHospitals));
  if (exportConfig?.mode !== "range") return events;
  const normalizedRange = exportConfigForMode("range", exportConfig);
  return filterEventsByExportRange(events, normalizedRange.startDate, normalizedRange.allFuture ? "" : normalizedRange.endDate);
}

function filterEventsByExportRange(events, startDateKey = "", endDateKey = "") {
  const startDate = startDateKey ? parseDateOnly(startDateKey) : null;
  const endDate = endDateKey ? parseDateOnly(endDateKey) : null;
  return [...events]
    .filter((event) => {
      const eventStart = parseDateOnly(String(event.start || "").slice(0, 10));
      const eventEnd = previewInclusiveEndDate(event, eventStart, parseDateOnly(String(event.end || "").slice(0, 10)));
      if (startDate && eventEnd < startDate) return false;
      if (endDate && eventStart > endDate) return false;
      return true;
    })
    .sort(comparePreviewEvents);
}

async function buildBrowserExportEvents(doctor, exportConfig = { mode: "full" }) {
  const baseData = await ensureBasePreviewData(doctor);
  return buildExportEventsFromBase(baseData, exportConfig);
}

function exportHospitalOptions() {
  return availableHospitalsForPreview(latestPreview?.events || []);
}

function normalizeExportHospitals(value) {
  return Array.isArray(value)
    ? [...new Set(value.map((item) => String(item || "").trim().toUpperCase()).filter(Boolean))]
    : [];
}

function matchesExportHospitals(event, selectedHospitals = []) {
  if (!selectedHospitals.length) return true;
  if (String(event?.source || "").trim().toUpperCase() === "CUSTOM") return true;
  return selectedHospitals.includes(String(event?.source || "").trim().toUpperCase());
}

function renderConflicts(items) {
  if (!items.length) {
    conflictsPanel.classList.add("hidden");
    conflictsList.innerHTML = "";
    return;
  }
  conflictsList.innerHTML = items.map((item) => `
    <article class="issue-card issue-ambiguous">
      <div>
        <strong>${escapeHtml(item.source)} · Week Starting ${escapeHtml(item.weekKey)}</strong>
        <p>Choose which import should overwrite this overlapping week.</p>
      </div>
      <label class="field">
        <span>Preferred import</span>
        <select data-conflict-key="${item.key}">
          ${item.options.map((option) => `<option value="${option.importId}" ${option.importId === item.selectedImportId ? "selected" : ""}>${escapeHtml(option.importName)}${option.addedAt ? ` · ${escapeHtml(formatTimestamp(option.addedAt))}` : ""}</option>`).join("")}
        </select>
      </label>
    </article>
  `).join("");
  conflictsPanel.classList.remove("hidden");
}

function renderIssues(items) {
  if (!items.length) {
    issuesPanel.classList.add("hidden");
    issuesList.innerHTML = "";
    return;
  }

  issuesList.innerHTML = items.map((item) => `
    <article class="issue-card issue-${item.status}" data-review-id="${item.id}" tabindex="0" role="button">
      <div>
        <strong>${formatIssueHeading(item)}</strong>
        <p>${escapeHtml(item.message)}</p>
        ${isShiftCodeIssue(item) ? `<p><button type="button" class="button button-secondary" data-preview-add-shift-code="${escapeHtml(item.id)}">${isCreatorAuthenticated() ? "Edit shift-code rule" : "Resolve shift code"}</button></p>` : `<p><button type="button" class="button button-secondary" data-open-review="${escapeHtml(item.id)}">View event</button></p>`}
      </div>
      <span>${escapeHtml(item.rawValue)}</span>
    </article>
  `).join("");
  issuesPanel.classList.remove("hidden");
}

function isShiftCodeIssue(item) {
  if (item?.resolutionType === "shift_code") return true;
  const message = String(item?.message || "").toLowerCase();
  return item?.status === "unknown" || message.includes("shift code not recognised") || message.includes("shift label not recognised");
}

async function reportPreviewIssues(items) {
  if (latestPreview?.derivedFromD1) return;
  if (!currentUserEmail && !adminViewingEmail) return;
  if (!items.length) return;
  for (const item of items) {
    if (!item?.message || isSuppressedIssue(item)) continue;
    const reviewItem = reviewIndex.get(item.id);
    const issue = {
      source: sanitizeIssueSource(item.source),
      date: String(item.startDay || "").trim(),
      rawValue: String(item.rawValue || "").trim(),
      code: String(item.code || "").trim().toUpperCase(),
      seniority: sanitizeRuleSeniority(item.seniority || reviewItem?.seniority),
      message: String(item.message || "").trim(),
      timeLabel: String(item.timeLabel || reviewItem?.timeLabel || "").trim(),
      suggestedTitle: String(item.suggestedTitle || reviewItem?.suggestedTitle || "").trim(),
      fingerprint: issueFingerprint(item.source, item.code || item.rawValue, item.seniority || reviewItem?.seniority),
    };
    if (!issue.fingerprint) continue;
    const fingerprint = `${activeCalendarOwnerId()}::${issue.fingerprint}`;
    if (reportedIssueFingerprints.has(fingerprint)) continue;
    reportedIssueFingerprints.add(fingerprint);
    await reportAccountError(issue, item.id || issue.fingerprint);
  }
}

async function reportPreviewConflicts(items) {
  if (!currentUserEmail && !adminViewingEmail) return;
  if (!items.length) return;
  for (const item of items) {
    const source = sanitizeIssueSource(item.source);
    const rawValue = String(item.key || `${item.source || ""}:${item.weekKey || ""}`).trim();
    if (!source || !rawValue) continue;
    const optionNames = (item.options || []).map((option) => option.importName).filter(Boolean);
    const issue = {
      source,
      date: String(item.weekKey || "").trim(),
      rawValue,
      message: `Multiple roster files overlap for ${source} week starting ${item.weekKey || "unknown"}: ${optionNames.join(" vs ") || "conflicting imports"}.`,
      timeLabel: "",
      suggestedTitle: item.selectedImportId ? `Selected source: ${selectedConflictImportName(item)}` : "",
      fingerprint: issueFingerprint(source, rawValue),
    };
    if (!issue.fingerprint) continue;
    const fingerprint = `${activeCalendarOwnerId()}::${issue.fingerprint}`;
    if (reportedIssueFingerprints.has(fingerprint)) continue;
    reportedIssueFingerprints.add(fingerprint);
    await reportAccountError(issue, issue.fingerprint);
  }
}

function selectedConflictImportName(item) {
  return (item.options || []).find((option) => option.importId === item.selectedImportId)?.importName || item.selectedImportId || "";
}

function indexReviewItems(items) {
  reviewIndex = new Map(items.map((item) => [item.id, item]));
}

function renderPreviewGrid(doctor, data) {
  const events = data.events || [];
  currentPreviewEvents = new Map(events.map((event) => [event.id, event]));
  updatePreviewDisplayExample(previewDisplayDraft || settings);
  const days = buildPreviewDays(events, data.previewStart, data.previewEnd);
  document.body.classList.add("has-calendar-preview");
  if (!days.length) {
    preview.innerHTML = `
      ${renderPreviewHeader(doctor, data)}
      <div class="preview-empty">No events match the current settings.</div>
    `;
    preview.classList.remove("hidden");
    previewSection.classList.remove("hidden");
    syncMobileChrome();
    return;
  }
  const weeks = chunkWeeks(days);
  const termSections = buildTermSections(weeks);

  preview.innerHTML = `
    ${renderPreviewHeader(doctor, data)}
    ${termSections}
  `;
  preview.classList.remove("hidden");
  previewSection.classList.remove("hidden");
  syncMobileChrome();
  if (pendingPreviewSnapToToday) {
    pendingPreviewSnapToToday = false;
    requestAnimationFrame(() => snapPreviewToCurrentMonth(false));
  }
}

function renderPreviewHeader(doctor, data) {
  const hospitalSelector = renderPreviewHospitalSelector(data.hospitals || []);
  return `
    <div class="preview-head">
      ${renderPreviewDoctorControl(doctor)}
      <div class="preview-toolbar">
        ${renderPreviewRangeControls(data.previewStart, data.previewEnd)}
        ${hospitalSelector || `<span class="preview-toolbar-spacer" aria-hidden="true"></span>`}
        <span class="preview-event-count">${data.count} events</span>
        ${canReturnToCreator()
          ? `<button type="button" class="button button-secondary preview-back-button" data-preview-back-to-creator>Back to creator</button>`
          : ""}
        <button type="button" class="button button-secondary preview-logout-button" data-preview-logout>Log out</button>
      </div>
    </div>
  `;
}

function renderPreviewDoctorControl(doctor) {
  const pickerOptions = doctorPickerOptions();
  if (pickerOptions.length > 1) {
    return `
      <label class="preview-doctor-control">
        <span>Doctor</span>
        <select data-preview-doctor-select>
          ${pickerOptions.map((option) => `
            <option value="${escapeHtml(option.key)}" ${option.key === doctor.key ? "selected" : ""}>
              ${escapeHtml(option.displayName)}
            </option>
          `).join("")}
        </select>
      </label>
    `;
  }

  return `
    <div class="preview-doctor-control">
      <span>Doctor</span>
      <strong>${escapeHtml(doctor?.displayName || currentAccount().realName || "Selected doctor")}</strong>
    </div>
  `;
}

function renderPreviewHospitalSelector(hospitals) {
  if (!hospitals || hospitals.length < 2) return "";
  return `
    <label class="preview-hospital-filter">
      <span>Hospital</span>
      <select data-preview-hospital-filter>
        <option value="all" ${settings.hospitalFilter === "all" ? "selected" : ""}>All hospitals</option>
        ${hospitals.map((code) => {
          const value = code.toLowerCase();
          return `<option value="${value}" ${settings.hospitalFilter === value ? "selected" : ""}>${escapeHtml(code)}</option>`;
        }).join("")}
      </select>
    </label>
  `;
}

function renderPreviewRangeControls(start, end) {
  const fromValue = start || "";
  const toValue = end || "";
  const terms = latestPreview ? availablePreviewTerms(latestPreview.events || []) : [];
  const selectedTermValue = selectedPreviewTermValue(terms, fromValue);
  return `
    <div class="preview-range-controls">
      <span class="preview-range-label">From</span>
      <button type="button" class="preview-range-button" data-range-trigger="from">
        ${escapeHtml(fromValue ? formatDate(fromValue) : "Set date")}
      </button>
      <input class="preview-range-input" type="date" value="${escapeHtml(fromValue)}" data-range-input="from" tabindex="-1" aria-hidden="true">
      <span class="preview-range-label">To</span>
      <button type="button" class="preview-range-button" data-range-trigger="to">
        ${escapeHtml(toValue ? formatDate(toValue) : "Set date")}
      </button>
      <input class="preview-range-input" type="date" value="${escapeHtml(toValue)}" data-range-input="to" tabindex="-1" aria-hidden="true">
      <button type="button" class="button button-secondary preview-today-button" data-range-today>Today</button>
      ${terms.length ? `
        <label class="preview-term-start-control">
          <span class="preview-range-label">From</span>
          <select class="preview-range-button preview-term-start-select" data-preview-term-start>
            ${terms.map((term) => `
              <option value="${escapeHtml(term.value)}" ${term.value === selectedTermValue ? "selected" : ""}>${escapeHtml(term.label)}</option>
            `).join("")}
          </select>
        </label>
      ` : ""}
    </div>
  `;
}

function buildPreviewDays(events, explicitStart = "", explicitEnd = "") {
  const eventMap = new Map();
  let firstDay = explicitStart ? parseDateOnly(explicitStart) : null;
  let lastDay = explicitEnd ? parseDateOnly(explicitEnd) : null;

  for (const event of events) {
    const startDate = parseDateOnly(event.start);
    const endDate = parseDateOnly(event.end);
    const inclusiveEnd = previewInclusiveEndDate(event, startDate, endDate);
    if (!firstDay || startDate < firstDay) firstDay = startDate;
    if (!lastDay || inclusiveEnd > lastDay) lastDay = inclusiveEnd;

    let cursor = new Date(startDate);
    while (cursor <= inclusiveEnd) {
      const key = formatDateKey(cursor);
      if (!eventMap.has(key)) eventMap.set(key, []);
      eventMap.get(key).push(event);
      cursor = addDays(cursor, 1);
    }
  }

  if (!firstDay || !lastDay) return [];
  const startMonday = mondayFor(firstDay);
  const endSunday = addDays(mondayFor(lastDay), 6);
  const days = [];
  for (let cursor = new Date(startMonday); cursor <= endSunday; cursor = addDays(cursor, 1)) {
    const key = formatDateKey(cursor);
    days.push({
      date: new Date(cursor),
      events: eventMap.get(key) || [],
    });
  }
  return days;
}

function chunkWeeks(days) {
  const weeks = [];
  for (let index = 0; index < days.length; index += 7) {
    weeks.push(days.slice(index, index + 7));
  }
  return weeks;
}

function renderDayCell(day) {
  const dayKey = formatDateKey(day.date);
  const dayIndex = (day.date.getDay() + 6) % 7;
  const weekendClass = dayIndex >= 5 ? " is-weekend" : "";
  const cards = day.events.length
    ? day.events.map((event) => renderPreviewChip(event, dayKey)).join("")
    : `<div class="preview-chip preview-chip-empty"></div>`;
  const currentDayClass = isCurrentDay(day.date) ? " is-current-day" : "";
  return `
    <div class="preview-cell${currentDayClass}${weekendClass}" data-day-index="${dayIndex}" data-add-date="${dayKey}">
      <div class="preview-date">${day.date.getDate()}</div>
      <div class="preview-stack">${cards}</div>
    </div>
  `;
}

function renderPreviewChip(event, dayKey) {
  const lines = [];
  const startKey = event.start.slice(0, 10);
  const mobileLabel = mobilePreviewTitleParts(event);
  const continuous = continuousEventDisplay(event, dayKey);
  if (settings.showNormalizedTitles) {
    const marker = event.isEditedImport ? '<span class="preview-chip-marker" aria-label="Imported event edited">*</span>' : "";
    lines.push(`<strong class="preview-chip-title-full">${escapeHtml(event.title)}${marker}</strong>`);
    lines.push(`
      <strong class="preview-chip-title-mobile">
        ${mobileLabel.prefix ? `<span class="preview-chip-source">${escapeHtml(mobileLabel.prefix)}</span>` : ""}
        <span>${escapeHtml(mobileLabel.label)}</span>${marker}
      </strong>
    `);
  }
  if (settings.showRawValues) {
    lines.push(`<span class="preview-chip-raw">${escapeHtml(event.rawValue)}</span>`);
  }
  const meta = [];
  if (!event.allDay && settings.showTimes && startKey === dayKey && event.timeLabel) meta.push(event.timeLabel);
  const metaMarkup = meta.length ? `<span class="preview-chip-meta">${escapeHtml(meta.join(" · "))}</span>` : "";
  const style = continuous.span > 1 ? ` style="--continuous-span: ${continuous.span};"` : "";
  return `<button type="button" class="preview-chip preview-chip-${eventTone(event)}${continuous.className}"${style} data-review-id="${event.id}" data-review-date="${dayKey}">${lines.join("")}${metaMarkup}</button>`;
}

function isMobileContinuousEvent(event) {
  if (!event?.allDay) return false;
  const start = parseDateOnly(event.start);
  const end = previewInclusiveEndDate(event, start, parseDateOnly(event.end));
  return end > start;
}

function continuousEventDisplay(event, dayKey) {
  if (!isMobileContinuousEvent(event)) return { className: "", span: 1 };
  const start = parseDateOnly(event.start);
  const end = previewInclusiveEndDate(event, start, parseDateOnly(event.end));
  const day = parseDateOnly(dayKey);
  const dayIndex = (day.getDay() + 6) % 7;
  const segmentStart = sameDateOnly(day, start) || dayIndex === 0;
  if (!segmentStart) return { className: " is-continuous-hidden", span: 1 };
  const weekStart = mondayFor(day);
  const weekEnd = addDays(weekStart, 6);
  const spanStart = start < weekStart ? weekStart : start;
  const spanEnd = end > weekEnd ? weekEnd : end;
  const span = Math.max(1, daysBetween(spanStart, spanEnd) + 1);
  const segmentEnd = sameDateOnly(spanEnd, end) || sameDateOnly(spanEnd, weekEnd);
  return {
    className: ` is-continuous is-continuous-start${segmentEnd ? " is-continuous-end" : ""}`,
    span,
  };
}

function mobilePreviewTitleParts(event) {
  const originalTitle = String(event?.title || "").trim();
  const prefixMatch = originalTitle.match(/^([A-Z]{1,6}):\s*(.+)$/);
  let prefix = prefixMatch ? `${prefixMatch[1]}:` : "";
  let label = prefixMatch ? prefixMatch[2] : originalTitle;
  const lower = label.toLowerCase();
  if (lower.includes("sick")) {
    label = "Sick";
    prefix = "";
  } else if (lower.includes("annual leave")) {
    label = "AL";
    prefix = "";
  } else if (lower.includes("conference leave")) {
    label = "CL";
    prefix = "";
  } else {
    label = label
      .replace(/\bClinical Support\b/gi, "CS")
      .replace(/\bFast Track\b/gi, "Fast")
      .replace(/\bFAST\b/g, "Fast")
      .replace(/\bAM\b/gi, "")
      .replace(/\bPM\b/gi, "")
      .replace(/\bshift\b/gi, "")
      .replace(/\s+/g, " ")
      .trim();
  }
  return { prefix, label: label || originalTitle || "Event" };
}

function renderMobileWeekSpans(week) {
  const weekStart = week[0]?.date;
  const weekEnd = week.at(-1)?.date;
  if (!weekStart || !weekEnd) return "";
  const unique = new Map();
  week.forEach((day) => {
    day.events.forEach((event) => {
      if (isMobileContinuousEvent(event)) unique.set(event.id, event);
    });
  });
  if (!unique.size) return "";
  const spans = [...unique.values()].map((event) => {
    const eventStart = parseDateOnly(event.start);
    const eventEnd = previewInclusiveEndDate(event, eventStart, parseDateOnly(event.end));
    const spanStart = eventStart < weekStart ? weekStart : eventStart;
    const spanEnd = eventEnd > weekEnd ? weekEnd : eventEnd;
    const startColumn = daysBetween(weekStart, spanStart) + 1;
    const endColumn = daysBetween(weekStart, spanEnd) + 2;
    const label = mobilePreviewTitleParts(event);
    return `
      <button type="button" class="preview-mobile-span preview-chip-${eventTone(event)}" style="grid-column: ${startColumn} / ${endColumn};" data-review-id="${event.id}" data-review-date="${formatDateKey(spanStart)}">
        ${label.prefix ? `<span class="preview-chip-source">${escapeHtml(label.prefix)}</span>` : ""}
        <span>${escapeHtml(label.label)}</span>
      </button>
    `;
  }).join("");
  return `<div class="preview-mobile-span-row">${spans}</div>`;
}

function sameDateOnly(left, right) {
  return Boolean(left && right)
    && left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate();
}

function daysBetween(start, end) {
  const left = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const right = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((right.getTime() - left.getTime()) / 86400000);
}

function eventTone(event) {
  const text = `${event.title || ""} ${event.rawValue || ""}`.toLowerCase();
  if (text.includes("annual") || text.includes("conference") || text.includes("leave")) return "leave";
  if (text.includes("phnw")) return "phnw";
  if (text.includes("clinical support") || /\bcs\b/.test(text) || /\bcso\b/.test(text)) return "cs";
  if (text.includes("night")) return "night";
  if (text.includes("pm") || text.includes("orange")) return "evening";
  if (text.includes("custom") || event.isCustom) return "custom";
  return "day";
}

function buildTermSections(weeks) {
  if (!weeks.length) return "";
  const sections = [];
  let current = null;

  weeks.forEach((week, index) => {
    const monday = week[0]?.date;
    const term = detectAustralianTerm(monday);
    if (!current || current.label !== term.label) {
      current = {
        label: term.label,
        weeks: [],
      };
      sections.push(current);
    }
    current.weeks.push({ week, index });
  });

  return sections.map((section) => renderTermSection(section)).join("");
}

function renderTermSection(section) {
  const headerCells = DAY_NAMES.map((day, index) => `<div class="preview-day-name${index >= 5 ? " is-weekend" : ""}" data-day-index="${index}" data-short-day="${escapeHtml(day[0])}">${day}</div>`).join("");
  const bodyRows = [];
  let lastMonthKey = "";
  const firstMonday = section.weeks[0]?.week?.[0]?.date;
  const lastSunday = section.weeks.at(-1)?.week?.at(-1)?.date;

  section.weeks.forEach(({ week, index }) => {
    const monday = week[0]?.date;
    const monthKey = monday ? `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, "0")}` : "";
    const isMonthStartWeek = Boolean(monthKey && monthKey !== lastMonthKey);
    if (monthKey && monthKey !== lastMonthKey) {
      bodyRows.push(`
        <div class="preview-month-row" data-month-key="${monthKey}">
          <span>${formatMonth(monday)}</span>
          ${canUseRosterInsights() ? `<button type="button" class="button button-secondary preview-month-button" data-insight-when="${formatDateKey(firstMonday)}" data-insight-when-end="${formatDateKey(lastSunday)}">When am I working with…?</button>` : ""}
        </div>
      `);
      lastMonthKey = monthKey;
    }
    bodyRows.push(`
      <div class="preview-week-label" ${isMonthStartWeek ? `data-month-start="${monthKey}"` : ""}>
        <strong>Week ${index + 1}</strong>
        <span>starting</span>
        <time datetime="${formatDateKey(monday)}">${formatLongDate(monday)}</time>
      </div>
      ${week.map((day) => renderDayCell(day)).join("")}
    `);
  });

  return `
    <section class="preview-term">
      <div class="preview-term-header">
        <div class="preview-term-title">${escapeHtml(section.label)}</div>
      </div>
      <div class="preview-grid">
        <div class="preview-week-label preview-week-label-head">Week</div>
        ${headerCells}
        ${bodyRows.join("")}
      </div>
    </section>
  `;
}

async function openWhoInsight(termStart, termEnd) {
  if (!canUseRosterInsights()) return;
  const date = defaultInsightDate(termStart, termEnd);
  insightsState = {
    mode: "who",
    termStart,
    termEnd,
    date,
  };
  await renderInsightsModal();
}

async function openWhoInsightForDate(date) {
  if (!canUseRosterInsights()) return;
  const selectedDate = String(date || "").slice(0, 10);
  if (!selectedDate) return;
  const range = currentCalendarInsightDateRange();
  insightsState = {
    mode: "who",
    termStart: range.start || selectedDate,
    termEnd: range.end || selectedDate,
    date: selectedDate,
  };
  await renderInsightsModal();
}

async function openWhenInsight(termStart, termEnd) {
  if (!canUseRosterInsights()) return;
  const range = currentCalendarInsightDateRange();
  const fromDate = formatDateKey(new Date());
  const toDate = range.end || termEnd || fromDate;
  insightsState = {
    mode: "when",
    termStart: range.start || termStart || fromDate,
    termEnd: toDate,
    fromDate,
    hospitalFilters: [],
    comparisonDoctorKey: "",
  };
  await renderInsightsModal();
}

async function openWhenInsightForDoctor(doctorKey) {
  if (!canUseRosterInsights()) return;
  const normalizedKey = normalizeRosterName(doctorKey);
  if (!normalizedKey) return;
  const range = currentCalendarInsightDateRange();
  const fromDate = formatDateKey(new Date());
  insightsState = {
    mode: "when",
    termStart: range.start || fromDate,
    termEnd: range.end || fromDate,
    fromDate,
    hospitalFilters: [],
    comparisonDoctorKey: normalizedKey,
  };
  await renderInsightsModal();
}

function closeInsightsModal() {
  insightsState = null;
  insightsModal.classList.add("hidden");
  insightsModal.setAttribute("aria-hidden", "true");
  insightsModalBody.innerHTML = "";
}

async function renderInsightsModal() {
  if (!insightsState) return;
  if (insightsState.mode === "who") {
    await renderWhoInsight();
  } else if (insightsState.mode === "when") {
    await renderWhenInsight();
  }
  insightsModal.classList.remove("hidden");
  insightsModal.setAttribute("aria-hidden", "false");
}

async function ensureInsightRosterAnalysis() {
  if (parsedRosterSources && (parsedRosterSources.mmc?.length || parsedRosterSources.ddh?.length || parsedRosterSources.casey?.length || parsedRosterSources.mch?.length)) return;
  if (hydrateInsightCacheFromSnapshot()) return;
  await hydrateInsightCacheFromServer();
}

async function hydrateInsightCacheFromServer() {
  if (!cloudAvailable || !latestPreview) return false;
  const range = deriveRangeBounds(buildResolvedPreviewEvents(latestPreview || { events: [] }).filter(isRosterShiftEvent));
  const startDate = range.start || latestPreview.events?.[0]?.start?.slice(0, 10) || "";
  const endDate = range.end || startDate;
  if (!startDate || !endDate) return false;
  try {
    const requestEmail = adminViewingEmail ? authUserEmail || currentUserEmail : currentUserEmail;
    const requestPassword = adminViewingEmail ? authUserPassword || currentUserPassword : currentUserPassword;
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryRosterInsights",
        email: requestEmail,
        password: requestPassword,
        startDate,
        endDate,
      }),
    });
    const data = await readJsonResponse(response, "Could not load roster insights.");
    if (!data.ok || data.unavailable || !Array.isArray(data.coworkers) || !data.coworkers.length) return false;
    const cache = new Map();
    const optionMap = new Map();
    for (const row of data.coworkers) {
      const key = normalizeRosterName(row.doctorKey || "");
      if (!key || !row.event) continue;
      if (!cache.has(key)) cache.set(key, []);
      cache.get(key).push(serializeEvent(row.event));
      if (!optionMap.has(key)) {
        optionMap.set(key, {
          key,
          displayName: row.displayName || key,
          sourceTypes: row.sourceType ? [String(row.sourceType).toLowerCase()] : [],
          aliases: row.sourceType ? [{ sourceType: String(row.sourceType).toLowerCase(), key, displayName: row.displayName || key }] : [],
        });
      }
    }
    if (!cache.size) return false;
    doctorAnalysisCacheKey = currentInsightCacheKey();
    doctorAnalysisCache = cache;
    insightDoctorOptionsCache = [...optionMap.values()].sort((left, right) => left.displayName.localeCompare(right.displayName));
    insightDoctorRoleCache = new Map();
    cacheCurrentInsightSnapshot();
    return true;
  } catch {
    return false;
  }
}

function selectedInsightDoctorKeys() {
  const doctor = selectedDoctor();
  return [...new Set([
    doctor?.key,
    ...(Array.isArray(doctor?.aliases) ? doctor.aliases.map((alias) => alias.key) : []),
  ].map(normalizeRosterName).filter(Boolean))];
}

async function fetchRosterInsightRows({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], doctorKeys = [], overlapDoctorKeys = [] } = {}) {
  if (!cloudAvailable || !startDate) return { ok: false, unavailable: true, rows: [] };
  const cacheKey = rosterInsightCacheKey({ startDate, endDate, sourceTypes, excludeDoctorKeys, doctorKeys, overlapDoctorKeys });
  if (visibleInsightWarmCache.has(cacheKey)) {
    return { ok: true, rows: visibleInsightWarmCache.get(cacheKey), elapsedMs: 0, cached: true };
  }
  const startedAt = performance.now();
  try {
    const requestEmail = adminViewingEmail ? authUserEmail || currentUserEmail : currentUserEmail;
    const requestPassword = adminViewingEmail ? authUserPassword || currentUserPassword : currentUserPassword;
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryRosterInsights",
        email: requestEmail,
        password: requestPassword,
        startDate,
        endDate,
        sourceTypes,
        excludeDoctorKeys,
        doctorKeys,
        overlapDoctorKeys,
      }),
    });
    const data = await readJsonResponse(response, "Could not load roster insights.");
    const elapsedMs = Math.round(performance.now() - startedAt);
    if (!data.ok || data.unavailable || !Array.isArray(data.coworkers)) {
      console.warn("Roster insight SQL lookup unavailable", { elapsedMs, startDate, endDate, sourceTypes, doctorKeys, overlapDoctorKeys });
      return { ok: false, unavailable: true, rows: [], elapsedMs };
    }
    if (elapsedMs > 1000) console.warn("Roster insight SQL lookup was slow", { elapsedMs, queryMs: data.queryMs, startDate, endDate, sourceTypes, doctorKeys, overlapDoctorKeys });
    visibleInsightWarmCache.set(cacheKey, data.coworkers);
    return { ok: true, rows: data.coworkers, elapsedMs, queryMs: data.queryMs };
  } catch (error) {
    const elapsedMs = Math.round(performance.now() - startedAt);
    console.warn("Roster insight SQL lookup failed", { elapsedMs, startDate, endDate, sourceTypes, doctorKeys, overlapDoctorKeys, error });
    return { ok: false, unavailable: true, rows: [], elapsedMs };
  }
}

async function fetchRosterOverlapDoctors({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], overlapDoctorKeys = [] } = {}) {
  if (!cloudAvailable || !startDate || !overlapDoctorKeys.length) return { ok: false, unavailable: true, doctors: [] };
  const cacheKey = rosterOverlapDoctorCacheKey({ startDate, endDate, sourceTypes, excludeDoctorKeys, overlapDoctorKeys });
  if (visibleInsightWarmCache.has(cacheKey)) {
    return { ok: true, doctors: visibleInsightWarmCache.get(cacheKey), elapsedMs: 0, cached: true };
  }
  const startedAt = performance.now();
  try {
    const requestEmail = adminViewingEmail ? authUserEmail || currentUserEmail : currentUserEmail;
    const requestPassword = adminViewingEmail ? authUserPassword || currentUserPassword : currentUserPassword;
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryRosterOverlapDoctors",
        email: requestEmail,
        password: requestPassword,
        startDate,
        endDate,
        sourceTypes,
        excludeDoctorKeys,
        overlapDoctorKeys,
      }),
    });
    const data = await readJsonResponse(response, "Could not load roster overlap doctors.");
    const elapsedMs = Math.round(performance.now() - startedAt);
    if (!data.ok || data.unavailable || !Array.isArray(data.doctors)) return { ok: false, unavailable: true, doctors: [], elapsedMs };
    if (elapsedMs > 1000) console.warn("Roster overlap doctor SQL lookup was slow", { elapsedMs, queryMs: data.queryMs, startDate, endDate, sourceTypes, overlapDoctorKeys });
    visibleInsightWarmCache.set(cacheKey, data.doctors);
    return { ok: true, doctors: data.doctors, elapsedMs, queryMs: data.queryMs };
  } catch (error) {
    const elapsedMs = Math.round(performance.now() - startedAt);
    console.warn("Roster overlap doctor SQL lookup failed", { elapsedMs, startDate, endDate, sourceTypes, overlapDoctorKeys, error });
    return { ok: false, unavailable: true, doctors: [], elapsedMs };
  }
}

function renderRosterInsightUnavailable() {
  return `<article class="issue-card"><p>Roster insight data is unavailable right now. Please try again shortly.</p></article>`;
}

function insightRowsToDoctorOptions(rows) {
  const doctors = new Map();
  for (const row of rows || []) {
    const key = normalizeRosterName(row.doctorKey || "");
    const sourceType = String(row.sourceType || "").toLowerCase();
    if (!key || !row.event) continue;
    const existing = doctors.get(key) || {
      key,
      displayName: row.displayName || key,
      sourceTypes: [],
      aliases: [],
    };
    if (sourceType && !existing.sourceTypes.includes(sourceType)) existing.sourceTypes.push(sourceType);
    if (sourceType && !existing.aliases.some((alias) => alias.sourceType === sourceType && alias.key === key)) {
      existing.aliases.push({ sourceType, key, displayName: row.displayName || key });
    }
    doctors.set(key, existing);
  }
  return [...doctors.values()].sort((left, right) => left.displayName.localeCompare(right.displayName));
}

function insightRowsToEventsByDoctor(rows) {
  const events = new Map();
  for (const row of rows || []) {
    const key = normalizeRosterName(row.doctorKey || "");
    if (!key || !row.event) continue;
    if (!events.has(key)) events.set(key, []);
    events.get(key).push(serializeEvent(row.event));
  }
  return events;
}

function insightWarmBaseKey() {
  return [
    activeWorkspaceOwnerKey(),
    selectedInsightDoctorKeys().join(","),
    currentCalendarRevision,
  ].join("|");
}

function stableInsightList(values = []) {
  return [...new Set((values || []).map((item) => String(item || "").trim()).filter(Boolean))]
    .sort()
    .join(",");
}

function rosterInsightCacheKey({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], doctorKeys = [], overlapDoctorKeys = [] } = {}) {
  return [
    "rows",
    insightWarmBaseKey(),
    startDate || "",
    endDate || startDate || "",
    stableInsightList(sourceTypes.map((item) => item.toLowerCase())),
    stableInsightList(excludeDoctorKeys.map(normalizeRosterName)),
    stableInsightList(doctorKeys.map(normalizeRosterName)),
    stableInsightList(overlapDoctorKeys.map(normalizeRosterName)),
  ].join("|");
}

function rosterOverlapDoctorCacheKey({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], overlapDoctorKeys = [] } = {}) {
  return [
    "overlapDoctors",
    insightWarmBaseKey(),
    startDate || "",
    endDate || startDate || "",
    stableInsightList(sourceTypes.map((item) => item.toLowerCase())),
    stableInsightList(excludeDoctorKeys.map(normalizeRosterName)),
    stableInsightList(overlapDoctorKeys.map(normalizeRosterName)),
  ].join("|");
}

function resetVisibleInsightWarmCache() {
  visibleInsightWarmCache = new Map();
  visibleInsightWarmKey = "";
}

async function renderWhoInsight() {
  const date = insightsState.date;
  const mine = selectedDoctorEventsForInsights(date, date).filter(isRosterShiftEvent).filter((event) => eventRosterDateKey(event) === date);
  const activeSources = new Set(mine.map(eventSourceCode).filter(Boolean));
  let coworkers = [];
  const serverResult = mine.length ? await fetchRosterInsightRows({
    startDate: date,
    endDate: date,
    excludeDoctorKeys: selectedInsightDoctorKeys(),
  }) : { ok: true, rows: [] };
  if (serverResult.ok) {
    const serverRows = serverResult.rows;
    const serverDoctors = insightRowsToDoctorOptions(serverRows);
    const serverEvents = insightRowsToEventsByDoctor(serverRows);
    coworkers = serverDoctors
      .map((doctor) => ({
        doctor,
        events: (serverEvents.get(doctor.key) || [])
          .filter(isRosterShiftEvent)
          .filter((event) => eventRosterDateKey(event) === date)
          .filter((event) => !activeSources.size || activeSources.has(eventSourceCode(event))),
      }))
      .filter((entry) => entry.events.length)
      .flatMap((entry) => buildWhoAssignments(entry.doctor, entry.events));
  }
  const teamAssignments = [
    ...buildSelectedWhoAssignments(mine),
    ...coworkers,
  ];
  const grouped = groupWhoAssignments(teamAssignments);

  insightsModalTitle.textContent = "Who";
  insightsModalSubtitle.textContent = "Doctors working on the same date as the selected calendar.";
  insightsModalBody.innerHTML = `
    <div class="insights-controls">
      <label class="field">
        <span>Date</span>
        <input type="date" value="${escapeHtml(date)}" min="${escapeHtml(insightsState.termStart)}" max="${escapeHtml(insightsState.termEnd)}" data-insights-who-date>
      </label>
    </div>
    <div class="issue-card">
      <strong>${escapeHtml(selectedDoctor()?.displayName || "Selected doctor")}</strong>
      <p>${mine.length ? escapeHtml(renderInsightShiftSummary(mine)) : "No rostered shifts for this date in the current calendar view."}</p>
    </div>
    ${!serverResult.ok
      ? renderRosterInsightUnavailable()
      : teamAssignments.length
      ? renderWhoGroups(grouped)
      : `<article class="issue-card"><p>No other doctors are rostered on this date.</p></article>`}
  `;
}

async function renderWhenInsight() {
  const hospitalFilters = Array.isArray(insightsState.hospitalFilters) ? insightsState.hospitalFilters : [];
  const fromDate = insightsState.fromDate || formatDateKey(new Date());
  const toDate = insightsState.termEnd || currentCalendarInsightDateRange().end || fromDate;
  const doctorResult = await fetchRosterOverlapDoctors({
    startDate: fromDate,
    endDate: toDate,
    sourceTypes: hospitalFilters.map((item) => item.toLowerCase()),
    excludeDoctorKeys: selectedInsightDoctorKeys(),
    overlapDoctorKeys: selectedInsightDoctorKeys(),
  });
  if (!doctorResult.ok) {
    insightsModalTitle.textContent = "When am I working with…?";
    insightsModalSubtitle.textContent = "Find future dates where both doctors are working from the selected date.";
    insightsModalBody.innerHTML = renderRosterInsightUnavailable();
    return;
  }
  const options = prioritizeDoctorOptions(insightRowsToDoctorOptions(doctorResult.doctors.map((doctor) => ({
    doctorKey: doctor.doctorKey,
    displayName: doctor.displayName,
    sourceType: doctor.sourceType,
    event: {},
  }))));
  const selectedKey = options.some((doctor) => doctor.key === insightsState.comparisonDoctorKey)
    ? insightsState.comparisonDoctorKey
    : options[0]?.key || "";
  insightsState.comparisonDoctorKey = selectedKey;
  if (!selectedKey) {
    renderWhenInsightResult({
      options,
      selectedComparison: null,
      mine: selectedDoctorEventsForInsights(fromDate, toDate, hospitalFilters).filter(isRosterShiftEvent),
      theirs: [],
      fromDate,
      toDate,
      hospitalFilters,
      hospitalOptions: [],
    });
    return;
  }
  const serverResult = await fetchRosterInsightRows({
    startDate: fromDate,
    endDate: toDate,
    sourceTypes: hospitalFilters.map((item) => item.toLowerCase()),
    doctorKeys: [selectedKey],
    overlapDoctorKeys: [],
  });
  if (serverResult.ok) {
    const serverRows = serverResult.rows;
    const selectedComparison = options.find((doctor) => doctor.key === selectedKey) || null;
    const serverEvents = insightRowsToEventsByDoctor(serverRows);
    const mine = selectedDoctorEventsForInsights(fromDate, toDate, hospitalFilters).filter(isRosterShiftEvent);
    const theirs = selectedComparison ? (serverEvents.get(selectedComparison.key) || []).filter(isRosterShiftEvent) : [];
    const hospitalOptions = availableHospitalsFromInsightEvents([...mine, ...[...serverEvents.values()].flat()]);
    renderWhenInsightResult({ options, selectedComparison, mine, theirs, fromDate, toDate, hospitalFilters, hospitalOptions });
    return;
  }
  insightsModalTitle.textContent = "When am I working with…?";
  insightsModalSubtitle.textContent = "Find future dates where both doctors are working from the selected date.";
  insightsModalBody.innerHTML = renderRosterInsightUnavailable();
}

function renderWhenInsightResult({ options, selectedComparison, mine, theirs, fromDate, toDate, hospitalFilters, hospitalOptions = null }) {
  const overlaps = buildOverlapDays(mine, theirs);
  const nextOverlapDate = chooseNextOverlapDate(overlaps);
  const availableHospitalOptions = hospitalOptions || availableHospitalsForInsightRange(fromDate, toDate);

  insightsModalTitle.textContent = "When am I working with…?";
  insightsModalSubtitle.textContent = "Find future dates where both doctors are working from the selected date.";
  insightsModalBody.innerHTML = `
    <div class="insights-controls">
      <label class="field">
        <span>Doctor</span>
        <select data-insights-when-doctor>
          ${options.map((doctor) => `
            <option value="${escapeHtml(doctor.key)}" ${doctor.key === selectedComparison?.key ? "selected" : ""}>${escapeHtml(doctor.displayName)}</option>
          `).join("")}
        </select>
      </label>
      <fieldset class="settings-group insights-hospital-group">
        <legend>Options</legend>
        <label class="field">
          <span>From</span>
          <input type="date" value="${escapeHtml(fromDate)}" min="${escapeHtml(insightsState.termStart || "")}" max="${escapeHtml(toDate)}" data-insights-when-from>
        </label>
        <p class="status">Leave all unticked to search every hospital.</p>
        <div class="toggle-list">
          ${availableHospitalOptions.map((hospital) => `
            <label class="toggle">
              <input type="checkbox" value="${escapeHtml(hospital)}" data-insights-when-hospital ${hospitalFilters.includes(hospital) ? "checked" : ""}>
              ${escapeHtml(hospital)}
            </label>
          `).join("")}
        </div>
      </fieldset>
    </div>
    ${selectedComparison
      ? overlaps.length
        ? overlaps.map((entry) => `
          <article class="issue-card${entry.date === nextOverlapDate ? " is-next-overlap" : ""}" ${entry.date === nextOverlapDate ? 'data-insight-next="true"' : ""}>
            ${renderInsightDateButton(entry.date, "data-insights-who-on-date")}
            <p><strong>Hospital:</strong> ${escapeHtml(entry.hospital || "Unknown")}</p>
            <p><strong>${escapeHtml(selectedDoctor()?.displayName || "Selected doctor")}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.mine))}</p>
            <p><strong>${escapeHtml(selectedComparison.displayName)}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.theirs))}</p>
          </article>
        `).join("")
        : `<article class="issue-card"><p>No overlapping working days were found from ${escapeHtml(formatDate(fromDate))}.</p></article>`
      : `<article class="issue-card"><p>No comparison doctors are available in these roster files.</p></article>`}
  `;
  const nextCard = insightsModalBody.querySelector("[data-insight-next='true']");
  if (nextCard) {
    requestAnimationFrame(() => nextCard.scrollIntoView({ block: "nearest", behavior: "smooth" }));
  }
}

function comparisonDoctorOptions(start = "", end = "", hospitalFilters = []) {
  const options = insightDoctorOptions();
  if (!options.length) return [];
  return prioritizeDoctorOptions(options)
    .filter((doctor) => doctor.key !== selectedDoctor()?.key)
    .filter((doctor) => comparisonDoctorEvents(doctor.key, start, end, hospitalFilters).some(isRosterShiftEvent));
}

function insightDoctorOptions() {
  if (parsedRosterSources && (parsedRosterSources.mmc?.length || parsedRosterSources.ddh?.length || parsedRosterSources.casey?.length || parsedRosterSources.mch?.length)) {
    return rosterDoctorOptions(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [], parsedRosterSources?.casey || [], parsedRosterSources?.mch || []);
  }
  return Array.isArray(insightDoctorOptionsCache) ? insightDoctorOptionsCache : [];
}

function availableInsightDateRange() {
  return currentCalendarInsightDateRange();
}

function currentCalendarInsightDateRange() {
  const currentEvents = buildResolvedPreviewEvents(latestPreview || { events: [] }).filter(isRosterShiftEvent);
  const range = deriveRangeBounds(currentEvents);
  if (range.start || range.end) return range;
  const fallback = deriveRangeBounds(latestPreview?.events || []);
  return {
    start: fallback.start || "",
    end: fallback.end || fallback.start || "",
  };
}

function legacyAvailableInsightDateRange() {
  const events = [
    ...buildResolvedPreviewEvents(latestPreview || { events: [] }),
    ...[...getDoctorAnalysisCache().values()].flat(),
  ].filter(isRosterShiftEvent);
  return deriveRangeBounds(events);
}

function selectedDoctorEventsForInsights(start, end, hospitalFilters = []) {
  if (!latestPreview) return [];
  return buildCurrentDoctorPreviewEvents(start, end, hospitalFilters);
}

function comparisonDoctorEvents(doctorKey, start, end, hospitalFilters = []) {
  const cache = getDoctorAnalysisCache();
  const events = cache.get(doctorKey) || [];
  return filterInsightEvents(events, start, end, hospitalFilters);
}

function buildCurrentDoctorPreviewEvents(start, end, hospitalFilters = []) {
  const events = buildResolvedPreviewEvents(latestPreview || { events: [] });
  return filterInsightEvents(events, start, end, hospitalFilters);
}

function filterInsightEvents(events, start, end, hospitalFilters = []) {
  const startDate = parseDateOnly(start);
  const endDate = parseDateOnly(end);
  return events
    .filter((event) => matchesInsightHospitalFilters(event, hospitalFilters))
    .filter((event) => eventOverlapsDateRange(event, startDate, endDate))
    .sort(comparePreviewEvents);
}

function matchesInsightHospitalFilters(event, hospitalFilters = []) {
  if (!hospitalFilters?.length) return true;
  const eventSources = eventSourceCodes(event);
  return eventSources.some((source) => hospitalFilters.includes(source));
}

function availableHospitalsForInsightRange(start, end) {
  const doctor = selectedDoctor();
  const seen = new Set();
  const options = doctor ? [doctor, ...comparisonDoctorOptions(start, end, [])] : comparisonDoctorOptions(start, end, []);
  for (const option of options) {
    const events = option.key === doctor?.key
      ? selectedDoctorEventsForInsights(start, end, [])
      : comparisonDoctorEvents(option.key, start, end, []);
    for (const event of events) {
      const code = eventSourceCode(event);
      if (code) seen.add(code);
    }
  }
  return [...seen].sort();
}

function availableHospitalsFromInsightEvents(events) {
  const seen = new Set();
  for (const event of events || []) {
    const code = eventSourceCode(event);
    if (code) seen.add(code);
  }
  return [...seen].sort();
}

function matchesPreviewHospitalFilter(event, hospitalFilter) {
  if (!hospitalFilter || hospitalFilter === "all") return true;
  const target = String(hospitalFilter).trim().toUpperCase();
  return eventSourceCodes(event).includes(target);
}

function buildOverlapDays(mine, theirs) {
  const mineByDay = indexEventsByDay(mine);
  const theirsByDay = indexEventsByDay(theirs);
  return [...mineByDay.keys()]
    .filter((date) => theirsByDay.has(date))
    .sort()
    .flatMap((date) => {
      const myEvents = mineByDay.get(date) || [];
      const theirEvents = theirsByDay.get(date) || [];
      const myHospitals = new Set(myEvents.map(eventSourceCode).filter(Boolean));
      const sharedHospitals = [...new Set(theirEvents.map(eventSourceCode).filter(Boolean))]
        .filter((hospital) => myHospitals.has(hospital));
      return sharedHospitals.map((hospital) => ({
        date,
        hospital,
        mine: myEvents.filter((event) => eventSourceCode(event) === hospital),
        theirs: theirEvents.filter((event) => eventSourceCode(event) === hospital),
      }));
    })
    .filter((entry) => entry.mine.length && entry.theirs.length);
}

function indexEventsByDay(events) {
  const map = new Map();
  for (const event of events) {
    const startDate = parseDateOnly(event.start);
    const endDate = previewInclusiveEndDate(event, startDate, parseDateOnly(event.end));
    for (let cursor = new Date(startDate); cursor <= endDate; cursor = addDays(cursor, 1)) {
      const key = formatDateKey(cursor);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(event);
    }
  }
  return map;
}

function renderInsightShiftSummary(events) {
  return [...events]
    .sort(compareInsightEvents)
    .map((event) => `${event.title}${event.allDay || !event.timeLabel ? "" : ` (${event.timeLabel})`}`)
    .join(" | ");
}

function insightShiftPeriodRank(events) {
  if (!events?.length) return 99;
  return Math.min(...events.map(insightEventPeriodRank));
}

function buildWhoAssignments(doctor, events) {
  if (!doctor) return [];
  const metadata = doctorMetadataForKey(doctor.key);
  return dedupeInsightEvents(events)
    .map((event) => buildWhoAssignment(doctor, metadata, event))
    .filter(Boolean);
}

function buildWhoAssignment(doctor, metadata, event) {
  const source = eventSourceCode(event);
  const role = metadata[source]?.role || metadata.any?.role || eventSeniorityRoleCode(event) || "";
  const activeRule = parserRuleForWhoEvent(event, source, role);
  const ruleTitle = activeRule ? parserRulePreviewTitle(activeRule) : "";
  const eventForGrouping = ruleTitle ? { ...event, title: ruleTitle } : event;
  const period = activeRule?.period ? whoPeriodLabel({ ...eventForGrouping, rawValue: activeRule.period }) : whoPeriodLabel(eventForGrouping);
  const rawTeam = activeRule?.base ? activeRule.base : whoTeamLabel(eventForGrouping);
  const isNightSsu = period === "Night" && rawTeam === "SSU";
  const isNightIc = isWhoNightIcShift({ event, period, rawTeam, rule: activeRule, ruleTitle });
  const team = whoDisplayTeamLabel({ period, rawTeam, isNightIc });
  return {
    doctorKey: doctor.key,
    doctorName: doctor.displayName,
    role,
    roleLabel: role || "",
    roleNote: isNightIc ? "IC" : "",
    roleRank: whoRoleRank(role),
    nightIcRank: isNightIc ? 0 : 1,
    source,
    period,
    team,
    teamRank: whoTeamRank(team, source),
    specialTime: whoSpecialTimeLabel(event, period),
    rawValue: String(event?.rawValue || "").trim(),
    ruleCode: activeRule?.code || parserRuleCodeFromRawValue(source, event?.rawValue || "") || incompleteShiftCodeFromTitle(source, event?.title || ""),
    suggestedTitle: ruleTitle || String(event?.title || "").trim(),
    timeLabel: event?.timeLabel || summarizeEventTimes(event?.start || "", event?.end || "", event?.allDay === true),
    date: eventRosterDateKey(event),
    location: String(event?.location || "").trim(),
    event,
  };
}

function whoDisplayTeamLabel({ period, rawTeam, isNightIc }) {
  if (period !== "Night") return rawTeam;
  if (isNightIc || rawTeam === "Night") return "Night main team";
  if (rawTeam === "Hub") return "Night Hub";
  if (rawTeam === "SSU" || rawTeam === "Night SSU") return "Night SSU";
  return rawTeam;
}

function isWhoNightIcShift({ event, period, rawTeam, rule, ruleTitle }) {
  if (period !== "Night") return false;
  const text = `${rawTeam || ""} ${rule?.base || ""} ${ruleTitle || ""} ${event?.title || ""} ${event?.rawValue || ""}`.toUpperCase();
  return /\bIC\b/.test(text) || /\bN1\b/.test(text);
}

function parserRuleForWhoEvent(event, source, role) {
  const normalizedSource = sanitizeIssueSource(source);
  const seniority = sanitizeRuleSeniority(event?.seniority || role);
  const code = parserRuleCodeFromRawValue(normalizedSource, event?.rawValue || "") || incompleteShiftCodeFromTitle(normalizedSource, event?.title || "");
  if (!normalizedSource || !seniority || !code) return null;
  return findParserExtensionRuleForSeniority(normalizedSource, seniority, code);
}

function buildSelectedWhoAssignments(events) {
  return buildWhoAssignments(selectedDoctor(), events)
    .filter((assignment) => isWhoTeamRole(assignment.team));
}

function isWhoTeamRole(team) {
  const normalized = String(team || "").trim().toLowerCase();
  return normalized && normalized !== "float" && normalized !== "rover";
}

function groupWhoAssignments(assignments) {
  const periods = new Map();
  for (const assignment of assignments) {
    if (!periods.has(assignment.period)) periods.set(assignment.period, []);
    periods.get(assignment.period).push(assignment);
  }
  return [...periods.entries()]
    .map(([period, items]) => ({
      period,
      teams: groupWhoTeams(items),
    }))
    .sort((left, right) => whoPeriodRank(left.period) - whoPeriodRank(right.period));
}

function groupWhoTeams(assignments) {
  const teams = new Map();
  for (const assignment of assignments) {
    if (!teams.has(assignment.team)) teams.set(assignment.team, []);
    teams.get(assignment.team).push(assignment);
  }
  return [...teams.entries()]
    .map(([team, items]) => ({
      team,
      items: [...items].sort(compareWhoAssignments),
    }))
    .sort((left, right) => {
      const teamDelta = whoTeamRank(left.team, left.items[0]?.source || "") - whoTeamRank(right.team, right.items[0]?.source || "");
      if (teamDelta !== 0) return teamDelta;
      return left.team.localeCompare(right.team);
    });
}

function compareWhoAssignments(left, right) {
  const nightIcDelta = (left.nightIcRank ?? 1) - (right.nightIcRank ?? 1);
  if (nightIcDelta !== 0) return nightIcDelta;
  const roleDelta = left.roleRank - right.roleRank;
  if (roleDelta !== 0) return roleDelta;
  return left.doctorName.localeCompare(right.doctorName);
}

function renderWhoGroups(groups) {
  return groups.map((group) => `
    <section class="who-period-group">
      <div class="who-period-divider"><span>${escapeHtml(group.period)}</span></div>
      ${group.teams.map((team) => `
        <article class="issue-card who-team-card">
          <strong class="who-team-title">${escapeHtml(team.team)}</strong>
          <div class="who-team-list">
            ${team.items.map((item) => `
              ${renderWhoTeamPerson(item, "data-insights-when-doctor-key")}
            `).join("")}
          </div>
        </article>
      `).join("")}
    </section>
  `).join("");
}

async function renderInlineWhoInsight(container, date, options = {}) {
  if (!container || !date) return;
  if (!canUseRosterInsights()) {
    container.classList.add("hidden");
    container.innerHTML = "";
    return;
  }
  container.classList.remove("hidden");
  container.dataset.inlineWhoDate = date;
  container.dataset.inlineWhoSource = String(options.source || "").toUpperCase();
  container.innerHTML = `<div class="event-inline-loading">Loading who is working with you...</div>`;
  const sourceFilter = String(options.source || "").toUpperCase();
  const sourceFilters = sourceFilter ? [sourceFilter] : [];
  const mine = selectedDoctorEventsForInsights(date, date, sourceFilters)
    .filter(isRosterShiftEvent)
    .filter((event) => eventRosterDateKey(event) === date);
  const activeSources = new Set((sourceFilter ? [sourceFilter] : mine.map(eventSourceCode)).filter(Boolean));
  let coworkers = [];
  const serverResult = mine.length ? await fetchRosterInsightRows({
    startDate: date,
    endDate: date,
    sourceTypes: sourceFilters.map((item) => item.toLowerCase()),
    excludeDoctorKeys: selectedInsightDoctorKeys(),
  }) : { ok: true, rows: [] };
  if (serverResult.ok) {
    const serverRows = serverResult.rows;
    const serverDoctors = insightRowsToDoctorOptions(serverRows);
    const serverEvents = insightRowsToEventsByDoctor(serverRows);
    coworkers = serverDoctors
      .map((doctor) => ({
        doctor,
        events: (serverEvents.get(doctor.key) || [])
          .filter(isRosterShiftEvent)
          .filter((event) => eventRosterDateKey(event) === date)
          .filter((event) => !activeSources.size || activeSources.has(eventSourceCode(event))),
      }))
      .filter((entry) => entry.events.length)
      .flatMap((entry) => buildWhoAssignments(entry.doctor, entry.events));
  }
  const teamAssignments = [
    ...buildSelectedWhoAssignments(mine),
    ...coworkers,
  ];
  container.innerHTML = `
    <div class="event-inline-head">
      <strong>Who else is working with me?</strong>
      <span>${escapeHtml(formatInsightDate(date))}${sourceFilter ? ` · ${escapeHtml(sourceFilter)}` : ""}</span>
    </div>
    <div class="issue-card event-inline-mine">
      <strong>${escapeHtml(selectedDoctor()?.displayName || "Selected doctor")}</strong>
      <p>${mine.length ? escapeHtml(renderInsightShiftSummary(mine)) : "No rostered shift found for this date."}</p>
    </div>
    ${!serverResult.ok
      ? renderRosterInsightUnavailable()
      : teamAssignments.length
      ? renderInlineWhoGroups(groupWhoAssignments(teamAssignments), date, sourceFilter)
      : `<article class="issue-card"><p>No other clinicians are rostered with you for this shift.</p></article>`}
  `;
}

function renderInlineWhoGroups(groups, date, sourceFilter = "") {
  return groups.map((group) => `
    <section class="who-period-group">
      <div class="who-period-divider"><span>${escapeHtml(group.period)}</span></div>
      ${group.teams.map((team) => `
        <article class="issue-card who-team-card">
          <strong class="who-team-title">${escapeHtml(team.team)}</strong>
          <div class="who-team-list">
            ${team.items.map((item) => `
              ${renderWhoTeamPerson(item, "data-inline-when-doctor")}
            `).join("")}
          </div>
        </article>
      `).join("")}
    </section>
  `).join("") + `<button type="button" class="button button-secondary event-inline-refresh" data-inline-back-who="${escapeHtml(date)}" data-inline-back-source="${escapeHtml(sourceFilter)}">Refresh list</button>`;
}

function renderWhoTeamPerson(item, doctorAttribute) {
  const roleParts = [item.roleLabel, item.roleNote].filter(Boolean);
  return `
    <div class="who-team-person">
      <button type="button" class="who-team-name" ${doctorAttribute}="${escapeHtml(item.doctorKey || "")}" title="Show future shifts with ${escapeHtml(item.doctorName)}">${escapeHtml(item.doctorName)}</button>
      <span class="who-team-meta">
        ${roleParts.length ? `<span class="who-team-role">${escapeHtml(roleParts.join(" · "))}</span>` : ""}
        ${item.specialTime ? `<span class="who-team-time">${escapeHtml(item.specialTime)}</span>` : ""}
      </span>
    </div>
  `;
}

async function renderInlineWhenInsight(container, doctorKey) {
  if (!container || !doctorKey) return;
  if (!canUseRosterInsights()) {
    container.classList.add("hidden");
    container.innerHTML = "";
    return;
  }
  container.classList.remove("hidden");
  container.innerHTML = `<div class="event-inline-loading">Loading future shifts together...</div>`;
  const fromDate = formatDateKey(new Date());
  const range = currentCalendarInsightDateRange();
  const toDate = range.end || fromDate;
  const serverResult = await fetchRosterInsightRows({ startDate: fromDate, endDate: toDate, doctorKeys: [doctorKey] });
  if (serverResult.ok) {
    const serverRows = serverResult.rows;
    const serverDoctors = prioritizeDoctorOptions(insightRowsToDoctorOptions(serverRows));
    const serverEvents = insightRowsToEventsByDoctor(serverRows);
    const selectedComparison = serverDoctors.find((doctor) => doctor.key === doctorKey) || null;
    const mine = selectedDoctorEventsForInsights(fromDate, toDate, []).filter(isRosterShiftEvent);
    const theirs = selectedComparison
      ? (serverEvents.get(selectedComparison.key) || []).filter(isRosterShiftEvent)
      : [];
    const overlaps = selectedComparison ? buildOverlapDays(mine, theirs) : [];
    renderInlineWhenInsightResult(container, { fromDate, selectedComparison, overlaps });
    return;
  }
  container.innerHTML = renderRosterInsightUnavailable();
}

function renderInlineWhenInsightResult(container, { fromDate, selectedComparison, overlaps }) {
  const backDate = container.dataset.inlineWhoDate || fromDate;
  const backSource = container.dataset.inlineWhoSource || "";
  container.innerHTML = `
    <div class="event-inline-head">
      <strong>${selectedComparison ? `When am I working with ${escapeHtml(selectedComparison.displayName)}?` : "Future shifts together"}</strong>
      <span>From ${escapeHtml(formatDate(fromDate))}</span>
    </div>
    <button type="button" class="button button-secondary event-inline-back" data-inline-back-who="${escapeHtml(backDate)}" data-inline-back-source="${escapeHtml(backSource)}">Back to this shift</button>
    ${selectedComparison
      ? overlaps.length
        ? overlaps.map((entry) => `
          <article class="issue-card">
            ${renderInsightDateButton(entry.date, "data-inline-who-on-date")}
            <p><strong>Hospital:</strong> ${escapeHtml(entry.hospital || "Unknown")}</p>
            <p><strong>${escapeHtml(selectedDoctor()?.displayName || "Selected doctor")}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.mine))}</p>
            <p><strong>${escapeHtml(selectedComparison.displayName)}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.theirs))}</p>
          </article>
        `).join("")
        : `<article class="issue-card"><p>No future overlapping working days were found.</p></article>`
      : `<article class="issue-card"><p>This clinician is not available in the comparison roster.</p></article>`}
  `;
}

async function openDoctorProfileFromInsight(doctorKey) {
  const normalizedKey = normalizeRosterName(doctorKey);
  if (!normalizedKey) return;
  const options = doctorPickerOptions();
  const localOption = options.find((doctor) => doctor.key === normalizedKey);
  if (localOption && options.length > 1) {
    if (canUseCreatorDoctorSwitcher() && cloudAvailable && !serverUsers.length) {
      await loadServerUsers();
    }
    closeInsightsModal();
    if (canUseCreatorDoctorSwitcher()) {
      await switchDoctorSelection(localOption.key, { resetRange: true });
    } else {
      doctorSelect.value = normalizedKey;
      clearPreviewData();
      saveCurrentSessionState();
      syncActionState();
      await updatePreview({ resetRange: true });
    }
    return;
  }
  setStatus("That doctor is not directly viewable from this account yet.", true);
}

function dedupeInsightEvents(events) {
  const seen = new Set();
  const deduped = [];
  for (const event of [...events].sort(compareInsightEvents)) {
    const marker = `${event.title}|${event.start}|${event.end}|${event.rawValue || ""}`;
    if (seen.has(marker)) continue;
    seen.add(marker);
    deduped.push(event);
  }
  return deduped;
}

function compareInsightEvents(left, right) {
  const periodDelta = insightEventPeriodRank(left) - insightEventPeriodRank(right);
  if (periodDelta !== 0) return periodDelta;
  const startDelta = String(left.start || "").localeCompare(String(right.start || ""));
  if (startDelta !== 0) return startDelta;
  return String(left.title || "").localeCompare(String(right.title || ""));
}

function insightEventPeriodRank(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  if (text.includes("night")) return 2;
  if (/\bpm\b/.test(text)) return 1;
  if (/\bam\b/.test(text)) return 0;

  const clock = extractTimePortion(event?.start || "");
  if (!clock) return 3;
  const [hoursText = "0", minutesText = "0"] = clock.split(":");
  const totalMinutes = Number(hoursText) * 60 + Number(minutesText);
  if (totalMinutes >= 20 * 60 || totalMinutes < 6 * 60) return 2;
  if (totalMinutes >= 12 * 60 + 1) return 1;
  return 0;
}

function whoPeriodLabel(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  if (text.includes("night")) return "Night";
  if (/\bpm\b/.test(text)) return "PM";
  if (/\bam\b/.test(text)) return "AM";
  const rank = insightEventPeriodRank(event);
  return rank === 2 ? "Night" : rank === 1 ? "PM" : "AM";
}

function whoPeriodRank(period) {
  if (period === "AM") return 0;
  if (period === "PM") return 1;
  if (period === "Night") return 2;
  return 3;
}

function whoTeamLabel(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  if (text.includes("avao")) return "AVAO";
  if (text.includes("green")) return "Green";
  if (text.includes("orange")) return "Orange";
  if (text.includes("amber")) return "Amber";
  if (text.includes("silver")) return "Silver";
  if (text.includes("resus")) return "Resus";
  if (text.includes("float") || text.includes("rover")) return "Float";
  if (text.includes("clinic")) return "Clinic";
  if (text.includes("fast")) return "Fast Track";
  if (text.includes("ssu")) return "SSU";
  if (text.includes("hith")) return "HITH";
  if (text.includes("vhh")) return "VHH";
  if (text.includes("paed")) return "Paeds";
  if (text.includes("extra")) return "Extra";
  return cleanWhoSourceTitle(event.title || "Other");
}

function cleanWhoSourceTitle(title) {
  return String(title || "")
    .replace(/^(MMC|DDH|Casey|MCH):\s*/i, "")
    .replace(/\s+(AM|PM)\b/i, "")
    .trim() || "Other";
}

function whoTeamRank(team, source) {
  const normalized = String(team || "").toLowerCase();
  const sourceCode = String(source || "").toUpperCase();
  const ranks = sourceCode === "DDH"
    ? ["avao", "orange", "silver", "resus", "float", "clinic", "fast track", "ssu", "hith", "vhh", "paeds", "extra", "other"]
    : ["green", "amber", "resus", "float", "clinic", "fast track", "night main team", "night hub", "night ssu", "ssu", "other"];
  const index = ranks.indexOf(normalized);
  return index >= 0 ? index : ranks.length;
}

function whoRoleRank(role) {
  const normalized = normalizeWhoRole(role);
  const ranks = {
    SMS: 0,
    CMO: 1,
    SR: 2,
    IR: 3,
    JR: 4,
    HMO: 5,
    I: 6,
    ENP: 7,
    AMP: 8,
  };
  return Object.prototype.hasOwnProperty.call(ranks, normalized) ? ranks[normalized] : 99;
}

function normalizeWhoRole(role) {
  const upper = String(role || "").trim().toUpperCase();
  if (!upper) return "";
  if (upper === "SMS" || upper.includes("SENIOR MEDICAL STAFF") || upper.includes("CONSULTANT") || upper.includes("STAFF SPECIALIST")) return "SMS";
  if (upper === "CMO" || upper.includes("CMO")) return "CMO";
  if (upper === "SR" || upper.includes("SENIOR REGISTRAR") || upper.includes("SENIOR REG")) return "SR";
  if (upper === "IR" || upper === "TR" || upper.includes("TRANSITIONAL") || upper.includes("INTERMEDIATE")) return "IR";
  if (upper === "JR" || upper.includes("JUNIOR REGISTRAR") || upper.includes("JUNIOR REG")) return "JR";
  if (upper === "H" || upper === "HMO" || upper.includes("HMO")) return "HMO";
  if (upper === "I" || upper.includes("INTERN")) return "I";
  if (upper === "ENP" || upper.includes("NURSE PRACTITIONER")) return "ENP";
  if (upper === "AMP" || upper.includes("PHYSIOTHERAPIST")) return "AMP";
  return upper;
}

function eventSeniorityRoleCode(event) {
  return normalizeWhoRole(event?.seniority || event?.role || "");
}

function whoSpecialTimeLabel(event, period) {
  if (event.allDay) return "";
  const start = extractTimePortion(event.start || "");
  const end = extractTimePortion(event.end || "");
  if (!start || !end) return "";
  const source = eventSourceCode(event);
  const standard = {
    MMC: { AM: new Set(["07:30", "08:00"]), PM: new Set(["14:30"]), Night: new Set() },
    DDH: { AM: new Set(["07:30", "08:00"]), PM: new Set(["14:30", "15:00"]), Night: new Set(["23:00"]) },
    CASEY: { AM: new Set(["07:30", "08:00"]), PM: new Set(["14:30"]), Night: new Set(["23:00"]) },
    MCH: { AM: new Set(["08:00"]), PM: new Set(["14:30", "15:00"]), Night: new Set(["23:00"]) },
  };
  const standardStarts = standard[source]?.[period] || new Set();
  return standardStarts.has(start) ? "" : `${start}-${end}`;
}

function eventRosterDateKey(event) {
  const start = String(event?.start || "");
  return start.slice(0, 10);
}

function eventSourceCode(event) {
  return eventSourceCodes(event)[0] || "";
}

function eventSourceCodes(event) {
  const values = Array.isArray(event?.sources) ? event.sources : [event?.source];
  const explicit = values
    .map((item) => normalizeEventSourceCode(item))
    .filter(Boolean);
  if (explicit.length) return [...new Set(explicit)];
  const titlePrefix = String(event?.title || "").match(/^(MMC|DDH|Casey|MCH):/i)?.[1];
  const titleCode = normalizeEventSourceCode(titlePrefix);
  return titleCode ? [titleCode] : [];
}

function normalizeEventSourceCode(value) {
  const code = String(value || "").trim().toUpperCase();
  if (code === "MMC" || code === "DDH" || code === "MCH") return code;
  if (code === "CASEY") return "CASEY";
  return "";
}

function isLeaveEvent(event) {
  return /\b(?:leave|conference|cme|annual|sick|personal|study|exam|sabbatical|parental|long service)\b/i.test(`${event?.title || ""} ${event?.rawValue || ""}`);
}

function isRosterShiftEvent(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  return !(
    isLeaveEvent(event)
    || text.includes("phnw")
    || text.includes("public holiday")
  );
}

function chooseNextOverlapDate(overlaps) {
  if (!overlaps.length) return "";
  const today = formatDateKey(new Date());
  return overlaps.find((entry) => entry.date >= today)?.date || overlaps[0].date;
}

function renderInsightDateButton(date, dataAttribute) {
  const safeDate = escapeHtml(date);
  return `
    <button type="button" class="insight-date-button" ${dataAttribute}="${safeDate}" title="Show everyone working with you on ${escapeHtml(formatInsightDate(date))}">
      ${escapeHtml(formatInsightDate(date))}
    </button>
  `;
}

function formatInsightDate(value) {
  return parseDateOnly(value).toLocaleDateString("en-AU", {
    weekday: "long",
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function defaultInsightDate(termStart, termEnd) {
  const today = formatDateKey(new Date());
  if (today < termStart) return termStart;
  if (today > termEnd) return termStart;
  return today;
}

function getDoctorAnalysisCache() {
  if (!parsedRosterSources || (!parsedRosterSources.mmc?.length && !parsedRosterSources.ddh?.length && !parsedRosterSources.casey?.length && !parsedRosterSources.mch?.length)) {
    if (!hydrateInsightCacheFromSnapshot()) {
      doctorAnalysisCacheKey = "";
      doctorAnalysisCache = new Map();
      insightDoctorOptionsCache = [];
      insightDoctorRoleCache = new Map();
    }
    return doctorAnalysisCache;
  }
  const cacheKey = currentInsightCacheKey();
  if (doctorAnalysisCacheKey === cacheKey && doctorAnalysisCache.size) return doctorAnalysisCache;
  const cache = new Map();
  const analysisSettings = {
    ...settings,
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
    includeLocations: false,
  };
  for (const doctor of rosterDoctorOptions(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [], parsedRosterSources?.casey || [], parsedRosterSources?.mch || [])) {
    const view = buildRosterView(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [], doctor.key, analysisSettings, {}, {}, [], parsedRosterSources?.casey || [], parsedRosterSources?.mch || []);
    cache.set(doctor.key, view.events);
  }
  doctorAnalysisCacheKey = cacheKey;
  doctorAnalysisCache = cache;
  insightDoctorOptionsCache = rosterDoctorOptions(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [], parsedRosterSources?.casey || [], parsedRosterSources?.mch || []);
  doctorRoleIndex = buildDoctorRoleIndex();
  insightDoctorRoleCache = new Map(doctorRoleIndex);
  cacheCurrentInsightSnapshot();
  return doctorAnalysisCache;
}

function clearDoctorAnalysisCache() {
  doctorAnalysisCacheKey = "";
  doctorAnalysisCache = new Map();
  insightDoctorOptionsCache = [];
  insightDoctorRoleCache = new Map();
  clearInsightWarmup();
}

function currentInsightCacheKey() {
  return JSON.stringify({
    imports: currentImportStateKey(),
    sourcePrefix: settings.showSourcePrefix,
    showAmPm: settings.showAmPm,
    includeAnnualLeave: settings.includeAnnualLeave,
    includeConferenceLeave: settings.includeConferenceLeave,
    includePublicHoliday: settings.includePublicHoliday,
    includeSickLeave: settings.includeSickLeave,
  });
}

function hydrateInsightCacheFromSnapshot(snapshot = currentSnapshot) {
  const cache = sanitizeInsightCache(snapshot?.insightCache);
  if (!cache || cache.key !== currentInsightCacheKey()) return false;
  doctorAnalysisCacheKey = cache.key;
  doctorAnalysisCache = new Map(Object.entries(cache.doctorEvents || {}));
  insightDoctorOptionsCache = cache.doctorOptions || [];
  insightDoctorRoleCache = new Map(Object.entries(cache.doctorRoles || {}));
  return Boolean(doctorAnalysisCache.size && insightDoctorOptionsCache.length);
}

function cacheCurrentInsightSnapshot() {
  if (!currentSnapshot || !doctorAnalysisCacheKey || !doctorAnalysisCache.size) return;
  currentSnapshot = sanitizeWorkspaceSnapshot({
    ...currentSnapshot,
    insightCache: buildInsightCachePayload(),
  });
  saveCurrentWorkspace();
}

function buildInsightCachePayload() {
  return {
    key: doctorAnalysisCacheKey,
    builtAt: new Date().toISOString(),
    doctorOptions: insightDoctorOptionsCache,
    doctorEvents: Object.fromEntries([...doctorAnalysisCache.entries()].map(([key, events]) => [key, events.map(serializeEvent)])),
    doctorRoles: Object.fromEntries(insightDoctorRoleCache.entries()),
  };
}

function scheduleInsightWarmup() {
  clearInsightWarmup();
  if (!canUseRosterInsights() || !latestPreview || !selectedDoctor()) return;
  const warmKey = [
    insightWarmBaseKey(),
    settings.dateFrom || "",
    settings.dateTo || "",
    settings.hospitalFilter || "all",
  ].join("|");
  if (visibleInsightWarmKey === warmKey && visibleInsightWarmCache.size) return;
  visibleInsightWarmKey = warmKey;
  const run = () => {
    insightWarmupTimer = 0;
    insightWarmupPromise = warmInsightData().catch((error) => {
      console.warn("Roster insight warmup failed", error);
    });
  };
  if (typeof requestIdleCallback === "function") {
    insightWarmupTimer = requestIdleCallback(run, { timeout: 1200 });
  } else {
    insightWarmupTimer = setTimeout(run, 250);
  }
}

function clearInsightWarmup() {
  if (typeof cancelIdleCallback === "function" && typeof insightWarmupTimer === "number" && insightWarmupTimer) {
    try {
      cancelIdleCallback(insightWarmupTimer);
    } catch {
      // Ignore unsupported idle callback cancellation.
    }
  } else {
    clearTimeout(insightWarmupTimer);
  }
  insightWarmupTimer = 0;
}

async function warmInsightData() {
  if (!canUseRosterInsights() || !latestPreview || !selectedDoctor()) return;
  const range = currentCalendarInsightDateRange();
  const startDate = settings.dateFrom || range.start || "";
  const endDate = settings.dateTo || range.end || startDate;
  if (!startDate || !endDate) return;
  const selectedKeys = selectedInsightDoctorKeys();
  const selectedEvents = selectedDoctorEventsForInsights(startDate, endDate).filter(isRosterShiftEvent);
  const dates = [...new Set(selectedEvents.map(eventRosterDateKey).filter(Boolean))]
    .sort()
    .slice(0, 42);
  await Promise.all(dates.map((date) => fetchRosterInsightRows({
    startDate: date,
    endDate: date,
    excludeDoctorKeys: selectedKeys,
  })));
  const dateSourcePairs = [];
  for (const event of selectedEvents) {
    const date = eventRosterDateKey(event);
    const source = eventSourceCode(event);
    if (date && source) dateSourcePairs.push(`${date}|${source}`);
  }
  await Promise.all([...new Set(dateSourcePairs)].slice(0, 64).map((pair) => {
    const [date, source] = pair.split("|");
    return fetchRosterInsightRows({
      startDate: date,
      endDate: date,
      sourceTypes: [source.toLowerCase()],
      excludeDoctorKeys: selectedKeys,
    });
  }));
  await fetchRosterOverlapDoctors({
    startDate,
    endDate,
    excludeDoctorKeys: selectedKeys,
    overlapDoctorKeys: selectedKeys,
  });
}

function syncActionState() {
  syncControlBarVisibility();
  const ready = Boolean(selectedDoctor());
  exportButton.disabled = !ready;
  mobileExportButton.disabled = !ready;
  syncMobileChrome();
}

function syncControlBarVisibility() {
  const loggedIn = Boolean(currentUserEmail && currentUserPassword);
  controlBar.classList.toggle("hidden", !loggedIn);
}

function createFormData(doctor = null) {
  return createFormDataForEntries(selectedFiles, doctor);
}

function createFormDataForEntries(entries, doctor = null) {
  const body = new FormData();
  for (const entry of entries) {
    if (!entry.file) continue;
    body.append("rosterFiles", entry.file);
    body.append("rosterFileId", entry.id);
    body.append("rosterFileAddedAt", entry.addedAt || "");
  }
  if (doctor) {
    body.append("doctorKey", doctor.key);
    body.append("doctorDisplay", doctor.displayName);
    body.append("doctorAliases", JSON.stringify(doctor.aliases || []));
  }
  body.append("settings", JSON.stringify(settings));
  body.append("overrides", JSON.stringify(cleanOverrides()));
  body.append("customEvents", JSON.stringify(customEventsForActiveCalendar()));
  body.append("conflictSelections", JSON.stringify(conflictSelections));
  return body;
}

function cleanOverrides() {
  const next = {};
  for (const [id, value] of Object.entries(overrides)) {
    const title = (value.title || "").trim();
    const include = value.include;
    const start = value.start || "";
    const end = value.end || "";
    const hasLocation = Object.prototype.hasOwnProperty.call(value, "location");
    const location = hasLocation ? value.location || "" : "";
    const allDay = value.allDay;
    if (!title && typeof include !== "boolean" && !start && !end && !hasLocation && typeof allDay !== "boolean") continue;
    next[id] = {};
    if (title) next[id].title = title;
    if (typeof include === "boolean") next[id].include = include;
    if (start) next[id].start = start;
    if (end) next[id].end = end;
    if (hasLocation) next[id].location = location;
    if (typeof allDay === "boolean") next[id].allDay = allDay;
  }
  return next;
}

function customEventToPreviewEvent(event) {
  if (event.allDay) {
    return {
      id: event.id,
      ownerEmail: normalizeEmail(event.ownerEmail),
      source: "Custom",
      title: event.title,
      allDay: true,
      start: event.startDate,
      end: formatDateKey(addDays(parseDateOnly(event.endDate), 1)),
      location: event.location || "",
      rawValue: "Custom event",
      timeLabel: "All day",
      monthKey: event.startDate.slice(0, 7),
      isEditedImport: false,
    };
  }

  const endDate = event.endDate && event.endDate !== event.startDate
    ? event.endDate
    : compareClockStrings(event.endTime, event.startTime) <= 0
      ? formatDateKey(addDays(parseDateOnly(event.startDate), 1))
      : event.startDate;
  const start = `${event.startDate}T${event.startTime}:00`;
  const end = `${endDate}T${event.endTime}:00`;
  return {
    id: event.id,
    ownerEmail: normalizeEmail(event.ownerEmail),
    source: "Custom",
    title: event.title,
    allDay: false,
    start,
    end,
    location: event.location || "",
    rawValue: "Custom event",
    timeLabel: `${event.startTime}-${event.endTime}`,
    monthKey: event.startDate.slice(0, 7),
    isEditedImport: false,
  };
}

function movePreviewEvent(id, targetDate) {
  const event = currentPreviewEvents.get(id);
  if (!event) return;
  if (isCustomPreviewEvent(event)) {
    const updated = shiftPreviewEventToDay(event, targetDate);
    customEvents = customEvents.map((item) => item.id === id && normalizeEmail(item.ownerEmail) === activeCalendarEmail() ? previewEventToCustomEvent(updated, item) : item);
  } else {
    const updated = shiftPreviewEventToDay(event, targetDate);
    syncImportedOverride(id, {
      start: updated.start,
      end: updated.end,
      allDay: updated.allDay,
      location: updated.location || "",
    });
  }
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Event moved.");
}

function shiftPreviewEventToDay(event, targetDate) {
  const startDate = event.start.slice(0, 10);
  const endDate = event.end.slice(0, 10);
  if (event.allDay) {
    const inclusiveEnd = previewInclusiveEndDate(event, parseDateOnly(startDate), parseDateOnly(endDate));
    const spanDays = diffDays(parseDateOnly(startDate), inclusiveEnd);
    const newStart = targetDate;
    const newEndInclusive = formatDateKey(addDays(parseDateOnly(targetDate), spanDays));
    return {
      ...event,
      start: newStart,
      end: formatDateKey(addDays(parseDateOnly(newEndInclusive), 1)),
      timeLabel: "All day",
    };
  }

  const endSpanDays = diffDays(parseDateOnly(startDate), parseDateOnly(endDate));
  const newEndDate = formatDateKey(addDays(parseDateOnly(targetDate), endSpanDays));
  const startClock = extractTimePortion(event.start);
  const endClock = extractTimePortion(event.end);
  return {
    ...event,
    start: `${targetDate}T${startClock}:00`,
    end: `${newEndDate}T${endClock}:00`,
    timeLabel: `${startClock}-${endClock}`,
  };
}

function previewEventToCustomEvent(event, existing = null) {
  return {
    id: event.id,
    ownerEmail: normalizeEmail(existing?.ownerEmail || event.ownerEmail || activeCalendarEmail()),
    title: event.title,
    startDate: event.start.slice(0, 10),
    endDate: event.allDay
      ? formatDateKey(addDays(parseDateOnly(event.end), -1))
      : event.end.slice(0, 10),
    allDay: event.allDay,
    startTime: event.allDay ? "" : extractTimePortion(event.start),
    endTime: event.allDay ? "" : extractTimePortion(event.end),
    location: event.location || "",
    include: existing?.include !== false,
  };
}

function copyPreviewEvent(id) {
  const event = currentPreviewEvents.get(id);
  if (!event) return;
  copiedEvent = { ...event };
  closeContextMenu();
  setStatus("Event copied.");
}

function pasteCopiedEvent(targetDate) {
  if (!copiedEvent) return;
  const shifted = shiftPreviewEventToDay({
    ...copiedEvent,
    id: newCustomEventId(),
    ownerEmail: activeCalendarEmail(),
    source: "Custom",
    isEditedImport: false,
  }, targetDate);
  closeContextMenu();
  openCustomEventModal(previewEventToCustomEvent(shifted), targetDate, { draft: true });
}

function deletePreviewEvent(id) {
  const event = currentPreviewEvents.get(id);
  if (!event) return;
  if (isCustomPreviewEvent(event)) {
    ensureEditableCustomEvent(event);
    removeCustomEventForActiveCalendar(id);
    if (openReviewId === id) closeReviewModal();
  } else {
    syncImportedOverride(id, { include: false });
    if (openReviewId === id) closeReviewModal();
  }
  closeContextMenu();
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Event deleted.");
}

function resetImportedEvent(id) {
  if (!hasImportedOverride(id)) return;
  delete overrides[id];
  closeContextMenu();
  rebuildClientPreview();
  saveCurrentSessionState();
  if (openReviewId === id) {
    openReviewModal(id);
  }
  setStatus("Imported event reset.");
}

function hasImportedOverride(id) {
  return Boolean(overrides[id] && Object.keys(overrides[id]).length);
}

function startPendingPreviewGesture(event, chip) {
  cancelPendingPreviewGesture();
  pendingPreviewGesture = {
    pointerId: event.pointerId,
    chip,
    startX: event.clientX,
    startY: event.clientY,
    timer: window.setTimeout(() => {
      if (!pendingPreviewGesture || pendingPreviewGesture.pointerId !== event.pointerId) return;
      const pending = pendingPreviewGesture;
      pendingPreviewGesture = null;
      suppressPreviewClickUntil = Date.now() + 350;
      startPreviewGesture({
        pointerId: pending.pointerId,
        clientX: pending.startX,
        clientY: pending.startY,
      }, pending.chip);
    }, 450),
  };
}

function updatePendingPreviewGesture(event) {
  const pending = pendingPreviewGesture;
  if (!pending) return;
  const dx = event.clientX - pending.startX;
  const dy = event.clientY - pending.startY;
  if (Math.hypot(dx, dy) > 10) cancelPendingPreviewGesture();
}

function cancelPendingPreviewGesture() {
  if (!pendingPreviewGesture) return;
  window.clearTimeout(pendingPreviewGesture.timer);
  pendingPreviewGesture = null;
}

function startPreviewGesture(event, chip) {
  const previewEvent = currentPreviewEvents.get(chip.dataset.reviewId);
  if (!previewEvent) return;
  previewGesture = {
    pointerId: event.pointerId,
    id: previewEvent.id,
    chip,
    sourceEvent: { ...previewEvent },
    startX: event.clientX,
    startY: event.clientY,
    originDay: previewEvent.start.slice(0, 10),
    hoverDay: previewEvent.start.slice(0, 10),
    slotOffset: 0,
    minuteOffset: 0,
    moved: false,
    timeShiftDisabled: false,
    autoShiftHandle: null,
    autoShiftAccumulator: 0,
    autoShiftLastTs: 0,
    autoShiftDirection: 0,
    autoShiftRate: 0,
    originalMetaText: chip.querySelector(".preview-chip-meta")?.textContent || "",
    originalMetaPresent: Boolean(chip.querySelector(".preview-chip-meta")),
  };
  chip.style.pointerEvents = "none";
  chip.setPointerCapture?.(event.pointerId);
}

function updatePreviewGesture(event) {
  const gesture = previewGesture;
  if (!gesture) return;
  const dx = event.clientX - gesture.startX;
  const dy = event.clientY - gesture.startY;
  gesture.moved = gesture.moved || Math.abs(dx) > 4 || Math.abs(dy) > 4;
  gesture.chip.classList.add("is-dragging");
  gesture.chip.style.transform = `translate(${dx}px, ${dy}px)`;

  const hoverDay = dayKeyAtPoint(event.clientX, event.clientY);
  gesture.hoverDay = hoverDay || "";
  if (hoverDay && hoverDay !== gesture.originDay && !gesture.timeShiftDisabled) {
    disableGestureTimeShift(gesture);
  }
  const slotOffset = calculateTimeShiftSlots(dy);
  const canTimeShift = (
    !gesture.sourceEvent.allDay &&
    !gesture.timeShiftDisabled &&
    hoverDay === gesture.originDay &&
    Math.abs(dy) >= Math.abs(dx) &&
    slotOffset !== 0 &&
    Math.abs(slotOffset) <= 6
  );

  if (canTimeShift) {
    clearDropTargets();
    applyPreviewGestureSlot(gesture, slotOffset);
    return;
  }

  if (Math.abs(slotOffset) > 6 && !gesture.timeShiftDisabled) {
    disableGestureTimeShift(gesture);
  } else if (gesture.slotOffset !== 0 || gesture.minuteOffset !== 0) {
    resetGestureTimeShift(gesture);
  }

  if (hoverDay && hoverDay !== gesture.originDay) {
    setDropTarget(hoverDay);
  } else {
    clearDropTargets();
  }
}

function finishPreviewGesture(event) {
  const gesture = previewGesture;
  if (!gesture) return;
  const hoverDay = dayKeyAtPoint(event.clientX, event.clientY) || gesture.hoverDay;
  const shouldSuppressClick = gesture.moved;
  stopGestureAutoShift(gesture);
  if (gesture.moved && gesture.minuteOffset !== 0 && hoverDay === gesture.originDay) {
    commitPreviewGestureTime(gesture);
    suppressPreviewClickUntil = Date.now() + 200;
  } else if (gesture.moved && hoverDay && hoverDay !== gesture.originDay) {
    movePreviewEvent(gesture.id, hoverDay);
    suppressPreviewClickUntil = Date.now() + 200;
  } else {
    restorePreviewGestureMeta(gesture);
    if (shouldSuppressClick) suppressPreviewClickUntil = Date.now() + 200;
  }
  teardownPreviewGesture();
}

function cancelPreviewGesture() {
  if (!previewGesture) return;
  stopGestureAutoShift(previewGesture);
  restorePreviewGestureMeta(previewGesture);
  teardownPreviewGesture();
}

function teardownPreviewGesture() {
  const gesture = previewGesture;
  if (!gesture) return;
  gesture.chip.classList.remove("is-dragging");
  gesture.chip.style.transform = "";
  gesture.chip.style.pointerEvents = "";
  clearDropTargets();
  previewGesture = null;
}

function dayKeyAtPoint(x, y) {
  const element = document.elementFromPoint(x, y);
  return element?.closest("[data-add-date]")?.dataset.addDate || "";
}

function setDropTarget(dayKey) {
  preview.querySelectorAll(".preview-cell.is-drop-target").forEach((node) => {
    if (node.dataset.addDate !== dayKey) node.classList.remove("is-drop-target");
  });
  const target = preview.querySelector(`[data-add-date="${dayKey}"]`);
  if (target) target.classList.add("is-drop-target");
}

function clearDropTargets() {
  preview.querySelectorAll(".preview-cell.is-drop-target").forEach((cell) => cell.classList.remove("is-drop-target"));
}

function calculateTimeShiftSlots(deltaY) {
  const direction = deltaY < 0 ? -1 : 1;
  const distance = Math.abs(deltaY);
  return direction * Math.floor(distance / 18);
}

function applyPreviewGestureSlot(gesture, slotOffset) {
  const previousAbsolute = Math.abs(gesture.slotOffset);
  const nextAbsolute = Math.abs(slotOffset);
  const direction = slotOffset < 0 ? -1 : 1;
  gesture.slotOffset = slotOffset;

  if (nextAbsolute <= 3) {
    stopGestureAutoShift(gesture);
    gesture.minuteOffset = slotOffset * 15;
    const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.minuteOffset);
    setPreviewChipMeta(gesture.chip, shifted, true);
    return;
  }

  if (previousAbsolute <= 3 || gesture.autoShiftDirection !== direction) {
    gesture.minuteOffset = direction * 60;
  }
  const rate = nextAbsolute === 4 ? 1 : nextAbsolute === 5 ? 2 : 4;
  startGestureAutoShift(gesture, direction, rate);
  const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.minuteOffset);
  setPreviewChipMeta(gesture.chip, shifted, true);
}

function startGestureAutoShift(gesture, direction, rate) {
  gesture.autoShiftDirection = direction;
  gesture.autoShiftRate = rate;
  if (gesture.autoShiftHandle) return;
  gesture.autoShiftAccumulator = 0;
  gesture.autoShiftLastTs = 0;
  const tick = (timestamp) => {
    if (!previewGesture || previewGesture !== gesture) return;
    if (gesture.timeShiftDisabled || Math.abs(gesture.slotOffset) < 4) {
      gesture.autoShiftHandle = null;
      gesture.autoShiftLastTs = 0;
      gesture.autoShiftAccumulator = 0;
      return;
    }
    if (!gesture.autoShiftLastTs) gesture.autoShiftLastTs = timestamp;
    const elapsed = timestamp - gesture.autoShiftLastTs;
    gesture.autoShiftLastTs = timestamp;
    gesture.autoShiftAccumulator += elapsed;
    const stepMs = 360 / gesture.autoShiftRate;
    while (gesture.autoShiftAccumulator >= stepMs) {
      gesture.autoShiftAccumulator -= stepMs;
      gesture.minuteOffset += gesture.autoShiftDirection * 15;
    }
    const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.minuteOffset);
    setPreviewChipMeta(gesture.chip, shifted, true);
    gesture.autoShiftHandle = requestAnimationFrame(tick);
  };
  gesture.autoShiftHandle = requestAnimationFrame(tick);
}

function stopGestureAutoShift(gesture) {
  if (gesture.autoShiftHandle) cancelAnimationFrame(gesture.autoShiftHandle);
  gesture.autoShiftHandle = null;
  gesture.autoShiftAccumulator = 0;
  gesture.autoShiftLastTs = 0;
  gesture.autoShiftDirection = 0;
  gesture.autoShiftRate = 0;
}

function resetGestureTimeShift(gesture) {
  stopGestureAutoShift(gesture);
  gesture.slotOffset = 0;
  gesture.minuteOffset = 0;
  restorePreviewGestureMeta(gesture);
}

function disableGestureTimeShift(gesture) {
  gesture.timeShiftDisabled = true;
  resetGestureTimeShift(gesture);
}

function restorePreviewGestureMeta(gesture) {
  const meta = gesture.chip.querySelector(".preview-chip-meta");
  if (gesture.originalMetaPresent) {
    if (meta) {
      meta.textContent = gesture.originalMetaText;
    } else {
      gesture.chip.insertAdjacentHTML("beforeend", `<span class="preview-chip-meta">${escapeHtml(gesture.originalMetaText)}</span>`);
    }
  } else if (meta) {
    meta.remove();
  }
}

function commitPreviewGestureTime(gesture) {
  const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.minuteOffset);
  if (shifted.source === "Custom") {
    customEvents = customEvents.map((item) => item.id === shifted.id && normalizeEmail(item.ownerEmail) === activeCalendarEmail() ? previewEventToCustomEvent(shifted, item) : item);
  } else {
    syncImportedOverride(shifted.id, {
      start: shifted.start,
      end: shifted.end,
      allDay: shifted.allDay,
    });
  }
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Event time updated.");
}

function setPreviewChipMeta(chip, event, forceTime = false) {
  const metaParts = [];
  if ((!event.allDay && settings.showTimes) || forceTime) metaParts.push(event.timeLabel);
  const text = metaParts.join(" · ");
  let meta = chip.querySelector(".preview-chip-meta");
  if (!text) {
    meta?.remove();
    return;
  }
  if (!meta) {
    meta = document.createElement("span");
    meta.className = "preview-chip-meta";
    chip.append(meta);
  }
  meta.textContent = text;
}

function shiftTimedEventByMinutes(event, minutes) {
  const start = addMinutesToDateTimeString(event.start, minutes);
  const end = addMinutesToDateTimeString(event.end, minutes);
  return {
    ...event,
    start,
    end,
    timeLabel: summarizeEventTimes(start, end, false),
  };
}

function addMinutesToDateTimeString(value, minutes) {
  const date = parseDateTimeString(value);
  date.setMinutes(date.getMinutes() + minutes);
  return formatDateTimeString(date);
}

function parseDateTimeString(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (!match) return new Date(value);
  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    0,
    0,
  );
}

function formatDateTimeString(date) {
  return `${formatDateKey(date)}T${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}:00`;
}

function openContextMenu(x, y, items) {
  contextMenu.innerHTML = items.map((item, index) => `<button type="button" class="context-menu-item" data-context-index="${index}">${escapeHtml(item.label)}</button>`).join("");
  contextMenu.dataset.items = JSON.stringify(items.map((item) => item.label));
  contextMenu.style.left = `${x}px`;
  contextMenu.style.top = `${y}px`;
  contextMenu.classList.remove("hidden");
  contextMenu.setAttribute("aria-hidden", "false");
  contextMenu.onclick = (event) => {
    const button = event.target.closest("[data-context-index]");
    if (!button) return;
    const item = items[Number(button.dataset.contextIndex)];
    closeContextMenu();
    item.action();
  };
}

function closeContextMenu() {
  contextMenu.classList.add("hidden");
  contextMenu.setAttribute("aria-hidden", "true");
  contextMenu.innerHTML = "";
}

function selectedDoctor() {
  const options = doctorPickerOptions();
  if (!options.length) return null;
  if (options.length === 1) return options[0];
  const preferredDoctorKey = preferredDoctorKeyForCurrentAccount();
  return options.find((doctor) => doctor.key === doctorSelect.value)
    || options.find((doctor) => doctor.key === preferredDoctorKey)
    || options[0];
}

function preferredDoctorKeyForCurrentAccount() {
  if (activeDoctorProfile?.doctorKey) return activeDoctorProfile.doctorKey;
  if (activeCalendarMode() === "claimed-account" && currentRosterClaims.length) return currentRosterClaims[0].key;
  if (currentUserEmail === OWNER_EMAIL && !adminViewingEmail && !activeDoctorProfile) return OWNER_DOCTOR_KEY;
  return "";
}

function preferredDoctorKeyForAccountEmail(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail) return "";
  if (targetEmail === OWNER_EMAIL) return OWNER_DOCTOR_KEY;
  const serverUser = serverUsers.map(normalizeServerUser).find((user) => user.email === targetEmail);
  const claims = sanitizeRosterClaims(serverUser?.claims || []);
  if (claims.length) return claims[0].key;
  const localUser = accountState.users.find((user) => normalizeEmail(user.email) === targetEmail);
  const localClaims = sanitizeRosterClaims(localUser?.claims || []);
  if (localClaims.length) return localClaims[0].key;
  const claimedDoctor = availableRosterDoctors.find((doctor) => normalizeEmail(doctor.accountEmail || doctor.claimedBy || "") === targetEmail);
  return normalizeRosterName(claimedDoctor?.key || "");
}

function claimedEmailForDoctorKey(doctorKey, displayName = "") {
  const normalizedKey = normalizeRosterName(doctorKey);
  const normalizedDisplayName = String(displayName || "").trim();
  if (!normalizedKey) return "";
  for (const user of serverUsers) {
    const claims = sanitizeRosterClaims(user?.claims || []);
    if (claims.some((claim) => claim.key === normalizedKey)) {
      return currentClaimedAccountEmail(user.email);
    }
  }
  if (normalizedDisplayName) {
    for (const user of serverUsers) {
      const claims = sanitizeRosterClaims(user?.claims || []);
      if (claims.some((claim) => likelySameRosterName(claim.displayName, normalizedDisplayName))) {
        return currentClaimedAccountEmail(user.email);
      }
      if (likelySameRosterName(user?.realName || "", normalizedDisplayName)) {
        return currentClaimedAccountEmail(user.email);
      }
    }
  }
  const claimedDoctor = availableRosterDoctors.find((doctor) => doctor.key === normalizedKey && doctor.claimedBy);
  return currentClaimedAccountEmail(claimedDoctor?.claimedBy || "");
}

function canUseDoctorPicker() {
  return isViewingCreatorAccount();
}

function canUseCreatorDoctorSwitcher() {
  return Boolean(isCreatorAuthenticated() && (canUseDoctorPicker() || adminViewingEmail || activeCalendarMode() === "doctor-profile"));
}

function canReturnToCreator() {
  return Boolean(isCreatorAuthenticated() && (adminViewingEmail || activeDoctorProfile));
}

function initialCalendarContext() {
  const email = normalizeEmail(currentUserEmail);
  return {
    mode: email === OWNER_EMAIL ? "creator-account" : "claimed-account",
    email,
    profile: null,
  };
}

function setActiveCalendarContext(mode, details = {}) {
  activeCalendarContext = {
    mode,
    email: normalizeEmail(details.email || currentUserEmail),
    profile: details.profile || null,
  };
}

function activeCalendarMode() {
  return activeCalendarContext?.mode || (activeDoctorProfile ? "doctor-profile" : currentUserEmail === OWNER_EMAIL && !adminViewingEmail ? "creator-account" : "claimed-account");
}

function rememberCreatorCalendarSourceRefs() {
  if (activeCalendarMode() !== "creator-account") return;
  const refs = sanitizeClientFileRefs(currentSnapshot?.fileRefs?.length ? currentSnapshot.fileRefs : selectedFiles.map(importRefForWorkspace));
  if (refs.length) creatorCalendarSourceFileRefs = refs;
}

function showSwitchOverlay(title, message) {
  if (!switchOverlay) return;
  switchOverlayTitle.textContent = title || "Switching…";
  switchOverlayMessage.textContent = message || "Loading calendar…";
  switchOverlay.classList.remove("hidden");
  switchOverlay.setAttribute("aria-hidden", "false");
}

function hideSwitchOverlay() {
  if (!switchOverlay) return;
  switchOverlay.classList.add("hidden");
  switchOverlay.setAttribute("aria-hidden", "true");
}

async function switchDoctorSelection(selectedKey, options = {}) {
  const resetRange = options.resetRange !== false;
  const normalizedSelectedKey = normalizeRosterName(selectedKey);
  doctorSelect.value = selectedKey;
  const canSwitchAsCreator = canUseCreatorDoctorSwitcher();
  if (canSwitchAsCreator && cloudAvailable && (!serverUsers.length || !availableRosterDoctors.length)) {
    await loadServerUsers();
  }
  const selectedOption = selectedDoctorOptionForKey(selectedKey);
  if (canSwitchAsCreator && normalizedSelectedKey === OWNER_DOCTOR_KEY) {
    try {
      showSwitchOverlay("Returning to creator...", "Restoring the creator calendar.");
      await returnToCreatorCalendar();
    } finally {
      hideSwitchOverlay();
    }
    return;
  }
  let resolvedAccount = null;
  if (canSwitchAsCreator && selectedOption) {
    try {
      resolvedAccount = await resolveDoctorAccountForSwitch(selectedOption);
    } catch (error) {
      setStatus(error.message || "Could not check whether that calendar is claimed.", true);
      return;
    }
  }
  const claimedEmail = normalizeEmail(resolvedAccount?.email || selectedOption?.accountEmail || claimedEmailForDoctorKey(selectedKey, selectedOption?.displayName || ""));
  if (canSwitchAsCreator && selectedOption) {
    showSwitchOverlay(
      `Switching to ${selectedOption.displayName}…`,
      resolvedAccount?.mode === "claimed-account" ? "Opening the linked account calendar." : "Opening the roster calendar and loading saved doctor-profile edits.",
    );
  }
  if (canSwitchAsCreator && selectedOption) {
    try {
      if (resolvedAccount?.mode === "claimed-account" && claimedEmail && claimedEmail !== currentUserEmail) {
        await enterUserAccount(claimedEmail);
      } else if (resolvedAccount?.mode === "claimed-account" && claimedEmail && claimedEmail === currentUserEmail && adminViewingEmail) {
        await enterUserAccount(claimedEmail);
      } else {
        await enterDoctorProfileView(selectedOption);
      }
    } finally {
      hideSwitchOverlay();
    }
    return;
  }
  clearPreviewData();
  saveCurrentSessionState();
  syncActionState();
  if (selectedDoctor()) await updatePreview({ resetRange });
}

async function resolveDoctorAccountForSwitch(doctor) {
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "resolveDoctorAccount",
      email: authUserEmail || currentUserEmail,
      password: authUserPassword || currentUserPassword,
      doctor: {
        key: doctor?.key || "",
        displayName: doctor?.displayName || "",
        sourceType: doctor?.sourceType || "",
        sourceTypes: normalizedDoctorSourceTypes(doctor),
        aliases: Array.isArray(doctor?.aliases) ? doctor.aliases : [],
      },
    }),
  });
  const data = await readJsonResponse(response, "Could not check whether that calendar is claimed.");
  const mode = data.mode === "claimed-account" && normalizeEmail(data.email) ? "claimed-account" : "doctor-profile";
  return {
    mode,
    email: mode === "claimed-account" ? normalizeEmail(data.email) : "",
  };
}

function selectedDoctorOptionForKey(selectedKey) {
  const normalizedKey = normalizeRosterName(selectedKey);
  const localOption = doctorPickerOptions().find((doctor) => doctor.key === normalizedKey) || null;
  const selectedDomOption = doctorSelect.selectedOptions?.[0] || null;
  if (localOption && (normalizedDoctorSourceTypes(localOption).length || selectedDomOption?.dataset.sourceTypes)) {
    const accountEmail = currentClaimedAccountEmail(localOption.accountEmail || selectedDomOption?.dataset.accountEmail || "");
    return {
      ...localOption,
      accountEmail,
      targetMode: accountEmail ? "claimed-account" : "doctor-profile",
      sourceTypes: normalizedDoctorSourceTypes(localOption).length
        ? normalizedDoctorSourceTypes(localOption)
        : String(selectedDomOption?.dataset.sourceTypes || "").split(",").map((item) => item.trim().toLowerCase()).filter(Boolean),
    };
  }
  if (!selectedDomOption || normalizeRosterName(selectedDomOption.value) !== normalizedKey) {
    if (!localOption) return null;
    const accountEmail = currentClaimedAccountEmail(localOption.accountEmail || "");
    return {
      ...localOption,
      accountEmail,
      targetMode: accountEmail ? "claimed-account" : "doctor-profile",
    };
  }
  const sourceTypes = String(selectedDomOption.dataset.sourceTypes || "")
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);
  const accountEmail = currentClaimedAccountEmail(selectedDomOption.dataset.accountEmail || "");
  return {
    key: normalizedKey,
    displayName: selectedDomOption.dataset.displayName || selectedDomOption.textContent.trim(),
    sourceTypes,
    accountEmail,
    targetMode: accountEmail ? "claimed-account" : "doctor-profile",
  };
}

function currentClaimedAccountEmail(email) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return "";
  if (normalizedEmail === currentUserEmail && (adminViewingEmail || currentUserEmail !== OWNER_EMAIL)) return normalizedEmail;
  if (serverUsers.map(normalizeServerUser).some((user) => user.email === normalizedEmail)) return normalizedEmail;
  return availableRosterDoctors.some((doctor) => normalizeEmail(doctor.accountEmail || doctor.claimedBy || "") === normalizedEmail)
    ? normalizedEmail
    : "";
}

function activeWorkspaceOwnerKey() {
  if (activeCalendarMode() === "doctor-profile" && activeDoctorProfile?.id) return `profile:${activeDoctorProfile.id}`;
  return viewedAccountEmail();
}

function activeCalendarOwnerId() {
  if (activeCalendarMode() === "doctor-profile") return activeDoctorProfile?.ownerId || "";
  return viewedAccountEmail();
}

function buildDoctorProfileId(doctor) {
  const key = normalizeRosterName(doctor?.key || "");
  const sources = doctorProfileSourceTypes(doctor).sort();
  return key && sources.length ? `${key}::${sources.join("+")}` : "";
}

function normalizedDoctorSourceTypes(doctor) {
  const values = Array.isArray(doctor?.sourceTypes) ? doctor.sourceTypes : [];
  if (doctor?.sourceType) values.push(doctor.sourceType);
  return [...new Set(values.map((item) => String(item || "").toLowerCase()).filter((item) => item === "mmc" || item === "ddh" || item === "casey" || item === "mch"))];
}

function doctorProfileSourceTypes(doctor) {
  const sourceTypes = normalizedDoctorSourceTypes(doctor);
  return sourceTypes.length ? sourceTypes : ["casey", "ddh", "mch", "mmc"];
}

function doctorOptionsForCurrentAccount(doctors) {
  const options = (doctors || []).map((doctor) => ({
    ...doctor,
    sourceTypes: normalizedDoctorSourceTypes(doctor),
  }));
  if (canUseCreatorDoctorSwitcher()) {
    const repositoryOptions = availableRosterDoctors.map((doctor) => ({
      ...doctor,
      sourceTypes: normalizedDoctorSourceTypes(doctor).length ? normalizedDoctorSourceTypes(doctor) : [String(doctor.sourceType || "").toLowerCase()].filter(Boolean),
    }));
    return buildCreatorDoctorOptions(dedupeDoctorOptions([...options, ...repositoryOptions]));
  }
  const claimedAliases = currentRosterClaims.map((claim) => ({
    sourceType: claim.sourceType,
    key: claim.key,
    displayName: claim.displayName,
  }));
  const claimMatches = currentRosterClaims.length
    ? options.filter((doctor) => currentRosterClaims.some((claim) => doctorMatchesRosterClaim(doctor, claim)))
    : [];
  const nameMatches = options.filter((doctor) => !claimMatches.includes(doctor) && doctorMatchesCurrentAccount(doctor));
  const matches = claimMatches.length ? claimMatches : nameMatches;
  if (!matches.length) return [];
  const aliases = [...claimedAliases, ...matches.flatMap((doctor) => {
    if (Array.isArray(doctor.aliases) && doctor.aliases.length) return doctor.aliases;
    const sourceTypes = doctor.sourceTypes.length ? doctor.sourceTypes : sourceTypesForClaimedDoctor(doctor.key);
    return sourceTypes.map((sourceType) => ({
      sourceType,
      key: doctor.key,
      displayName: doctor.displayName,
    }));
  })];
  const dedupedAliases = dedupeDoctorAliases(aliases);
  const primary = dedupedAliases[0] || matches[0];
  const displayName = currentAccount().realName || primary.displayName || matches[0].displayName;
  return [{
    key: primary.key || matches[0].key,
    displayName,
    aliases: dedupedAliases,
    sourceTypes: [...new Set(dedupedAliases.map((alias) => alias.sourceType))],
  }];
}

function doctorPickerOptions() {
  if (!canUseCreatorDoctorSwitcher()) return doctorOptions;
  const repositoryOptions = availableRosterDoctors.map((doctor) => ({
    ...doctor,
    sourceTypes: normalizedDoctorSourceTypes(doctor).length
      ? normalizedDoctorSourceTypes(doctor)
      : [String(doctor.sourceType || "").toLowerCase()].filter(Boolean),
  }));
  const preferredDoctorKey = preferredDoctorKeyForCurrentAccount();
  const preferredDoctor = preferredDoctorKey ? [{
    key: preferredDoctorKey,
    displayName: preferredDoctorKey === OWNER_DOCTOR_KEY
      ? formatRosterDisplayName(OWNER_DOCTOR_KEY)
      : formatRosterDisplayName(preferredDoctorKey),
    sourceTypes: [],
    ownerRoute: preferredDoctorKey === OWNER_DOCTOR_KEY,
  }] : [];
  const fallbackOptions = repositoryOptions.length ? [] : doctorOptions;
  return buildCreatorDoctorOptions(dedupeDoctorOptions([...preferredDoctor, ...repositoryOptions, ...fallbackOptions]));
}

function dedupeDoctorOptions(options) {
  const dedupedByIdentity = new Map();
  for (const doctor of options || []) {
    if (!doctor?.key) continue;
    const identity = doctorIdentityKey(doctor);
    if (!identity) continue;
    const existing = dedupedByIdentity.get(identity);
    dedupedByIdentity.set(identity, existing ? mergeDoctorOption(existing, doctor) : { ...doctor });
  }
  return [...dedupedByIdentity.values()];
}

function doctorIdentityKey(doctor) {
  if (doctor?.key === OWNER_DOCTOR_KEY || rosterIdentityKey(doctor?.displayName || doctor?.key) === rosterIdentityKey(OWNER_DOCTOR_KEY)) {
    return `owner:${OWNER_DOCTOR_KEY}`;
  }
  return rosterIdentityKey(doctor?.displayName || doctor?.key) || normalizeRosterName(doctor?.key || "");
}

function mergeDoctorOption(existing, incoming) {
  const existingAliases = Array.isArray(existing.aliases) ? existing.aliases : [];
  const incomingAliases = Array.isArray(incoming.aliases) ? incoming.aliases : [];
  const sourceTypes = [...new Set([
    ...normalizedDoctorSourceTypes(existing),
    ...normalizedDoctorSourceTypes(incoming),
  ])];
  const ownerRoute = existing.ownerRoute === true || incoming.ownerRoute === true;
  const preferred = ownerRoute
    ? (existing.ownerRoute ? existing : incoming.ownerRoute ? incoming : existing)
    : existing;
  return {
    ...existing,
    ...preferred,
    key: ownerRoute ? OWNER_DOCTOR_KEY : (preferred.key || existing.key || incoming.key),
    displayName: ownerRoute ? formatRosterDisplayName(OWNER_DOCTOR_KEY) : (preferred.displayName || existing.displayName || incoming.displayName),
    sourceType: preferred.sourceType || existing.sourceType || incoming.sourceType || "",
    sourceTypes,
    aliases: dedupeDoctorAliases([...existingAliases, ...incomingAliases]),
    accountEmail: ownerRoute ? "" : (existing.accountEmail || incoming.accountEmail || ""),
    claimedBy: ownerRoute ? "" : (existing.claimedBy || incoming.claimedBy || ""),
    claimedByName: ownerRoute ? "" : (existing.claimedByName || incoming.claimedByName || ""),
    ownerRoute,
  };
}

function buildCreatorDoctorOptions(options) {
  return prioritizeDoctorOptions(
    options.map((doctor) => ({
      ...doctor,
      accountEmail: doctor.key === OWNER_DOCTOR_KEY || doctor.ownerRoute ? "" : claimedEmailForDoctorKey(doctor.key, doctor.displayName),
    })),
  );
}

function prioritizeDoctorOptions(options) {
  return [...options].sort((left, right) => {
    const leftOwner = left.key === OWNER_DOCTOR_KEY ? 1 : 0;
    const rightOwner = right.key === OWNER_DOCTOR_KEY ? 1 : 0;
    if (leftOwner !== rightOwner) return rightOwner - leftOwner;
    return left.displayName.localeCompare(right.displayName);
  });
}

function doctorMatchesCurrentAccount(doctor) {
  const claimKeys = new Set(currentRosterClaims.map((claim) => claim.key));
  const claimIdentityKeys = new Set(currentRosterClaims.flatMap((claim) => [rosterIdentityKey(claim.displayName), rosterIdentityKey(claim.key)]).filter(Boolean));
  if (claimKeys.has(doctor.key)) return true;
  if (claimIdentityKeys.has(rosterIdentityKey(doctor.displayName || doctor.key))) return true;
  return likelySameRosterName(currentAccount().realName, doctor.displayName);
}

function doctorMatchesRosterClaim(doctor, claim) {
  const normalizedClaim = {
    sourceType: String(claim?.sourceType || "").toLowerCase(),
    key: normalizeRosterName(claim?.key || ""),
    displayName: String(claim?.displayName || "").trim(),
  };
  if (!normalizedClaim.sourceType || !normalizedClaim.key) return false;
  const aliases = Array.isArray(doctor?.aliases) && doctor.aliases.length
    ? doctor.aliases
    : normalizedDoctorSourceTypes(doctor).map((sourceType) => ({
        sourceType,
        key: doctor?.key || "",
        displayName: doctor?.displayName || "",
      }));
  return aliases.some((alias) => {
    const aliasSource = String(alias?.sourceType || "").toLowerCase();
    if (aliasSource && aliasSource !== normalizedClaim.sourceType) return false;
    const aliasKey = normalizeRosterName(alias?.key || "");
    if (aliasKey && aliasKey === normalizedClaim.key) return true;
    return rosterIdentityKey(alias?.displayName || aliasKey) === rosterIdentityKey(normalizedClaim.displayName || normalizedClaim.key);
  });
}

function sourceTypesForClaimedDoctor(key) {
  return currentRosterClaims.filter((claim) => claim.key === key).map((claim) => claim.sourceType);
}

function dedupeDoctorAliases(aliases) {
  const seen = new Set();
  return aliases.filter((alias) => {
    if (!alias.sourceType || !alias.key) return false;
    const marker = `${alias.sourceType}:${alias.key}`;
    if (seen.has(marker)) return false;
    seen.add(marker);
    return true;
  });
}

function nameTokenMatch(left, right) {
  const leftTokens = rosterNameTokens(left);
  const rightTokens = rosterNameTokens(right);
  if (leftTokens.length < 2 || rightTokens.length < 2) return false;
  const rightSet = new Set(rightTokens);
  return leftTokens.every((token) => rightSet.has(token));
}

function likelySameRosterName(left, right) {
  if (nameTokenMatch(left, right)) return true;
  const leftTokens = rosterNameTokens(left);
  const rightTokens = rosterNameTokens(right);
  if (leftTokens.length < 2 || rightTokens.length < 2) return false;
  if (leftTokens.at(-1) !== rightTokens.at(-1)) return false;
  const leftFirst = leftTokens[0] || "";
  const rightFirst = rightTokens[0] || "";
  return leftFirst.length >= 3 && rightFirst.length >= 3 && (leftFirst.startsWith(rightFirst) || rightFirst.startsWith(leftFirst));
}

function rosterNameTokens(value) {
  return rosterIdentityKey(value)
    .split(/\s+/)
    .filter(Boolean);
}

function rosterIdentityKey(value) {
  const raw = String(value || "");
  const stripped = raw
    .replace(/[^A-Za-z0-9,]+/g, " ")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/^(DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A PROF|ASSOC PROF)\s+/, "");
  const parts = stripped.split(/\s*,\s*/).filter(Boolean);
  return parts.length === 2 ? `${parts[1]} ${parts[0]}`.trim() : stripped.replace(/,/g, "");
}

function formatRosterDisplayName(value) {
  const identity = rosterIdentityKey(value);
  const tokens = identity.split(" ").filter(Boolean);
  if (!tokens.length) return String(value || "").trim();
  return tokens.map((token, index) => index === tokens.length - 1 ? token : toDisplayNameToken(token)).join(" ");
}

function toDisplayNameToken(value) {
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
}

function doctorMetadataForKey(doctorKey) {
  if (insightDoctorRoleCache.size) return insightDoctorRoleCache.get(doctorKey) || { any: { role: "" } };
  if (!doctorRoleIndex) doctorRoleIndex = buildDoctorRoleIndex();
  return doctorRoleIndex.get(doctorKey) || { any: { role: "" } };
}

function buildDoctorRoleIndex() {
  const index = new Map();
  if (!parsedRosterSources) return index;
  for (const entry of parsedRosterSources.mmc || []) {
    collectMmcDoctorRoles(entry.workbook, index);
  }
  for (const entry of parsedRosterSources.ddh || []) {
    collectDdhDoctorRoles(entry.workbook, index);
  }
  for (const entry of parsedRosterSources.casey || []) {
    collectCaseyDoctorRoles(entry.workbook, index);
  }
  for (const entry of parsedRosterSources.mch || []) {
    collectMchDoctorRoles(entry.workbook, index);
  }
  return index;
}

function collectMmcDoctorRoles(workbook, index) {
  const roleMap = new Map([
    ["SMS", "SMS"],
    ["CMO", "CMO"],
    ["SENIOR REG", "SR"],
    ["INTERMEDIATE REG", "IR"],
    ["JUNIOR REG", "JR"],
    ["HMO", "HMO"],
    ["HMO MUST BE 111", "HMO"],
    ["HMO - MUST BE 111", "HMO"],
    ["ENP", "ENP"],
    ["AMP", "AMP"],
    ["INTERN", "I"],
  ]);
  for (const sheetName of workbook?.SheetNames || []) {
    if (!String(sheetName).startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    const range = decodeSheetRange(sheet);
    let currentRole = "";
    for (let row = 1; row <= range.e.r + 1; row += 1) {
      const marker = cleanSheetCell(sheet, row, 3).replace(/\s+/g, " ").trim().toUpperCase();
      if (roleMap.has(marker)) currentRole = roleMap.get(marker);
      const name = cleanSheetCell(sheet, row, 4);
      if (!looksLikeRosterPerson(name)) continue;
      assignDoctorRole(index, normalizeRosterName(name), "MMC", currentRole);
    }
  }
}

function collectDdhDoctorRoles(workbook, index) {
  const sectionMap = new Map([
    ["SENIOR MEDICAL STAFF", "SMS"],
    ["SENIOR REGISTRARS", "SR"],
    ["REGISTRAR", "SR"],
    ["REGISTRARS", "SR"],
    ["CMO'S", "CMO"],
    ["CMOS", "CMO"],
    ["JUNIOR REGISTRARS", "JR"],
    ["ED HMO'S", "HMO"],
    ["HMO'S", "HMO"],
    ["INTERNS", "I"],
    ["ENP", "ENP"],
    ["NURSE PRACTITIONERS", "ENP"],
    ["NURSE PRAC. CANDIDATES", "ENP"],
    ["AMP", "AMP"],
    ["AMP'S", "AMP"],
    ["PHYSIOTHERAPIST", "AMP"],
    ["PHYSIOTHERAPISTS", "AMP"],
  ]);
  const sheet = workbook?.Sheets?.[workbook?.SheetNames?.[0]];
  if (!sheet) return;
  const range = decodeSheetRange(sheet);
  let currentRole = "";
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    const value = cleanSheetCell(sheet, row, 1).replace(/\s+/g, " ").trim();
    if (!value) continue;
    const upper = value.toUpperCase();
    if (sectionMap.has(upper)) {
      currentRole = sectionMap.get(upper);
      continue;
    }
    if (isDdhHmoSectionHeading(upper)) {
      currentRole = "HMO";
      continue;
    }
    if (!looksLikeRosterPerson(value)) continue;
    assignDoctorRole(index, normalizeRosterName(value), "DDH", currentRole);
  }
}

function collectCaseyDoctorRoles(workbook, index) {
  const sectionMap = new Map([
    ["GERIATRICIAN", "SMS"],
    ["SNR CMOS", "CMO"],
    ["SENIOR CMOS", "CMO"],
    ["ED GPS", "CMO"],
    ["SNR REGS", "SR"],
    ["SENIOR REGS", "SR"],
    ["JNR REGS", "JR"],
    ["JUNIOR REGS", "JR"],
    ["PAEDS ED HMOS", "HMO"],
    ["HMOS", "HMO"],
    ["INTERNS", "I"],
    ["AMP COVER", "AMP"],
  ]);
  for (const sheetName of workbook?.SheetNames || []) {
    if (sheetName === "Personal Roster" || sheetName === "DayRoster") continue;
    const sheet = workbook.Sheets[sheetName];
    if (!/^TERM\s+\d+,/i.test(cleanSheetCell(sheet, 1, 1))) continue;
    const range = decodeSheetRange(sheet);
    let currentRole = "SMS";
    for (let row = 3; row <= range.e.r + 1; row += 1) {
      const value = cleanSheetCell(sheet, row, 1).replace(/\s+/g, " ").trim();
      if (!value) continue;
      const upper = value.toUpperCase();
      if (upper === "ROSTERED STAFF") break;
      if (sectionMap.has(upper)) {
        currentRole = sectionMap.get(upper);
        continue;
      }
      if (upper.startsWith("ONCALL")) continue;
      const parsed = value.match(/^\(([^)]+)\)\s*(.+)$/);
      const name = (parsed?.[2] || value).replace(/\s+-\s+LOCUM(?:\s+SMS)?$/i, "").replace(/\s+\(7\/fn\)$/i, "").trim();
      const role = parsed ? roleCodeFromRosterPrefix(parsed[1]) : currentRole;
      if (!looksLikeRosterPerson(name)) continue;
      assignDoctorRole(index, normalizeRosterName(name), "CASEY", role);
    }
  }
}

function collectMchDoctorRoles(workbook, index) {
  for (const sheetName of workbook?.SheetNames || []) {
    if (!/^Week\s+\d+$/i.test(sheetName) || sheetName === "Week 0") continue;
    const sheet = workbook.Sheets[sheetName];
    if (!String(cleanSheetCell(sheet, 4, 4)).toUpperCase().includes("PAEDIATRIC EMERGENCY DEPARTMENT ROSTER")) continue;
    const ranges = [
      [21, 49],
      [54, 84],
      [86, 102],
    ];
    for (const [startRow, endRow] of ranges) {
      for (let row = startRow; row <= endRow; row += 1) {
        const name = cleanSheetCell(sheet, row, 4).replace(/\s*\([^)]*\)\s*$/g, "").replace(/\s+/g, " ").trim();
        if (!looksLikeRosterPerson(name) || /^locum$/i.test(name)) continue;
        assignDoctorRole(index, normalizeRosterName(name), "MCH", roleCodeForMchRole(cleanSheetCell(sheet, row, 1)));
      }
    }
  }
}

function roleCodeForMchRole(value) {
  const upper = String(value || "").toUpperCase();
  if (upper.includes("CONSULTANT") || upper.includes("STAFF SPECIALIST")) return "SMS";
  if (upper.includes("FELLOW")) return "SR";
  if (upper.includes("REGISTRAR")) return "JR";
  if (upper.includes("HMO")) return "HMO";
  if (upper.includes("INTERN")) return "I";
  return "";
}

function roleCodeFromRosterPrefix(prefix) {
  const upper = String(prefix || "").trim().toUpperCase();
  if (upper === "SR") return "SR";
  if (upper === "TR") return "IR";
  if (upper === "JR") return "JR";
  if (upper === "I") return "I";
  if (upper === "H" || upper === "HMO" || upper === "ED HMO" || upper === "PAEDS HMO") return "HMO";
  if (upper === "GERI") return "SMS";
  if (upper === "AMP") return "AMP";
  return "";
}

function assignDoctorRole(index, doctorKey, source, role) {
  if (!doctorKey) return;
  if (!index.has(doctorKey)) index.set(doctorKey, { any: { role: "" } });
  const entry = index.get(doctorKey);
  if (!entry[source]) entry[source] = { role: "" };
  if (!entry[source].role && role) entry[source].role = role;
  if (!entry.any.role && role) entry.any.role = role;
}

function decodeSheetRange(sheet) {
  return XLSX.utils.decode_range(sheet?.["!ref"] || "A1:A1");
}

function cleanSheetCell(sheet, row, col) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  return String(sheet?.[address]?.v ?? "").trim();
}

function looksLikeRosterPerson(value) {
  const cleaned = String(value || "").trim();
  if (cleaned.length < 5 || !cleaned.includes(" ") || /^\d/.test(cleaned)) return false;
  if (/^(WEEK|DATE|HMO|SMS|CMO|ENP|AMP|INTERN|SENIOR|JUNIOR|REGISTRAR|GERIATRICIAN)/i.test(cleaned)) return false;
  return /[A-Za-z]/.test(cleaned);
}

function isDdhHmoSectionHeading(value) {
  return /^ED HMO/i.test(value) || /^HMO\b/i.test(value);
}

function resetDerivedState(options = {}) {
  const preserveSession = options.preserveSession === true;
  const preservedAvailableDoctors = isCreatorAuthenticated() ? availableRosterDoctors : [];
  const preservedSession = preserveSession ? buildActiveSessionState() : null;
  doctorOptions = [];
  detectedSources = {};
  availableRosterDoctors = preservedAvailableDoctors;
  overrides = preserveSession ? sanitizeOverrideState(preservedSession?.overrides) : {};
  customEvents = preserveSession ? sanitizeActiveCalendarCustomEvents(preservedSession?.customEvents) : [];
  restoredSessionState = preserveSession ? preservedSession : null;
  doctorRoleIndex = null;
  activeDoctorProfile = null;
  undoHistory = [];
  redoHistory = [];
  lastHistorySignature = "";
  clearDoctorAnalysisCache();
  resetVisibleInsightWarmCache();
  lastRosterPersistence = null;
  closeInsightsModal();
  settings = defaultSettings();
  renderSettings();
  doctorSelect.innerHTML = "";
  doctorName.textContent = "";
  doctorName.classList.add("hidden");
  doctorSelect.classList.add("hidden");
  doctorSection.classList.add("hidden");
  controlBar.classList.add("hidden");
  mobileActionBar.classList.add("hidden");
  closeSettingsPanel({ commit: false });
  claimSection.classList.add("hidden");
  clearPreviewData();
}

function resetTransientCalendarData() {
  parsedRosterSources = null;
  parsedImportDoctors = new Map();
  doctorRoleIndex = null;
  restoredSessionState = null;
  clearDoctorAnalysisCache();
  resetVisibleInsightWarmCache();
  clearPreviewData();
}

function clearPreviewData() {
  latestPreview = null;
  reviewIndex = new Map();
  currentPreviewEvents = new Map();
  availablePreviewHospitals = [];
  reportedIssueFingerprints = new Set();
  document.body.classList.remove("has-calendar-preview");
  issuesPanel.classList.add("hidden");
  conflictsPanel.classList.add("hidden");
  previewSection.classList.add("hidden");
  preview.innerHTML = "";
  preview.classList.add("hidden");
  issuesList.innerHTML = "";
  conflictsList.innerHTML = "";
  closeReviewModal();
  closeCustomEventModal();
  closeContextMenu();
}

function ensureOverride(id) {
  if (!overrides[id]) overrides[id] = {};
  return overrides[id];
}

function removeEmptyOverride(id) {
  const value = overrides[id];
  if (!value) return;
  const hasLocation = Object.prototype.hasOwnProperty.call(value, "location");
  if (!value.title && typeof value.include !== "boolean" && !value.start && !value.end && !hasLocation && typeof value.allDay !== "boolean") {
    delete overrides[id];
  }
}

function getBaseImportedEvent(id) {
  return latestPreview?.events?.find((event) => event.id === id) || null;
}

function syncImportedOverride(id, patch) {
  const baseEvent = getBaseImportedEvent(id);
  const reviewItem = reviewIndex.get(id);
  const next = ensureOverride(id);

  const nextTitle = patch.title ?? next.title ?? "";
  const baseTitle = reviewItem?.suggestedTitle || baseEvent?.title || "";
  next.title = nextTitle && nextTitle !== baseTitle ? nextTitle : "";

  if (typeof patch.include === "boolean") {
    next.include = patch.include !== (reviewItem?.include ?? true) ? patch.include : undefined;
  }

  const nextStart = patch.start ?? next.start ?? "";
  next.start = nextStart && nextStart !== baseEvent?.start ? nextStart : "";

  const nextEnd = patch.end ?? next.end ?? "";
  next.end = nextEnd && nextEnd !== baseEvent?.end ? nextEnd : "";

  const nextAllDay = typeof patch.allDay === "boolean" ? patch.allDay : next.allDay;
  next.allDay = typeof nextAllDay === "boolean" && nextAllDay !== baseEvent?.allDay ? nextAllDay : undefined;

  if (Object.prototype.hasOwnProperty.call(patch, "location") || Object.prototype.hasOwnProperty.call(next, "location")) {
    const nextLocation = Object.prototype.hasOwnProperty.call(patch, "location") ? patch.location : next.location;
    if ((nextLocation || "") !== (baseEvent?.location || "")) {
      next.location = nextLocation || "";
    } else {
      delete next.location;
    }
  }

  removeEmptyOverride(id);
}

function fileFingerprint(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

function previewInclusiveEndDate(event, startDate, endDate) {
  if (event.allDay) {
    return addDays(endDate, -1);
  }
  if (endDate > startDate) {
    return startDate;
  }
  if (endDate <= startDate) {
    return startDate;
  }
  const endClock = extractTimePortion(event.end);
  if (endClock === "00:00") {
    return addDays(endDate, -1);
  }
  return endDate;
}

function parseDateOnly(value) {
  const [year, month, day] = value.slice(0, 10).split("-").map(Number);
  return new Date(year, month - 1, day);
}

function mondayFor(date) {
  const copy = new Date(date);
  const day = copy.getDay();
  const delta = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + delta);
  return copy;
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function formatDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatLongDate(date) {
  return date.toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function formatDate(value) {
  return parseDateOnly(value).toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function formatMonth(date) {
  return date.toLocaleDateString("en-AU", {
    month: "long",
    year: "numeric",
  });
}

function isCurrentDay(date) {
  const today = new Date();
  return date.getFullYear() === today.getFullYear()
    && date.getMonth() === today.getMonth()
    && date.getDate() === today.getDate();
}

function deriveRangeBounds(events) {
  if (!events.length) return { start: "", end: "" };
  let start = events[0].start.slice(0, 10);
  let end = formatDateKey(previewInclusiveEndDate(events[0], parseDateOnly(events[0].start), parseDateOnly(events[0].end)));
  for (const event of events.slice(1)) {
    const eventStart = event.start.slice(0, 10);
    const eventEnd = formatDateKey(previewInclusiveEndDate(event, parseDateOnly(event.start), parseDateOnly(event.end)));
    if (eventStart < start) start = eventStart;
    if (eventEnd > end) end = eventEnd;
  }
  return { start, end };
}

function deriveDefaultPreviewRange(events) {
  const today = new Date();
  const currentTerm = australianTermForDate(today);
  const defaultStartDate = defaultPreviewStartForTerm(currentTerm, today);
  const defaultStart = formatDateKey(defaultStartDate);
  const currentTermEnd = formatDateKey(addDays(currentTerm.end, -1));
  const eventRange = deriveRangeBounds(events || []);
  const latestEventTerm = eventRange.end ? australianTermForDate(parseDateOnly(eventRange.end)) : currentTerm;
  const latestEventTermEnd = formatDateKey(addDays(latestEventTerm.end, -1));
  return {
    start: defaultStart,
    end: maxDateKey(latestEventTermEnd, currentTermEnd),
  };
}

function shouldApplyDefaultPreviewRange(options, events) {
  if (options.resetRange || !settings.dateFrom || !settings.dateTo) return true;
  const defaultRange = deriveDefaultPreviewRange(events || []);
  if (!defaultRange.start || !defaultRange.end) return false;
  if (settings.dateFrom === defaultRange.start && settings.dateTo === defaultRange.end) return true;
  const eventRange = deriveRangeBounds(events || []);
  return Boolean(eventRange.start && eventRange.end && settings.dateFrom === eventRange.start && settings.dateTo === eventRange.end);
}

function boundedPreviewStart(value, defaultStart) {
  if (!defaultStart) return value || "";
  if (!value) return defaultStart;
  return value;
}

function boundedPreviewEnd(value, defaultEnd) {
  if (!defaultEnd) return value || "";
  if (!value) return defaultEnd;
  return maxDateKey(value, defaultEnd);
}

function maxDateKey(left, right) {
  if (!left) return right || "";
  if (!right) return left || "";
  return left > right ? left : right;
}

function eventOverlapsDateRange(event, rangeStart, rangeEnd) {
  const eventStart = parseDateOnly(event.start);
  const eventEnd = previewInclusiveEndDate(event, eventStart, parseDateOnly(event.end));
  return eventStart <= rangeEnd && eventEnd >= rangeStart;
}

function filterEventsByPreviewRange(events, start, end) {
  if (!start && !end) return events;
  const rangeStart = start ? parseDateOnly(start) : null;
  const rangeEnd = end ? parseDateOnly(end) : null;
  return events.filter((event) => {
    const eventStart = parseDateOnly(event.start);
    const eventEnd = previewInclusiveEndDate(event, eventStart, parseDateOnly(event.end));
    if (rangeStart && eventEnd < rangeStart) return false;
    if (rangeEnd && eventStart > rangeEnd) return false;
    return true;
  });
}

function formatPreviewRange(start, end) {
  if (!start || !end) return "";
  return `${start} to ${end}`;
}

function openPreviewRangePicker(which) {
  const input = preview.querySelector(`[data-range-input="${which}"]`);
  if (!input) return;
  if (typeof input.showPicker === "function") {
    input.showPicker();
    return;
  }
  input.focus({ preventScroll: true });
  input.click();
}

function applyPreviewRangeChange(which, value) {
  if (!value) return;
  if (which === "from") {
    settings.dateFrom = value;
    if (settings.dateTo && settings.dateTo < value) settings.dateTo = value;
  } else {
    settings.dateTo = value;
    if (settings.dateFrom && settings.dateFrom > value) settings.dateFrom = value;
  }
  if (settingsInputs.dateFrom) settingsInputs.dateFrom.value = settings.dateFrom;
  if (settingsInputs.dateTo) settingsInputs.dateTo.value = settings.dateTo;
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Preview range updated.");
}

function applyPreviewTermStart(value) {
  if (!value) return;
  settings.dateFrom = value;
  if (settings.dateTo && settings.dateTo < value) settings.dateTo = value;
  if (settingsInputs.dateFrom) settingsInputs.dateFrom.value = settings.dateFrom;
  if (settingsInputs.dateTo) settingsInputs.dateTo.value = settings.dateTo;
  pendingPreviewSnapToToday = true;
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Preview start term updated.");
}

function defaultPreviewStartForTerm(term, today = new Date()) {
  const oneMonthAfterTermStart = addDays(term.start, 28);
  if (today >= term.start && today < oneMonthAfterTermStart) {
    return addDays(term.start, -28);
  }
  return term.start;
}

function targetForCurrentPreviewMonth() {
  const todayKey = formatDateKey(new Date());
  const todayCell = preview.querySelector(`[data-add-date="${todayKey}"]`);
  if (todayCell) {
    const term = todayCell.closest(".preview-term");
    let rowPointer = todayCell;
    let weekLabel = null;
    while (rowPointer) {
      if (rowPointer.classList?.contains("preview-week-label")) {
        weekLabel = rowPointer;
        break;
      }
      rowPointer = rowPointer.previousElementSibling;
    }
    let monthRow = weekLabel?.previousElementSibling || rowPointer?.previousElementSibling || null;
    while (monthRow && !monthRow.classList?.contains("preview-month-row")) {
      monthRow = monthRow.previousElementSibling;
    }
    const grid = todayCell.closest(".preview-grid");
    const firstMonthRow = grid?.querySelector(".preview-month-row") || null;
    const target = monthRow
      ? (monthRow === firstMonthRow ? term?.querySelector(".preview-term-header") || monthRow : monthRow)
      : term?.querySelector(".preview-term-header") || weekLabel || todayCell;
    return monthRow || target;
  }
  const todayMonthKey = todayKey.slice(0, 7);
  return preview.querySelector(`[data-month-key="${todayMonthKey}"]`);
}

function scrollPreviewTarget(target, smooth = true) {
  if (!target) return false;
  if (isMobileLayout()) {
    syncMobileViewportInsets();
    const rawAnchor = getComputedStyle(document.documentElement).getPropertyValue("--mobile-top-anchor");
    const anchor = Number.parseFloat(rawAnchor) || 36;
    const nextTop = Math.max(0, window.scrollY + target.getBoundingClientRect().top - anchor);
    window.scrollTo({ top: nextTop, behavior: smooth ? "smooth" : "auto" });
    return true;
  }
  const scroller = previewSection;
  const header = preview.querySelector(".preview-head");
  if (!scroller) return false;
  const bannerBottom = header?.getBoundingClientRect().bottom || 0;
  const targetTop = target.getBoundingClientRect().top;
  const nextTop = Math.max(0, scroller.scrollTop + (targetTop - bannerBottom));
  scroller.scrollTo({ top: nextTop, behavior: smooth ? "smooth" : "auto" });
  return true;
}

function snapPreviewToCurrentMonth(smooth = true) {
  scrollPreviewTarget(targetForCurrentPreviewMonth(), smooth);
}

function buildEventOverridePatch(event, item, override = {}) {
  if (!event) return null;
  const allDay = typeof override.allDay === "boolean" ? override.allDay : event.allDay;
  const start = override.start || event.start;
  const end = override.end || event.end;
  const title = (override.title || "").trim() || item?.suggestedTitle || event.title;
  const location = Object.prototype.hasOwnProperty.call(override, "location") ? override.location || "" : event.location;
  const sourceTitle = item?.suggestedTitle || event.title;
  return {
    ...event,
    title,
    start,
    end,
    allDay,
    location: location || "",
    timeLabel: summarizeEventTimes(start, end, allDay),
    isEditedImport: (
      title !== sourceTitle ||
      start !== event.start ||
      end !== event.end ||
      allDay !== event.allDay ||
      (location || "") !== (event.location || "")
    ),
  };
}

function detectAustralianTerm(date) {
  return { label: australianTermForDate(date).label };
}

function australianTermForDate(date) {
  const year = date.getFullYear();
  const candidates = [
    buildAustralianTerm(year, 1, 1),
    buildAustralianTerm(year, 2, 4),
    buildAustralianTerm(year, 3, 7),
    buildAustralianTerm(year, 4, 10),
    buildAustralianTerm(year - 1, 4, 10),
  ];
  const match = candidates.find((term) => date >= term.start && date < term.end);
  return match || buildAustralianTerm(year, 1, 1);
}

function nextAustralianTerm(term) {
  const nextTermNumber = term.termNumber === 4 ? 1 : term.termNumber + 1;
  const nextYear = term.termNumber === 4 ? term.year + 1 : term.year;
  return buildAustralianTerm(nextYear, nextTermNumber, startMonthIndexForTerm(nextTermNumber));
}

function buildAustralianTerm(year, termNumber, startMonthIndex) {
  const start = firstMondayOfMonth(year, startMonthIndex);
  const end = addDays(start, 91);
  return {
    label: `Term ${termNumber}`,
    year,
    termNumber,
    start,
    end,
  };
}

function formatAustralianTermLabel(term) {
  return `Term ${term.termNumber} ${term.year}`;
}

function startMonthIndexForTerm(termNumber) {
  return [1, 4, 7, 10][Math.max(0, Math.min(3, termNumber - 1))];
}

function firstMondayOfMonth(year, monthIndex) {
  const date = new Date(year, monthIndex, 1);
  const day = date.getDay();
  const delta = day === 0 ? 1 : day === 1 ? 0 : 8 - day;
  date.setDate(date.getDate() + delta);
  return date;
}

function formatTimestamp(value) {
  if (!value) return "-";
  const formatted = new Date(value).toLocaleString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
  return formatted.replace(/,\s0(\d:\d{2}\s?pm)$/i, ", $1");
}

function formatIssueHeading(item) {
  const status = item.status === "unknown"
    ? "Unknown"
    : item.status === "deleted"
      ? "Deleted"
      : "Review";
  return `${status} · ${item.source} · ${formatDate(item.startDay || item.date)}`;
}

function parserRuleCodeForIssue(issue) {
  const code = String(issue?.code || "").trim().toUpperCase();
  if (code) return code;
  return parserRuleCodeFromRawValue(issue?.source, issue?.rawValue);
}

function parserRuleCodeFromRawValue(sourceValue, rawValue) {
  const source = sanitizeIssueSource(sourceValue);
  const text = String(rawValue || "").trim();
  const upper = text.toUpperCase();
  if (source === "DDH") return normalizeDdhParserRuleCodeText(text);
  const timedCode = parserRuleCodeFromTimedRawValue(source, upper);
  if (timedCode) return timedCode;
  return upper;
}

function parserRuleCodeFromTimedRawValue(source, value) {
  const prefixMatch = String(value || "").match(/^\s*(\d{2}):?(\d{2})\s*[-–]\s*(\d{2}):?(\d{2})(?:\s*(.+?))?\s*$/);
  if (prefixMatch) {
    const label = String(prefixMatch[5] || "").trim().toUpperCase();
    if (label) return label;
    return inferParserRulePeriodCode([Number(prefixMatch[1]), Number(prefixMatch[2])], source);
  }
  const suffixMatch = String(value || "").match(/^\s*(.+?)\s+(\d{2}):?(\d{2})\s*[-–]\s*(\d{2}):?(\d{2})\s*$/);
  return suffixMatch ? suffixMatch[1].trim().toUpperCase() : "";
}

function inferParserRulePeriodCode(startHm, source = "") {
  const [hour] = startHm;
  if (hour >= 22 || hour < 6) return "NIGHT";
  if (source === "MCH" ? hour >= 12 : hour >= 14) return "PM";
  return "AM";
}

function normalizeDdhParserRuleCodeText(value) {
  const text = String(value || "").trim();
  const upper = text.toUpperCase();
  const aliases = new Map([
    ["CLINICAL SUPPORT", "CS"],
    ["SSU SMS", "SSU"],
    ["ORANGE PM (ON-CALL)", "ORANGE PM"],
    ["PM FAST IC", "FAST PM"],
    ["ORANGE AM IC", "ORANGE AM"],
    ["ONSITE CS", "CS ONSITE"],
  ]);
  return aliases.get(upper) || upper;
}

function parserRulePeriodForIssue(issue) {
  const title = String(issue?.suggestedTitle || "").trim();
  const match = title.match(/\b(AM|PM|NIGHT)\b/i);
  return match ? match[1].toUpperCase() : "";
}

function parserRuleBaseForIssue(issue) {
  const source = sanitizeIssueSource(issue?.source);
  let title = String(issue?.suggestedTitle || "").trim();
  if (!title) return "";
  if (source && title.toUpperCase().startsWith(`${source}: `)) {
    title = title.slice(source.length + 2).trim();
  }
  title = title.replace(/\b(AM|PM|NIGHT)\b/gi, " ").replace(/\bFLOAT\b/gi, " ").replace(/\s+/g, " ").trim();
  return title;
}

function parserRuleSuffixForIssue(issue) {
  const title = String(issue?.suggestedTitle || "").trim();
  if (/\bFLOAT\b/i.test(title)) return "Float";
  return "";
}

function timeRangeParts(value) {
  const match = String(value || "").match(/(\d{2}:\d{2})-(\d{2}:\d{2})/);
  return {
    start: match?.[1] || "09:00",
    end: match?.[2] || "17:00",
  };
}

function defaultLocationForIssueSource(source) {
  if (sanitizeIssueSource(source) === "MMC") return settings.defaultLocationMmc || DEFAULT_MMC_LOCATION;
  if (sanitizeIssueSource(source) === "DDH") return settings.defaultLocationDdh || DEFAULT_DDH_LOCATION;
  if (sanitizeIssueSource(source) === "Casey") return settings.defaultLocationCasey || DEFAULT_CASEY_LOCATION;
  if (sanitizeIssueSource(source) === "MCH") return settings.defaultLocationMch || DEFAULT_MCH_LOCATION;
  return "";
}

function comparePreviewEvents(left, right) {
  const leftDate = left.start.slice(0, 10);
  const rightDate = right.start.slice(0, 10);
  if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
  if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
  return left.title.localeCompare(right.title);
}

function summarizeEvents(events) {
  const first = events[0].start.slice(0, 10);
  const lastEvent = events[events.length - 1];
  const last = lastEvent.allDay ? addDays(parseDateOnly(lastEvent.end), -1) : parseDateOnly(lastEvent.end);
  return `${first} to ${formatDateKey(last)}`;
}

function summarizeEventTimes(start, end, allDay) {
  if (allDay) return "All day";
  return `${extractTimePortion(start)}-${extractTimePortion(end)}`;
}

function openReviewModal(id, selectedDay = "") {
  const previewEvent = currentPreviewEvents.get(id);
  if (isCustomPreviewEvent(previewEvent)) {
    const customEvent = ensureEditableCustomEvent(previewEvent);
    if (customEvent) {
      openCustomEventModal(customEvent, selectedDay || customEvent.startDate);
    }
    return;
  }
  const item = reviewIndex.get(id);
  if (!item) {
    const customEvent = customEventsForActiveCalendar().find((entry) => entry.id === id);
    if (customEvent) {
      openCustomEventModal(customEvent, selectedDay || customEvent.startDate);
    }
    return;
  }
  openReviewId = id;
  const event = currentPreviewEvents.get(id) || buildEventOverridePatch(getBaseImportedEvent(id), item, overrides[id] || {});
  const overrideValue = escapeHtml((overrides[id]?.title ?? item.overrideTitle ?? ""));
  const includeValue = typeof overrides[id]?.include === "boolean" ? overrides[id].include : item.include;
  const startDate = event?.start?.slice(0, 10) || item.startDay;
  const endDate = event?.allDay
    ? formatDateKey(addDays(parseDateOnly(event.end), -1))
    : event?.end?.slice(0, 10) || item.endDay;
  const insightDate = selectedDay || startDate;
  const allDay = event?.allDay ?? item.allDay;
  const startTime = event?.allDay ? "" : extractTimePortion(event?.start || "");
  const endTime = event?.allDay ? "" : extractTimePortion(event?.end || "");
  const preset = detectLocationPreset(event?.location || item.location || "");
  const warnings = item.warnings.length
    ? `<ul class="review-warnings">${item.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("")}</ul>`
    : "";
  const badge = item.status === "ok" ? "" : `<span class="review-badge review-badge-${item.status}">${item.status}</span>`;
  const resetButton = hasImportedOverride(id)
    ? `<div class="modal-actions"><button type="button" class="button button-secondary" data-override-reset="${id}">Reset Imported Event</button></div>`
    : "";

  reviewModalBody.innerHTML = `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>${escapeHtml(item.source)} · ${formatDate(item.startDay)}</strong>
          <span>${escapeHtml(item.rawValue)}</span>
        </div>
        ${badge}
      </div>
      <div class="review-body">
        <label class="field">
          <span>Normalised result</span>
          <input
            type="text"
            value="${overrideValue}"
            placeholder="${escapeHtml(item.suggestedTitle)}"
            data-override-title="${item.id}"
          >
        </label>
        <label class="toggle review-toggle">
          <input type="checkbox" ${includeValue ? "checked" : ""} ${item.exportable ? "" : "disabled"} data-override-include="${item.id}">
          Include in export
        </label>
        <div class="custom-event-grid event-date-grid">
          <label class="field">
            <span>Start date</span>
            <input type="date" value="${startDate}" data-override-start-date="${item.id}">
          </label>
          <label class="field">
            <span>End date</span>
            <input type="date" value="${endDate}" data-override-end-date="${item.id}">
          </label>
        </div>
        <label class="toggle review-toggle">
          <input type="checkbox" ${allDay ? "checked" : ""} data-override-all-day="${item.id}">
          All day
        </label>
        <div class="custom-event-grid event-time-grid ${allDay ? "hidden" : ""}" data-override-time-fields="${item.id}">
          <label class="field">
            <span>Start time</span>
            <input type="text" inputmode="numeric" value="${formatEditorTime(startTime || "09:00")}" data-override-start-time="${item.id}">
          </label>
          <label class="field">
            <span>End time</span>
            <input type="text" inputmode="numeric" value="${formatEditorTime(endTime || "17:00")}" data-override-end-time="${item.id}">
          </label>
        </div>
        <div class="custom-event-grid">
          <label class="field">
            <span>Location</span>
            <select data-override-location-mode="${item.id}">
              ${buildLocationOptionMarkup(preset.mode, item.source)}
            </select>
          </label>
          <label class="field ${preset.mode === "custom" ? "" : "hidden"}" data-override-custom-location-field="${item.id}">
            <span>Custom location</span>
            <input type="text" value="${escapeHtml(preset.customValue)}" data-override-custom-location="${item.id}">
          </label>
        </div>
      </div>
      <div class="review-meta">
        <span>Suggested title: ${escapeHtml(item.suggestedTitle || "No normalised result")}</span>
        ${item.timeLabel ? `<span>Times: ${escapeHtml(item.timeLabel)}</span>` : ""}
        ${item.location ? `<span>Location: ${escapeHtml(item.location)}</span>` : ""}
      </div>
      ${canUseRosterInsights() && !isLeaveEvent(event) ? `<section class="event-inline-insight" data-review-who-panel aria-live="polite"></section>` : ""}
      ${resetButton}
      ${warnings}
    </article>
  `;
  reviewModal.classList.remove("hidden");
  reviewModal.setAttribute("aria-hidden", "false");
  if (canUseRosterInsights() && !isLeaveEvent(event)) {
    void renderInlineWhoInsight(reviewModalBody.querySelector("[data-review-who-panel]"), insightDate, { source: eventSourceCode(event) });
  }
}

function closeReviewModal() {
  openReviewId = "";
  reviewModal.classList.add("hidden");
  reviewModal.setAttribute("aria-hidden", "true");
  reviewModalBody.innerHTML = "";
}

function openCustomEventModal(event = null, presetDate = null, options = {}) {
  populateLocationOptions();
  const now = presetDate || latestPreview?.events?.[0]?.start?.slice(0, 10) || formatDateKey(new Date());
  customEventId.value = event?.id || "";
  customEventTitle.value = event?.title || "";
  customEventStartDate.value = event?.startDate || now;
  customEventEndDate.value = event?.endDate || event?.startDate || now;
  customEventAllDay.checked = event?.allDay ?? false;
  customEventStartTime.value = formatEditorTime(event?.startTime || "09:00");
  customEventEndTime.value = formatEditorTime(event?.endTime || "10:00");
  const preset = detectLocationPreset(event?.location || "");
  customEventLocationMode.value = preset.mode;
  customEventCustomLocation.value = preset.customValue;
  customEventCustomLocationField.classList.toggle("hidden", preset.mode !== "custom");
  customEventTimeFields.classList.toggle("hidden", customEventAllDay.checked);
  customEventDeleteButton.classList.toggle("hidden", !event || options.draft === true);
  if (customEventWhoPanel) {
    // Custom events are not rostered shifts, so coworker lookup is not meaningful here.
    customEventWhoPanel.innerHTML = "";
    customEventWhoPanel.classList.add("hidden");
  }
  customEventModal.classList.remove("hidden");
  customEventModal.setAttribute("aria-hidden", "false");
}

function closeCustomEventModal() {
  customEventModal.classList.add("hidden");
  customEventModal.setAttribute("aria-hidden", "true");
  customEventForm.reset();
  customEventDeleteButton.classList.add("hidden");
  customEventCustomLocationField.classList.add("hidden");
  customEventTimeFields.classList.remove("hidden");
  if (customEventWhoPanel) {
    customEventWhoPanel.classList.add("hidden");
    customEventWhoPanel.innerHTML = "";
  }
}

function populateLocationOptions() {
  customEventLocationMode.innerHTML = buildLocationOptionMarkup();
}

function buildLocationOptionMarkup(selectedMode = "", source = "") {
  const options = [];
  const sourceTypes = new Set(locationOptionSourceTypes(source));
  if (detectedSources.mmc?.length || sourceTypes.has("mmc")) options.push({ value: "mmc", label: "MMC Car Park" });
  if (detectedSources.ddh?.length || sourceTypes.has("ddh")) options.push({ value: "ddh", label: "DDH Car Park" });
  if (detectedSources.casey?.length || sourceTypes.has("casey")) options.push({ value: "casey", label: "Casey Hospital" });
  if (detectedSources.mch?.length || sourceTypes.has("mch")) options.push({ value: "mch", label: "MCH" });
  options.push({ value: "offsite", label: "Off-site" });
  options.push({ value: "custom", label: "Custom location" });
  return options.map((option) => `<option value="${option.value}" ${option.value === selectedMode ? "selected" : ""}>${option.label}</option>`).join("");
}

function locationOptionSourceTypes(source = "") {
  const explicit = String(source || "").trim().toLowerCase();
  if (["mmc", "ddh", "casey", "mch"].includes(explicit)) return [explicit];
  return recognizedHospitalTypesForActiveAccount();
}

function detectLocationPreset(location) {
  if (!location) return { mode: "offsite", customValue: "" };
  if (location === settings.defaultLocationMmc || location === DEFAULT_MMC_LOCATION) return { mode: "mmc", customValue: "" };
  if (location === settings.defaultLocationDdh || location === DEFAULT_DDH_LOCATION) return { mode: "ddh", customValue: "" };
  if (location === settings.defaultLocationCasey || location === DEFAULT_CASEY_LOCATION) return { mode: "casey", customValue: "" };
  if (location === settings.defaultLocationMch || location === DEFAULT_MCH_LOCATION) return { mode: "mch", customValue: "" };
  return { mode: "custom", customValue: location };
}

function readCustomEventForm() {
  const title = customEventTitle.value.trim();
  const startDate = customEventStartDate.value;
  const endDate = customEventEndDate.value || startDate;
  const allDay = customEventAllDay.checked;
  const startTime = parseEditorTimeInput(customEventStartTime.value);
  const endTime = parseEditorTimeInput(customEventEndTime.value);
  const location = resolveCustomEventLocation();

  if (!title) {
    setStatus("Manual events need a title.", true);
    return null;
  }
  if (!startDate || !endDate) {
    setStatus("Manual events need a start and end date.", true);
    return null;
  }
  if (!allDay && (!startTime || !endTime)) {
    setStatus("Timed manual events need both a start and end time.", true);
    return null;
  }

  return {
    id: customEventId.value || newCustomEventId(),
    ownerEmail: activeCalendarEmail(),
    title,
    startDate,
    endDate,
    allDay,
    startTime,
    endTime,
    location,
    include: true,
  };
}

function newCustomEventId() {
  if (globalThis.crypto?.randomUUID) return `custom-${globalThis.crypto.randomUUID()}`;
  return `custom-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function resolveCustomEventLocation() {
  if (customEventLocationMode.value === "mmc") return settings.defaultLocationMmc || DEFAULT_MMC_LOCATION;
  if (customEventLocationMode.value === "ddh") return settings.defaultLocationDdh || DEFAULT_DDH_LOCATION;
  if (customEventLocationMode.value === "casey") return settings.defaultLocationCasey || DEFAULT_CASEY_LOCATION;
  if (customEventLocationMode.value === "mch") return settings.defaultLocationMch || DEFAULT_MCH_LOCATION;
  if (customEventLocationMode.value === "custom") return customEventCustomLocation.value.trim();
  return "";
}

function applyImportedEventFormState(id) {
  const baseEvent = getBaseImportedEvent(id);
  const startDate = reviewModalBody.querySelector(`[data-override-start-date="${id}"]`)?.value || baseEvent?.start?.slice(0, 10) || "";
  const rawEndDate = reviewModalBody.querySelector(`[data-override-end-date="${id}"]`)?.value || startDate;
  const endDateInput = rawEndDate < startDate ? startDate : rawEndDate;
  const allDay = reviewModalBody.querySelector(`[data-override-all-day="${id}"]`)?.checked ?? baseEvent?.allDay ?? false;
  const startTime = parseEditorTimeInput(reviewModalBody.querySelector(`[data-override-start-time="${id}"]`)?.value)
    || extractTimePortion(baseEvent?.start || "")
    || "09:00";
  const endTime = parseEditorTimeInput(reviewModalBody.querySelector(`[data-override-end-time="${id}"]`)?.value)
    || extractTimePortion(baseEvent?.end || "")
    || "17:00";
  const timeFields = reviewModalBody.querySelector(`[data-override-time-fields="${id}"]`);
  if (timeFields) timeFields.classList.toggle("hidden", allDay);

  const endDate = !allDay && compareClockStrings(endTime, startTime) <= 0 && endDateInput === startDate
    ? formatDateKey(addDays(parseDateOnly(startDate), 1))
    : endDateInput;

  syncImportedOverride(id, {
    start: allDay ? startDate : `${startDate}T${startTime}:00`,
    end: allDay ? formatDateKey(addDays(parseDateOnly(endDateInput), 1)) : `${endDate}T${endTime}:00`,
    allDay,
    location: resolveImportedLocation(id),
  });

  const customLocationField = reviewModalBody.querySelector(`[data-override-custom-location-field="${id}"]`);
  if (customLocationField) {
    const mode = reviewModalBody.querySelector(`[data-override-location-mode="${id}"]`)?.value || "offsite";
    customLocationField.classList.toggle("hidden", mode !== "custom");
  }
}

function resolveImportedLocation(id) {
  const mode = reviewModalBody.querySelector(`[data-override-location-mode="${id}"]`)?.value || "offsite";
  if (mode === "mmc") return settings.defaultLocationMmc || DEFAULT_MMC_LOCATION;
  if (mode === "ddh") return settings.defaultLocationDdh || DEFAULT_DDH_LOCATION;
  if (mode === "casey") return settings.defaultLocationCasey || DEFAULT_CASEY_LOCATION;
  if (mode === "mch") return settings.defaultLocationMch || DEFAULT_MCH_LOCATION;
  if (mode === "custom") {
    return reviewModalBody.querySelector(`[data-override-custom-location="${id}"]`)?.value.trim() || "";
  }
  return "";
}

function extractTimePortion(value) {
  const match = String(value).match(/T(\d{2}:\d{2})/);
  return match ? match[1] : "";
}

function formatEditorTime(value) {
  const parsed = parseClockParts(value);
  if (!parsed) return "";
  const suffix = parsed.hours >= 12 ? "pm" : "am";
  const hour12 = parsed.hours % 12 || 12;
  const hourLabel = suffix === "am" && hour12 < 10 ? String(hour12).padStart(2, "0") : String(hour12);
  return `${hourLabel}:${String(parsed.minutes).padStart(2, "0")} ${suffix}`;
}

function parseEditorTimeInput(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return "";
  const compact = text.replace(/\s+/g, "");
  const twelveHour = compact.match(/^(\d{1,2})(?::?(\d{2}))?(am|pm)$/);
  if (twelveHour) {
    let hours = Number(twelveHour[1]);
    const minutes = Number(twelveHour[2] || 0);
    if (hours < 1 || hours > 12 || minutes > 59) return "";
    if (twelveHour[3] === "pm" && hours !== 12) hours += 12;
    if (twelveHour[3] === "am" && hours === 12) hours = 0;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  const twentyFourHour = compact.match(/^(\d{1,2})(?::?(\d{2}))$/);
  if (!twentyFourHour) return "";
  const hours = Number(twentyFourHour[1]);
  const minutes = Number(twentyFourHour[2]);
  if (hours > 23 || minutes > 59) return "";
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function parseClockParts(value) {
  const match = String(value || "").match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours > 23 || minutes > 59) return null;
  return { hours, minutes };
}

function isClockString(value) {
  return /^\d{2}:\d{2}$/.test(String(value || "").trim());
}

function compareClockStrings(left, right) {
  return left.localeCompare(right);
}

function diffDays(start, end) {
  return Math.round((end.getTime() - start.getTime()) / 86400000);
}

function summarizeDetectedSources(imports) {
  return {
    mmc: imports.filter((item) => item.sourceType === "mmc").map((item) => item.name),
    ddh: imports.filter((item) => item.sourceType === "ddh").map((item) => item.name),
    casey: imports.filter((item) => item.sourceType === "casey").map((item) => item.name),
    mch: imports.filter((item) => item.sourceType === "mch").map((item) => item.name),
  };
}

function latestImportTimestamp() {
  if (!selectedFiles.length) return "";
  return selectedFiles.reduce((latest, entry) => !latest || (entry.addedAt || "") > latest ? entry.addedAt || "" : latest, "");
}

function openFilesModal() {
  renderFilesList();
  filesModal.classList.remove("hidden");
  filesModal.setAttribute("aria-hidden", "false");
  if (isCreatorAuthenticated() && cloudAvailable) {
    void refreshCalendarStoreStatus({ silent: true });
  }
}

function closeFilesModal() {
  filesModal.classList.add("hidden");
  filesModal.setAttribute("aria-hidden", "true");
}

function openExportModal() {
  if (!selectedFiles.length) {
    setStatus("Add at least one roster file first.", true);
    return;
  }
  if (!selectedDoctor()) {
    setStatus("Choose a doctor before exporting.", true);
    return;
  }
  setPendingExportMode("full");
  pendingExportRange = defaultExportRangeState();
  pendingExportHospitals = exportHospitalOptions();
  renderExportModal();
  exportModal.classList.remove("hidden");
  exportModal.setAttribute("aria-hidden", "false");
}

function closeExportModal() {
  exportModal.classList.add("hidden");
  exportModal.setAttribute("aria-hidden", "true");
  exportModalBody.innerHTML = "";
}

function setPendingExportMode(mode) {
  pendingExportMode = mode === "range" ? "range" : "full";
}

function renderExportModal() {
  const doctor = selectedDoctor();
  const hospitalOptions = exportHospitalOptions();
  const canCopySubscription = canCopySubscriptionUrl();
  exportModalBody.innerHTML = `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>${escapeHtml(doctor?.displayName || "Selected doctor")}</strong>
          <span>Save a one-off export, or open/copy a live subscription feed.</span>
        </div>
      </div>
      <div class="export-variant-picker">
        <button type="button" class="button ${pendingExportMode === "full" ? "button-primary" : "button-secondary"}" data-export-mode="full">Full calendar</button>
        <button type="button" class="button ${pendingExportMode === "range" ? "button-primary" : "button-secondary"}" data-export-mode="range">Date range</button>
      </div>
      ${pendingExportMode === "range" ? `
        <div class="export-range-panel">
          <label class="toggle">
            <input type="checkbox" data-export-all-future ${pendingExportRange.allFuture ? "checked" : ""}>
            All future events
          </label>
          <div class="settings-subgrid${pendingExportRange.allFuture ? "" : " event-date-grid"}">
            <label class="field">
              <span>Start date</span>
              <input type="date" value="${escapeHtml(pendingExportRange.startDate)}" data-export-range-input="start">
            </label>
            ${pendingExportRange.allFuture ? "" : `
              <label class="field">
                <span>Finish date</span>
                <input type="date" value="${escapeHtml(pendingExportRange.endDate)}" data-export-range-input="end" min="${escapeHtml(pendingExportRange.startDate || "")}">
              </label>
            `}
          </div>
        </div>
      ` : ""}
      ${hospitalOptions.length > 1 ? `
        <div class="export-hospital-panel">
          <span>Hospitals</span>
          <div class="export-hospital-picker">
            ${hospitalOptions.map((hospital) => `
              <button type="button" class="button ${pendingExportHospitals.includes(hospital) ? "button-primary" : "button-secondary"}" data-export-hospital="${escapeHtml(hospital)}">${escapeHtml(displaySourceCode(hospital))}</button>
            `).join("")}
          </div>
        </div>
      ` : ""}
      <div class="export-actions-grid">
        <button type="button" class="button button-primary" data-export-action="download">Save as .ics file</button>
        <button type="button" class="button button-secondary" data-export-action="apple">Open in Apple Calendar</button>
        <button type="button" class="button button-secondary" data-export-action="copy" ${canCopySubscription ? "" : "disabled"}>Copy URL</button>
      </div>
      ${canCopySubscription ? "" : `<p class="status">Subscription URLs are available only for claimed user accounts and the creator account, not unclaimed doctor profiles.</p>`}
    </article>
  `;
}

async function handleExportAction(action) {
  const doctor = selectedDoctor();
  if (!doctor) {
    setStatus("Choose a doctor before exporting.", true);
    return;
  }
  const exportConfig = exportConfigForMode(pendingExportMode, pendingExportRange);
  if (exportConfig.mode === "range" && !exportConfig.startDate) {
    setStatus("Choose a start date for the export range.", true);
    return;
  }
  if (exportConfig.mode === "range" && !exportConfig.allFuture && !exportConfig.endDate) {
    setStatus("Choose a finish date or use all future events.", true);
    return;
  }
  if (exportConfig.mode === "range" && !exportConfig.allFuture && exportConfig.endDate < exportConfig.startDate) {
    setStatus("The finish date must be on or after the start date.", true);
    return;
  }
  if (action === "download" && exportHospitalOptions().length > 1 && !pendingExportHospitals.length) {
    setStatus("Choose at least one hospital for this export.", true);
    return;
  }
  try {
    if (currentSnapshotStale) {
      setStatus("Refreshing calendar...");
      await refreshSnapshotInBackground();
    }
    if (action === "download") {
      setStatus("Building calendar file...");
      const events = await buildBrowserExportEvents(doctor, { ...exportConfig, hospitals: pendingExportHospitals });
      if (!events.length) {
        throw new Error("No calendar events were found for that export selection.");
      }
      const ics = exportIcs(events, doctor.displayName);
      downloadIcs(ics, `${doctor.displayName} roster.ics`);
      closeExportModal();
      setStatus("Calendar file ready.");
      return;
    }
    if (action === "apple") {
      if (!canCopySubscriptionUrl()) {
        setStatus("Subscription links are not available for this calendar.", true);
        return;
      }
      setStatus("Saving subscription feed...");
      const snapshot = snapshotCloudSavePayload();
      snapshot.session = {
        ...snapshot.session,
        exportRange: normalizeExportRangeState(exportConfig.mode === "range" ? exportConfig : defaultExportRangeState()),
      };
      await saveCloudState(snapshot);
      const url = subscriptionUrl("webcal", exportConfig.mode === "range" ? "range" : "full");
      if (!url) throw new Error("No subscription link is available for this account yet.");
      closeExportModal();
      window.location.href = url;
      setStatus("Opening Apple Calendar subscription...");
      return;
    }
    if (!canCopySubscriptionUrl()) {
      setStatus("Subscription links are not available for this calendar.", true);
      return;
    }
    const url = subscriptionUrl("https", exportConfig.mode === "range" ? "range" : "full");
    if (!url) throw new Error("No subscription link is available for this account yet.");
    await navigator.clipboard.writeText(url);
    closeExportModal();
    setStatus("Subscription URL copied.");
    const snapshot = snapshotCloudSavePayload();
    snapshot.session = {
      ...snapshot.session,
      exportRange: normalizeExportRangeState(exportConfig.mode === "range" ? exportConfig : defaultExportRangeState()),
    };
    saveCloudState(snapshot).catch(() => setStatus("Subscription URL copied, but the saved feed range could not be updated.", true));
  } catch (error) {
    setStatus(error.message || "Export failed.", true);
  }
}

function downloadIcs(ics, filename) {
  const payload = new Blob([ics], { type: "text/calendar; charset=utf-8" });
  const url = URL.createObjectURL(payload);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function hasActiveExportFilters() {
  if (settings.hospitalFilter && settings.hospitalFilter !== "all") return true;
  if (!latestPreview?.events?.length) return false;
  const range = deriveRangeBounds(latestPreview.events);
  return Boolean(
    (settings.dateFrom && range.start && settings.dateFrom !== range.start)
    || (settings.dateTo && range.end && settings.dateTo !== range.end),
  );
}

function closeAccountsModal() {
  accountsModal.classList.add("hidden");
  accountsModal.setAttribute("aria-hidden", "true");
}

function loadAccountState() {
  try {
    const stored = JSON.parse(localStorage.getItem(ACCOUNT_STATE_KEY) || "null");
    if (stored && Array.isArray(stored.users) && stored.currentEmail) {
      return {
        ...stored,
        users: stored.users.map((user) => ({
          email: normalizeEmail(user.email),
          realName: "",
          claims: [],
          ...user,
          email: normalizeEmail(user.email),
          role: user.role || (normalizeEmail(user.email) === OWNER_EMAIL ? "owner" : "user"),
          claims: sanitizeRosterClaims(user.claims),
        })).filter((user) => user.email),
      };
    }
  } catch {
    // Ignore invalid local state.
  }
  return {
    currentEmail: OWNER_EMAIL,
    users: [
      { email: OWNER_EMAIL, realName: "Richard Haydon", password: "", role: "owner", claims: [] },
    ],
  };
}

function saveAccountState() {
  localStorage.setItem(ACCOUNT_STATE_KEY, JSON.stringify(accountState));
  syncAccountsButton();
  renderAccountsModal();
}

function ensureLocalAccountLogin(email, password, options = {}) {
  const realName = String(options.realName || "").trim();
  const existing = accountState.users.find((user) => user.email === email);
  if (!existing) {
    accountState.users.push({
      email,
      realName,
      password,
      role: email === OWNER_EMAIL ? "owner" : "user",
      claims: [],
    });
  } else {
    existing.password = password || existing.password || "";
    existing.role = existing.role || (email === OWNER_EMAIL ? "owner" : "user");
    if (realName) existing.realName = realName;
  }
  accountState.currentEmail = email;
  saveAccountState();
}

function sanitizeRosterClaims(claims) {
  if (!Array.isArray(claims)) return [];
  return claims
    .map((claim) => ({
      key: normalizeRosterName(claim?.key || ""),
      displayName: String(claim?.displayName || "").trim(),
      sourceType: String(claim?.sourceType || "").toLowerCase(),
      matchedAt: String(claim?.matchedAt || ""),
    }))
    .filter((claim) => claim.key && claim.displayName && claim.sourceType);
}

function sanitizeAvailableRosterDoctors(doctors) {
  if (!Array.isArray(doctors)) return [];
  return doctors
    .map((doctor) => ({
      key: normalizeRosterName(doctor?.key || ""),
      displayName: formatRosterDisplayName(doctor?.displayName || doctor?.key || ""),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
      sourceTypes: Array.isArray(doctor?.sourceTypes) ? doctor.sourceTypes.map((item) => String(item || "").toLowerCase()).filter(Boolean) : [],
      aliases: Array.isArray(doctor?.aliases)
        ? doctor.aliases.map((alias) => ({
            key: normalizeRosterName(alias?.key || ""),
            displayName: formatRosterDisplayName(alias?.displayName || alias?.key || ""),
            sourceType: String(alias?.sourceType || "").toLowerCase(),
          })).filter((alias) => alias.key && alias.displayName && alias.sourceType)
        : [],
      claimedBy: normalizeEmail(doctor?.claimedBy || ""),
      claimedByName: String(doctor?.claimedByName || "").trim(),
      accountEmail: normalizeEmail(doctor?.accountEmail || doctor?.claimedBy || ""),
    }))
    .filter((doctor) => doctor.key && doctor.displayName && doctor.sourceType);
}

function normalizeRosterName(value) {
  return String(value || "")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function currentAccount() {
  const email = viewedAccountEmail() || accountState.currentEmail;
  const serverAccount = serverUsers.map(normalizeServerUser).find((user) => user.email === email);
  return serverAccount || accountState.users.find((user) => user.email === email) || {
    email,
    realName: "",
    password: "",
    role: currentUserRole === "creator" ? "owner" : "user",
    claims: [],
  };
}

function viewedAccountEmail() {
  return normalizeEmail(adminViewingEmail || currentUserEmail);
}

function authenticatedAccountEmail() {
  return normalizeEmail(authUserEmail || currentUserEmail);
}

function isViewingCreatorAccount() {
  return activeCalendarMode() === "creator-account"
    && normalizeEmail(currentUserEmail) === OWNER_EMAIL
    && !adminViewingEmail
    && !activeDoctorProfile;
}

function isOwnerAccount() {
  return isViewingCreatorAccount();
}

function canUseRosterInsights() {
  if (activeCalendarMode() === "doctor-profile") return isCreatorAuthenticated();
  if (currentUserRole === "creator" && !adminViewingEmail) return true;
  return currentInsightsEnabled === true;
}

function canRemoveImports() {
  return isViewingCreatorAccount();
}

function isCreatorAuthenticated() {
  return normalizeEmail(authUserEmail || currentUserEmail) === OWNER_EMAIL && Boolean(authUserPassword || currentUserPassword);
}

function syncAccountsButton() {
  const ownerView = isViewingCreatorAccount();
  const issueCount = ownerView ? adminIssueCount() : 0;
  accountsButton.innerHTML = ownerView
    ? `Admin${issueCount ? `<span class="notification-badge">${issueCount}</span>` : ""}`
    : "Account";
}

function renderAccountsModal() {
  const me = currentAccount();
  const ownerView = isViewingCreatorAccount();
  accountsModalTitle.textContent = ownerView ? "Admin" : "Account";
  accountsModalSubtitle.textContent = ownerView
    ? "Review user issues, manage accounts, and update the owner account."
    : "Manage your account details.";

  const serverOtherUsers = serverUsers
    .map(normalizeServerUser)
    .filter((user) => user.email !== me.email);
  const localOtherUsers = accountState.users.filter((user) => user.email !== me.email);
  const otherUsers = serverOtherUsers.length ? serverOtherUsers : localOtherUsers;
  const availableUserSeniorities = [...new Set(otherUsers.flatMap((user) => normalizeServerUser(user).seniorities || []))].sort();
  if (adminUserSeniorityFilter && !availableUserSeniorities.includes(adminUserSeniorityFilter)) adminUserSeniorityFilter = "";
  const filteredOtherUsers = adminUserSeniorityFilter
    ? otherUsers.filter((user) => normalizeServerUser(user).seniorities.includes(adminUserSeniorityFilter))
    : otherUsers;
  const linkedNames = renderLinkedRosterNames(currentRosterClaims, currentSuggestedClaims);
  if (ownerView && !["errors", "system", "users", "files", "owner"].includes(currentAdminTab)) currentAdminTab = "system";
  const issueCount = adminIssueCount();
  const adminTabs = ownerView ? `
    <div class="admin-tabs" role="tablist" aria-label="Admin sections">
      <button type="button" class="entrance-tab ${currentAdminTab === "system" ? "is-active" : ""}" data-admin-tab="system">System</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "errors" ? "is-active" : ""}" data-admin-tab="errors">Errors${issueCount ? `<span class="notification-badge">${issueCount}</span>` : ""}</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "users" ? "is-active" : ""}" data-admin-tab="users">Users</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "files" ? "is-active" : ""}" data-admin-tab="files">Files</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "owner" ? "is-active" : ""}" data-admin-tab="owner">Account</button>
    </div>
  ` : "";
  const ownerCard = `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>${ownerView ? "Account" : "Your account"}</strong>
          <span>${escapeHtml(me.realName || "Name not set")} · ${escapeHtml(me.email)}</span>
        </div>
      </div>
      <form class="review-body" data-account-form>
        <label class="field">
          <span>Preferred display name</span>
          <input type="text" value="${escapeHtml(me.realName || "")}" data-account-real-name placeholder="Name shown on calendar banner">
        </label>
        <label class="field">
          <span>Email address</span>
          <input type="email" value="${escapeHtml(me.email)}" data-account-email ${ownerView ? "readonly" : "readonly"}>
        </label>
        <label class="field">
          <span>Password</span>
          <input type="password" value="${escapeHtml(me.password || "")}" data-account-password placeholder="Update password">
        </label>
        <div class="modal-actions">
          <button type="submit" class="button button-primary">Save password</button>
          ${me.email !== OWNER_EMAIL ? `<button type="button" class="button button-danger" data-delete-account="${escapeHtml(me.email)}">Delete account</button>` : ""}
        </div>
      </form>
      ${renderAccountHospitalLocationsCard()}
      ${linkedNames}
      ${ownerView ? "" : renderFilesMarkup({
        canRemove: false,
        canAdd: true,
        heading: "Files used to generate your calendar...",
        description: "These roster files currently feed your calendar.",
      })}
    </article>
  `;
  const usersCard = ownerView ? `
      <article class="review-card">
        <div class="review-top">
          <div>
            <strong>Create user account</strong>
            <span>Create an account and enter it immediately for setup or testing.</span>
          </div>
        </div>
        <form class="review-body" data-create-account-form>
          <label class="field">
            <span>Real name</span>
            <input type="text" data-create-real-name placeholder="Name shown to the user" autocomplete="name">
          </label>
          <label class="field">
            <span>Email address</span>
            <input type="email" data-create-email placeholder="doctor@example.com" autocomplete="email">
          </label>
          <label class="field">
            <span>Temporary password</span>
            <input type="password" data-create-password placeholder="Temporary password" autocomplete="new-password">
          </label>
          <div class="modal-actions">
            <button type="submit" class="button button-primary">Create and enter account</button>
          </div>
        </form>
      </article>
      <article class="review-card">
        <div class="review-top">
          <div>
            <strong>Other users</strong>
            <span>${filteredOtherUsers.length ? `${filteredOtherUsers.length} account${filteredOtherUsers.length === 1 ? "" : "s"}` : otherUsers.length ? "No matching users." : "No other users have logged in yet."}</span>
          </div>
        </div>
        <label class="field">
          <span>Filter by seniority</span>
          <select data-admin-user-seniority-filter>
            <option value="">All seniorities</option>
            ${availableUserSeniorities.map((seniority) => `<option value="${escapeHtml(seniority)}" ${seniority === adminUserSeniorityFilter ? "selected" : ""}>${escapeHtml(seniority)}</option>`).join("")}
          </select>
        </label>
        <div class="issues-list">
          ${filteredOtherUsers.length ? filteredOtherUsers.map((user) => `
            <article class="issue-card account-user-card">
              <div>
                <strong>${escapeHtml(user.realName || "Name not set")}</strong>
                <p>${escapeHtml(user.email)} · ${user.role === "owner" ? "Creator" : "Standard user"} · ${formatUserSites(user)} · storage limit: latest 6 months active</p>
                ${renderLinkedRosterNames(user.claims || [], [], { compact: true, email: user.email })}
              </div>
              ${user.role === "owner" ? "" : `
                <label class="toggle review-toggle">
                  <input type="checkbox" ${user.insightsEnabled ? "checked" : ""} data-toggle-user-insights="${escapeHtml(user.email)}">
                  Allow “Who/When am I working with?” tools
                </label>
              `}
              <div class="account-actions">
                <button type="button" class="button button-secondary" data-enter-account="${escapeHtml(user.email)}">Enter account</button>
                ${user.role === "owner" ? "" : `<button type="button" class="button button-secondary button-small" data-edit-roster-claims="${escapeHtml(user.email)}">Edit</button>`}
                ${user.email !== OWNER_EMAIL ? `<button type="button" class="button button-danger" data-delete-account="${escapeHtml(user.email)}">Delete</button>` : ""}
              </div>
              ${user.role === "owner" ? "" : `
                <div class="account-claim-editor hidden" data-claim-editor="${escapeHtml(user.email)}">
                  <select data-admin-claim-select="${escapeHtml(user.email)}">
                    <option value="">Add roster name...</option>
                    ${availableRosterDoctors.map((doctor, index) => `<option value="${index}">${escapeHtml(`${doctor.displayName} (${doctor.sourceType.toUpperCase()})${doctor.claimedBy && doctor.claimedBy !== user.email ? ` - claimed by ${doctor.claimedBy}` : ""}`)}</option>`).join("")}
                  </select>
                  <button type="button" class="button button-secondary" data-admin-add-claim="${escapeHtml(user.email)}">Add roster name</button>
                </div>
              `}
            </article>
          `).join("") : `<article class="issue-card"><p>${otherUsers.length ? "No users match this seniority." : "No additional users yet."}</p></article>`}
        </div>
      </article>
    ` : "";
  const errorsCard = ownerView ? renderAdminErrorsCard(serverOtherUsers) : "";
  const systemCard = ownerView ? renderSystemAdminCard() : "";
  const filesCard = ownerView ? renderFilesMarkup({
    canRemove: canRemoveImports(),
    canAdd: true,
    heading: "Files",
    description: "Files currently used to generate the creator calendar.",
  }) : "";
  const adminBody = ownerView
    ? (currentAdminTab === "errors"
        ? errorsCard
        : currentAdminTab === "system"
          ? systemCard
          : currentAdminTab === "users"
            ? usersCard
            : currentAdminTab === "files"
              ? filesCard
              : ownerCard)
    : ownerCard;
  accountsBody.innerHTML = `${adminTabs}${adminBody}`;
}

function renderSystemAdminCard() {
  return `
    <div class="issues-list">
      ${renderCalendarStoreCard()}
      ${renderParserRulesCard()}
    </div>
  `;
}

function renderAccountHospitalLocationsCard() {
  const sourceTypes = recognizedHospitalTypesForActiveAccount();
  const rows = sourceTypes.map((sourceType) => {
    const config = hospitalLocationConfig(sourceType);
    return `
      <label class="field">
        <span>${escapeHtml(config.label)}</span>
        <input type="text" value="${escapeHtml(settings[config.settingKey] || config.defaultValue)}" data-account-location-key="${config.settingKey}">
      </label>
    `;
  }).join("");
  return `
    <section class="review-body account-hospital-locations">
      <div class="section-head">
        <h4>Recognised hospitals &amp; default locations</h4>
        <p>${sourceTypes.length ? "These hospitals are linked to this account's roster data." : "No linked hospitals recognised yet."}</p>
      </div>
      ${rows || `<article class="issue-card"><p>Link a roster identity to expose hospital defaults here.</p></article>`}
    </section>
  `;
}

const ACCOUNT_HOSPITAL_LOCATION_ORDER = ["mmc", "ddh", "mch", "casey"];

function recognizedHospitalTypesForActiveAccount() {
  if (isViewingCreatorAccount()) {
    return ACCOUNT_HOSPITAL_LOCATION_ORDER.filter((sourceType) => detectedSources[sourceType]?.length);
  }
  const claimTypes = currentRosterClaims.map((claim) => String(claim.sourceType || "").toLowerCase());
  const doctorTypes = normalizedDoctorSourceTypes(selectedDoctor());
  const linkedTypes = new Set([...claimTypes, ...doctorTypes]);
  return ACCOUNT_HOSPITAL_LOCATION_ORDER.filter((sourceType) => linkedTypes.has(sourceType));
}

function hospitalLocationConfig(sourceType) {
  return {
    mmc: { label: "MMC", settingKey: "defaultLocationMmc", defaultValue: DEFAULT_MMC_LOCATION },
    ddh: { label: "DDH", settingKey: "defaultLocationDdh", defaultValue: DEFAULT_DDH_LOCATION },
    casey: { label: "Casey", settingKey: "defaultLocationCasey", defaultValue: DEFAULT_CASEY_LOCATION },
    mch: { label: "MCH", settingKey: "defaultLocationMch", defaultValue: DEFAULT_MCH_LOCATION },
  }[sourceType];
}

function updateDefaultLocationSetting(settingKey, value) {
  if (!["defaultLocationMmc", "defaultLocationDdh", "defaultLocationCasey", "defaultLocationMch"].includes(settingKey)) return;
  settings[settingKey] = String(value || "").trim() || defaultSettings()[settingKey];
  renderSettings();
  saveCurrentSessionState();
  if (latestPreview) rebuildClientPreview();
  setStatus("Default export locations updated.");
}

function renderCalendarStoreCard() {
  const status = calendarStoreStatus;
  const unavailable = status?.unavailable === true;
  const selectedPersistence = summarizeSelectedRosterPersistence(selectedFiles, status);
  const missingSelectedFiles = selectedPersistence.missingEntries.slice(0, 3);
  const retainedSourceCount = (status?.files || []).filter((file) => file.rawSourceAvailable === true).length;
  const retainedSourceTotal = (status?.files || []).length;
  const retainedSourceDetail = status && !unavailable
    ? `${retainedSourceCount}/${retainedSourceTotal} source file${retainedSourceTotal === 1 ? "" : "s"} retained.`
    : "";
  const checkedAt = status?.checkedAt ? `Last checked ${formatTimestamp(status.checkedAt)}.` : "Not checked yet.";
  const syncSummary = rosterSyncSummary();
  const serverSyncedCount = Number(status?.populated || 0);
  const serverExpectedCount = Number(status?.total || 0);
  const serverStatusComplete = Boolean(status && !unavailable && serverExpectedCount > 0 && serverSyncedCount === serverExpectedCount);
  const statusErrorDetail = calendarStoreStatusError
    ? `${status ? `${serverSyncedCount}/${serverExpectedCount || selectedPersistence.expectedCount || "?"} roster files confirmed from last good status. ` : ""}Status check failed: ${calendarStoreStatusError}. ${checkedAt}`
    : "";
  const detail = syncSummary
    ? `${syncSummary}. ${checkedAt}`
    : calendarStoreStatusError
    ? statusErrorDetail
    : unavailable
      ? "Roster database is unavailable to this deployment."
      : serverStatusComplete
        ? `${serverSyncedCount} roster file${serverSyncedCount === 1 ? "" : "s"} synced. ${retainedSourceDetail} ${checkedAt}`
        : status && selectedPersistence.expectedCount > 0 && selectedPersistence.complete
          ? `${selectedPersistence.persistedCount} roster file${selectedPersistence.persistedCount === 1 ? "" : "s"} synced. ${retainedSourceDetail} ${checkedAt}`
        : status
          ? `Sync issue detected: ${serverSyncedCount}/${serverExpectedCount} roster files confirmed. ${retainedSourceDetail} ${checkedAt}`
          : "Roster database status not checked yet.";
  return `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>Roster database</strong>
          <span>${escapeHtml(detail)}</span>
        </div>
      </div>
      <div class="review-body">
        ${missingSelectedFiles.length ? `
          <div class="issues-list">
            ${missingSelectedFiles.map((entry) => `
              <article class="issue-card">
                <div>
                  <strong>${escapeHtml(entry.name)}</strong>
                  <p>Roster file not yet synced.</p>
                </div>
              </article>
            `).join("")}
          </div>
        ` : ""}
        <div class="modal-actions">
          <button type="button" class="button button-secondary" data-refresh-calendar-store>Check status</button>
          <button type="button" class="button button-secondary" data-replace-active-rosters>Rebuild from roster files</button>
          <button type="button" class="button button-secondary" data-view-console>${adminConsoleOpen ? "Hide console" : "View console"}</button>
        </div>
        ${adminConsoleOpen ? renderAdminConsoleMarkup() : ""}
      </div>
    </article>
  `;
}

function renderAdminConsoleMarkup() {
  const body = adminConsoleLoading
    ? `<article class="issue-card"><p>Loading console...</p></article>`
    : adminConsoleMessages.length
      ? adminConsoleMessages.map((entry) => `
          <article class="issue-card console-entry ${entry.isError ? "is-error" : ""}">
            <div>
              <strong>${escapeHtml(formatConsoleTimestamp(entry.createdAt))}</strong>
              <p>${escapeHtml(entry.message)}</p>
            </div>
          </article>
        `).join("")
      : `<article class="issue-card"><p>No console messages stored yet.</p></article>`;
  return `<div class="issues-list console-history">${body}</div>`;
}

function formatConsoleTimestamp(value) {
  const date = new Date(value || "");
  return Number.isNaN(date.getTime()) ? "Unknown time" : date.toLocaleString();
}

function renderParserRulesCard() {
  const groups = [
    { source: "MMC", rules: sanitizeParserExtensionRuleList(parserExtensions?.mmc, "MMC") },
    { source: "DDH", rules: sanitizeParserExtensionRuleList(parserExtensions?.ddh, "DDH") },
    { source: "Casey", rules: sanitizeParserExtensionRuleList(parserExtensions?.casey, "Casey") },
    { source: "MCH", rules: sanitizeParserExtensionRuleList(parserExtensions?.mch, "MCH") },
  ];
  const allUnknownIssues = collectUnknownShiftIssues();
  const unknownIssues = allUnknownIssues;
  const unknownSources = new Set(allUnknownIssues.map((item) => item.source));
  return `
    <div class="issues-list">
      ${parserRuleSuggestions.length ? `
        <article class="review-card">
          <div class="review-top">
            <div>
              <strong>User suggestions</strong>
              <span>${parserRuleSuggestions.length} pending suggestion${parserRuleSuggestions.length === 1 ? "" : "s"}</span>
            </div>
          </div>
          <div class="issues-list">
            ${parserRuleSuggestions.map((suggestion) => `
              <article class="issue-card">
                <div>
                  <strong>${escapeHtml(suggestion.rule.source)} · ${escapeHtml(suggestion.rule.seniority)} · ${escapeHtml(suggestion.rule.code)}</strong>
                  <p>${escapeHtml(suggestion.realName || suggestion.email)} suggested ${escapeHtml(parserRulePreviewTitle(suggestion.rule))}</p>
                  <p>${escapeHtml(parserRulePreviewMeta(suggestion.rule))}${suggestion.updatedAt ? ` · ${escapeHtml(formatTimestamp(suggestion.updatedAt))}` : ""}</p>
                </div>
                <div class="account-actions">
                  <button type="button" class="button button-secondary" data-parser-suggestion-action="approveGlobal" data-suggestion-id="${escapeHtml(suggestion.id)}">Approve globally</button>
                  <button type="button" class="button button-secondary" data-parser-suggestion-action="approveUser" data-suggestion-id="${escapeHtml(suggestion.id)}">Approve for user</button>
                  <button type="button" class="button button-secondary" data-parser-suggestion-action="overwrite" data-suggestion-id="${escapeHtml(suggestion.id)}">Overwrite</button>
                  <button type="button" class="button button-secondary" data-parser-suggestion-action="reject" data-suggestion-id="${escapeHtml(suggestion.id)}">Reject</button>
                </div>
              </article>
            `).join("")}
          </div>
        </article>
      ` : ""}
      <article class="review-card">
        <div class="review-top">
          <div>
            <strong>Missing / unresolved shift codes</strong>
            <span>${unknownIssues.length ? `${unknownIssues.length} code${unknownIssues.length === 1 ? "" : "s"} needing review` : "No unresolved shift codes."}</span>
          </div>
          <button type="button" class="button button-secondary button-small" data-add-manual-shift-code>Add</button>
        </div>
        <div class="issues-list">
          ${unknownIssues.length ? unknownIssues.map((item) => `
            <article class="issue-card">
              <div>
                <strong>${escapeHtml(item.source)} · ${escapeHtml(item.code)}</strong>
                <p>${escapeHtml(item.seniorityLabel)}</p>
                <p>${escapeHtml(item.message || "Shift code not recognised.")}</p>
                <p>${escapeHtml(item.sample)}${item.count > 1 ? ` · seen ${item.count} times` : ""}</p>
              </div>
              <div class="account-actions">
                <button type="button" class="button button-secondary" data-add-shift-code="${escapeHtml(item.email)}" data-error-id="${escapeHtml(item.id)}" data-shift-code-seniorities="${escapeHtml(item.seniorities.join("|"))}">Edit shift code</button>
                <button type="button" class="button button-secondary" data-ignore-shift-code="${escapeHtml(item.email)}" data-error-id="${escapeHtml(item.id)}" data-shift-code-seniorities="${escapeHtml(item.seniorities.join("|"))}">Ignore</button>
              </div>
            </article>
          `).join("") : `<article class="issue-card"><p>No missing or unresolved shift codes need review.</p></article>`}
        </div>
      </article>
      <article class="issue-card">
        <div>
          <strong>Shift code rules</strong>
          <p>Review and edit parser rules by hospital and seniority.</p>
        </div>
        ${groups.map((group) => `
          <details class="issue-card" data-parser-rule-source-section="${escapeHtml(group.source)}">
            <summary><strong>${group.source}${unknownSources.has(group.source) ? " *" : ""}</strong> · ${visibleParserRules(group.rules).length} rule${visibleParserRules(group.rules).length === 1 ? "" : "s"}</summary>
            <div class="issues-list">
              ${allUnknownIssues.filter((item) => item.source === group.source).map((item) => `
                <article class="issue-card issue-unknown">
                  <div>
                    <strong>${escapeHtml(item.code)} · Unrecognised</strong>
                    <p>${escapeHtml(item.seniorityLabel)} · ${escapeHtml(item.message || "Shift code not recognised.")}</p>
                  </div>
                  <div class="account-actions">
                    <button type="button" class="button button-secondary" data-add-shift-code="${escapeHtml(item.email)}" data-error-id="${escapeHtml(item.id)}" data-shift-code-seniorities="${escapeHtml(item.seniorities.join("|"))}">Edit shift code</button>
                  </div>
                </article>
              `).join("")}
              ${parserRuleSeniorityDisplayOrder().map((seniority) => {
                const rules = visibleParserRules(group.rules).filter((rule) => rule.seniority === seniority);
                if (!rules.length) return "";
                return `
                  <details class="issue-card" data-parser-rule-seniority-section="${escapeHtml(`${group.source}|${seniority}`)}">
                    <summary><strong>${escapeHtml(seniority)}</strong> · ${rules.length} rule${rules.length === 1 ? "" : "s"}</summary>
                    <div class="issues-list">
                      ${rules.map((rule) => `
                        <article class="issue-card" data-parser-rule-card="${escapeHtml(parserRuleFocusKey(rule))}">
                          <div>
                            <strong>${escapeHtml(rule.code)}${rule.includeAsShift === false ? " · Hidden" : ""}</strong>
                            <p>${escapeHtml(parserRulePreviewTitle(rule))}</p>
                            <p>${escapeHtml(parserRulePreviewMeta(rule))}</p>
                          </div>
                          <div class="account-actions">
                            <button type="button" class="button button-secondary" data-edit-parser-rule="${escapeHtml(rule.code)}" data-edit-parser-source="${escapeHtml(rule.source)}" data-edit-parser-seniority="${escapeHtml(rule.seniority)}">Edit</button>
                            <button type="button" class="button button-secondary" data-delete-parser-rule="${escapeHtml(rule.code)}" data-delete-parser-source="${escapeHtml(rule.source)}" data-delete-parser-seniority="${escapeHtml(rule.seniority)}">Delete</button>
                          </div>
                        </article>
                      `).join("")}
                    </div>
                  </details>
                `;
              }).join("")}
              ${ignoredParserRules(group.rules).length ? `
                <details class="issue-card" data-parser-rule-ignored-section="${escapeHtml(group.source)}">
                  <summary><strong>Ignored shift codes</strong> · ${ignoredParserRules(group.rules).length} code${ignoredParserRules(group.rules).length === 1 ? "" : "s"}</summary>
                  <div class="issues-list">
                    ${ignoredParserRules(group.rules).map((rule) => `
                      <article class="issue-card" data-parser-rule-card="${escapeHtml(parserRuleFocusKey(rule))}">
                        <div>
                          <strong>${escapeHtml(rule.code)} · Ignored</strong>
                          <p>${escapeHtml(rule.seniority)}</p>
                          <p>${escapeHtml(parserRulePreviewMeta(rule))}</p>
                        </div>
                        <div class="account-actions">
                          <button type="button" class="button button-secondary" data-edit-parser-rule="${escapeHtml(rule.code)}" data-edit-parser-source="${escapeHtml(rule.source)}" data-edit-parser-seniority="${escapeHtml(rule.seniority)}">Edit</button>
                          <button type="button" class="button button-secondary" data-delete-parser-rule="${escapeHtml(rule.code)}" data-delete-parser-source="${escapeHtml(rule.source)}" data-delete-parser-seniority="${escapeHtml(rule.seniority)}">Delete</button>
                        </div>
                      </article>
                    `).join("")}
                  </div>
                </details>
              ` : ""}
            </div>
          </details>
        `).join("")}
      </article>
    </div>
  `;
}

function visibleParserRules(rules = []) {
  return (rules || []).filter((rule) => rule.ignore !== true && rule.kind !== "ignore");
}

function ignoredParserRules(rules = []) {
  return (rules || []).filter((rule) => rule.ignore === true || rule.kind === "ignore");
}

function collectUnknownShiftIssues() {
  const byKey = new Map();
  for (const user of serverUsers.map(normalizeServerUser)) {
    for (const issue of user.adminIssues || []) {
      const source = sanitizeIssueSource(issue.source);
      const seniority = sanitizeRuleSeniority(issue.seniority);
      const code = parserRuleCodeForIssue(issue);
      if (!source || !code) continue;
      if (isKnownResolvedShiftCodeValue(source, issue.rawValue, issue.suggestedTitle)) continue;
      if (isShiftCodeResolvedByActiveRules({ source, seniority, code, rawValue: code })) continue;
      const key = `${source}|${code}`;
      const existing = byKey.get(key);
      if (existing) {
        existing.count += issue.count || 1;
        existing.seniorities = addUniqueSeniority(existing.seniorities, seniority);
        existing.seniorityLabel = formatShiftCodeSeniorities(existing.seniorities);
        if ((issue.lastSeenAt || "") > (existing.lastSeenAt || "")) existing.lastSeenAt = issue.lastSeenAt || "";
        continue;
      }
      const seniorities = addUniqueSeniority([], seniority);
      byKey.set(key, {
        source,
        seniority,
        seniorities,
        seniorityLabel: formatShiftCodeSeniorities(seniorities),
        code,
        id: issue.id || issue.fingerprint || "",
        email: user.email,
        message: issue.message || "",
        sample: `${user.realName || user.email} · ${formatDate(issue.startDay || issue.date || "")} · ${issue.rawValue || code}`,
        count: issue.count || 1,
        lastSeenAt: issue.lastSeenAt || "",
      });
    }
  }
  return [...byKey.values()].sort((left, right) => {
    if (left.source !== right.source) return left.source.localeCompare(right.source);
    return left.code.localeCompare(right.code);
  });
}

function addUniqueSeniority(items = [], seniority = "") {
  const normalized = sanitizeRuleSeniority(seniority);
  return [...new Set([...items, normalized])].sort((left, right) => senioritySortRank(left) - senioritySortRank(right));
}

function senioritySortRank(value) {
  const order = parserRuleSeniorityDisplayOrder();
  const index = order.indexOf(sanitizeRuleSeniority(value));
  return index >= 0 ? index : order.length;
}

function formatShiftCodeSeniorities(items = []) {
  const seniorities = [...new Set(items.map(sanitizeRuleSeniority).filter(Boolean))].sort((left, right) => senioritySortRank(left) - senioritySortRank(right));
  if (!seniorities.length) return "Unknown seniority";
  if (seniorities.length === 1) return seniorities[0];
  return `${seniorities.length} seniorities: ${seniorities.join(", ")}`;
}

function splitShiftCodeSeniorities(value = "") {
  return String(value || "")
    .split("|")
    .map(sanitizeRuleSeniority)
    .filter(Boolean);
}

function parserRuleSeniorityDisplayOrder() {
  const values = parserRuleSeniorities();
  return ["Unknown", ...values.filter((item) => item !== "Unknown")];
}

function renderLinkedRosterNames(claims, suggestedClaims = [], options = {}) {
  const items = sanitizeRosterClaims(claims);
  const suggestions = sanitizeRosterClaims(suggestedClaims);
  if (!items.length && !suggestions.length) {
    return `<p class="status">No roster names are linked to this account yet.</p>`;
  }
  return `
    <div class="issues-list account-claim-list">
      ${items.map((claim) => `
        <article class="issue-card account-claim-card">
          <div>
            <strong>${escapeHtml(claim.sourceType.toUpperCase())}</strong>
            <p>${escapeHtml(claim.displayName)}</p>
          </div>
          ${options.creatorTools || !options.compact ? `<button type="button" class="button button-secondary button-small account-claim-remove" data-remove-roster-claim="${escapeHtml(claim.sourceType)}:${escapeHtml(claim.key)}" data-claim-email="${escapeHtml(options.email || currentUserEmail)}">Remove / this is wrong</button>` : ""}
        </article>
      `).join("")}
      ${suggestions.map((claim, index) => `
        <article class="issue-card">
          <div>
            <strong>Suggested ${escapeHtml(claim.sourceType.toUpperCase())}</strong>
            <p>${escapeHtml(claim.displayName)}</p>
          </div>
          <div class="account-actions">
            <button type="button" class="button button-primary" data-confirm-suggested-claim="${index}">Confirm</button>
            <button type="button" class="button button-secondary" data-reject-suggested-claim="${index}">This is wrong</button>
          </div>
        </article>
      `).join("")}
      ${!options.compact ? `<button type="button" class="button button-secondary" data-report-roster-identity>Report roster name problem</button>` : ""}
    </div>
  `;
}

function adminIssueCount() {
  return serverUsers
    .map(normalizeServerUser)
    .filter((user) => user.email !== OWNER_EMAIL)
    .reduce((total, user) => total + ((user.adminIssues || []).length || 0), 0);
}

function renderAdminErrorsCard(users) {
  const issueUsers = users.filter((user) => (user.adminIssues || []).length);
  return `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>Errors</strong>
          <span>${issueUsers.length ? `${adminIssueCount()} issue${adminIssueCount() === 1 ? "" : "s"} across ${issueUsers.length} user${issueUsers.length === 1 ? "" : "s"}` : "No user issues queued."}</span>
        </div>
      </div>
      <div class="issues-list">
        ${issueUsers.length ? issueUsers.map((user) => `
          <article class="issue-card">
            <div>
              <strong>${escapeHtml(user.realName || "Name not set")}</strong>
              <p>${escapeHtml(user.email)}</p>
            </div>
            <div class="issues-list">
              ${(user.adminIssues || []).map((issue) => `
                <article class="issue-card">
                  <div>
                    <strong>${escapeHtml(`${user.realName || user.email} · ${issue.source || "Roster"} · ${formatDate(issue.date || issue.startDay || "")}`)}</strong>
                    <p>${escapeHtml(issue.message)}</p>
                    <p>Raw code/value: ${escapeHtml(issue.rawValue || "Unknown")}</p>
                    ${issue.timeLabel ? `<p>Explicit roster time: ${escapeHtml(issue.timeLabel)}</p>` : ""}
                    ${issue.suggestedTitle ? `<p>Suggested normalised result: ${escapeHtml(issue.suggestedTitle)}</p>` : `<p>Suggested normalised result: No normalised result</p>`}
                    <p>${escapeHtml(formatTimestamp(issue.lastSeenAt))}${issue.count > 1 ? ` · seen ${issue.count} times` : ""}</p>
                  </div>
                  <div class="account-actions">
                    <button type="button" class="button button-secondary" data-enter-account="${escapeHtml(user.email)}">Enter account</button>
                    <button type="button" class="button button-secondary" data-add-shift-code="${escapeHtml(user.email)}" data-error-id="${escapeHtml(issue.id)}">Add shift code</button>
                    <button type="button" class="button button-secondary" data-clear-admin-errors="${escapeHtml(user.email)}" data-error-id="${escapeHtml(issue.id)}">Clear</button>
                    <button type="button" class="button button-secondary" data-ignore-admin-error="${escapeHtml(user.email)}" data-error-id="${escapeHtml(issue.id)}">Ignore forever</button>
                  </div>
                </article>
              `).join("")}
            </div>
            <div class="account-actions">
              <button type="button" class="button button-secondary" data-clear-admin-errors="${escapeHtml(user.email)}">Clear all for ${escapeHtml(user.realName || user.email)}</button>
            </div>
          </article>
        `).join("") : `<article class="issue-card"><p>No user errors to review.</p></article>`}
      </div>
    </article>
  `;
}

function subscriptionUrl(protocol = "https", view = "full") {
  if (!currentSubscription?.token) return "";
  const url = new URL("/api/feed", window.location.origin);
  url.searchParams.set("token", currentSubscription.token);
  url.searchParams.set("view", view === "range" ? "range" : "full");
  if (protocol === "webcal") {
    return url.toString().replace(/^https?:/i, "webcal:");
  }
  return url.toString();
}


function canCopySubscriptionUrl() {
  return Boolean(currentSubscription?.enabled && activeCalendarMode() !== "doctor-profile");
}

async function copySubscriptionLink(kind = "https") {
  const url = subscriptionUrl(kind === "webcal" ? "webcal" : "https", "full");
  if (!url) {
    setStatus("No subscription link is available for this account yet.", true);
    return;
  }
  try {
    await navigator.clipboard.writeText(url);
    setStatus(`${kind === "webcal" ? "webcal" : "HTTPS"} subscription link copied.`);
  } catch {
    setStatus("Could not copy the subscription link.", true);
  }
}

function formatUserSites(user) {
  const sites = Array.isArray(user.sites) ? user.sites.filter(Boolean) : [];
  return sites.length ? `sites: ${sites.join(", ")}` : "no linked sites";
}

async function clearAdminErrors(email, errorId = "") {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to clear user errors.", true);
    return;
  }
  const issueBeforeClear = errorId ? findAdminIssue(email, errorId) : null;
  const issuesBeforeClear = (serverUsers.map(normalizeServerUser).find((item) => item.email === normalizeEmail(email))?.adminIssues || []).map((issue) => issue?.fingerprint).filter(Boolean);
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "clearUserError",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail: email,
        errorId,
      }),
    });
    await readJsonResponse(response, "Could not clear user errors.");
    await loadServerUsers();
    renderAccountsModal();
    if (normalizeEmail(email) === currentUserEmail) {
      if (errorId) {
        const fingerprint = issueBeforeClear?.fingerprint;
        if (fingerprint) dismissedIssueFingerprints.add(fingerprint);
      } else {
        for (const fingerprint of issuesBeforeClear) {
          dismissedIssueFingerprints.add(fingerprint);
        }
      }
      rebuildClientPreview();
    }
    syncAccountsButton();
    setStatus(errorId ? "User warning dismissed." : "All user warnings dismissed for that account.");
  } catch (error) {
    setStatus(error.message || "Could not clear user errors.", true);
  }
}

async function ignoreAdminErrorForever(email, errorId = "") {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to ignore parser warnings.", true);
    return;
  }
  const issue = findAdminIssue(email, errorId);
  if (!issue?.fingerprint) {
    setStatus("Could not find that parser warning.", true);
    return;
  }
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "ignoreUserErrorForever",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        fingerprint: issue.fingerprint,
      }),
    });
    await readJsonResponse(response, "Could not ignore that parser warning.");
    ignoredIssueFingerprints.add(issue.fingerprint);
    await loadServerUsers();
    renderAccountsModal();
    syncAccountsButton();
    rebuildClientPreview();
    setStatus("Parser warning ignored forever.");
  } catch (error) {
    setStatus(error.message || "Could not ignore that parser warning.", true);
  }
}

async function handleParserSuggestionAction(action, suggestionId) {
  if (action === "overwrite") {
    openParserRuleModalFromSuggestion(suggestionId);
    return;
  }
  const suggestion = parserRuleSuggestions.find((item) => item.id === suggestionId);
  if (!suggestion) {
    setStatus("Could not find that suggestion.", true);
    return;
  }
  if (action === "approveGlobal" || action === "approveUser" || action === "reject") {
    await decideParserRuleSuggestion(suggestionId, action, suggestion.rule);
  }
}

async function decideParserRuleSuggestion(suggestionId, decision, rule) {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to review suggestions.", true);
    return;
  }
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "decideParserRuleSuggestion",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        suggestionId,
        decision,
        rule,
      }),
    });
    const data = await readJsonResponse(response, "Could not update that suggestion.");
    parserExtensions = sanitizeParserExtensions(data.parserExtensions || parserExtensions);
    globalParserExtensions = sanitizeParserExtensions(data.parserExtensions || globalParserExtensions);
    parserRuleSuggestions = sanitizeParserRuleSuggestions(data.suggestions);
    setParserExtensions(parserExtensions);
    await loadServerUsers();
    renderAccountsModal();
    const labels = {
      approveGlobal: "Suggestion approved globally.",
      approveUser: "Suggestion approved for that user.",
      reject: "Suggestion rejected.",
    };
    setStatus(labels[decision] || "Suggestion updated.");
  } catch (error) {
    setStatus(error.message || "Could not update that suggestion.", true);
  }
}

function findAdminIssue(email, errorId = "") {
  const user = serverUsers.map(normalizeServerUser).find((item) => item.email === normalizeEmail(email));
  if (!user) return null;
  return (user.adminIssues || []).find((issue) => issue.id === errorId || issue.fingerprint === errorId) || null;
}

function findParserExtensionRule(source, code) {
  return findParserExtensionRuleForSeniority(source, "", code);
}

function findParserExtensionRuleForSeniority(source, seniority, code) {
  const sourceKey = sanitizeIssueSource(source).toLowerCase();
  const normalizedSeniority = sanitizeRuleSeniority(seniority);
  const normalizedCode = String(code || "").trim().toUpperCase();
  if (!sourceKey || !normalizedCode) return null;
  return sanitizeParserExtensionRuleList(parserExtensions?.[sourceKey], sourceKey.toUpperCase())
    .find((rule) => rule.code === normalizedCode && (!seniority || rule.seniority === normalizedSeniority)) || null;
}

function parserRuleFocusKey(rule) {
  return `${rule.source}|${rule.seniority}|${rule.code}`;
}

function parserRuleExistsForIssue(issue) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = parserRuleCodeForIssue(issue);
  return isShiftCodeResolvedByActiveRules({ source, seniority, code });
}

function isShiftCodeResolvedByActiveRules(issue) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = parserRuleCodeForIssue(issue);
  if (!source || !code) return false;
  if (seniority !== "Unknown") return Boolean(findParserExtensionRuleForSeniority(source, seniority, code));
  const sourceKey = source.toLowerCase();
  return sanitizeParserExtensionRuleList(parserExtensions?.[sourceKey], source)
    .some((rule) => rule.code === code);
}

function isKnownResolvedShiftCodeValue(sourceValue, rawValue, normalizedTitle = "") {
  const source = sanitizeIssueSource(sourceValue);
  const code = parserRuleCodeFromRawValue(source, rawValue);
  if (!source || !code) return false;
  if (["AM", "PM", "NIGHT"].includes(code)) return true;
  if (code === "PHNW") return true;
  if (source === "MCH" && ["CS", "OCS", "0CS", "CSOS"].includes(code)) return true;
  if (source === "DDH") {
    if (["CS", "CS ONSITE", "SSU"].includes(code)) return true;
    if (/^(ORANGE|SILVER|FAST|AVAO|ROVER)\s+(AM|PM)$/.test(code)) return true;
  }
  const titleCode = incompleteShiftCodeFromTitle(source, normalizedTitle);
  return Boolean(titleCode && isShiftCodeResolvedByActiveRules({ source, seniority: "Unknown", code: titleCode }));
}

function previewIssueWithReviewContext(issue) {
  const reviewItem = issue?.id ? reviewIndex.get(issue.id) : null;
  if (!reviewItem) return issue;
  return {
    ...issue,
    seniority: issue?.seniority || reviewItem.seniority,
    rawValue: issue?.rawValue || reviewItem.rawValue,
    suggestedTitle: issue?.suggestedTitle || reviewItem.normalizedTitle || reviewItem.suggestedTitle,
    normalizedTitle: issue?.normalizedTitle || reviewItem.normalizedTitle || reviewItem.suggestedTitle,
    timeLabel: issue?.timeLabel || reviewItem.timeLabel,
  };
}

function shouldShowPreviewIssue(issue) {
  if (!issue) return false;
  const reviewItem = issue?.id ? reviewIndex.get(issue.id) : null;
  const override = overrides[issue.id] || {};
  const include = typeof override.include === "boolean" ? override.include : reviewItem?.include ?? true;
  if (!include || isSuppressedIssue(issue)) return false;
  const contextualIssue = previewIssueWithReviewContext(issue);
  if (!isShiftCodeIssue(contextualIssue)) return true;
  const normalizedTitle = contextualIssue?.normalizedTitle || contextualIssue?.suggestedTitle || "";
  if (isKnownResolvedShiftCodeValue(contextualIssue?.source, contextualIssue?.rawValue, normalizedTitle)) return false;
  return !isShiftCodeResolvedByActiveRules(contextualIssue);
}

function pruneResolvedLatestPreviewIssues() {
  if (!latestPreview) return;
  latestPreview = {
    ...latestPreview,
    issues: (latestPreview.issues || []).filter(shouldShowPreviewIssue),
  };
}

function issueMatchesSavedParserRule(issue, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedRule) return false;
  return sanitizeIssueSource(issue?.source) === normalizedRule.source
    && sanitizeRuleSeniority(issue?.seniority) === normalizedRule.seniority
    && parserRuleCodeForIssue(issue) === normalizedRule.code;
}

function parserRulesEquivalent(left, right) {
  const leftRule = sanitizeParserExtensionRule(left);
  const rightRule = sanitizeParserExtensionRule(right);
  if (!leftRule || !rightRule) return false;
  return leftRule.source === rightRule.source
    && leftRule.code === rightRule.code
    && leftRule.kind === rightRule.kind
    && leftRule.base === rightRule.base
    && leftRule.period === rightRule.period
    && leftRule.suffix === rightRule.suffix
    && leftRule.allDay === rightRule.allDay
    && leftRule.startTime === rightRule.startTime
    && leftRule.endTime === rightRule.endTime
    && leftRule.location === rightRule.location
    && leftRule.includeAsShift === rightRule.includeAsShift
    && leftRule.ignore === rightRule.ignore;
}

function matchingParserRuleGroup(rule) {
  const normalized = sanitizeParserExtensionRule(rule);
  if (!normalized) return [];
  const sourceKey = normalized.source.toLowerCase();
  return sanitizeParserExtensionRuleList(parserExtensions?.[sourceKey], normalized.source)
    .filter((item) => parserRulesEquivalent(item, normalized));
}

function parserRulePreviewTitle(rule, sourceSettings = settings) {
  if (!rule) return "";
  if (rule.ignore === true || rule.kind === "ignore") return `${rule.source}: ${rule.code}`;
  if (String(rule.base || "").trim().toLowerCase() === "hub" && String(rule.period || "").trim().toUpperCase() === "NIGHT") {
    return sourceSettings.showSourcePrefix ? `${rule.source}: Night Hub` : "Night Hub";
  }
  const parts = [];
  if (rule.base) parts.push(rule.base);
  if (sourceSettings.showAmPm && rule.period) parts.push(rule.period === "NIGHT" ? "Night" : rule.period);
  if (rule.suffix) parts.push(rule.suffix);
  const core = parts.join(" ").trim();
  if (!core) return "";
  return sourceSettings.showSourcePrefix ? `${rule.source}: ${core}` : core;
}

function parserRulePreviewMeta(rule) {
  if (!rule) return "";
  if (rule.ignore === true || rule.kind === "ignore") return "Ignored shift code";
  if (rule.includeAsShift === false) return "Hidden from calendar";
  const meta = [];
  meta.push(rule.allDay ? "All day" : summarizeEventTimes(`2000-01-01T${rule.startTime}:00`, `2000-01-01T${rule.endTime}:00`, false));
  if (rule.location) meta.push(rule.location);
  return meta.join(" · ");
}

function renderParserRulePreview() {
  if (!parserRulePreview) return;
  const source = sanitizeIssueSource(parserRuleSource?.value);
  const seniorities = selectedParserRuleSeniorities();
  const ignore = parserRuleIgnore?.checked === true;
  const rule = sanitizeParserExtensionRule({
    source,
    seniority: seniorities[0],
    code: parserRuleCode?.value,
    kind: ignore ? "ignore" : "shift",
    ignore,
    base: parserRuleBase?.value,
    period: parserRulePeriod?.value,
    suffix: parserRuleSuffix?.value,
    allDay: parserRuleAllDay?.checked,
    startTime: parseEditorTimeInput(parserRuleStartTime?.value || ""),
    endTime: parseEditorTimeInput(parserRuleEndTime?.value || ""),
    location: parserRuleLocation?.value,
    includeAsShift: ignore ? false : parserRuleIncludeAsShift?.checked,
  }, source);
  if (!rule) {
    parserRulePreview.innerHTML = `
      <div>
        <strong>Preview</strong>
        <p>Fill in the rule to preview the final calendar output before saving.</p>
      </div>
    `;
    return;
  }
  parserRulePreview.innerHTML = `
    <div>
      <strong>Preview</strong>
      <p>${escapeHtml(parserRulePreviewTitle(rule))}</p>
      <p>${escapeHtml(seniorities.length > 1 ? `${seniorities.length} seniorities` : rule.seniority)} · ${escapeHtml(parserRulePreviewMeta(rule))}</p>
    </div>
  `;
}

function syncParserRuleIgnoreControls() {
  const ignore = parserRuleIgnore?.checked === true;
  if (parserRuleIncludeAsShift) {
    if (ignore) parserRuleIncludeAsShift.checked = false;
    parserRuleIncludeAsShift.disabled = ignore;
  }
  if (parserRuleBase) parserRuleBase.required = !ignore;
  if (parserRuleAllDay) parserRuleAllDay.disabled = ignore;
  if (parserRuleStartTime) parserRuleStartTime.disabled = ignore;
  if (parserRuleEndTime) parserRuleEndTime.disabled = ignore;
  if (ignore && parserRuleTimeFields) parserRuleTimeFields.classList.add("hidden");
  if (!ignore && parserRuleTimeFields) parserRuleTimeFields.classList.toggle("hidden", parserRuleAllDay?.checked === true);
}

function setParserRuleSourceRawReadonly(readonly = true) {
  if (parserRuleSource) parserRuleSource.readOnly = readonly;
  if (parserRuleRawValue) parserRuleRawValue.readOnly = readonly;
}

function setParserRuleModalIssueFields(issue, selectedSeniorities = [], options = {}) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const rawValue = String(issue?.rawValue || issue?.code || "").trim();
  const code = parserRuleCodeForIssue({ ...issue, rawValue });
  parserRuleIssueId.value = issue?.fingerprint || issue?.id || issueFingerprint(source, rawValue || code, seniority);
  parserRuleSource.value = source;
  parserRuleRawValue.value = rawValue;
  populateParserRuleSeniorityOptions(selectedSeniorities.length ? selectedSeniorities : [seniority], options.allowMultipleSeniorities !== false);
  parserRuleOriginalSeniority.value = seniority;
  parserRuleOriginalCode.value = code;
  parserRuleCode.value = code;
  parserRuleBase.value = parserRuleBaseForIssue(issue);
  parserRulePeriod.value = parserRulePeriodForIssue(issue);
  parserRuleSuffix.value = parserRuleSuffixForIssue(issue);
  parserRuleAllDay.checked = !issue?.timeLabel || issue.timeLabel === "All day";
  parserRuleStartTime.value = parserRuleAllDay.checked ? "" : timeRangeParts(issue?.timeLabel).start;
  parserRuleEndTime.value = parserRuleAllDay.checked ? "" : timeRangeParts(issue?.timeLabel).end;
  parserRuleLocation.value = issue?.location || defaultLocationForIssueSource(source);
  if (parserRuleIgnore) parserRuleIgnore.checked = options.ignore === true || issue?.ignore === true || issue?.kind === "ignore";
  parserRuleIncludeAsShift.checked = true;
  parserRuleTimeFields.classList.toggle("hidden", parserRuleAllDay.checked);
  setParserRuleSourceRawReadonly(options.readonlySourceRaw !== false);
  syncParserRuleIgnoreControls();
}

function openParserRuleModalFromSyntheticIssue(issue, options = {}) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = String(issue?.code || parserRuleCodeForIssue(issue)).trim().toUpperCase();
  if (!source || !seniority || !code) {
    setStatus("Could not prepare that shift-code rule.", true);
    return;
  }
  parserRuleSaveContext = {
    mode: options.mode || (isCreatorAuthenticated() ? "global" : "local"),
    suggestionId: "",
    targetEmail: options.targetEmail || adminViewingEmail || currentUserEmail,
    replacementTargets: [],
  };
  setParserRuleModalIssueFields({ ...issue, source, seniority, code }, [seniority], {
    allowMultipleSeniorities: options.allowMultipleSeniorities,
    readonlySourceRaw: options.readonlySourceRaw,
  });
  parserRuleModalTitle.textContent = options.title || "Add shift code";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function openManualParserRuleModal() {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to add shift-code rules.", true);
    return;
  }
  parserRuleSaveContext = { mode: "global", suggestionId: "", targetEmail: "", replacementTargets: [] };
  parserRuleIssueId.value = "";
  parserRuleOriginalCode.value = "";
  parserRuleOriginalSeniority.value = "Unknown";
  parserRuleSource.value = "";
  parserRuleRawValue.value = "";
  parserRuleCode.value = "";
  populateParserRuleSeniorityOptions(["Unknown"], true);
  parserRuleBase.value = "";
  parserRulePeriod.value = "";
  parserRuleSuffix.value = "";
  parserRuleAllDay.checked = false;
  if (parserRuleIgnore) parserRuleIgnore.checked = false;
  parserRuleIncludeAsShift.checked = true;
  parserRuleStartTime.value = "";
  parserRuleEndTime.value = "";
  parserRuleLocation.value = "";
  parserRuleTimeFields.classList.remove("hidden");
  setParserRuleSourceRawReadonly(false);
  syncParserRuleIgnoreControls();
  parserRuleModalTitle.textContent = "Add shift code";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function openParserRuleModal(email, errorId = "", selectedSeniorities = [], options = {}) {
  const issue = findAdminIssue(email, errorId);
  if (!issue) {
    setStatus("Could not find that parser warning.", true);
    return;
  }
  parserRuleSaveContext = { mode: "global", suggestionId: "", targetEmail: normalizeEmail(email), replacementTargets: [] };
  setParserRuleModalIssueFields(issue, selectedSeniorities.length ? selectedSeniorities : [issue.seniority], options);
  parserRuleModalTitle.textContent = "Edit shift code";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function openParserRuleModalFromPreviewIssue(issueId = "") {
  const issue = previewIssueWithReviewContext((latestPreview?.issues || []).find((item) => item.id === issueId) || null);
  if (!issue || !shouldShowPreviewIssue(issue)) {
    pruneResolvedLatestPreviewIssues();
    rebuildClientPreview();
    setStatus("That parser warning has already been resolved.");
    return;
  }
  parserRuleSaveContext = {
    mode: isCreatorAuthenticated() ? "global" : "local",
    suggestionId: "",
    targetEmail: adminViewingEmail || currentUserEmail,
    replacementTargets: [],
  };
  const reviewItem = reviewIndex.get(issue.id);
  setParserRuleModalIssueFields({
    ...issue,
    fingerprint: issueFingerprint(issue.source, issue.rawValue, issue.seniority || reviewItem?.seniority),
    seniority: issue.seniority || reviewItem?.seniority,
  }, [issue.seniority || reviewItem?.seniority], { allowMultipleSeniorities: isCreatorAuthenticated() });
  parserRuleModalTitle.textContent = isCreatorAuthenticated() ? "Edit shift code" : "Resolve shift code";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function openParserRuleModalFromRule(source, seniority, code) {
  const rule = findParserExtensionRuleForSeniority(source, seniority, code);
  if (!rule) {
    setStatus("Could not find that saved shift-code rule.", true);
    return;
  }
  const group = matchingParserRuleGroup(rule);
  parserRuleSaveContext = {
    mode: "globalEdit",
    suggestionId: "",
    targetEmail: "",
    replacementTargets: group.map((item) => ({ source: item.source, seniority: item.seniority, code: item.code })),
  };
  parserRuleIssueId.value = "";
  parserRuleSource.value = rule.source;
  parserRuleRawValue.value = rule.code;
  parserRuleOriginalCode.value = rule.code;
  populateParserRuleSeniorityOptions(group.length ? group.map((item) => item.seniority) : [rule.seniority], true);
  parserRuleOriginalSeniority.value = rule.seniority;
  parserRuleCode.value = rule.code;
  parserRuleBase.value = rule.base;
  parserRulePeriod.value = rule.period;
  parserRuleSuffix.value = rule.suffix;
  parserRuleAllDay.checked = rule.allDay;
  if (parserRuleIgnore) parserRuleIgnore.checked = rule.ignore === true || rule.kind === "ignore";
  parserRuleIncludeAsShift.checked = rule.includeAsShift !== false;
  parserRuleStartTime.value = rule.allDay ? "" : rule.startTime;
  parserRuleEndTime.value = rule.allDay ? "" : rule.endTime;
  parserRuleLocation.value = rule.location || "";
  parserRuleTimeFields.classList.toggle("hidden", parserRuleAllDay.checked);
  setParserRuleSourceRawReadonly(true);
  syncParserRuleIgnoreControls();
  parserRuleModalTitle.textContent = "Edit shift code";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function openParserRuleModalFromSuggestion(suggestionId = "") {
  const suggestion = parserRuleSuggestions.find((item) => item.id === suggestionId);
  const rule = sanitizeParserExtensionRule(suggestion?.rule);
  if (!suggestion || !rule) {
    setStatus("Could not find that suggestion.", true);
    return;
  }
  parserRuleSaveContext = { mode: "suggestionOverwrite", suggestionId, targetEmail: suggestion.email };
  parserRuleIssueId.value = suggestion.fingerprint || issueFingerprint(rule.source, rule.code, rule.seniority);
  parserRuleSource.value = rule.source;
  parserRuleRawValue.value = suggestion.rawValue || rule.code;
  parserRuleOriginalCode.value = rule.code;
  populateParserRuleSeniorityOptions([rule.seniority], false);
  parserRuleOriginalSeniority.value = rule.seniority;
  parserRuleCode.value = rule.code;
  parserRuleBase.value = rule.base;
  parserRulePeriod.value = rule.period;
  parserRuleSuffix.value = rule.suffix;
  parserRuleAllDay.checked = rule.allDay;
  if (parserRuleIgnore) parserRuleIgnore.checked = rule.ignore === true || rule.kind === "ignore";
  parserRuleIncludeAsShift.checked = rule.includeAsShift !== false;
  parserRuleStartTime.value = rule.allDay ? "" : rule.startTime;
  parserRuleEndTime.value = rule.allDay ? "" : rule.endTime;
  parserRuleLocation.value = rule.location || "";
  parserRuleTimeFields.classList.toggle("hidden", parserRuleAllDay.checked);
  setParserRuleSourceRawReadonly(true);
  syncParserRuleIgnoreControls();
  parserRuleModalTitle.textContent = "Overwrite suggestion";
  renderParserRulePreview();
  parserRuleModal.classList.remove("hidden");
  parserRuleModal.setAttribute("aria-hidden", "false");
}

function closeParserRuleModal() {
  parserRuleModal?.classList.add("hidden");
  parserRuleModal?.setAttribute("aria-hidden", "true");
  parserRuleForm?.reset();
  if (parserRuleIncludeAsShift) parserRuleIncludeAsShift.disabled = false;
  if (parserRuleBase) parserRuleBase.required = true;
  if (parserRuleAllDay) parserRuleAllDay.disabled = false;
  if (parserRuleStartTime) parserRuleStartTime.disabled = false;
  if (parserRuleEndTime) parserRuleEndTime.disabled = false;
  parserRuleSaveContext = { mode: "global", suggestionId: "", targetEmail: "", replacementTargets: [] };
  parserRuleTimeFields?.classList.remove("hidden");
  setParserRuleSourceRawReadonly(true);
  if (parserRulePreview) {
    parserRulePreview.innerHTML = `
      <div>
        <strong>Preview</strong>
        <p>Fill in the rule to preview the final calendar output before saving.</p>
      </div>
    `;
  }
}

function populateParserRuleSeniorityOptions(selected = [], allowMultiple = false) {
  if (!parserRuleSeniority) return;
  const selectedValues = Array.isArray(selected) ? selected : [selected];
  const normalizedSelected = new Set(selectedValues.map(sanitizeRuleSeniority).filter(Boolean));
  if (!normalizedSelected.size) normalizedSelected.add("Unknown");
  const seniorities = parserRuleSeniorities();
  const concreteSeniorities = seniorities.filter((item) => item !== "Unknown");
  const inputType = allowMultiple ? "checkbox" : "radio";
  const allChecked = allowMultiple && concreteSeniorities.every((seniority) => normalizedSelected.has(seniority));
  const allOption = allowMultiple ? `
      <label class="toggle">
        <input type="checkbox" name="parserRuleSeniorityAll" value="all" ${allChecked ? "checked" : ""}>
        All
      </label>
    ` : "";
  parserRuleSeniority.innerHTML = allOption + seniorities
    .map((seniority) => `
      <label class="toggle">
        <input type="${inputType}" name="parserRuleSeniorityOption" value="${escapeHtml(seniority)}" ${normalizedSelected.has(seniority) ? "checked" : ""}>
        ${escapeHtml(seniority)}
      </label>
    `)
    .join("");
  normalizeParserRuleSenioritySelection();
}

function selectedParserRuleSeniorities() {
  if (!parserRuleSeniority) return [];
  return [...parserRuleSeniority.querySelectorAll('input[name="parserRuleSeniorityOption"]:checked')]
    .map((input) => sanitizeRuleSeniority(input.value))
    .filter(Boolean);
}

function normalizeParserRuleSenioritySelection(trigger = null) {
  if (!parserRuleSeniority) return;
  const allInput = parserRuleSeniority.querySelector('input[name="parserRuleSeniorityAll"]');
  const options = [...parserRuleSeniority.querySelectorAll('input[name="parserRuleSeniorityOption"]')];
  const unknown = options.find((input) => sanitizeRuleSeniority(input.value) === "Unknown");
  const concrete = options.filter((input) => sanitizeRuleSeniority(input.value) !== "Unknown");
  if (allInput && trigger === allInput) {
    if (allInput.checked) {
      concrete.forEach((input) => { input.checked = true; });
      if (unknown) unknown.checked = false;
    } else {
      concrete.forEach((input) => { input.checked = false; });
      if (unknown) unknown.checked = true;
    }
  } else if (trigger && concrete.includes(trigger) && trigger.checked) {
    if (unknown) unknown.checked = false;
  } else if (trigger === unknown && unknown?.checked) {
    concrete.forEach((input) => { input.checked = false; });
  }
  const checkedConcrete = concrete.filter((input) => input.checked);
  if (checkedConcrete.length) {
    if (unknown) unknown.checked = false;
  } else if (unknown) {
    unknown.checked = true;
  }
  if (allInput) allInput.checked = concrete.length > 0 && concrete.every((input) => input.checked);
}

async function saveParserRuleFromModal() {
  const saveMode = parserRuleSaveContext.mode || "global";
  if (saveMode !== "local" && !isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to add shift codes.", true);
    return;
  }
  const source = sanitizeIssueSource(parserRuleSource.value);
  const rawValue = String(parserRuleRawValue.value || "").trim();
  const previousCode = String(parserRuleOriginalCode?.value || "").trim().toUpperCase();
  const previousSeniority = sanitizeRuleSeniority(parserRuleOriginalSeniority?.value);
  const code = String(parserRuleCode.value || "").trim().toUpperCase();
  const selectedSeniorities = selectedParserRuleSeniorities();
  const seniority = selectedSeniorities[0] || "";
  const base = String(parserRuleBase.value || "").trim();
  const period = String(parserRulePeriod.value || "").trim().toUpperCase();
  const suffix = String(parserRuleSuffix.value || "").trim();
  const ignore = parserRuleIgnore?.checked === true;
  const allDay = parserRuleAllDay.checked;
  const startTime = parseEditorTimeInput(parserRuleStartTime.value);
  const endTime = parseEditorTimeInput(parserRuleEndTime.value);
  const location = String(parserRuleLocation.value || "").trim();
  const includeAsShift = ignore ? false : parserRuleIncludeAsShift?.checked !== false;
  const fingerprint = sanitizeIssueFingerprint(parserRuleIssueId.value);
  if (!source || !selectedSeniorities.length || !code || (!ignore && !base)) {
    setStatus("Source, seniority, shift code, and base title are required.", true);
    return;
  }
  if (!ignore && !allDay && (!startTime || !endTime)) {
    setStatus("Timed shift-code rules need both a start and end time.", true);
    return;
  }
  const ruleTemplate = {
    source,
    code,
    kind: ignore ? "ignore" : "shift",
    ignore,
    base,
    period,
    suffix,
    allDay: ignore ? true : allDay,
    startTime: ignore ? "" : startTime,
    endTime: ignore ? "" : endTime,
    location,
    includeAsShift,
  };
  const rules = selectedSeniorities.map((item) => ({ ...ruleTemplate, seniority: item }));
  const rule = rules[0];
  try {
    if (saveMode === "local") {
      const response = await fetch("/api/state", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          action: "saveLocalParserExtensionRule",
          email: currentUserEmail,
          password: currentUserPassword,
          targetEmail: parserRuleSaveContext.targetEmail || currentUserEmail,
          fingerprint,
          rawValue,
          rule,
        }),
      });
      const data = await readJsonResponse(response, "Could not save your shift-code resolution.");
      applyIssueConfig(data.issueConfig);
      closeParserRuleModal();
      if (selectedFiles.length) {
        parsedRosterSources = null;
        await analyzeFiles({ preserveVisiblePreview: true });
      } else if (parsedRosterSources) {
        await updatePreview();
      } else if (latestPreview) {
        latestPreview = {
          ...latestPreview,
          issues: (latestPreview.issues || []).filter((issue) => sanitizeIssueFingerprint(issueFingerprint(issue.source, issue.rawValue, issue.seniority)) !== fingerprint),
        };
        rebuildClientPreview();
      }
      await refreshActiveWhoInsightSurfaces();
      setStatus(ignore ? "Shift code ignored for your calendar." : "Shift code resolved for your calendar and sent to Admin.");
      return;
    }
    if (saveMode === "suggestionOverwrite") {
      await decideParserRuleSuggestion(parserRuleSaveContext.suggestionId, "approveUser", rule);
      closeParserRuleModal();
      await refreshActiveWhoInsightSurfaces();
      setStatus("Suggestion overwritten for that user.");
      return;
    }
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "saveParserExtensionRule",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        fingerprint,
        source,
        rawValue,
        previousCode,
        previousSeniority,
        replacementTargets: saveMode === "globalEdit" ? parserRuleSaveContext.replacementTargets || [] : [],
        rules,
        rule,
      }),
    });
    const data = await readJsonResponse(response, "Could not save the shift-code rule.");
    parserExtensions = sanitizeParserExtensions(data.parserExtensions);
    setParserExtensions(parserExtensions);
    pruneResolvedLatestPreviewIssues();
    if (fingerprint) ignoredIssueFingerprints.delete(fingerprint);
    closeParserRuleModal();
    await loadServerUsers();
    renderAccountsModal();
    if (parsedRosterSources) {
      await updatePreview();
    } else if (latestPreview) {
      latestPreview = {
        ...latestPreview,
        issues: (latestPreview.issues || []).filter((issue) => !rules.some((savedRule) => issueMatchesSavedParserRule(issue, savedRule))),
        review: (latestPreview.review || []).map((item) => rules.some((savedRule) => issueMatchesSavedParserRule(item, savedRule)) ? {
          ...item,
          status: "ok",
          warnings: [],
        } : item),
      };
      rebuildClientPreview();
    } else if (selectedFiles.length) {
      parsedRosterSources = null;
      await analyzeFiles({ preserveVisiblePreview: true });
    }
    await refreshActiveWhoInsightSurfaces();
    setStatus(ignore ? "Shift code ignored." : includeAsShift ? "Shift code added to the parser." : "Shift code hidden from calendar.");
  } catch (error) {
    setStatus(error.message || "Could not save the shift-code rule.", true);
  }
}

async function refreshActiveWhoInsightSurfaces() {
  const panels = [...document.querySelectorAll(".event-inline-insight[data-inline-who-date]")]
    .filter((panel) => !panel.classList.contains("hidden"));
  await Promise.all(panels.map((panel) => renderInlineWhoInsight(panel, panel.dataset.inlineWhoDate || "", {
    source: panel.dataset.inlineWhoSource || "",
  })));
  if (!insightsModal.classList.contains("hidden") && insightsState?.mode === "who") {
    await renderWhoInsight();
  }
}

async function deleteParserRule(source, seniority, code) {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to delete shift-code rules.", true);
    return;
  }
  const rule = findParserExtensionRuleForSeniority(source, seniority, code);
  if (!rule) {
    setStatus("Could not find that saved shift-code rule.", true);
    return;
  }
  const confirmed = window.confirm(`Delete the ${rule.source} ${rule.seniority} shift-code rule for ${rule.code}? Roster entries will be rechecked and may return as Unknown.`);
  if (!confirmed) return;
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "deleteParserExtensionRule",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        rule: {
          source: rule.source,
          seniority: rule.seniority,
          code: rule.code,
        },
      }),
    });
    const data = await readJsonResponse(response, "Could not delete the shift-code rule.");
    parserExtensions = sanitizeParserExtensions(data.parserExtensions);
    setParserExtensions(parserExtensions);
    await loadServerUsers();
    renderAccountsModal();
    if (selectedFiles.length) {
      parsedRosterSources = null;
      await analyzeFiles({ preserveVisiblePreview: true });
    } else if (latestPreview) {
      await updatePreview();
    }
    setStatus("Shift-code rule deleted. Matching roster entries have been rechecked.");
  } catch (error) {
    setStatus(error.message || "Could not delete the shift-code rule.", true);
  }
}

async function reportAccountError(issue, errorId = "") {
  if (!issue?.message || !cloudAvailable || !currentUserEmail || !currentUserPassword) return;
  try {
    await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "reportUserError",
        email: adminViewingEmail ? authUserEmail || currentUserEmail : currentUserEmail,
        password: adminViewingEmail ? authUserPassword || currentUserPassword : currentUserPassword,
        targetEmail: adminViewingEmail ? currentUserEmail : "",
        errorId,
        message: `${formatIssueHeading(issue)} — ${issue.message}`,
        issue,
      }),
    });
    if (isCreatorAuthenticated()) {
      await loadServerUsers();
      syncAccountsButton();
      if (!accountsModal.classList.contains("hidden") && currentAdminTab === "errors") renderAccountsModal();
    }
  } catch {
    // Keep UI responsive if error reporting fails.
  }
}

async function updateAccountDetails(email, patch) {
  accountState.users = accountState.users.map((user) => user.email === email ? {
    ...user,
    password: patch.password || user.password || "",
    realName: patch.realName || user.realName || "",
  } : user);
  saveAccountState();
  if (cloudAvailable) {
    try {
      const response = await fetch("/api/state", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          action: "updateAccount",
          email: adminViewingEmail ? authenticatedAccountEmail() : viewedAccountEmail(),
          password: adminViewingEmail ? authUserPassword || currentUserPassword : currentUserPassword,
          targetEmail: adminViewingEmail ? viewedAccountEmail() : "",
          realName: patch.realName || "",
          newPassword: patch.password || "",
        }),
      });
      const data = await readJsonResponse(response, "Could not update account.");
      currentRosterClaims = sanitizeRosterClaims(data.claims || currentRosterClaims);
      currentSuggestedClaims = sanitizeRosterClaims(data.suggestedClaims || data.nameMatches || currentSuggestedClaims);
      currentSnapshot = null;
      if (normalizeEmail(email) === currentUserEmail && patch.password) currentUserPassword = patch.password;
      if (isCreatorAuthenticated()) await loadServerUsers();
    } catch (error) {
      setStatus(error.message || "Could not update account.", true);
      return;
    }
  }
  renderLoginState();
  if (latestPreview) rebuildClientPreview();
  setStatus("Account details updated.");
}

function addLocalAccount() {
  const nextEmail = `user${accountState.users.length}@example.com`;
  accountState.users.push({ email: nextEmail, realName: "", password: "", role: "user" });
  saveAccountState();
  setStatus("User added to local account list.");
}

async function createAccountFromOwner(formElement) {
  if (!isCreatorAuthenticated()) {
    setStatus("Creator authentication is required to create accounts.", true);
    return;
  }
  const realName = formElement.querySelector("[data-create-real-name]")?.value.trim() || "";
  const email = normalizeEmail(formElement.querySelector("[data-create-email]")?.value || "");
  const password = formElement.querySelector("[data-create-password]")?.value || "";
  if (!realName || !email || !password) {
    setStatus("Enter a real name, email address, and temporary password.", true);
    return;
  }

  setStatus(`Creating ${email}...`);
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "adminCreateUser",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail: email,
        targetRealName: realName,
        targetPassword: password,
      }),
    });
    const data = await readJsonResponse(response, "Could not create account.");
    if (data.user) {
      serverUsers = [
        ...serverUsers.filter((user) => normalizeServerUser(user).email !== email),
        data.user,
      ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
    }
    formElement.reset();
    await loadServerUsers();
    await enterUserAccount(email);
  } catch (error) {
    if (error.message === "Cloud storage is not configured.") {
      setStatus(serverStorageRequiredMessage(), true);
      return;
    }
    setStatus(error.message || "Could not create account.", true);
  }
}

async function setUserInsightsEnabled(email, enabled) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || !isCreatorAuthenticated()) return;
  const previousUsers = serverUsers.map((user) => ({ ...normalizeServerUser(user) }));
  serverUsers = serverUsers.map((user) => {
    const normalized = normalizeServerUser(user);
    return normalized.email === targetEmail ? { ...normalized, insightsEnabled: enabled === true } : normalized;
  });
  renderAccountsModal();
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setUserInsightsEnabled",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail,
        insightsEnabled: enabled === true,
      }),
    });
    const data = await readJsonResponse(response, "Could not update user feature access.");
    if (data.user) {
      serverUsers = [
        ...serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail),
        data.user,
      ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
      renderAccountsModal();
    }
    if (targetEmail === currentUserEmail) {
      currentInsightsEnabled = enabled === true;
      rebuildClientPreview();
    }
    setStatus(enabled ? "Working-with tools enabled for that user." : "Working-with tools disabled for that user.");
  } catch (error) {
    serverUsers = previousUsers;
    renderAccountsModal();
    setStatus(error.message || "Could not update user feature access.", true);
  }
}

async function confirmSuggestedClaim(index) {
  const claim = currentSuggestedClaims[index];
  if (!claim) return;
  await claimSelectedRosterName(claim);
}

async function rejectSuggestedClaim(index) {
  const claim = currentSuggestedClaims[index];
  if (!claim) return;
  currentSuggestedClaims = currentSuggestedClaims.filter((_, itemIndex) => itemIndex !== index);
  latestNameMatches = currentSuggestedClaims;
  await reportRosterIdentityIssue(`Rejected suggested roster name ${claim.displayName} (${claim.sourceType.toUpperCase()}).`);
  renderAccountsModal();
  renderClaimSection();
}

async function removeRosterClaim(marker, claimEmail = "") {
  const [sourceType, ...keyParts] = String(marker || "").split(":");
  const key = keyParts.join(":");
  if (!sourceType || !key) return;
  const targetEmail = normalizeEmail(claimEmail || currentUserEmail);
  const creatorAction = isCreatorAuthenticated() && targetEmail && targetEmail !== currentUserEmail;
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "removeRosterClaim",
        email: creatorAction ? authUserEmail || currentUserEmail : currentUserEmail,
        password: creatorAction ? authUserPassword || currentUserPassword : currentUserPassword,
        targetEmail: creatorAction ? targetEmail : "",
        claim: { sourceType, key },
      }),
    });
    const data = await readJsonResponse(response, "Could not remove roster name.");
    if (targetEmail === currentUserEmail) {
      currentRosterClaims = sanitizeRosterClaims(data.claims || []);
      currentSnapshot = null;
      await replaceStoredImports([]);
      selectedFiles = [];
      renderDoctorState();
    }
    await reportRosterIdentityIssue(`Removed wrong roster name ${key} (${sourceType.toUpperCase()}).`, targetEmail);
    if (isCreatorAuthenticated()) await loadServerUsers();
    renderAccountsModal();
    setStatus("Roster name removed.");
  } catch (error) {
    setStatus(error.message || "Could not remove roster name.", true);
  }
}

async function reportRosterIdentityIssue(message = "Roster name match needs review.", email = currentUserEmail) {
  if (!cloudAvailable && !isCreatorAuthenticated()) return;
  const targetEmail = normalizeEmail(email || currentUserEmail);
  const creatorAction = isCreatorAuthenticated() && targetEmail && targetEmail !== currentUserEmail;
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "reportRosterIdentityIssue",
        email: creatorAction ? authUserEmail || currentUserEmail : currentUserEmail,
        password: creatorAction ? authUserPassword || currentUserPassword : currentUserPassword,
        targetEmail: creatorAction ? targetEmail : "",
        message,
      }),
    });
    await readJsonResponse(response, "Could not report roster name problem.");
    if (isCreatorAuthenticated()) await loadServerUsers();
    syncAccountsButton();
    setStatus("Roster name problem reported.");
  } catch (error) {
    setStatus(error.message || "Could not report roster name problem.", true);
  }
}

async function addAdminRosterClaim(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || !isCreatorAuthenticated()) return;
  const select = [...accountsBody.querySelectorAll("[data-admin-claim-select]")]
    .find((item) => normalizeEmail(item.dataset.adminClaimSelect) === targetEmail);
  const selected = select ? availableRosterDoctors[Number(select.value)] : null;
  if (!selected) {
    setStatus("Choose a roster name to add.", true);
    return;
  }
  const user = serverUsers.map(normalizeServerUser).find((item) => item.email === targetEmail);
  const nextClaims = [
    ...(user?.claims || []),
    selected,
  ];
  await saveAdminRosterClaims(targetEmail, nextClaims);
}

function toggleAdminRosterClaimControls(email) {
  const targetEmail = normalizeEmail(email);
  const editor = [...accountsBody.querySelectorAll("[data-claim-editor]")]
    .find((item) => normalizeEmail(item.dataset.claimEditor) === targetEmail);
  const select = [...accountsBody.querySelectorAll("[data-admin-claim-select]")]
    .find((item) => normalizeEmail(item.dataset.adminClaimSelect) === targetEmail);
  if (editor) {
    const isHidden = editor.classList.toggle("hidden");
    if (!isHidden && select) select.focus();
    return;
  }
  if (select) {
    select.focus();
  }
}

async function saveAdminRosterClaims(email, claims) {
  const targetEmail = normalizeEmail(email);
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setAccountRosterClaims",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail,
        claims,
      }),
    });
    const data = await readJsonResponse(response, "Could not update roster names.");
    if (data.user) {
      serverUsers = [
        ...serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail),
        data.user,
      ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
    }
    await loadServerUsers();
    renderAccountsModal();
    setStatus("Roster names updated.");
  } catch (error) {
    setStatus(error.message || "Could not update roster names.", true);
  }
}

function removeLocalAccount(email) {
  accountState.users = accountState.users.filter((user) => user.email !== email);
  if (!accountState.users.some((user) => user.email === accountState.currentEmail)) {
    accountState.currentEmail = "";
  }
  saveAccountState();
  setStatus("User removed from local account list.");
}

async function deleteAccount(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail) return;
  if (targetEmail === OWNER_EMAIL) {
    setStatus("The creator account cannot be deleted from the app.", true);
    return;
  }

  const deletingViewedAccount = targetEmail === viewedAccountEmail();
  const creatorDeletingSwitchedUser = deletingViewedAccount && Boolean(adminViewingEmail) && isCreatorAuthenticated();
  const confirmed = window.confirm(`Delete account ${targetEmail}? This removes the account login and saved workspace. This cannot be undone.`);
  if (!confirmed) return;

  const creatorCanDelete = isCreatorAuthenticated();
  const requestEmail = creatorCanDelete ? authUserEmail || currentUserEmail : currentUserEmail;
  const requestPassword = creatorCanDelete ? authUserPassword || currentUserPassword : currentUserPassword;
  cancelScheduledCloudStateSave();

  try {
    if (cloudAvailable) {
      const response = await fetch("/api/state", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          action: "deleteAccount",
          email: requestEmail,
          password: requestPassword,
          targetEmail: creatorCanDelete ? targetEmail : "",
        }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.error || "Account deletion failed.");
    }

    deleteLocalAccountData(targetEmail);
    clearDeletedAccountClaims(targetEmail);
    if (creatorCanDelete) await loadServerUsers();

    if (creatorDeletingSwitchedUser) {
      closeAccountsModal();
      await returnToCreatorCalendar();
      setStatus(`Deleted ${targetEmail}.`);
      return;
    }

    if (deletingViewedAccount) {
      closeAccountsModal();
      localStorage.removeItem(CURRENT_EMAIL_KEY);
      sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
      localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
      currentUserEmail = "";
      currentUserPassword = "";
      authUserEmail = "";
      authUserPassword = "";
      adminViewingEmail = "";
      currentUserRole = "user";
      cloudAvailable = false;
      setActiveCalendarContext("claimed-account", { email: "" });
      await clearLocalWorkspace();
      renderLoginState();
      openLoginModal();
      setEntranceStatus("Account deleted.");
      setStatus("Account deleted.");
      return;
    }

    renderAccountsModal();
    renderDoctorState();
    setStatus(`Deleted ${targetEmail}.`);
  } catch (error) {
    setStatus(error.message || "Account deletion failed.", true);
  }
}

function deleteLocalAccountData(email) {
  accountState.users = accountState.users.filter((user) => user.email !== email);
  if (!accountState.users.some((user) => user.email === accountState.currentEmail)) {
    accountState.currentEmail = "";
  }
  saveAccountState();
  const store = loadWorkspaceStore();
  delete store[email];
  saveWorkspaceStore(store);
}

function clearDeletedAccountClaims(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail) return;
  const deletedUser = serverUsers.map(normalizeServerUser).find((item) => item.email === targetEmail)
    || accountState.users.find((item) => normalizeEmail(item.email) === targetEmail);
  const deletedClaims = sanitizeRosterClaims(deletedUser?.claims || []);
  serverUsers = serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail);
  availableRosterDoctors = clearClaimedDoctorMetadata(availableRosterDoctors, targetEmail, deletedClaims);
  doctorOptions = clearClaimedDoctorMetadata(doctorOptions, targetEmail, deletedClaims);
  if (currentSnapshot) {
    currentSnapshot = sanitizeWorkspaceSnapshot({
      ...currentSnapshot,
      doctorOptions: clearClaimedDoctorMetadata(currentSnapshot.doctorOptions || [], targetEmail, deletedClaims),
      insightCache: currentSnapshot.insightCache
        ? {
            ...currentSnapshot.insightCache,
            doctorOptions: clearClaimedDoctorMetadata(currentSnapshot.insightCache.doctorOptions || [], targetEmail, deletedClaims),
          }
        : currentSnapshot.insightCache,
    });
  }
}

function clearClaimedDoctorMetadata(doctors, deletedEmail, deletedClaims = []) {
  if (!Array.isArray(doctors)) return [];
  const normalizedEmail = normalizeEmail(deletedEmail);
  const claimMarkers = new Set(sanitizeRosterClaims(deletedClaims).map((claim) => `${claim.sourceType}:${claim.key}`));
  const claimKeys = new Set(sanitizeRosterClaims(deletedClaims).map((claim) => claim.key));
  return doctors.map((doctor) => {
    const sourceTypes = normalizedDoctorSourceTypes(doctor);
    const markers = sourceTypes.map((sourceType) => `${sourceType}:${doctor.key}`);
    const claimedByDeletedEmail = normalizedEmail && normalizeEmail(doctor.accountEmail || doctor.claimedBy || "") === normalizedEmail;
    const matchesDeletedClaim = markers.some((marker) => claimMarkers.has(marker)) || claimKeys.has(doctor.key);
    if (!claimedByDeletedEmail && !matchesDeletedClaim) return doctor;
    const cleaned = { ...doctor };
    delete cleaned.accountEmail;
    delete cleaned.claimedBy;
    delete cleaned.claimedByName;
    return cleaned;
  });
}

function accountCalendarContextForEmail(email) {
  const targetEmail = normalizeEmail(email);
  return calendarSnapshotContext({
    mode: targetEmail === OWNER_EMAIL ? "creator-account" : "claimed-account",
    ownerEmail: targetEmail,
    doctorKey: preferredDoctorKeyForAccountEmail(targetEmail),
  });
}

async function validateClaimedAccountCalendarInBackground(context = {}, options = {}) {
  const targetEmail = normalizeEmail(context.ownerEmail || context.ownerId);
  const cachedSnapshot = options.preserveRenderedSnapshot ? currentSnapshot : null;
  const cachedRevision = cachedSnapshot?.calendarRevision || "";
  await restoreCloudState({
    adminTargetEmail: targetEmail === OWNER_EMAIL ? "" : targetEmail,
    preserveSessionOnFailure: true,
    deferHydration: true,
    accountSwitchStartedAt: options.accountSwitchStartedAt,
  });
  if (cachedSnapshot?.preview && (!currentSnapshot?.preview || currentSnapshot.cacheKey !== cachedSnapshot.cacheKey)) {
    currentSnapshot = cachedSnapshot;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = cachedSnapshot.cachedAt || cachedSnapshot.preview?.lastParsed || "";
  }
  await hydrateAuthenticatedWorkspace({
    adminTargetEmail: targetEmail === OWNER_EMAIL ? "" : targetEmail,
    includeBootstrap: true,
    accountSwitchStartedAt: options.accountSwitchStartedAt,
    cachedRevision,
  }, 0);
  renderLoginState();
}

async function validateDoctorProfileCalendarInBackground(doctor, previousState, options = {}) {
  const result = await loadUnclaimedDoctorCalendar(doctor, previousState, {
    profile: options.profile,
    cachedRevision: options.cachedRevision || "",
  });
  if (result) {
    await commitCalendarLoad(result, { saveInBackground: true });
  } else {
    renderLoginState();
    setStatus("Calendar is up to date.");
  }
}

async function enterUserAccount(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || (!isOwnerAccount() && !isCreatorAuthenticated())) return;
  const previousState = captureCalendarViewState();
  const accountSwitchStartedAt = performance.now();
  const creatorEmail = authUserEmail || currentUserEmail;
  const creatorPassword = authUserPassword || currentUserPassword;
  if (normalizeEmail(creatorEmail) !== OWNER_EMAIL || !creatorPassword) {
    setStatus("Creator authentication is required to enter another account.", true);
    return;
  }
  queueBackgroundCloudStateSave(capturePendingCloudStateSave() || creatorCalendarSavePayload() || snapshotCloudSavePayload(), { delayMs: 1500 });

  closeAccountsModal();
  authUserEmail = creatorEmail;
  authUserPassword = creatorPassword;
  adminViewingEmail = targetEmail;
  activeDoctorProfile = null;
  setActiveCalendarContext(targetEmail === OWNER_EMAIL ? "creator-account" : "claimed-account", { email: targetEmail });
  currentUserEmail = targetEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = targetEmail === OWNER_EMAIL ? "creator" : "user";
  forceConsoleSkin();
  setStatus(`Entering ${targetEmail}...`);
  const targetContext = accountCalendarContextForEmail(targetEmail);
  const renderedCachedSnapshot = renderCachedCalendarSnapshotForContext(targetContext, { accountSwitchStartedAt });
  renderLoginState();
  try {
    const validation = validateClaimedAccountCalendarInBackground(targetContext, {
      accountSwitchStartedAt,
      preserveRenderedSnapshot: renderedCachedSnapshot,
    });
    if (renderedCachedSnapshot) {
      void validation.catch((error) => {
        setStatus(normalizeAuthMessage(error.message || `Could not update ${targetEmail}.`), true);
      });
    } else {
      await validation;
    }
  } catch (error) {
    restoreCalendarViewState(previousState);
    renderLoginState();
    setStatus(normalizeAuthMessage(error.message || `Could not enter ${targetEmail}.`), true);
  }
}

async function enterDoctorProfileView(doctor) {
  if (!isOwnerAccount() && !isCreatorAuthenticated()) return;
  const previousState = captureCalendarViewState();
  const creatorEmail = authUserEmail || currentUserEmail;
  const creatorPassword = authUserPassword || currentUserPassword;
  queueBackgroundCloudStateSave(capturePendingCloudStateSave() || snapshotCloudSavePayload(), { delayMs: 1500 });
  const profile = doctorProfileForDoctor(doctor);
  const targetContext = profile ? calendarSnapshotContext({
    mode: "doctor-profile",
    ownerId: profile.ownerId,
    doctorKey: profile.doctorKey,
  }) : null;
  adminViewingEmail = "";
  currentUserEmail = creatorEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = "creator";
  activeDoctorProfile = profile;
  if (profile) setActiveCalendarContext("doctor-profile", { email: currentUserEmail, profile });
  localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  setStatus(`Opening ${doctor.displayName}...`);
  const renderedCachedSnapshot = targetContext
    ? renderCachedCalendarSnapshotForContext(targetContext, { accountSwitchStartedAt: performance.now() })
    : false;
  renderLoginState();
  try {
    const validation = validateDoctorProfileCalendarInBackground(doctor, previousState, {
      profile,
      cachedRevision: renderedCachedSnapshot ? currentSnapshot?.calendarRevision || "" : "",
      renderedCachedSnapshot,
    });
    if (renderedCachedSnapshot) {
      void validation.catch((error) => {
        setStatus(error.message || `Could not update ${doctor.displayName}.`, true);
      });
    } else {
      await validation;
    }
  } catch (error) {
    restoreCalendarViewState(previousState);
    renderLoginState();
    setStatus(error.message || `Could not open ${doctor.displayName}.`, true);
  }
}

function doctorProfileForDoctor(doctor) {
  const profileId = buildDoctorProfileId(doctor);
  if (!profileId) return null;
  return {
    id: profileId,
    ownerId: `doctor-profile:${profileId}`,
    doctorKey: doctor.key,
    displayName: doctor.displayName,
    sourceTypes: doctorProfileSourceTypes(doctor),
    aliases: Array.isArray(doctor.aliases) ? doctor.aliases : [],
  };
}

async function loadUnclaimedDoctorCalendar(doctor, sourceContext, options = {}) {
  const profile = options.profile || doctorProfileForDoctor(doctor);
  if (!profile?.id) {
    throw new Error(`Could not open ${doctor?.displayName || "that clinician"} because their roster source was not available.`);
  }
  const profileData = await fetchDoctorProfileState(profile, { cachedRevision: options.cachedRevision || "" });
  if (profileData.snapshotCurrent === true) {
    currentCalendarRevision = String(profileData.calendarRevision || currentCalendarRevision || "");
    if (currentSnapshot) currentSnapshot.calendarRevision = currentCalendarRevision;
    if (currentSnapshot) saveCalendarSnapshotCacheForContext(currentSnapshot, {
      mode: "doctor-profile",
      ownerId: profile.ownerId,
      doctorKey: profile.doctorKey,
    });
    return null;
  }
  const snapshot = sanitizeWorkspaceSnapshot(profileData.snapshot);
  if (snapshot && profileData.calendarRevision) snapshot.calendarRevision = String(profileData.calendarRevision);
  const snapshotUsable = snapshot?.preview && snapshot?.doctorOptions?.length && profileData.snapshotStale !== true;
  if (snapshotUsable) {
    return {
      mode: "doctor-profile",
      ownerId: profile.ownerId,
      doctor: { ...doctor, sourceTypes: profile.sourceTypes },
      profile,
      imports: importRefsToClientEntries(snapshot.fileRefs || []),
      snapshot,
      session: snapshot.session || profileData.profile?.state?.session || {},
      preview: snapshot.preview,
      doctorOptions: snapshot.doctorOptions || [],
      detectedSources: snapshot.detectedSources || {},
      calendarRevision: profileData.calendarRevision || "",
    };
  }

  const cached = buildUnclaimedPreviewFromSnapshotCache(doctor, profile, sourceContext, profileData.profile?.state?.session || {});
  if (cached) return doctorProfileLoadResultFromCached(profile, cached);
  throw new Error(`${doctor.displayName} could not be loaded from the roster database. Rebuild the roster database from Admin > Files if this persists.`);
}

function hasDoctorProfileImportCandidates(sourceContext) {
  return Boolean(
    sourceContext?.creatorCalendarSourceFileRefs?.length
    || sourceContext?.currentSnapshot?.fileRefs?.length
    || sourceContext?.selectedFiles?.length
    || creatorCalendarSourceFileRefs.length,
  );
}

function doctorProfileLoadResultFromCached(profile, cached) {
  return {
    mode: "doctor-profile",
    ownerId: profile.ownerId,
    doctor: cached.doctor,
    profile,
    imports: cached.imports,
    snapshot: null,
    session: cached.session,
    preview: cached.preview,
    doctorOptions: cached.doctorOptions,
    detectedSources: cached.detectedSources,
  };
}

function buildUnclaimedPreviewFromSnapshotCache(doctor, profile, sourceContext, session = {}) {
  const cache = sanitizeInsightCache(sourceContext?.currentSnapshot?.insightCache);
  if (!cache?.doctorEvents) return null;
  const requestedKeys = [profile.doctorKey, doctor?.key, ...(doctor?.aliases || []).map((alias) => alias.key)]
    .map(normalizeRosterName)
    .filter(Boolean);
  const matchedKey = requestedKeys.find((key) => Object.prototype.hasOwnProperty.call(cache.doctorEvents, key));
  if (!matchedKey) return null;
  const cachedDoctor = cache.doctorOptions.find((item) => item.key === matchedKey)
    || sourceContext?.doctorOptions?.find((item) => normalizeRosterName(item.key) === matchedKey)
    || doctor;
  const selected = {
    ...cachedDoctor,
    key: profile.doctorKey,
    displayName: profile.displayName,
    sourceTypes: profile.sourceTypes,
  };
  const events = (cache.doctorEvents[matchedKey] || [])
    .filter((event) => event && typeof event === "object")
    .map(serializeEvent);
  const review = events.map(reviewItemForCachedEvent);
  return {
    doctor: selected,
    imports: (sourceContext?.selectedFiles || []).map((entry) => ({ ...entry })),
    doctorOptions: buildCreatorDoctorOptions(cache.doctorOptions.length ? cache.doctorOptions : sourceContext?.doctorOptions || [selected]),
    detectedSources: sourceContext?.currentSnapshot?.detectedSources || sourceContext?.detectedSources || {},
    session: {
      ...session,
      doctorKey: selected.key,
      settings: session?.settings || { ...settings },
    },
    preview: {
      ...previewSummary(events),
      events,
      review,
      issues: [],
      conflicts: [],
      imports: [],
      sources: sourceContext?.currentSnapshot?.preview?.sources || [],
      lastParsed: new Date().toISOString(),
    },
  };
}

function reviewItemForCachedEvent(event) {
  return {
    id: event.id,
    source: event.source,
    seniority: event.seniority || "",
    startDay: event.start?.slice(0, 10) || "",
    endDay: event.end?.slice(0, 10) || event.start?.slice(0, 10) || "",
    rawValue: event.rawValue || event.title || "",
    normalizedTitle: event.title,
    suggestedTitle: event.title,
    overrideTitle: "",
    status: "cached",
    warnings: [],
    include: true,
    exportable: true,
    location: event.location || "",
    allDay: event.allDay === true,
    timeLabel: event.timeLabel || summarizeEventTimes(event.start, event.end, event.allDay === true),
  };
}

async function fetchDoctorProfileState(profile, options = {}) {
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "loadDoctorProfile",
      email: authUserEmail || currentUserEmail,
      password: authUserPassword || currentUserPassword,
      profileId: profile.id,
      doctorKey: profile.doctorKey,
      displayName: profile.displayName,
      sourceTypes: profile.sourceTypes,
      aliases: profile.aliases,
      cachedRevision: options.cachedRevision || "",
    }),
  });
  const data = await readJsonResponse(response, "Doctor profile load failed.");
  applyIssueConfig(data.issueConfig);
  return data;
}

async function loadUnclaimedSourceImports(doctor, sourceContext, profile) {
  const refs = uniqueClientFileRefs([
    ...(sourceContext?.creatorCalendarSourceFileRefs || []),
    ...(sourceContext?.currentSnapshot?.fileRefs || []),
    ...((sourceContext?.selectedFiles || []).map(importRefForWorkspace)),
    ...creatorCalendarSourceFileRefs,
  ]);
  if (refs.length) {
    const imports = await loadCloudImportsByRefs(refs).catch(() => []);
    if (imports.length) return imports;
  }
  const previousFiles = sourceContext?.selectedFiles || [];
  if (previousFiles.length) {
    const restored = await loadStoredImportsByRefs(previousFiles.map(importRefForWorkspace)).catch(() => []);
    if (restored.length) return restored;
    const inMemory = previousFiles.filter((entry) => entry.file).map((entry) => ({ ...entry }));
    if (inMemory.length) return inMemory;
  }
  const creatorImports = await loadCreatorAccountImports().catch(() => []);
  if (creatorImports.length) return creatorImports;
  return await loadDoctorProfileImportsForProfile(profile).catch(() => []);
}

function sanitizeClientFileRefs(refs) {
  if (!Array.isArray(refs)) return [];
  return refs.map(importRefForWorkspace).filter((ref) => ref.id);
}

function uniqueClientFileRefs(refs) {
  const seen = new Set();
  const unique = [];
  for (const ref of sanitizeClientFileRefs(refs)) {
    const marker = ref.id || `${ref.name}:${ref.size}:${ref.lastModified}`;
    if (seen.has(marker)) continue;
    seen.add(marker);
    unique.push(ref);
  }
  return unique;
}

async function loadCloudImportsByRefs(refs) {
  return [];
}

async function loadCreatorAccountImports() {
  return [];
}

async function loadDoctorProfileImportsForProfile(profile) {
  return [];
}

async function buildUnclaimedPreviewFromImports(doctor, profile, imports, session = {}) {
  const doctorForBuild = {
    ...doctor,
    key: profile.doctorKey,
    displayName: profile.displayName,
    sourceTypes: profile.sourceTypes,
  };
  const parsed = await parseRosterEntriesLenient(imports, doctorForBuild);
  const parsedDoctors = doctorsByImportId(parsed.sources);
  const allDoctors = buildCreatorDoctorOptions(rosterDoctorOptions(parsed.sources.mmc, parsed.sources.ddh, parsed.sources.casey, parsed.sources.mch));
  const selected = allDoctors.find((item) => item.key === doctorForBuild.key) || doctorForBuild;
  const buildSettings = {
    ...defaultSettings(),
    ...settings,
    ...(session?.settings || {}),
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
  };
  const view = buildRosterView(
    parsed.sources.mmc,
    parsed.sources.ddh,
    selected.key,
    buildSettings,
    sanitizeOverrideState(session?.overrides),
    session?.conflictSelections || {},
    selected.aliases || [],
    parsed.sources.casey,
    parsed.sources.mch,
  );
  const events = view.events;
  return {
    doctor: selected,
    parsedSources: parsed.sources,
    parsedDoctors,
    doctorOptions: allDoctors,
    detectedSources: summarizeDetectedSources(sourceImports(parsed.sources)),
    session: {
      ...session,
      doctorKey: selected.key,
      settings: session?.settings || { ...settings },
    },
    preview: {
      ...previewSummary(events),
      events: events.map(serializeEvent),
      review: view.reviewItems.map(serializeReviewItem),
      issues: view.issues,
      conflicts: view.conflicts.map(serializeConflict),
      imports: view.imports,
      sources: sourceNames(parsed.sources),
      lastParsed: new Date().toISOString(),
    },
  };
}

async function commitCalendarLoad(result, options = {}) {
  if (!result || result.mode !== "doctor-profile") {
    throw new Error("Unsupported calendar load result.");
  }
  activeDoctorProfile = result.profile;
  setActiveCalendarContext("doctor-profile", { email: currentUserEmail, profile: activeDoctorProfile });
  resetTransientCalendarData();
  selectedFiles = result.imports || [];
  currentSnapshot = result.snapshot;
  currentCalendarRevision = String(result.calendarRevision || currentSnapshot?.calendarRevision || currentCalendarRevision || "");
  if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = result.snapshot?.builtAt || "";
  restoredSessionState = result.session || {};
  doctorOptions = result.doctorOptions?.length ? result.doctorOptions : [result.doctor];
  detectedSources = result.detectedSources || {};
  parsedRosterSources = result.parsedSources || null;
  parsedImportDoctors = result.parsedDoctors || new Map();
  latestPreview = JSON.parse(JSON.stringify(result.preview));
  applySessionState(restoredSessionState, { inheritedSettings: rosterDefaultSettings() });
  if (result.doctor?.key) restoredSessionState.doctorKey = result.doctor.key;
  renderSettings();
  renderFilesList();
  renderDoctorState();
  if (doctorOptions.length > 1 && result.doctor?.key) doctorSelect.value = result.doctor.key;
  indexReviewItems(latestPreview.review || []);
  rebuildClientPreview();
  scheduleInsightWarmup();
  cacheCurrentSnapshot(buildActiveSessionState());
  saveCalendarSnapshotCacheForContext(currentSnapshot, {
    mode: "doctor-profile",
    ownerId: result.profile?.ownerId,
    doctorKey: result.profile?.doctorKey,
  });
  saveCurrentWorkspace();
  if (options.saveInBackground) {
    queueBackgroundCloudStateSave(snapshotCloudSavePayload());
  } else {
    await saveCloudState();
  }
  renderLoginState();
  setStatus("Calendar loaded.");
}

function captureCalendarViewState() {
  return {
    adminViewingEmail,
    activeDoctorProfile,
    activeCalendarContext: activeCalendarContext ? { ...activeCalendarContext } : null,
    currentUserEmail,
    currentUserPassword,
    currentUserRole,
    selectedFiles: selectedFiles.map((entry) => ({ ...entry })),
    creatorCalendarSourceFileRefs: creatorCalendarSourceFileRefs.map((entry) => ({ ...entry })),
    currentSnapshot: currentSnapshot ? JSON.parse(JSON.stringify(currentSnapshot)) : null,
    currentSnapshotStale,
    currentSnapshotBuiltAt,
    restoredSessionState: restoredSessionState ? JSON.parse(JSON.stringify(restoredSessionState)) : null,
    doctorOptions: doctorOptions.map((doctor) => ({ ...doctor })),
    detectedSources: JSON.parse(JSON.stringify(detectedSources || {})),
    latestPreview: latestPreview ? JSON.parse(JSON.stringify(latestPreview)) : null,
  };
}

function restoreCalendarViewState(state) {
  if (!state) return;
  adminViewingEmail = state.adminViewingEmail;
  activeDoctorProfile = state.activeDoctorProfile;
  activeCalendarContext = state.activeCalendarContext;
  currentUserEmail = state.currentUserEmail;
  currentUserPassword = state.currentUserPassword;
  currentUserRole = state.currentUserRole;
  if (currentUserEmail) localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  if (currentUserPassword) sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  selectedFiles = state.selectedFiles || [];
  creatorCalendarSourceFileRefs = state.creatorCalendarSourceFileRefs || creatorCalendarSourceFileRefs;
  currentSnapshot = state.currentSnapshot;
  currentSnapshotStale = state.currentSnapshotStale;
  currentSnapshotBuiltAt = state.currentSnapshotBuiltAt;
  restoredSessionState = state.restoredSessionState;
  doctorOptions = state.doctorOptions || [];
  detectedSources = state.detectedSources || {};
  latestPreview = state.latestPreview;
  if (latestPreview) {
    renderWorkspaceFromSnapshot({
      preview: latestPreview,
      session: restoredSessionState || {},
      doctorOptions,
      detectedSources,
      fileRefs: selectedFiles.map(importRefForWorkspace),
    }, restoredSessionState || {});
  } else {
    clearPreviewData();
    renderFilesList();
    renderDoctorState();
  }
}

async function exitDoctorProfileView() {
  try {
    await flushCloudStateSave().catch(() => {});
    cancelScheduledCloudStateSave();
    await saveCloudState();
  } catch {
    // Keep local state even if cloud save fails.
  }
  activeDoctorProfile = null;
  setActiveCalendarContext("creator-account", { email: currentUserEmail });
  clearPreviewData();
  restoredSessionState = loadCurrentSessionState();
  currentSnapshot = sanitizeWorkspaceSnapshot(loadCurrentWorkspace()?.snapshot);
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = "";
  await bootstrapImports();
  renderLoginState();
}

async function returnToCreatorCalendar() {
  await returnToCreatorAccount();
}

async function returnToCreatorAccount() {
  const previousState = captureCalendarViewState();
  const accountSwitchStartedAt = performance.now();
  const creatorEmail = authUserEmail || OWNER_EMAIL;
  const creatorPassword = authUserPassword || currentUserPassword;
  queueBackgroundCloudStateSave(capturePendingCloudStateSave() || snapshotCloudSavePayload(), { delayMs: 1500 });
  adminViewingEmail = "";
  activeDoctorProfile = null;
  currentUserEmail = creatorEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = "creator";
  setActiveCalendarContext("creator-account", { email: currentUserEmail });
  localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  forceConsoleSkin();
  setStatus("Returning to creator account...");
  forceCreatorDoctorSession();
  const targetContext = accountCalendarContextForEmail(OWNER_EMAIL);
  const renderedCachedSnapshot = renderCachedCalendarSnapshotForContext(targetContext, { accountSwitchStartedAt });
  renderLoginState();
  const validateCreator = async () => {
    await validateClaimedAccountCalendarInBackground(targetContext, {
      accountSwitchStartedAt,
      preserveRenderedSnapshot: renderedCachedSnapshot,
    });
    forceCreatorDoctorSession();
    if (currentSnapshot) {
      const snapshotDoctorKey = normalizeRosterName(currentSnapshot.session?.doctorKey || "");
      currentSnapshot = sanitizeWorkspaceSnapshot({
        ...currentSnapshot,
        session: {
          ...(currentSnapshot.session || {}),
          doctorKey: OWNER_DOCTOR_KEY,
        },
      });
      if (snapshotDoctorKey && snapshotDoctorKey !== OWNER_DOCTOR_KEY) {
        currentSnapshot = null;
        currentSnapshotStale = false;
        currentSnapshotBuiltAt = "";
      }
    }
    if (doctorSelect && doctorPickerOptions().some((doctor) => doctor.key === OWNER_DOCTOR_KEY)) {
      doctorSelect.value = OWNER_DOCTOR_KEY;
    }
    if (selectedDoctor()?.key !== OWNER_DOCTOR_KEY) {
      clearPreviewData();
      await updatePreview({ resetRange: false });
    }
    renderLoginState();
  };
  if (renderedCachedSnapshot) {
    void validateCreator().catch((error) => {
      setStatus(normalizeAuthMessage(error.message || "Could not update the creator calendar."), true);
    });
  } else {
    try {
      await validateCreator();
    } catch (error) {
      restoreCalendarViewState(previousState);
      renderLoginState();
      setStatus(normalizeAuthMessage(error.message || "Could not return to creator account."), true);
    }
  }
}

async function clearLocalWorkspace() {
  selectedFiles = [];
  currentSnapshot = null;
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = "";
  resetDerivedState();
  renderFilesList();
}

function applyShiftColours(sourceSettings = settings) {
  const mappings = {
    day: "shiftColorDay",
    evening: "shiftColorEvening",
    night: "shiftColorNight",
    cs: "shiftColorCs",
    leave: "shiftColorLeave",
    custom: "shiftColorCustom",
    phnw: "shiftColorPhnw",
  };
  for (const [tone, field] of Object.entries(mappings)) {
    const colour = isHexColour(sourceSettings[field]) ? sourceSettings[field] : defaultShiftColourForField(field);
    const rgb = hexToRgb(colour);
    document.documentElement.style.setProperty(`--chip-${tone}-text`, colour);
    document.documentElement.style.setProperty(`--chip-${tone}-bg-strong`, `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.24)`);
    document.documentElement.style.setProperty(`--chip-${tone}-bg-soft`, `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.12)`);
  }
}

function applyCurrentDayHighlight(sourceSettings = settings, options = {}) {
  const borderColour = isHexColour(sourceSettings.currentDayBorderColor) ? sourceSettings.currentDayBorderColor : "#c44949";
  const backgroundColour = isHexColour(sourceSettings.currentDayBackgroundColor) ? sourceSettings.currentDayBackgroundColor : borderColour;
  const borderOpacity = normalizeOpacity(sourceSettings.currentDayBorderOpacity, 20);
  const backgroundOpacity = normalizeOpacity(sourceSettings.currentDayBackgroundOpacity, 20);
  const fillStyle = sourceSettings.currentDayFillStyle === "solid" ? "solid" : "gradient";
  const borderWidth = Math.max(1, Math.min(4, Number(sourceSettings.currentDayBorderWidth || 2)));
  const direction = String(sourceSettings.currentDayGradientDirection || "90deg");
  const borderRgb = hexToRgb(borderColour);
  const backgroundRgb = hexToRgb(backgroundColour);
  const fillColour = `rgba(${backgroundRgb.r}, ${backgroundRgb.g}, ${backgroundRgb.b}, ${backgroundOpacity})`;
  const fadeColour = `rgba(${backgroundRgb.r}, ${backgroundRgb.g}, ${backgroundRgb.b}, 0.04)`;
  const fillSurface = fillStyle === "solid"
    ? fillColour
    : direction === "radial"
      ? `radial-gradient(circle, ${fillColour}, ${fadeColour} 72%, rgba(255, 255, 255, 0.82))`
      : `linear-gradient(${direction || "90deg"}, ${fillColour}, ${fadeColour} 76%, rgba(255, 255, 255, 0.82))`;
  if (!options.previewOnly) {
    document.documentElement.style.setProperty("--today-border-color", `rgba(${borderRgb.r}, ${borderRgb.g}, ${borderRgb.b}, ${borderOpacity})`);
    document.documentElement.style.setProperty("--today-border-width", `${borderWidth}px`);
    document.documentElement.style.setProperty("--today-fill-color", fillColour);
    document.documentElement.style.setProperty("--today-fill-surface", fillSurface);
  }
  if (currentDayPreview) {
    currentDayPreview.style.borderColor = `rgba(${borderRgb.r}, ${borderRgb.g}, ${borderRgb.b}, ${borderOpacity})`;
    currentDayPreview.style.borderWidth = `${borderWidth}px`;
    currentDayPreview.style.background = fillSurface;
  }
}

function applyWeekendShade(sourceSettings = settings, options = {}) {
  const previewWeekendDays = document.querySelectorAll(".settings-calendar-day.is-weekend");
  if (sourceSettings.weekendShadeEnabled === false) {
    if (!options.previewOnly) {
      document.documentElement.style.setProperty("--weekend-shade-surface", "transparent");
    }
    previewWeekendDays.forEach((day) => {
      day.style.background = "transparent";
    });
    return;
  }
  const colour = isHexColour(sourceSettings.weekendShadeColor) ? sourceSettings.weekendShadeColor : "#e5e7eb";
  const opacity = normalizeOpacity(sourceSettings.weekendShadeOpacity, 30);
  const rgb = hexToRgb(colour);
  const shade = `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${opacity})`;
  if (!options.previewOnly) {
    document.documentElement.style.setProperty("--weekend-shade-surface", shade);
  }
  previewWeekendDays.forEach((day) => {
    day.style.background = shade;
  });
}

function syncPreviewStyleControls(sourceSettings = settings) {
  for (const output of document.querySelectorAll("[data-opacity-output]")) {
    const field = output.dataset.opacityOutput;
    output.textContent = `${Math.round(Number(sourceSettings[field] || 0))}%`;
  }
  const fillMode = sourceSettings.currentDayFillStyle === "gradient" ? "gradient" : "solid";
  settingsPanel?.querySelectorAll("[data-fill-mode]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.fillMode === fillMode);
  });
  settingsPanel?.querySelectorAll("[data-gradient-direction]").forEach((button) => {
    const active = fillMode === "gradient" && button.dataset.gradientDirection === sourceSettings.currentDayGradientDirection;
    button.classList.toggle("is-active", active);
    button.disabled = false;
  });
  settingsPanel?.querySelector(".direction-grid")?.classList.toggle("is-muted", fillMode !== "gradient");
  const weekendDisabled = sourceSettings.weekendShadeEnabled === false;
  settingsPanel?.querySelector(".weekend-card")?.classList.toggle("is-disabled", weekendDisabled);
  ["weekendShadeColor"].forEach((field) => {
    if (settingsInputs[field]) settingsInputs[field].disabled = weekendDisabled;
  });
  settingsPanel?.querySelectorAll('[data-opacity-stepper="weekendShadeOpacity"] button').forEach((button) => {
    button.disabled = weekendDisabled;
  });
}

function updatePreviewDisplayExample(sourceSettings = settings) {
  const chip = document.querySelector(".settings-example-chip");
  const title = document.querySelector("#settingsExampleTitle");
  const raw = document.querySelector("#settingsExampleRaw");
  const time = document.querySelector("#settingsExampleTime");
  if (!title || !raw || !time) return;
  const event = settingsExampleEvent();
  if (chip) {
    chip.classList.remove("preview-chip-day", "preview-chip-evening", "preview-chip-night", "preview-chip-cs", "preview-chip-leave", "preview-chip-custom", "preview-chip-phnw");
    chip.classList.add(`preview-chip-${event.tone}`);
  }
  title.textContent = event.title;
  raw.textContent = event.rawValue;
  time.textContent = event.timeLabel;
  title.classList.toggle("hidden", !sourceSettings.showNormalizedTitles);
  raw.classList.toggle("hidden", !sourceSettings.showRawValues);
  time.classList.toggle("hidden", !sourceSettings.showTimes || !event.timeLabel);
}

function settingsExampleEvent() {
  const events = currentPreviewEvents.size
    ? [...currentPreviewEvents.values()]
    : latestPreview
      ? buildFilteredPreviewEvents(latestPreview, settings)
      : [];
  const usable = events.filter((event) => {
    const raw = String(event?.rawValue || "").trim();
    return raw && raw !== "Custom event" && String(event?.title || "").trim();
  });
  const preferred = usable.find((event) => !event.allDay && event.timeLabel)
    || usable.find((event) => event.timeLabel)
    || usable[0];
  return {
    title: String(preferred?.title || "MMC: SSU PM").trim(),
    rawValue: String(preferred?.rawValue || "1430-0000 PSSC").trim(),
    timeLabel: String(preferred?.timeLabel || "2:30 pm - 12:00 am").trim(),
    tone: preferred ? eventTone(preferred) : "day",
  };
}

function defaultColourForField(field) {
  if (field === "weekendShadeColor") return "#e5e7eb";
  return defaultShiftColourForField(field);
}

function defaultShiftColourForField(field) {
  const key = field.replace(/^shiftColor/, "").toLowerCase();
  return SHIFT_COLOUR_DEFAULTS[key] || SHIFT_COLOUR_DEFAULTS.day;
}

function isHexColour(value) {
  return /^#[0-9a-f]{6}$/i.test(String(value || ""));
}

function hexToRgb(hex) {
  const clean = String(hex || "#000000").replace("#", "");
  return {
    r: Number.parseInt(clean.slice(0, 2), 16),
    g: Number.parseInt(clean.slice(2, 4), 16),
    b: Number.parseInt(clean.slice(4, 6), 16),
  };
}

function normalizeOpacity(value, fallbackPercent) {
  const numeric = Number.parseFloat(String(value ?? fallbackPercent));
  if (!Number.isFinite(numeric)) return fallbackPercent / 100;
  return Math.max(0, Math.min(100, numeric)) / 100;
}

function forceConsoleSkin() {
  document.body.dataset.skin = "console";
}

function hideLoadingScreen() {
  loadingScreen?.classList.add("hidden");
}

function setEntranceStatus(message, isError = false) {
  entranceStatus.textContent = message;
  entranceStatus.dataset.error = isError ? "true" : "false";
}

function loadCurrentUserEmail() {
  return normalizeEmail(localStorage.getItem(CURRENT_EMAIL_KEY));
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function sanitizeSubscription(value) {
  if (!value || typeof value !== "object") return null;
  const token = String(value.token || "").trim();
  if (!token) return null;
  return {
    token,
    enabled: value.enabled === true,
  };
}

function applyIssueConfig(value) {
  const config = value && typeof value === "object" ? value : {};
  parserExtensions = sanitizeParserExtensions(config.parserExtensions);
  globalParserExtensions = sanitizeParserExtensions(config.globalParserExtensions || config.parserExtensions);
  localParserExtensions = sanitizeParserExtensions(config.localParserExtensions);
  parserRuleSuggestions = sanitizeParserRuleSuggestions(config.parserRuleSuggestions);
  dismissedIssueFingerprints = new Set(sanitizeIssueFingerprintList(config.dismissedFingerprints));
  ignoredIssueFingerprints = new Set(sanitizeIssueFingerprintList(config.ignoredFingerprints));
  setParserExtensions(parserExtensions);
}

function sanitizeParserExtensions(value) {
  const input = value && typeof value === "object" ? value : {};
  const defaults = parserRuleDefaults();
  const removed = sanitizeParserRuleRemovals(input._removed);
  return {
    mmc: applyParserRuleRemovals(mergeParserRuleLists(defaults.mmc, sanitizeParserExtensionRuleList(input.mmc, "MMC")), removed),
    ddh: applyParserRuleRemovals(mergeParserRuleLists(defaults.ddh, sanitizeParserExtensionRuleList(input.ddh, "DDH")), removed),
    casey: applyParserRuleRemovals(mergeParserRuleLists(defaults.casey, sanitizeParserExtensionRuleList(input.casey, "Casey")), removed),
    mch: applyParserRuleRemovals(mergeParserRuleLists(defaults.mch, sanitizeParserExtensionRuleList(input.mch, "MCH")), removed),
    _removed: removed,
  };
}

function sanitizeParserRuleRemovals(items) {
  if (!Array.isArray(items)) return [];
  const byKey = new Map();
  for (const item of items) {
    const source = sanitizeIssueSource(item?.source);
    const seniority = sanitizeRuleSeniority(item?.seniority);
    const code = String(item?.code || "").trim().toUpperCase();
    if (!source || !code) continue;
    byKey.set(`${source}|${seniority}|${code}`, { source, seniority, code });
  }
  return [...byKey.values()].sort(compareParserRules);
}

function applyParserRuleRemovals(rules, removals) {
  const removedKeys = new Set((removals || []).map(parserRuleStorageKey));
  return (rules || []).filter((rule) => !removedKeys.has(parserRuleStorageKey(rule)));
}

function mergeParserRuleLists(defaults, overrides) {
  const byKey = new Map();
  for (const rule of defaults) byKey.set(parserRuleStorageKey(rule), rule);
  for (const rule of overrides) byKey.set(parserRuleStorageKey(rule), rule);
  return [...byKey.values()].sort(compareParserRules);
}

function parserRuleStorageKey(rule) {
  return `${rule.source}|${rule.seniority}|${rule.code}`;
}

function compareParserRules(left, right) {
  const leftRank = parserRuleSeniorities().indexOf(left.seniority);
  const rightRank = parserRuleSeniorities().indexOf(right.seniority);
  const rankDelta = (leftRank < 0 ? 99 : leftRank) - (rightRank < 0 ? 99 : rightRank);
  if (rankDelta) return rankDelta;
  return left.code.localeCompare(right.code);
}

function sanitizeParserExtensionRuleList(items, source) {
  if (!Array.isArray(items)) return [];
  return items
    .map((item) => sanitizeParserExtensionRule(item, source))
    .filter((item) => !isObsoleteSeededParserRule(item))
    .filter(Boolean);
}

function sanitizeParserExtensionRule(item, forcedSource = "") {
  if (!item || typeof item !== "object") return null;
  const source = sanitizeIssueSource(forcedSource || item.source);
  const seniority = sanitizeRuleSeniority(item.seniority);
  const code = normalizeParserExtensionRuleCode(source, item.code || "");
  const ignore = item.ignore === true || String(item.kind || "").trim().toLowerCase() === "ignore";
  const base = String(item.base || "").trim();
  const period = String(item.period || "").trim().toUpperCase();
  const suffix = String(item.suffix || "").trim();
  const allDay = item.allDay === true;
  const startTime = String(item.startTime || "").trim();
  const endTime = String(item.endTime || "").trim();
  const location = String(item.location || "").trim();
  if (!source || !code || (!ignore && !base)) return null;
  if (isRestrictedClinicalSupportRule({ seniority, code, base })) return null;
  if (!ignore && !allDay && (!isClockString(startTime) || !isClockString(endTime))) return null;
  return {
    source,
    seniority,
    code,
    kind: ignore ? "ignore" : String(item.kind || "shift").trim().toLowerCase(),
    base,
    period,
    suffix,
    allDay: ignore ? true : allDay,
    startTime: ignore || allDay ? "" : startTime,
    endTime: ignore || allDay ? "" : endTime,
    location,
    includeAsShift: ignore ? false : item.includeAsShift !== false,
    ignore,
  };
}

function isRestrictedClinicalSupportRule(rule) {
  const seniority = sanitizeRuleSeniority(rule?.seniority);
  if (seniority === "SMS" || seniority === "CMO") return false;
  const code = String(rule?.code || "").trim().toUpperCase();
  const base = String(rule?.base || "").trim().toUpperCase();
  return code === "CS"
    || code === "CSO"
    || code === "CS ONSITE"
    || code === "CLIN SUPP"
    || code === "CLINICAL SUPP"
    || base === "CS"
    || base === "CSO"
    || base === "CS ONSITE";
}

function normalizeParserExtensionRuleCode(source, value) {
  const text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (source === "MMC" || source === "Casey" || source === "MCH") {
    return parserRuleCodeFromRawValue(source, text);
  }
  return text;
}

function sanitizeParserRuleSuggestions(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => sanitizeParserRuleSuggestion(item))
    .filter(Boolean)
    .sort((left, right) => (right.updatedAt || "").localeCompare(left.updatedAt || ""));
}

function sanitizeParserRuleSuggestion(item) {
  const rule = sanitizeParserExtensionRule(item?.rule);
  const email = normalizeEmail(item?.email);
  if (!rule || !email) return null;
  return {
    id: String(item.id || `${email}::${rule.source}|${rule.seniority}|${rule.code}`).trim(),
    email,
    realName: String(item.realName || "").trim(),
    fingerprint: sanitizeIssueFingerprint(item.fingerprint || issueFingerprint(rule.source, rule.code, rule.seniority)),
    rawValue: String(item.rawValue || rule.code || "").trim(),
    rule,
    status: "pending",
    createdAt: String(item.createdAt || ""),
    updatedAt: String(item.updatedAt || ""),
  };
}

function sanitizeRuleSeniority(value) {
  const text = String(value || "").trim();
  const upper = text.toUpperCase();
  const aliases = new Map([
    ["SR", "Senior Registrar"],
    ["SENIOR REG", "Senior Registrar"],
    ["SENIOR REGISTRAR", "Senior Registrar"],
    ["IR", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REG", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REGISTRAR", "Transitional/Intermediate Registrar"],
    ["TRANSITIONAL REGISTRAR", "Transitional/Intermediate Registrar"],
    ["JR", "Junior Registrar"],
    ["JUNIOR REG", "Junior Registrar"],
    ["JUNIOR REGISTRAR", "Junior Registrar"],
    ["I", "Intern"],
    ["INTERN", "Intern"],
  ]);
  if (aliases.has(upper)) return aliases.get(upper);
  return parserRuleSeniorities().find((item) => item.toUpperCase() === upper) || "Unknown";
}

function isObsoleteSeededParserRule(rule) {
  if (!rule || rule.source !== "MMC") return false;
  if (rule.code === "CS" || rule.code === "CSO") return false;
  const impossibleFloat = new Set(["ACR", "PCR", "ARR", "PRR", "ASSR", "PSSR"]);
  if (impossibleFloat.has(rule.code) && isOldDefaultMmcRule(rule)) return true;
  return false;
}

function isConsultantStyleMmcCode(code) {
  const text = String(code || "").trim().toUpperCase();
  return /^[AP][GARC][CR]$/.test(text) || /^[AP]SS[CR]$/.test(text);
}

function isOldDefaultMmcRule(rule) {
  const expected = oldDefaultMmcRuleShape(rule.code);
  if (!expected) return false;
  return rule.base === expected.base
    && rule.period === expected.period
    && rule.suffix === expected.suffix
    && rule.allDay === false
    && rule.startTime === expected.startTime
    && rule.endTime === expected.endTime
    && rule.includeAsShift !== false;
}

function oldDefaultMmcRuleShape(code) {
  const text = String(code || "").trim().toUpperCase();
  const teamMap = { G: "Green", A: "Amber", R: "Resus", C: "Clinic" };
  const teamMatch = text.match(/^([AP])([GARC])([CR])$/);
  if (teamMatch) {
    return {
      base: teamMap[teamMatch[2]],
      period: teamMatch[1] === "A" ? "AM" : "PM",
      suffix: teamMatch[3] === "R" ? "Float" : "",
      startTime: teamMatch[1] === "A" ? "08:00" : "14:30",
      endTime: teamMatch[1] === "A" ? "17:30" : "00:00",
    };
  }
  const ssuMatch = text.match(/^([AP])SS([CR])$/);
  if (!ssuMatch) return null;
  return {
    base: "SSU",
    period: ssuMatch[1] === "A" ? "AM" : "PM",
    suffix: ssuMatch[2] === "R" ? "Float" : "",
    startTime: ssuMatch[1] === "A" ? "07:30" : "14:30",
    endTime: ssuMatch[1] === "A" ? "17:30" : "00:00",
  };
}

function sanitizeIssueFingerprintList(items) {
  if (!Array.isArray(items)) return [];
  return [...new Set(items.map((item) => sanitizeIssueFingerprint(item)).filter(Boolean))];
}

function sanitizeIssueFingerprint(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const [source, ...rest] = text.split("::");
  if (rest.length >= 2) {
    const seniority = sanitizeRuleSeniority(rest[0]);
    const rawValue = rest.slice(1).join("::");
    return issueFingerprint(source, rawValue, seniority);
  }
  return issueFingerprint(source, rest.join("::"));
}

function sanitizeIssueSource(value) {
  const source = String(value || "").trim().toUpperCase();
  if (source === "MMC" || source === "DDH") return source;
  if (source === "CASEY") return "Casey";
  if (source === "MCH") return "MCH";
  return "";
}

function issueFingerprint(source, rawValue, seniority = "") {
  const normalizedSource = sanitizeIssueSource(source);
  const normalizedSeniority = seniority ? sanitizeRuleSeniority(seniority) : "";
  const normalizedRaw = String(rawValue || "").trim();
  return normalizedSource && normalizedRaw ? `${normalizedSource}::${normalizedSeniority ? `${normalizedSeniority}::` : ""}${normalizedRaw}` : "";
}

function sanitizeWorkspaceSnapshot(value) {
  if (!value || typeof value !== "object" || !value.preview) return null;
  const snapshot = {
    preview: JSON.parse(JSON.stringify(value.preview)),
    session: value.session && typeof value.session === "object" ? JSON.parse(JSON.stringify(value.session)) : {},
    doctorOptions: Array.isArray(value.doctorOptions)
      ? value.doctorOptions.map((doctor) => ({
          ...doctor,
          key: normalizeRosterName(doctor?.key || ""),
          displayName: String(doctor?.displayName || "").trim(),
          sourceTypes: Array.isArray(doctor?.sourceTypes) ? doctor.sourceTypes.map((item) => String(item || "").toLowerCase()).filter(Boolean) : [],
          aliases: Array.isArray(doctor?.aliases) ? doctor.aliases.map((alias) => ({
            key: normalizeRosterName(alias?.key || ""),
            displayName: String(alias?.displayName || "").trim(),
            sourceType: String(alias?.sourceType || "").toLowerCase(),
          })).filter((alias) => alias.key && alias.displayName) : [],
          accountEmail: normalizeEmail(doctor?.accountEmail || doctor?.claimedBy || ""),
          claimedBy: normalizeEmail(doctor?.claimedBy || ""),
          claimedByName: String(doctor?.claimedByName || "").trim(),
        })).filter((doctor) => doctor.key && doctor.displayName)
      : [],
    detectedSources: {
      mmc: Array.isArray(value.detectedSources?.mmc) ? value.detectedSources.mmc.map((item) => String(item || "")).filter(Boolean) : [],
      ddh: Array.isArray(value.detectedSources?.ddh) ? value.detectedSources.ddh.map((item) => String(item || "")).filter(Boolean) : [],
      casey: Array.isArray(value.detectedSources?.casey) ? value.detectedSources.casey.map((item) => String(item || "")).filter(Boolean) : [],
      mch: Array.isArray(value.detectedSources?.mch) ? value.detectedSources.mch.map((item) => String(item || "")).filter(Boolean) : [],
    },
    fileRefs: Array.isArray(value.fileRefs) ? value.fileRefs.map(importRefForWorkspace).filter((item) => item.id) : [],
    subscriptionFeeds: value.subscriptionFeeds && typeof value.subscriptionFeeds === "object" ? JSON.parse(JSON.stringify(value.subscriptionFeeds)) : {},
    insightCache: sanitizeInsightCache(value.insightCache),
    profileCoverage: value.profileCoverage && typeof value.profileCoverage === "object" ? JSON.parse(JSON.stringify(value.profileCoverage)) : null,
  };
  snapshot.calendarRevision = String(value.calendarRevision || "");
  snapshot.cacheKey = String(value.cacheKey || "");
  snapshot.cachedAt = String(value.cachedAt || "");
  return snapshot;
}

function sanitizeInsightCache(value) {
  if (!value || typeof value !== "object") return null;
  const doctorEvents = {};
  for (const [key, events] of Object.entries(value.doctorEvents || {})) {
    const normalizedKey = normalizeRosterName(key);
    if (!normalizedKey || !Array.isArray(events)) continue;
    doctorEvents[normalizedKey] = events
      .filter((event) => event && typeof event === "object")
      .map((event) => JSON.parse(JSON.stringify(event)));
  }
  const doctorOptions = Array.isArray(value.doctorOptions)
    ? value.doctorOptions.map((doctor) => ({
        ...doctor,
        key: normalizeRosterName(doctor?.key || ""),
        displayName: String(doctor?.displayName || "").trim(),
        sourceTypes: Array.isArray(doctor?.sourceTypes) ? doctor.sourceTypes.map((item) => String(item || "").toLowerCase()).filter(Boolean) : [],
        aliases: Array.isArray(doctor?.aliases) ? doctor.aliases.map((alias) => ({
          key: normalizeRosterName(alias?.key || ""),
          displayName: String(alias?.displayName || "").trim(),
          sourceType: String(alias?.sourceType || "").toLowerCase(),
        })).filter((alias) => alias.key && alias.displayName) : [],
      })).filter((doctor) => doctor.key && doctor.displayName)
    : [];
  const doctorRoles = {};
  for (const [key, roles] of Object.entries(value.doctorRoles || {})) {
    const normalizedKey = normalizeRosterName(key);
    if (normalizedKey && roles && typeof roles === "object") {
      doctorRoles[normalizedKey] = JSON.parse(JSON.stringify(roles));
    }
  }
  if (!String(value.key || "") || !Object.keys(doctorEvents).length || !doctorOptions.length) return null;
  return {
    key: String(value.key),
    builtAt: String(value.builtAt || ""),
    doctorOptions,
    doctorEvents,
    doctorRoles,
  };
}

function importRefForWorkspace(entry) {
  return {
    id: String(entry?.id || entry?.repoId || ""),
    name: String(entry?.name || "roster.xlsx"),
    size: Number(entry?.size || 0),
    lastModified: Number(entry?.lastModified || 0),
    addedAt: String(entry?.addedAt || ""),
    sourceType: String(entry?.sourceType || "pending"),
  };
}

function importRefsToClientEntries(refs = []) {
  return (Array.isArray(refs) ? refs : [])
    .map((entry) => ({
      ...importRefForWorkspace(entry),
      repoId: String(entry?.repoId || entry?.id || ""),
    }))
    .filter((entry) => entry.id);
}

function normalizeSavedExportRange(value) {
  const normalized = normalizeExportRangeState(value);
  return {
    startDate: normalized.startDate,
    endDate: normalized.endDate,
    allFuture: normalized.allFuture !== false,
  };
}

function normalizeServerUser(value) {
  if (typeof value === "string") {
    return {
      email: value,
      realName: "",
      role: value === OWNER_EMAIL ? "owner" : "user",
      sites: [],
      claims: [],
      insightsEnabled: false,
      adminIssues: [],
      issuesCount: 0,
    };
  }
  const email = normalizeEmail(value?.email);
  const role = value?.role || (email === OWNER_EMAIL ? "owner" : "user");
  return {
    email,
    realName: String(value?.realName || "").trim(),
    role: role === "creator" ? "owner" : role,
    sites: Array.isArray(value?.sites) ? value.sites : [],
    seniorities: Array.isArray(value?.seniorities) ? value.seniorities.map((item) => String(item || "").trim()).filter(Boolean) : [],
    claims: sanitizeRosterClaims(value?.claims || []),
    insightsEnabled: role === "owner" || role === "creator" || value?.insightsEnabled === true,
    adminIssues: Array.isArray(value?.adminIssues) ? value.adminIssues : [],
    issuesCount: Number(value?.issuesCount || 0),
  };
}

function normalizeAuthMessage(message) {
  const normalized = String(message || "").trim();
  if (!normalized) return "Incorrect username or password.";
  if (
    normalized === "Incorrect password."
    || normalized === "Account not found."
    || normalized === "Account not found. Create an account first."
  ) {
    return "Incorrect username or password.";
  }
  return normalized;
}

function openLoginModal(prefillEmail = currentUserEmail || "") {
  setEntranceTab("login");
  forceConsoleSkin();
  loginEmail.value = prefillEmail;
  loginPassword.value = "";
  if (stayLoggedIn) stayLoggedIn.checked = Boolean(localStorage.getItem(PERSISTENT_PASSWORD_KEY));
  entrancePage.classList.remove("hidden");
  appShell.classList.add("hidden");
  mobileActionBar.classList.add("hidden");
  setTimeout(() => loginEmail.focus(), 0);
}

function closeLoginModal() {
  entrancePage.classList.add("hidden");
  appShell.classList.remove("hidden");
}

async function logoutCurrentUser() {
  try {
    await flushCloudStateSave();
  } catch {
    // Keep logout moving even if cloud persistence fails.
  }
  cancelScheduledCloudStateSave();
  localStorage.removeItem(CURRENT_EMAIL_KEY);
  sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
  localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
  currentUserEmail = "";
  currentUserPassword = "";
  authUserEmail = "";
  authUserPassword = "";
  adminViewingEmail = "";
  currentUserRole = "user";
  cloudAvailable = false;
  setActiveCalendarContext("claimed-account", { email: "" });
  currentRosterClaims = [];
  latestNameMatches = [];
  availableRosterDoctors = [];
  currentSubscription = null;
  currentInsightsEnabled = false;
  currentSuggestedClaims = [];
  selectedFiles = [];
  resetDerivedState();
  renderLoginState();
  openLoginModal();
  setStatus("Log in to load a roster workspace.");
}

function renderLoginState() {
  const loggedIn = Boolean(currentUserEmail && currentUserPassword);
  loginBar.classList.toggle("hidden", !loggedIn);
  appShell.classList.toggle("hidden", !loggedIn);
  entrancePage.classList.toggle("hidden", loggedIn);
  if (!loggedIn) mobileActionBar.classList.add("hidden");
  const me = currentAccount();
  const displayName = me.realName ? `${me.realName} · ` : "";
  const viewingText = activeDoctorProfile
    ? `Viewing as ${activeDoctorProfile.displayName} · doctor profile`
    : adminViewingEmail
      ? `Viewing as ${displayName}${currentUserEmail}`
      : `${displayName}${currentUserEmail}`;
  loginIdentity.textContent = loggedIn
    ? `${viewingText} · ${currentUserRole === "creator" ? "Creator" : "Standard account"}${cloudAvailable ? " · Cloud sync on" : " · Cloud sync required"}`
    : "";
  backToCreatorButton.classList.toggle("hidden", !canReturnToCreator());
  syncAccountsButton();
  syncActionState();
  syncMobileChrome();
}

async function loginWithEmail(email, password, options = {}) {
  const previousEmail = currentUserEmail;
  const loginStartedAt = performance.now();
  try {
    await flushCloudStateSave().catch(() => {});
    cancelScheduledCloudStateSave();
    ensureLocalAccountLogin(email, password, options);
    currentUserEmail = normalizeEmail(email);
    currentUserPassword = password;
    authUserEmail = currentUserEmail;
    authUserPassword = currentUserPassword;
    adminViewingEmail = "";
    currentUserRole = currentUserEmail === OWNER_EMAIL ? "creator" : "user";
    setActiveCalendarContext(currentUserRole === "creator" ? "creator-account" : "claimed-account", { email: currentUserEmail });
    localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
    sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
    if (options.stayLoggedIn) {
      localStorage.setItem(PERSISTENT_PASSWORD_KEY, currentUserPassword);
    } else if (!options.adminTargetEmail) {
      localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
    }
    forceConsoleSkin();
    setStatus("Loading account workspace...");
    setEntranceStatus("Loading account workspace...");
    if (previousEmail !== currentUserEmail) {
      await clearLocalWorkspace();
    }
    await restoreCloudState({ ...options, deferHydration: true, loginStartedAt });
    if (!currentUserEmail) return;
    renderLoginState();
    closeLoginModal();
    setEntranceStatus("");
    markLoginPhase("shellRendered", loginStartedAt);
    const renderedCachedSnapshot = renderCachedCalendarSnapshot({ loginStartedAt });
    setStatus(renderedCachedSnapshot ? "Checking calendar for updates..." : "Loading calendar...");
    void hydrateAuthenticatedWorkspace({ ...options, includeBootstrap: true }, loginStartedAt);
  } catch (error) {
    const message = normalizeAuthMessage(error.message || "Login failed.");
    setEntranceStatus(message, true);
    setStatus(message, true);
  }
}

async function restoreCloudState(options = {}) {
  if (!currentUserEmail) return;
  try {
    const adminTargetEmail = normalizeEmail(options.adminTargetEmail);
    setActiveCalendarContext(adminTargetEmail
      ? "claimed-account"
      : currentUserEmail === OWNER_EMAIL
        ? "creator-account"
        : "claimed-account", { email: adminTargetEmail || currentUserEmail });
    const requestEmail = adminTargetEmail ? authUserEmail : currentUserEmail;
    const requestPassword = adminTargetEmail ? authUserPassword : currentUserPassword;
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: adminTargetEmail ? "adminLoadUser" : "login",
        email: requestEmail,
        password: requestPassword,
        targetEmail: adminTargetEmail,
        mode: options.mode || "login",
        realName: options.realName || "",
      }),
    });
    const data = await readJsonResponse(response, "Login failed.");
    await applyCloudStateData(data);
    markLoginPhase("authenticated", options.loginStartedAt);
    markAccountSwitchPhase("adminLoadUser", options.accountSwitchStartedAt);
    if (!options.deferHydration) await hydrateAuthenticatedWorkspace({ ...options, includeBootstrap: false }, options.loginStartedAt);
  } catch (error) {
    cancelScheduledCloudStateSave();
    const attemptedEmail = currentUserEmail;
    const message = error.message === "Cloud storage is not configured."
      ? serverStorageRequiredMessage()
      : normalizeAuthMessage(error.message || "Login failed.");
    if (options.preserveSessionOnFailure) {
      setStatus(message, true);
      throw new Error(message);
    }
    cloudAvailable = false;
    currentSubscription = null;
    currentInsightsEnabled = false;
    localStorage.removeItem(CURRENT_EMAIL_KEY);
    sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
    localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
    if (options.mode === "create") {
      accountState.users = accountState.users.filter((user) => user.email !== currentUserEmail);
      saveAccountState();
    }
    currentUserEmail = "";
    currentUserPassword = "";
    renderLoginState();
    openLoginModal(attemptedEmail);
    setStatus(message, true);
    setEntranceStatus(message, true);
  }
}

async function hydrateAuthenticatedWorkspace(options = {}, loginStartedAt = 0) {
  if (!currentUserEmail) return;
  try {
    const adminTargetEmail = normalizeEmail(options.adminTargetEmail);
    if (adminTargetEmail && adminTargetEmail !== OWNER_EMAIL && !currentRosterClaims.length) {
      await resolveCurrentAccountClaims(adminTargetEmail);
      markLoginPhase("claimsResolved", loginStartedAt);
      markAccountSwitchPhase("claimsResolved", options.accountSwitchStartedAt);
    } else if (!adminTargetEmail && currentUserEmail !== OWNER_EMAIL && !currentRosterClaims.length) {
      await resolveCurrentAccountClaims();
      markLoginPhase("claimsResolved", loginStartedAt);
    }
    if (!adminTargetEmail && currentUserEmail === OWNER_EMAIL) {
      forceCreatorDoctorSession();
    }
    const cachedRevision = options.cachedRevision || (currentSnapshot?.cacheKey === currentCalendarSnapshotCacheKey()
      ? currentSnapshot.calendarRevision || currentCalendarRevision || ""
      : "");
    const loadedFreshCalendar = await loadCloudCalendarEvents({ adminTargetEmail, cachedRevision });
    markLoginPhase("calendarLoaded", loginStartedAt);
    markAccountSwitchPhase("calendarLoaded", options.accountSwitchStartedAt);
    if (options.includeBootstrap !== false) {
      if (loadedFreshCalendar || currentSnapshot?.preview) {
        await bootstrapImports();
      } else {
        await bootstrapImports();
      }
      markLoginPhase("workspaceRendered", loginStartedAt);
      markAccountSwitchPhase("workspaceRendered", options.accountSwitchStartedAt);
    }
    if (latestNameMatches.length) {
      const sites = [...new Set(latestNameMatches.map((claim) => claim.sourceType.toUpperCase()))].join(", ");
      setStatus(`Suggested roster name${latestNameMatches.length === 1 ? "" : "s"} for ${sites || "uploaded rosters"}. Please confirm in Account.`);
    }
    if (isCreatorAuthenticated()) {
      void loadServerUsers();
    }
  } catch (error) {
    const message = normalizeAuthMessage(error.message || "Workspace hydration failed.");
    setStatus(message, true);
    console.warn("Post-login workspace hydration failed", { message, email: currentUserEmail, error, timings: lastLoginTimings });
  }
}

function markLoginPhase(phase, loginStartedAt = 0) {
  if (!loginStartedAt) return;
  if (!lastLoginTimings || lastLoginTimings.startedAt !== loginStartedAt) {
    lastLoginTimings = { startedAt: loginStartedAt };
  }
  lastLoginTimings[phase] = Math.round(performance.now() - loginStartedAt);
  window.__rosterLoginTimings = { ...lastLoginTimings };
  if (phase === "workspaceRendered") {
    console.info("Login timings", window.__rosterLoginTimings);
  }
}

function markAccountSwitchPhase(phase, accountSwitchStartedAt = 0) {
  if (!accountSwitchStartedAt) return;
  if (!lastAccountSwitchTimings || lastAccountSwitchTimings.startedAt !== accountSwitchStartedAt) {
    lastAccountSwitchTimings = { startedAt: accountSwitchStartedAt };
  }
  lastAccountSwitchTimings[phase] = Math.round(performance.now() - accountSwitchStartedAt);
  window.__rosterAccountSwitchTimings = { ...lastAccountSwitchTimings };
  if (phase === "workspaceRendered") {
    console.info("Account switch timings", window.__rosterAccountSwitchTimings);
  }
}

async function resolveCurrentAccountClaims(targetEmailOverride = "") {
  const targetEmail = normalizeEmail(targetEmailOverride);
  if (!currentUserEmail || (!targetEmail && currentUserEmail === OWNER_EMAIL)) return;
  const requestEmail = targetEmail ? authUserEmail : currentUserEmail;
  const requestPassword = targetEmail ? authUserPassword : currentUserPassword;
  if (!requestEmail || !requestPassword) return;
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "resolveAccountClaims",
      email: requestEmail,
      password: requestPassword,
      targetEmail,
    }),
  });
  const data = await readJsonResponse(response, "Account claim resolution failed.");
  await applyCloudStateData(data);
}

async function restoreDoctorProfileState() {
  if (!activeDoctorProfile || !cloudAvailable) {
    currentSnapshot = null;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = "";
    restoredSessionState = loadCurrentSessionState();
    return;
  }
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "loadDoctorProfile",
      email: authUserEmail || currentUserEmail,
      password: authUserPassword || currentUserPassword,
      profileId: activeDoctorProfile.id,
      doctorKey: activeDoctorProfile.doctorKey,
      displayName: activeDoctorProfile.displayName,
      sourceTypes: activeDoctorProfile.sourceTypes,
    }),
  });
  const data = await readJsonResponse(response, "Doctor profile load failed.");
  applyIssueConfig(data.issueConfig);
  currentCalendarRevision = String(data.calendarRevision || currentCalendarRevision || "");
  currentSnapshot = sanitizeWorkspaceSnapshot(data.snapshot);
  if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
  currentSnapshotStale = data.snapshotStale === true;
  currentSnapshotBuiltAt = String(data.snapshotBuiltAt || "");
  selectedFiles = importRefsToClientEntries(currentSnapshot?.fileRefs || []);
  restoredSessionState = currentSnapshot?.session || data.profile?.state?.session || loadCurrentSessionState() || null;
  if (!selectedFiles.length) {
    await loadDoctorProfileImportsIntoWorkspace();
  }
  saveWorkspaceSnapshotForEmail(activeWorkspaceOwnerKey(), {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: restoredSessionState || {},
    snapshot: currentSnapshot,
  });
}

async function loadDoctorProfileImportsIntoWorkspace() {
  selectedFiles = [];
}

async function applyCloudStateData(data) {
  cloudAvailable = data.cloudAvailable === true;
  currentCalendarRevision = String(data.calendarRevision || currentCalendarRevision || "");
  currentUserRole = data.role || currentUserRole;
  currentInsightsEnabled = currentUserRole === "creator" || data.insightsEnabled === true;
  currentRosterClaims = sanitizeRosterClaims(data.claims || []);
  currentSuggestedClaims = sanitizeRosterClaims(data.suggestedClaims || data.nameMatches || []);
  latestNameMatches = currentSuggestedClaims;
  const incomingAvailableDoctors = sanitizeAvailableRosterDoctors(data.availableDoctors || []);
  if (incomingAvailableDoctors.length || !isCreatorAuthenticated()) {
    availableRosterDoctors = incomingAvailableDoctors;
  }
  currentSubscription = sanitizeSubscription(data.subscription);
  applyIssueConfig(data.issueConfig);
  if (data.realName) {
    const localAccount = accountState.users.find((user) => user.email === currentUserEmail);
    if (localAccount) {
      localAccount.realName = data.realName;
    } else {
      accountState.users.push({
        email: currentUserEmail,
        realName: data.realName,
        password: "",
        role: currentUserEmail === OWNER_EMAIL ? "owner" : "user",
      });
    }
    saveAccountState();
  }
  if (!cloudAvailable) return;
  if (!data.state) {
    selectedFiles = [];
    currentSnapshot = null;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = "";
    restoredSessionState = null;
    await replaceStoredImports([]);
    clearWorkspaceStoreEntry(currentUserEmail);
    return;
  }
  currentSnapshot = sanitizeWorkspaceSnapshot(data.snapshot);
  if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
  if (currentSnapshot && shouldRebuildAccountSnapshot(currentSnapshot)) {
    currentSnapshot = null;
  }
  currentSnapshotStale = data.snapshotStale === true;
  currentSnapshotBuiltAt = String(data.snapshotBuiltAt || "");
  selectedFiles = importRefsToClientEntries(data.state.imports || currentSnapshot?.fileRefs || []);
  restoredSessionState = currentSnapshot?.session || (data.state.session && typeof data.state.session === "object" ? data.state.session : null);
  rememberCreatorCalendarSourceRefs();
  saveWorkspaceSnapshotForEmail(activeWorkspaceOwnerKey(), {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: restoredSessionState || {},
    snapshot: currentSnapshot,
  });
}

async function loadCloudCalendarEvents(options = {}) {
  if (!cloudAvailable) return false;
  const adminTargetEmail = normalizeEmail(options.adminTargetEmail);
  const requestEmail = adminTargetEmail ? authUserEmail : currentUserEmail;
  const requestPassword = adminTargetEmail ? authUserPassword : currentUserPassword;
  if (!requestEmail || !requestPassword) return false;
  const preferredDoctorKey = normalizeRosterName(
    restoredSessionState?.doctorKey
    || selectedDoctor()?.key
    || (isCreatorAuthenticated() && !adminTargetEmail ? OWNER_DOCTOR_KEY : currentRosterClaims[0]?.key)
    || "",
  );
  const range = cloudCalendarEventRange();
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "loadCalendarEvents",
      email: requestEmail,
      password: requestPassword,
      targetEmail: adminTargetEmail,
      doctorKey: preferredDoctorKey,
      startDate: range.startDate,
      endDate: range.endDate,
      cachedRevision: options.cachedRevision || "",
    }),
  });
  const data = await readJsonResponse(response, "Calendar load failed.");
  currentCalendarRevision = String(data.calendarRevision || currentCalendarRevision || "");
  lastCacheDiagnostics = (data.diagnostics && typeof data.diagnostics === "object") ? data.diagnostics : null;
  updateVersionBar();
  if (data.snapshotCurrent === true) {
    if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
    if (currentSnapshot) saveCalendarSnapshotCache(currentSnapshot);
    return Boolean(currentSnapshot);
  }
  currentSnapshot = sanitizeWorkspaceSnapshot(clearCloudLoadedSnapshotFilters(data.snapshot));
  if (currentSnapshot) {
    currentSnapshot.calendarRevision = currentCalendarRevision;
    currentSnapshot.cacheKey = currentCalendarSnapshotCacheKey();
  }
  currentSnapshotStale = data.snapshotStale === true;
  currentSnapshotBuiltAt = String(data.snapshotBuiltAt || "");
  if (!currentSnapshot) return false;
  selectedFiles = importRefsToClientEntries(currentSnapshot.fileRefs || selectedFiles.map(importRefForWorkspace));
  restoredSessionState = currentSnapshot.session || clearCloudLoadedSessionFilters(restoredSessionState || {});
  saveWorkspaceSnapshotForEmail(activeWorkspaceOwnerKey(), {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: restoredSessionState || {},
    snapshot: currentSnapshot,
  });
  saveCalendarSnapshotCache(currentSnapshot);
  return true;
}

function cloudCalendarEventRange() {
  const today = new Date();
  const startDate = `${today.getFullYear()}-01-01`;
  const endDate = formatDateKey(new Date(today.getFullYear() + 1, 0, 31));
  return {
    startDate,
    endDate: endDate < startDate ? startDate : endDate,
  };
}

function clearCloudLoadedSnapshotFilters(snapshot) {
  if (!snapshot || typeof snapshot !== "object") return snapshot;
  return {
    ...snapshot,
    session: clearCloudLoadedSessionFilters(snapshot.session || {}),
  };
}

function clearCloudLoadedSessionFilters(session) {
  const settingsState = session?.settings && typeof session.settings === "object" ? session.settings : {};
  return {
    ...(session || {}),
    settings: {
      ...settingsState,
      dateFrom: "",
      dateTo: "",
      hospitalFilter: "all",
    },
  };
}

function serverStorageRequiredMessage() {
  return "Server storage is not configured. Add the D1 ROSTER_DB binding to the Pages project, redeploy, then log in again.";
}

function scheduleCloudStateSave() {
  if (!currentUserEmail) return;
  cancelScheduledCloudStateSave();
  const snapshot = snapshotCloudSavePayload();
  pendingCloudSaveSnapshot = snapshot;
  cloudSaveTimer = setTimeout(() => {
    const queued = pendingCloudSaveSnapshot;
    pendingCloudSaveSnapshot = null;
    saveCloudState(queued || snapshot).catch((error) => {
      if (!error?.isRosterPersistenceError) cloudAvailable = false;
      renderLoginState();
      setStatus(error.message || "Cloud save failed.", true);
    });
  }, 700);
}

function cancelScheduledCloudStateSave() {
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = 0;
  pendingCloudSaveSnapshot = null;
}

async function flushCloudStateSave() {
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = 0;
  const snapshot = pendingCloudSaveSnapshot;
  pendingCloudSaveSnapshot = null;
  if (!snapshot) return;
  await saveCloudState(snapshot);
}

function capturePendingCloudStateSave() {
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = 0;
  const snapshot = pendingCloudSaveSnapshot;
  pendingCloudSaveSnapshot = null;
  return snapshot;
}

function queueBackgroundCloudStateSave(snapshot = null, options = {}) {
  const payload = snapshot || snapshotCloudSavePayload();
  if (!payload) return;
  const run = () => saveCloudState(payload).catch((error) => {
    if (!error?.isRosterPersistenceError) cloudAvailable = false;
    renderLoginState();
    setStatus(error.message || "Cloud save failed.", true);
  });
  const delayMs = Math.max(0, Number(options.delayMs || 0));
  if (delayMs) {
    window.setTimeout(run, delayMs);
  } else {
    run();
  }
}

function shouldRebuildAccountSnapshot(snapshot) {
  if (currentUserRole === "creator" || activeCalendarMode() === "doctor-profile") return false;
  if (!currentRosterClaims.length) {
    return Boolean(snapshot?.doctorOptions?.length || snapshot?.preview?.events?.length);
  }
  const claimMarkers = new Set(currentRosterClaims.map((claim) => `${claim.sourceType}:${claim.key}`));
  const snapshotMarkers = new Set();
  for (const doctor of snapshot?.doctorOptions || []) {
    const aliases = Array.isArray(doctor?.aliases) && doctor.aliases.length
      ? doctor.aliases
      : normalizedDoctorSourceTypes(doctor).map((sourceType) => ({
          sourceType,
          key: doctor?.key || "",
          displayName: doctor?.displayName || "",
        }));
    for (const alias of aliases) {
      const sourceType = String(alias?.sourceType || "").toLowerCase();
      const key = normalizeRosterName(alias?.key || "");
      if (sourceType && key) snapshotMarkers.add(`${sourceType}:${key}`);
    }
  }
  return [...claimMarkers].some((marker) => !snapshotMarkers.has(marker));
}

function snapshotCloudSavePayload() {
  if (activeCalendarMode() === "doctor-profile") {
    return {
      accountEmail: currentUserEmail,
      requestEmail: authUserEmail || currentUserEmail,
      requestPassword: authUserPassword || currentUserPassword,
      doctorProfile: { ...activeDoctorProfile },
      imports: [],
      session: buildActiveSessionState(),
    };
  }
  return {
    accountEmail: viewedAccountEmail(),
    requestEmail: adminViewingEmail ? authenticatedAccountEmail() : viewedAccountEmail(),
    requestPassword: adminViewingEmail ? authUserPassword : currentUserPassword,
    targetEmail: adminViewingEmail ? viewedAccountEmail() : "",
    imports: selectedFiles.map((entry) => ({ ...entry })),
    session: buildActiveSessionState(),
    removedImportIds: [],
  };
}

function forceCreatorDoctorSession() {
  restoredSessionState = {
    ...(restoredSessionState || {}),
    doctorKey: OWNER_DOCTOR_KEY,
  };
  if (currentSnapshot) {
    currentSnapshot = sanitizeWorkspaceSnapshot({
      ...currentSnapshot,
      session: {
        ...(currentSnapshot.session || {}),
        doctorKey: OWNER_DOCTOR_KEY,
      },
    });
  }
}

function creatorCalendarSavePayload() {
  if (currentUserEmail !== OWNER_EMAIL || adminViewingEmail || activeDoctorProfile) return null;
  return {
    accountEmail: currentUserEmail,
    requestEmail: currentUserEmail,
    requestPassword: currentUserPassword,
    targetEmail: "",
    imports: selectedFiles.map((entry) => ({ ...entry })),
    session: {
      ...buildActiveSessionState(),
      doctorKey: OWNER_DOCTOR_KEY,
    },
    removedImportIds: [],
  };
}

async function saveCloudState(snapshot = null) {
  const task = () => saveCloudStateNow(snapshot);
  const queued = cloudStateSaveQueue.then(task, task);
  cloudStateSaveQueue = queued.catch(() => {});
  return await queued;
}

function savePayloadMatchesActiveCalendar(payload) {
  if (payload?.doctorProfile) {
    return activeCalendarMode() === "doctor-profile" && activeDoctorProfile?.id === payload.doctorProfile.id;
  }
  return normalizeEmail(payload?.accountEmail || "") === viewedAccountEmail() && !activeDoctorProfile;
}

async function saveCloudStateNow(snapshot = null) {
  const payload = snapshot || snapshotCloudSavePayload();
  if (!payload.accountEmail || !payload.requestEmail || !payload.requestPassword || !cloudAvailable) return;
  const snapshotPayload = await buildWorkspaceSnapshotPayload(payload.session);
  const shouldApplySavedSnapshot = savePayloadMatchesActiveCalendar(payload);
  if (payload.doctorProfile) {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "saveDoctorProfile",
        email: payload.requestEmail,
        password: payload.requestPassword,
        profileId: payload.doctorProfile.id,
        doctorKey: payload.doctorProfile.doctorKey,
        displayName: payload.doctorProfile.displayName,
        sourceTypes: payload.doctorProfile.sourceTypes,
        state: { version: 1, imports: [], session: payload.session },
        snapshot: snapshotPayload,
      }),
    });
    const data = await readJsonResponse(response, "Doctor profile save failed.");
    const savedCalendarRevision = String(data.calendarRevision || "");
    if (shouldApplySavedSnapshot && savedCalendarRevision) currentCalendarRevision = savedCalendarRevision;
    if (snapshotPayload && shouldApplySavedSnapshot) {
      currentSnapshot = sanitizeWorkspaceSnapshot(snapshotPayload);
      if (currentSnapshot) currentSnapshot.calendarRevision = currentCalendarRevision;
      currentSnapshotStale = false;
      currentSnapshotBuiltAt = new Date().toISOString();
      rememberCreatorCalendarSourceRefs();
      saveCalendarSnapshotCache(currentSnapshot);
    }
    renderLoginState();
    return;
  }
  const state = await buildCloudState(payload.imports, payload.session);
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "save",
      email: payload.requestEmail,
      password: payload.requestPassword,
      targetEmail: payload.targetEmail,
      state,
      snapshot: snapshotPayload,
      removedImportIds: payload.removedImportIds || [],
    }),
  });
  const data = await readJsonResponse(response, "Cloud save failed.");
  const savedCalendarRevision = String(data.calendarRevision || "");
  if (shouldApplySavedSnapshot && savedCalendarRevision) currentCalendarRevision = savedCalendarRevision;
  if (shouldApplySavedSnapshot && data.claims && payload.accountEmail === currentUserEmail) currentRosterClaims = sanitizeRosterClaims(data.claims);
  if (snapshotPayload && shouldApplySavedSnapshot) {
    currentSnapshot = sanitizeWorkspaceSnapshot(snapshotPayload);
    if (currentSnapshot) currentSnapshot.calendarRevision = currentCalendarRevision;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = new Date().toISOString();
    rememberCreatorCalendarSourceRefs();
    saveCalendarSnapshotCache(currentSnapshot);
  }
  renderLoginState();
}

async function replaceActiveRostersWithCurrentUploads() {
  if (!isCreatorAuthenticated()) return;
  try {
    setStatus("Rebuilding roster database from roster files...");
    if (!selectedFiles.length) await refreshCalendarStoreStatus({ silent: true });
    const sourceEntries = selectedFiles.length ? selectedFiles : retainedRosterEntriesFromStatus();
    if (!sourceEntries.length) {
      throw new Error("Rebuild requires at least one retained roster file. Re-upload the missing source files first if none are listed.");
    }
    const rebuildEntries = [];
    const unavailableNames = [];
    for (const entry of sourceEntries) {
      const retained = await ensureRosterEntrySource(entry).catch(() => null);
      if (!retained?.file) unavailableNames.push(entry.name);
      else rebuildEntries.push(retained);
    }
    if (unavailableNames.length) {
      throw new Error(`Could not rebuild ${unavailableNames.join(", ")} because the retained source file is missing. Re-upload ${unavailableNames.length === 1 ? "it" : "them"} once.`);
    }
    await saveSelectedRosterFilesToD1(rebuildEntries, { force: true, retainSources: false });
    const keepFileIds = rebuildEntries.map((entry) => entry.id);
    calendarStoreStatus = {
      ...await calendarStoreRequest("replaceActiveRosterFiles", {
        keepFileIds,
        selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
      }),
      checkedAt: new Date().toISOString(),
    };
    calendarStoreStatusError = "";
    replaceActiveCalendarCustomEvents(customEventsForActiveCalendar());
    await loadCloudCalendarEvents();
    if (currentSnapshot) renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
    renderFileSurfaces();
    setStatus(`Roster database rebuilt from ${keepFileIds.length} roster file${keepFileIds.length === 1 ? "" : "s"}.`);
  } catch (error) {
    setStatus(error.message || "Could not replace active rosters.", true);
  }
}

function retainedRosterEntriesFromStatus() {
  return (calendarStoreStatus?.files || [])
    .filter((file) => file?.id && file.rawSourceAvailable === true)
    .map((file) => ({
      id: file.id,
      repoId: file.id,
      name: file.name,
      sourceType: file.sourceType,
      size: file.size || 0,
      lastModified: file.lastModified || 0,
      addedAt: file.uploadedAt || "",
    }));
}

async function reparseRosterFile(id) {
  const statusEntry = (calendarStoreStatus?.files || []).find((file) => file.id === id);
  const entry = selectedFiles.find((item) => item.id === id) || (statusEntry ? {
    id: statusEntry.id,
    repoId: statusEntry.id,
    name: statusEntry.name,
    sourceType: statusEntry.sourceType,
    addedAt: "",
  } : null);
  if (!entry) return;
  try {
    setStatus(`Reparsing ${entry.name}...`);
    const retained = await ensureRosterEntrySource(entry);
    if (!retained?.file) throw new Error(`${entry.name} has no retained source file. Re-upload it once to enable reparsing.`);
    await saveSelectedRosterFilesToD1([retained], { force: true });
    await refreshCalendarStoreStatus({ silent: true });
    const reparsed = (calendarStoreStatus?.files || []).find((file) => file.id === entry.id);
    if (!reparsed || Number(reparsed.eventCount || 0) <= 0) {
      setRosterSyncState(entry, "failed", "Reparse produced 0 events.");
      renderFileSurfaces();
      throw new Error(`${entry.name} reparse produced 0 events.`);
    }
    await loadCloudCalendarEvents();
    if (currentSnapshot) renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
    renderFileSurfaces();
    setStatus(`${entry.name} reparsed.`);
  } catch (error) {
    setStatus(error.message || `Could not reparse ${entry.name}.`, true);
  }
}


async function buildWorkspaceSnapshotPayload(session = buildActiveSessionState()) {
  if (!latestPreview || !selectedDoctor()) return null;
  return {
    preview: JSON.parse(JSON.stringify(latestPreview)),
    session: JSON.parse(JSON.stringify(session || {})),
    doctorOptions: JSON.parse(JSON.stringify(doctorOptions || [])),
    detectedSources: JSON.parse(JSON.stringify(detectedSources || {})),
    fileRefs: selectedFiles.map(importRefForWorkspace),
    subscriptionFeeds: await buildSubscriptionFeeds(session),
    insightCache: doctorAnalysisCacheKey && doctorAnalysisCache.size ? buildInsightCachePayload() : currentSnapshot?.insightCache || null,
    calendarRevision: currentCalendarRevision,
  };
}

function cacheCurrentSnapshot(session = buildActiveSessionState()) {
  if (!latestPreview || !selectedDoctor()) return;
  currentSnapshot = sanitizeWorkspaceSnapshot({
    preview: JSON.parse(JSON.stringify(latestPreview)),
    session: JSON.parse(JSON.stringify(session || {})),
    doctorOptions: JSON.parse(JSON.stringify(doctorOptions || [])),
    detectedSources: JSON.parse(JSON.stringify(detectedSources || {})),
    fileRefs: selectedFiles.map(importRefForWorkspace),
    subscriptionFeeds: currentSnapshot?.subscriptionFeeds || {},
    insightCache: doctorAnalysisCacheKey && doctorAnalysisCache.size ? buildInsightCachePayload() : currentSnapshot?.insightCache || null,
    calendarRevision: currentCalendarRevision,
    cacheKey: currentCalendarSnapshotCacheKey(),
    cachedAt: new Date().toISOString(),
  });
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = new Date().toISOString();
  rememberCreatorCalendarSourceRefs();
  saveCalendarSnapshotCache(currentSnapshot);
}

async function loadServerUsers() {
  const requestEmail = adminViewingEmail ? authUserEmail : currentUserEmail;
  const requestPassword = adminViewingEmail ? authUserPassword : currentUserPassword;
  if (!requestEmail || !requestPassword || normalizeEmail(requestEmail) !== OWNER_EMAIL || !cloudAvailable) return;
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "listUsers",
        email: requestEmail,
        password: requestPassword,
      }),
    });
    const data = await readJsonResponse(response, "Could not load users.");
    serverUsers = data.users || [];
    availableRosterDoctors = sanitizeAvailableRosterDoctors(data.availableDoctors || availableRosterDoctors);
    applyIssueConfig(data.issueConfig);
    syncAccountsButton();
  } catch {
    // Keep the last available local list.
  }
}

async function refreshCalendarStoreStatus(options = {}) {
  if (!isCreatorAuthenticated() || !cloudAvailable) return;
  try {
    const data = await calendarStoreRequest("calendarStoreStatus", {
      selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
      expectedFileIds: selectedFiles.map((entry) => entry.id),
    });
    calendarStoreStatus = { ...data, checkedAt: new Date().toISOString() };
    if (isCreatorAuthenticated() && Array.isArray(data.files) && data.files.length && !selectedFiles.some((entry) => entry.file)) {
      selectedFiles = importRefsToClientEntries(data.files);
      rememberCreatorCalendarSourceRefs();
      renderFilesList();
    }
    calendarStoreStatusError = "";
    if (!options.silent) setStatus("Roster database status checked.");
  } catch (error) {
    calendarStoreStatusError = error.message || "Could not check roster database status.";
    if (!options.silent) setStatus(calendarStoreStatusError, true);
  }
  if (!accountsModal.classList.contains("hidden") && currentAdminTab === "system") renderAccountsModal();
}

async function toggleAdminConsole() {
  adminConsoleOpen = !adminConsoleOpen;
  if (!adminConsoleOpen) {
    renderAccountsModal();
    return;
  }
  adminConsoleLoading = true;
  renderAccountsModal();
  try {
    const data = await calendarStoreRequest("consoleMessages");
    adminConsoleMessages = Array.isArray(data.messages) ? data.messages : [];
  } catch (error) {
    adminConsoleMessages = [];
    setStatus(error.message || "Could not load console history.", true);
  } finally {
    adminConsoleLoading = false;
    if (!accountsModal.classList.contains("hidden") && currentAdminTab === "system") renderAccountsModal();
  }
}

async function persistConsoleMessage(message, isError) {
  if (!cloudAvailable || !currentUserEmail || !currentUserPassword) return;
  try {
    await calendarStoreRequest("appendConsoleMessage", { message, isError: isError === true });
  } catch {
    // Console persistence must not interfere with the foreground workflow.
  }
}

async function buildDerivedCalendarFilePayload(importEntry, statusFile = {}) {
  const parsed = await parseRosterEntries([importEntry], null);
  const sources = parsed.sources || {};
  const sourceType = sources.mmc?.length ? "mmc"
    : sources.ddh?.length ? "ddh"
      : sources.casey?.length ? "casey"
        : sources.mch?.length ? "mch"
          : String(statusFile.sourceType || importEntry.sourceType || "").toLowerCase();
  const sourceEntries = {
    mmc: sourceType === "mmc" ? sources.mmc || [] : [],
    ddh: sourceType === "ddh" ? sources.ddh || [] : [],
    casey: sourceType === "casey" ? sources.casey || [] : [],
    mch: sourceType === "mch" ? sources.mch || [] : [],
  };
  const doctors = rosterDoctorOptions(sourceEntries.mmc, sourceEntries.ddh, sourceEntries.casey, sourceEntries.mch)
    .flatMap((doctor) => {
      const aliases = Array.isArray(doctor.aliases) && doctor.aliases.length
        ? doctor.aliases.filter((alias) => String(alias.sourceType || "").toLowerCase() === sourceType)
        : [{ key: doctor.key, displayName: doctor.displayName, sourceType }];
      return aliases.map((alias) => ({
        key: normalizeRosterName(alias.key || doctor.key),
        displayName: alias.displayName || doctor.displayName,
        sourceType,
      }));
    });
  const uniqueDoctors = [];
  const seenDoctors = new Set();
  for (const doctor of doctors) {
    const marker = `${doctor.sourceType}:${doctor.key}`;
    if (seenDoctors.has(marker)) continue;
    seenDoctors.add(marker);
    uniqueDoctors.push(doctor);
  }
  const analysisSettings = {
    ...rosterDefaultSettings(),
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
    includeLocations: true,
  };
  const eventsByDoctor = {};
  const issuesByDoctor = {};
  let eventCount = 0;
  for (const doctor of uniqueDoctors) {
    const view = buildRosterView(
      sourceEntries.mmc,
      sourceEntries.ddh,
      doctor.key,
      analysisSettings,
      {},
      {},
      [],
      sourceEntries.casey,
      sourceEntries.mch,
    );
    const events = view.events.map(serializeEvent);
    eventsByDoctor[doctor.key] = events;
    issuesByDoctor[doctor.key] = (view.issues || []).map((issue) => ({
      id: issue.id,
      source: issue.source,
      seniority: issue.seniority,
      startDay: issue.startDay,
      rawValue: issue.rawValue,
      status: issue.status,
      message: issue.message,
      resolutionType: issue.resolutionType,
      suggestedTitle: issue.suggestedTitle,
      timeLabel: issue.timeLabel,
    }));
    eventCount += events.length;
  }
  assertDerivedCalendarFilePayload(importEntry, uniqueDoctors, eventCount);
  return {
    file: {
      ...importRefForWorkspace(importEntry),
      id: statusFile.id || importEntry.id,
      sourceType,
    },
    selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
    doctors: uniqueDoctors,
    eventsByDoctor,
    issuesByDoctor,
    eventCount,
  };
}

function assertDerivedCalendarFilePayload(importEntry, doctors, eventCount) {
  const name = importEntry?.name || "Uploaded roster";
  if (!doctors.length) {
    throw new Error(`${name} could not be indexed: no doctors were found.`);
  }
  if (!eventCount) {
    throw new Error(`${name} could not be indexed: no events were found.`);
  }
  if (eventCount < doctors.length) {
    throw new Error(`${name} could not be indexed reliably: only ${eventCount} events were found for ${doctors.length} doctors.`);
  }
}

async function calendarStoreRequest(action, extra = {}) {
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action,
      email: authUserEmail || currentUserEmail,
      password: authUserPassword || currentUserPassword,
      ...extra,
    }),
  });
  return await readJsonResponse(response, "SQL calendar store request failed.");
}

async function buildCloudState(imports = selectedFiles, session = buildActiveSessionState()) {
  await saveSelectedRosterFilesToD1(imports);
  const subscriptionFeeds = await buildSubscriptionFeeds(session);
  return {
    version: 1,
    imports: imports.map(importRefForWorkspace),
    session,
    subscriptionFeeds,
  };
}

async function saveSelectedRosterFilesToD1(imports = selectedFiles, options = {}) {
  if (!cloudAvailable || !currentUserEmail) return emptyRosterPersistenceSummary();
  const entries = (imports || []).filter((entry) => entry?.file);
  if (!entries.length) return emptyRosterPersistenceSummary();
  const expectedFileIds = entries.map((entry) => entry.id);
  const persistedIds = new Set(calendarStoreStatus?.expectedFiles?.persistedFileIds || []);
  const failedIds = new Set([...rosterSyncStates.entries()].filter(([, state]) => state.status === "failed").map(([id]) => id));
  const entriesToSave = options.force === true
    ? entries
    : entries.filter((entry) => !persistedIds.has(entry.id) || failedIds.has(entry.id));
  const saveResults = [];
  let latestStatus = calendarStoreStatus;
  if (!entriesToSave.length) return summarizeRosterPersistence(entries, latestStatus, saveResults);
  beginRosterSync(entriesToSave, options.force === true ? "rebuild" : "sync");
  for (const entry of entriesToSave) {
    try {
      if (options.retainSources !== false) {
        setRosterSyncState(entry, "uploading-source");
        await retainRosterSource(entry);
      }
      setRosterSyncState(entry, "parsing");
      const payload = await buildDerivedCalendarFilePayload(entry, entry);
      setRosterSyncState(entry, "saving");
      latestStatus = await calendarStoreRequest("saveDerivedCalendarFile", {
        ...payload,
        expectedFileIds,
        skipStatus: true,
      });
      latestStatus = mergeLightweightRosterStatus(latestStatus, payload.file, latestStatus.fileStatus, expectedFileIds);
      saveResults.push({ entry, ok: true });
      setRosterSyncState(entry, "synced");
    } catch (error) {
      saveResults.push({ entry, ok: false, error });
      setRosterSyncState(entry, "failed", error?.message || "D1 save failed.");
    }
  }
  if (isCreatorAuthenticated()) {
    try {
      latestStatus = await calendarStoreRequest("calendarStoreStatus", {
        selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
        expectedFileIds,
      });
      calendarStoreStatusError = "";
    } catch (error) {
      calendarStoreStatusError = error.message || "Could not check roster database status.";
      // Keep the most recent save response if the creator status refresh fails.
    }
  }
  if (latestStatus) calendarStoreStatus = { ...latestStatus, checkedAt: new Date().toISOString() };
  invalidateCalendarSnapshotCache();
  const summary = summarizeRosterPersistence(entries, calendarStoreStatus, saveResults);
  lastRosterPersistence = summary;
  renderFileSurfaces();
  if (!summary.complete) {
    finishRosterSync();
    const error = new Error(rosterPersistenceFailureMessage(summary));
    error.isRosterPersistenceError = true;
    throw error;
  }
  setStatus(`Calendar saved to D1. ${summary.persistedCount}/${summary.expectedCount} roster file${summary.expectedCount === 1 ? "" : "s"} confirmed.`);
  finishRosterSync();
  return summary;
}

function mergeLightweightRosterStatus(status, file, fileStatus, expectedFileIds = []) {
  const base = status && Array.isArray(status.files) ? status : calendarStoreStatus || { files: [] };
  const normalizedFile = {
    id: fileStatus?.id || file?.id,
    name: fileStatus?.name || file?.name || "roster.xlsx",
    sourceType: fileStatus?.sourceType || file?.sourceType || "",
    expectedDoctors: Number(fileStatus?.indexedDoctors || 0),
    indexedDoctors: Number(fileStatus?.indexedDoctors || 0),
    eventCount: Number(fileStatus?.eventCount || 0),
    selectedDoctorEventCount: 0,
    rawSourceAvailable: true,
    status: fileStatus?.status || (Number(fileStatus?.eventCount || 0) > 0 ? "populated" : "missing"),
  };
  const files = [...(base.files || []).filter((item) => item.id !== normalizedFile.id), normalizedFile]
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || "")));
  const populated = files.filter((item) => item.status === "populated").length;
  const partial = files.filter((item) => item.status === "partial").length;
  const expectedIds = expectedFileIds.length ? expectedFileIds : files.map((item) => item.id);
  return {
    ...base,
    ok: true,
    total: files.length,
    populated,
    partial,
    remaining: Math.max(0, files.length - populated),
    eventCount: files.reduce((total, item) => total + Number(item.eventCount || 0), 0),
    files,
    expectedFiles: {
      expectedCount: expectedIds.length,
      expectedFileIds: expectedIds,
      persistedCount: files.filter((item) => expectedIds.includes(item.id)).length,
      populatedCount: files.filter((item) => expectedIds.includes(item.id) && Number(item.eventCount || 0) > 0).length,
      activeCount: files.filter((item) => expectedIds.includes(item.id) && Number(item.eventCount || 0) > 0).length,
      persistedFileIds: files.filter((item) => expectedIds.includes(item.id)).map((item) => item.id),
      populatedFileIds: files.filter((item) => expectedIds.includes(item.id) && Number(item.eventCount || 0) > 0).map((item) => item.id),
      activeFileIds: files.filter((item) => expectedIds.includes(item.id) && Number(item.eventCount || 0) > 0).map((item) => item.id),
      missingFileIds: expectedIds.filter((id) => !files.some((item) => item.id === id)),
    },
  };
}


function beginRosterSync(entries, mode = "sync") {
  for (const entry of entries) setRosterSyncState(entry, "pending", "", mode);
  scheduleRosterSyncRefresh();
}

function setRosterSyncState(entry, status, message = "", mode = "sync") {
  rosterSyncStates.set(entry.id, { status, message, mode, name: entry.name });
  renderFileSurfaces();
  if (!accountsModal.classList.contains("hidden") && currentAdminTab === "system") renderAccountsModal();
}

function finishRosterSync() {
  if (![...rosterSyncStates.values()].some((state) => ["pending", "parsing", "saving"].includes(state.status))) {
    clearInterval(rosterSyncRefreshTimer);
    rosterSyncRefreshTimer = 0;
  }
}

function scheduleRosterSyncRefresh() {
  if (rosterSyncRefreshTimer) return;
  rosterSyncRefreshTimer = setInterval(() => void refreshCalendarStoreStatus({ silent: true }), 5000);
}

function rosterSyncLabel(entry) {
  const state = rosterSyncStates.get(entry.id);
  if (!state || state.status === "synced") return "";
  const labels = { pending: "Queued to sync", "uploading-source": "Retaining source file…", parsing: "Parsing roster file…", saving: "Saving to roster database…", failed: `Sync failed${state.message ? `: ${state.message}` : ""}` };
  return `<span>${escapeHtml(labels[state.status] || "")}</span>`;
}

function rosterSyncSummary() {
  const states = [...rosterSyncStates.values()];
  const active = states.filter((state) => ["pending", "parsing", "saving"].includes(state.status));
  if (!active.length) return "";
  const done = states.filter((state) => state.status === "synced").length;
  const total = states.length;
  const verb = states.some((state) => state.mode === "rebuild") ? "Rebuilding" : "Syncing";
  return `${verb} ${done}/${total} roster files`;
}

function emptyRosterPersistenceSummary() {
  return {
    expectedCount: 0,
    persistedCount: 0,
    activeCount: 0,
    expectedFileIds: [],
    persistedFileIds: [],
    missingEntries: [],
    failedEntries: [],
    complete: true,
  };
}

function summarizeSelectedRosterPersistence(imports = selectedFiles, status = calendarStoreStatus) {
  return summarizeRosterPersistence(
    (imports || []).filter((entry) => entry?.file),
    status,
    [],
  );
}

function summarizeRosterPersistence(entries = [], status = null, saveResults = []) {
  const expectedEntries = (entries || []).filter((entry) => entry?.id);
  const expectedFileIds = expectedEntries.map((entry) => entry.id);
  const expectedFileIdSet = new Set(expectedFileIds);
  const statusExpected = status?.expectedFiles && typeof status.expectedFiles === "object"
    ? status.expectedFiles
    : null;
  const statusExpectedIds = Array.isArray(statusExpected?.expectedFileIds) ? statusExpected.expectedFileIds : [];
  const statusMatchesEntries = Boolean(statusExpected)
    && statusExpectedIds.length === expectedFileIds.length
    && statusExpectedIds.every((id) => expectedFileIdSet.has(id));
  const persistedFileIds = statusMatchesEntries
    ? (statusExpected.populatedFileIds || []).filter((id) => expectedFileIdSet.has(id))
    : (status?.files || []).map((file) => file.id).filter((id) => expectedFileIdSet.has(id));
  const persistedSet = new Set(persistedFileIds);
  const activeFileIds = statusMatchesEntries
    ? (statusExpected.activeFileIds || []).filter((id) => expectedFileIdSet.has(id))
    : (status?.files || []).filter((file) => file.active !== false).map((file) => file.id).filter((id) => expectedFileIdSet.has(id));
  const failedEntries = saveResults
    .filter((result) => result && result.ok === false && result.entry?.id)
    .map((result) => ({
      id: result.entry.id,
      name: result.entry.name,
      message: result.error?.message || "D1 save failed.",
    }));
  const missingEntries = expectedEntries
    .filter((entry) => !persistedSet.has(entry.id))
    .map((entry) => ({ id: entry.id, name: entry.name }));
  return {
    expectedCount: expectedEntries.length,
    persistedCount: persistedFileIds.length,
    activeCount: activeFileIds.length,
    expectedFileIds,
    persistedFileIds,
    missingEntries,
    failedEntries,
    complete: expectedEntries.length === persistedFileIds.length,
  };
}

function hasUnconfirmedLocalRosterFiles() {
  const localEntries = selectedFiles.filter((entry) => entry?.file);
  if (!localEntries.length) return false;
  const summary = summarizeSelectedRosterPersistence(localEntries, calendarStoreStatus);
  if (!lastRosterPersistence) return !summary.complete;
  const currentIds = localEntries.map((entry) => entry.id).sort().join("|");
  const persistedIds = [...(lastRosterPersistence.expectedFileIds || [])].sort().join("|");
  return currentIds !== persistedIds || !lastRosterPersistence.complete || !summary.complete;
}

function rosterPersistenceFailureMessage(summary) {
  const missingNames = summary.missingEntries.map((entry) => entry.name);
  const failedNames = summary.failedEntries.map((entry) => entry.name);
  const missingDetail = missingNames.length ? ` Missing from D1: ${[...new Set(missingNames)].join(", ")}.` : "";
  const failedDetail = failedNames.length ? ` Failed to save this upload: ${[...new Set(failedNames)].join(", ")}.` : "";
  return `Roster save incomplete: ${summary.persistedCount}/${summary.expectedCount} selected roster file${summary.expectedCount === 1 ? "" : "s"} confirmed in D1.${missingDetail}${failedDetail}`;
}

function buildActiveSessionState() {
  return {
    doctorKey: doctorOptions.length > 1 ? doctorSelect.value : doctorOptions[0]?.key || "",
    settings: { ...settings },
    exportRange: normalizeSavedExportRange(pendingExportRange),
    overrides: cleanOverrides(),
    customEvents: customEventsForActiveCalendar(),
    conflictSelections: { ...conflictSelections },
    hadPreview: Boolean(latestPreview),
    savedAt: new Date().toISOString(),
  };
}

async function buildSubscriptionFeeds(session = buildActiveSessionState()) {
  if (activeCalendarMode() === "doctor-profile") return {};
  const doctor = selectedDoctor();
  if (!doctor) return {};
  const fullEvents = latestPreview?.events?.length
    ? buildExportEventsFromBase(latestPreview, { mode: "full" })
    : await buildBrowserExportEvents(doctor, { mode: "full" }).catch(() => []);
  const rangeConfig = exportConfigForMode("range", session?.exportRange || defaultExportRangeState());
  const rangeEvents = rangeConfig.startDate
    ? (latestPreview?.events?.length
      ? buildExportEventsFromBase(latestPreview, rangeConfig)
      : await buildBrowserExportEvents(doctor, rangeConfig).catch(() => []))
    : [];
  if (!fullEvents.length && !rangeEvents.length) return {};
  return {
    full: fullEvents.length ? {
      doctorKey: doctor.key,
      doctorDisplay: doctor.displayName,
      generatedAt: new Date().toISOString(),
      ics: exportIcs(fullEvents, doctor.displayName),
    } : null,
    range: rangeEvents.length ? {
      doctorKey: doctor.key,
      doctorDisplay: doctor.displayName,
      startDate: rangeConfig.startDate,
      endDate: rangeConfig.allFuture ? "" : rangeConfig.endDate,
      allFuture: rangeConfig.allFuture !== false,
      generatedAt: new Date().toISOString(),
      ics: exportIcs(rangeEvents, doctor.displayName),
    } : null,
  };
}

async function serializeCloudImports(imports) {
  return await Promise.all(imports.map(async (entry) => {
    if (!entry.file) {
      return importRefForWorkspace(entry);
    }
    return {
      id: entry.id,
      name: entry.name,
      size: entry.size,
      lastModified: entry.lastModified,
      addedAt: entry.addedAt,
      sourceType: entry.sourceType,
      doctors: parsedImportDoctors.get(entry.id) || [],
      type: entry.file?.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      dataUrl: await fileToDataUrl(entry.file),
    };
  }));
}

async function deserializeCloudImports(imports) {
  const entries = [];
  for (const item of imports) {
    if (!item?.dataUrl || !item?.name) continue;
    const blob = await dataUrlToBlob(item.dataUrl);
    entries.push({
      id: item.id || `${item.name}:${item.size || blob.size}:${item.lastModified || Date.now()}`,
      name: item.name,
      size: item.size || blob.size,
      lastModified: item.lastModified || Date.now(),
      addedAt: item.addedAt || new Date().toISOString(),
      sourceType: item.sourceType || "pending",
      file: new File([blob], item.name, { type: item.type || blob.type, lastModified: item.lastModified || Date.now() }),
    });
  }
  return entries;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
}

async function dataUrlToBlob(dataUrl) {
  const response = await fetch(dataUrl);
  return await response.blob();
}

async function retainRosterSource(entry) {
  if (!entry?.file) return;
  await calendarStoreRequest("uploadRawRosterFile", {
    file: importRefForWorkspace(entry),
    type: entry.file.type || "application/octet-stream",
    dataUrl: await fileToDataUrl(entry.file),
  });
}

async function ensureRosterEntrySource(entry) {
  if (entry?.file) return entry;
  const raw = await calendarStoreRequest("fetchRawRosterFile", { fileId: entry.id });
  if (!raw?.dataUrl) return entry;
  const blob = await dataUrlToBlob(raw.dataUrl);
  const hydrated = {
    ...entry,
    file: new File([blob], entry.name, {
      type: raw.type || blob.type || "application/octet-stream",
      lastModified: entry.lastModified || Date.now(),
    }),
  };
  selectedFiles = selectedFiles.map((item) => item.id === hydrated.id ? hydrated : item);
  return hydrated;
}

function currentImportStateKey() {
  if (!selectedFiles.length) return "";
  return selectedFiles.map((entry) => entry.id).sort().join("|");
}

function loadSessionStore() {
  try {
    return JSON.parse(localStorage.getItem(SESSION_STATE_KEY) || "{}");
  } catch {
    return {};
  }
}

function saveSessionStore(store) {
  localStorage.setItem(SESSION_STATE_KEY, JSON.stringify(store));
}

function loadWorkspaceStore() {
  try {
    const store = JSON.parse(localStorage.getItem(ACCOUNT_WORKSPACES_KEY) || "{}");
    if (Object.values(store).some((workspace) => workspace?.snapshot)) {
      try {
        saveWorkspaceStore(store);
      } catch {
        // Use the in-memory migration for this run even if storage cleanup fails.
      }
    }
    return lightweightWorkspaceStore(store);
  } catch {
    return {};
  }
}

function saveWorkspaceStore(store) {
  const lightweight = lightweightWorkspaceStore(store);
  try {
    localStorage.setItem(ACCOUNT_WORKSPACES_KEY, JSON.stringify(lightweight));
  } catch (error) {
    if (!isStorageQuotaError(error)) throw error;
    const sessionOnly = {};
    for (const [key, workspace] of Object.entries(lightweight)) {
      sessionOnly[key] = {
        fileRefs: [],
        session: workspace?.session || {},
      };
    }
    localStorage.setItem(ACCOUNT_WORKSPACES_KEY, JSON.stringify(sessionOnly));
  }
}

function lightweightWorkspaceStore(store) {
  const next = {};
  for (const [key, workspace] of Object.entries(store || {})) {
    next[key] = {
      fileRefs: Array.isArray(workspace?.fileRefs) ? workspace.fileRefs.map(importRefForWorkspace).filter((item) => item.id) : [],
      session: workspace?.session && typeof workspace.session === "object" ? workspace.session : {},
    };
  }
  return next;
}

function isStorageQuotaError(error) {
  return error?.name === "QuotaExceededError"
    || error?.name === "NS_ERROR_DOM_QUOTA_REACHED"
    || error?.code === 22
    || error?.code === 1014;
}

function loadCalendarSnapshotCacheStore() {
  try {
    const store = JSON.parse(localStorage.getItem(CALENDAR_SNAPSHOT_CACHE_KEY) || "{}");
    return store && typeof store === "object" ? store : {};
  } catch {
    return {};
  }
}

function saveCalendarSnapshotCacheStore(store) {
  try {
    localStorage.setItem(CALENDAR_SNAPSHOT_CACHE_KEY, JSON.stringify(store || {}));
  } catch (error) {
    if (!isStorageQuotaError(error)) throw error;
    const entries = Object.entries(store || {})
      .sort((left, right) => String(right[1]?.cachedAt || "").localeCompare(String(left[1]?.cachedAt || "")))
      .slice(0, 3);
    localStorage.setItem(CALENDAR_SNAPSHOT_CACHE_KEY, JSON.stringify(Object.fromEntries(entries)));
  }
}

function calendarSnapshotContext(options = {}) {
  const mode = options.mode || activeCalendarMode();
  const owner = mode === "doctor-profile"
    ? String(options.ownerId || activeDoctorProfile?.ownerId || activeWorkspaceOwnerKey() || "").trim()
    : normalizeEmail(options.ownerEmail || options.ownerId || activeWorkspaceOwnerKey() || viewedAccountEmail() || currentUserEmail);
  const doctorKey = normalizeRosterName(
    options.doctorKey
    || restoredSessionState?.doctorKey
    || selectedDoctor()?.key
    || (options.ownerEmail ? preferredDoctorKeyForAccountEmail(options.ownerEmail) : preferredDoctorKeyForCurrentAccount())
    || "",
  );
  const range = options.range || cloudCalendarEventRange();
  return {
    mode,
    ownerId: owner,
    ownerEmail: mode === "doctor-profile" ? "" : owner,
    doctorKey,
    range,
  };
}

function calendarSnapshotCacheKeyForContext(context = {}) {
  const normalized = calendarSnapshotContext(context);
  if (!normalized.ownerId) return "";
  return [normalized.mode, normalized.ownerId, normalized.doctorKey, normalized.range.startDate || "", normalized.range.endDate || ""].join("|");
}

function currentCalendarSnapshotCacheKey(options = {}) {
  return calendarSnapshotCacheKeyForContext(options);
}

function loadCachedCalendarSnapshotForContext(context = {}) {
  const key = calendarSnapshotCacheKeyForContext(context);
  if (!key) return null;
  const entry = loadCalendarSnapshotCacheStore()[key];
  const snapshot = sanitizeWorkspaceSnapshot(entry?.snapshot || entry);
  if (!snapshot?.preview) return null;
  snapshot.cacheKey = key;
  snapshot.calendarRevision = String(entry?.calendarRevision || snapshot.calendarRevision || "");
  snapshot.cachedAt = String(entry?.cachedAt || snapshot.cachedAt || "");
  return snapshot;
}

function loadCachedCalendarSnapshot() {
  return loadCachedCalendarSnapshotForContext();
}

function saveCalendarSnapshotCacheForContext(snapshot = currentSnapshot, context = {}) {
  const sanitized = sanitizeWorkspaceSnapshot(snapshot);
  if (!sanitized?.preview) return;
  const key = calendarSnapshotCacheKeyForContext({
    ...context,
    doctorKey: sanitized.session?.doctorKey || context.doctorKey || selectedDoctor()?.key || "",
  });
  if (!key) return;
  const store = loadCalendarSnapshotCacheStore();
  sanitized.cacheKey = key;
  sanitized.calendarRevision = String(sanitized.calendarRevision || currentCalendarRevision || "");
  sanitized.cachedAt = new Date().toISOString();
  store[key] = {
    calendarRevision: sanitized.calendarRevision,
    cachedAt: sanitized.cachedAt,
    snapshot: sanitized,
  };
  saveCalendarSnapshotCacheStore(store);
}

function saveCalendarSnapshotCache(snapshot = currentSnapshot) {
  saveCalendarSnapshotCacheForContext(snapshot);
}

function invalidateCalendarSnapshotCache() {
  try {
    localStorage.removeItem(CALENDAR_SNAPSHOT_CACHE_KEY);
  } catch {
    // Cache invalidation must not block the foreground workflow.
  }
}

function renderCachedCalendarSnapshot(options = {}) {
  return renderCachedCalendarSnapshotForContext(calendarSnapshotContext(options), options);
}

function renderCachedCalendarSnapshotForContext(context = {}, options = {}) {
  const cached = loadCachedCalendarSnapshotForContext(context);
  if (!cached?.preview) return false;
  currentSnapshot = cached;
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = cached.cachedAt || cached.preview?.lastParsed || "";
  currentCalendarRevision = cached.calendarRevision || currentCalendarRevision;
  selectedFiles = importRefsToClientEntries(cached.fileRefs || []);
  renderWorkspaceFromSnapshot(cached, cached.session || {});
  markLoginPhase("cachedCalendarRendered", options.loginStartedAt);
  markAccountSwitchPhase("cachedCalendarRendered", options.accountSwitchStartedAt);
  setStatus("Calendar loaded from cache. Checking for updates...");
  return true;
}

function saveWorkspaceSnapshotForEmail(email, snapshot) {
  if (!email) return;
  const store = loadWorkspaceStore();
  store[email] = {
    fileRefs: Array.isArray(snapshot?.fileRefs) ? snapshot.fileRefs.map(importRefForWorkspace).filter((item) => item.id) : [],
    session: snapshot?.session && typeof snapshot.session === "object" ? snapshot.session : {},
  };
  saveWorkspaceStore(store);
}

function clearWorkspaceStoreEntry(email) {
  if (!email) return;
  const store = loadWorkspaceStore();
  delete store[email];
  saveWorkspaceStore(store);
}

function currentWorkspaceSnapshot() {
  return {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: {
      doctorKey: doctorOptions.length > 1 ? doctorSelect.value : doctorOptions[0]?.key || "",
      settings: { ...settings },
      overrides: cleanOverrides(),
      customEvents: customEventsForActiveCalendar(),
      conflictSelections: { ...conflictSelections },
      exportRange: normalizeSavedExportRange(pendingExportRange),
      hadPreview: Boolean(latestPreview),
      savedAt: new Date().toISOString(),
    },
  };
}

function currentHistorySnapshot() {
  return {
    doctorKey: doctorOptions.length > 1 ? doctorSelect.value : doctorOptions[0]?.key || "",
    settings: { ...settings },
    overrides: cleanOverrides(),
    customEvents: customEventsForActiveCalendar(),
    conflictSelections: { ...conflictSelections },
    hadPreview: Boolean(latestPreview),
  };
}

function historySnapshotSignature(snapshot) {
  return JSON.stringify({
    ...snapshot,
    customEvents: sanitizeCustomEvents(snapshot.customEvents || [], activeCalendarEmail()),
  });
}

function recordHistorySnapshot() {
  if (applyingHistory) return;
  const snapshot = currentHistorySnapshot();
  const signature = historySnapshotSignature(snapshot);
  if (signature === lastHistorySignature) return;
  undoHistory.push(snapshot);
  if (undoHistory.length > 150) undoHistory.shift();
  redoHistory = [];
  lastHistorySignature = signature;
}

function loadCurrentWorkspace() {
  const key = activeWorkspaceOwnerKey();
  if (!key) return null;
  const store = loadWorkspaceStore();
  return store[key] || null;
}

function saveCurrentWorkspace() {
  const key = activeWorkspaceOwnerKey();
  if (!key) return;
  saveWorkspaceSnapshotForEmail(key, currentWorkspaceSnapshot());
}

function loadCurrentSessionState() {
  const workspace = loadCurrentWorkspace();
  return workspace?.session || null;
}

function saveCurrentSessionState() {
  try {
    saveCurrentWorkspace();
    scheduleCloudStateSave();
    recordHistorySnapshot();
  } catch {
    // Ignore persistence failures for session-only state.
  }
}

function applySessionState(session, options = {}) {
  const inheritedSettings = options.inheritedSettings && typeof options.inheritedSettings === "object"
    ? options.inheritedSettings
    : {};
  settings = {
    ...defaultSettings(),
    ...settings,
    ...inheritedSettings,
    ...(session?.settings || {}),
  };
  overrides = sanitizeOverrideState(session?.overrides);
  pendingExportRange = normalizeSavedExportRange(session?.exportRange || defaultExportRangeState());
  customEvents = sanitizeActiveCalendarCustomEvents(session?.customEvents);
  conflictSelections = {
    ...loadConflictSelections(),
    ...(session?.conflictSelections || {}),
  };
}

function sanitizeOverrideState(value) {
  if (!value || typeof value !== "object") return {};
  return JSON.parse(JSON.stringify(value));
}

function activeCalendarEmail() {
  return normalizeEmail(activeCalendarOwnerId());
}

function customEventsForActiveCalendar() {
  return sanitizeActiveCalendarCustomEvents(customEvents);
}

function isCustomPreviewEvent(event) {
  return String(event?.source || "").toLowerCase() === "custom";
}

function ensureEditableCustomEvent(event) {
  if (!isCustomPreviewEvent(event)) return null;
  const existing = customEventsForActiveCalendar().find((entry) => entry.id === event.id);
  if (existing) return existing;
  const restored = previewEventToCustomEvent(event);
  replaceActiveCalendarCustomEvents([
    ...customEventsForActiveCalendar(),
    restored,
  ]);
  saveCurrentSessionState();
  return customEventsForActiveCalendar().find((entry) => entry.id === event.id) || null;
}

function reconcileMaterializedPreviewCustomEvents() {
  if (latestPreview?.customEventsMaterialized !== true) return;
  const restored = (latestPreview.events || [])
    .filter(isCustomPreviewEvent)
    .map((event) => previewEventToCustomEvent(event));
  if (!restored.length) return;
  const before = customEventsForActiveCalendar().length;
  replaceActiveCalendarCustomEvents([
    ...customEventsForActiveCalendar(),
    ...restored,
  ]);
  if (customEventsForActiveCalendar().length !== before) {
    saveCurrentSessionState();
  }
}

function sanitizeActiveCalendarCustomEvents(items) {
  const ownerEmail = activeCalendarEmail();
  return sanitizeCustomEvents(items, ownerEmail).filter((item) => item.ownerEmail === ownerEmail);
}

function removeCustomEventForActiveCalendar(id) {
  const ownerEmail = activeCalendarEmail();
  customEvents = customEvents.filter((item) => !(item.id === id && normalizeEmail(item.ownerEmail) === ownerEmail));
}

function replaceActiveCalendarCustomEvents(items) {
  const ownerEmail = activeCalendarEmail();
  const preserved = customEvents.filter((item) => normalizeEmail(item.ownerEmail) !== ownerEmail);
  customEvents = [
    ...preserved,
    ...sanitizeCustomEvents(items, ownerEmail),
  ];
}

function sanitizeCustomEvents(items, defaultOwnerEmail = "") {
  if (!Array.isArray(items)) return [];
  const fallbackOwnerEmail = normalizeEmail(defaultOwnerEmail);
  const events = items
    .filter((item) => item && item.id && item.title && item.startDate && item.endDate)
    .map((item) => ({
      id: String(item.id),
      ownerEmail: normalizeEmail(item.ownerEmail || fallbackOwnerEmail),
      title: String(item.title),
      startDate: String(item.startDate),
      endDate: String(item.endDate),
      allDay: Boolean(item.allDay),
      startTime: item.allDay ? "" : String(item.startTime || ""),
      endTime: item.allDay ? "" : String(item.endTime || ""),
      location: String(item.location || ""),
      include: item.include !== false,
    }))
    .filter((item) => item.ownerEmail);
  return latestCustomEventsByIdentity(events);
}

function latestCustomEventsById(events) {
  const byId = new Map();
  for (const event of events || []) {
    byId.delete(event.id);
    byId.set(event.id, event);
  }
  return [...byId.values()];
}

function latestCustomEventsByIdentity(events) {
  const byIdentity = new Map();
  for (const event of latestCustomEventsById(events)) {
    const key = [
      normalizeEmail(event.ownerEmail),
      event.title,
      event.startDate,
      event.endDate,
      event.allDay ? "all-day" : `${event.startTime}|${event.endTime}`,
      event.location,
    ].join("|");
    byIdentity.delete(key);
    byIdentity.set(key, event);
  }
  return [...byIdentity.values()];
}

async function applyHistorySnapshot(snapshot) {
  if (!snapshot) return;
  const previousDoctorKey = selectedDoctor()?.key || "";
  applyingHistory = true;
  try {
    settings = {
      ...defaultSettings(),
      ...(snapshot.settings || {}),
    };
    overrides = sanitizeOverrideState(snapshot.overrides);
    replaceActiveCalendarCustomEvents(snapshot.customEvents || []);
    conflictSelections = {
      ...loadConflictSelections(),
      ...(snapshot.conflictSelections || {}),
    };
    saveConflictSelections();
    if (doctorOptions.length > 1 && snapshot.doctorKey && doctorOptions.some((doctor) => doctor.key === snapshot.doctorKey)) {
      doctorSelect.value = snapshot.doctorKey;
    }
    renderSettings();
    syncActionState();
    if (!snapshot.hadPreview) {
      clearPreviewData();
    } else if ((selectedDoctor()?.key || "") !== previousDoctorKey) {
      clearPreviewData();
      await updatePreview({ resetRange: false });
    } else {
      rebuildClientPreview();
    }
  } finally {
    applyingHistory = false;
    lastHistorySignature = historySnapshotSignature(snapshot);
  }
}

async function undoHistoryAction() {
  if (undoHistory.length < 2) return;
  const current = undoHistory.pop();
  const previous = undoHistory[undoHistory.length - 1];
  if (current) redoHistory.push(current);
  await applyHistorySnapshot(previous);
  setStatus("Undid the last calendar change.");
}

async function redoHistoryAction() {
  if (!redoHistory.length) return;
  const snapshot = redoHistory.pop();
  if (!snapshot) return;
  undoHistory.push(snapshot);
  await applyHistorySnapshot(snapshot);
  setStatus("Redid the calendar change.");
}

function isTextEditingTarget(target) {
  if (!(target instanceof Element)) return false;
  return Boolean(target.closest("input, textarea, select, [contenteditable='true']"));
}

function handleHistoryShortcut(event) {
  if (isTextEditingTarget(event.target)) return false;
  const key = event.key.toLowerCase();
  const metaOrCtrl = event.metaKey || event.ctrlKey;
  if (!metaOrCtrl || event.altKey) return false;
  if (key === "z") {
    event.preventDefault();
    if (event.shiftKey) {
      void redoHistoryAction();
    } else {
      void undoHistoryAction();
    }
    return true;
  }
  if (event.ctrlKey && key === "y") {
    event.preventDefault();
    void redoHistoryAction();
    return true;
  }
  return false;
}

const DB_NAME = "roster-converter";
const DB_VERSION = 1;
const IMPORT_STORE = "imports";
const CONFLICT_SELECTIONS_KEY = "roster-conflict-selections";

async function openImportsDb() {
  if (!("indexedDB" in window)) {
    throw new Error("Browser storage is unavailable.");
  }
  return await new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(IMPORT_STORE)) {
        db.createObjectStore(IMPORT_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open import storage."));
  });
}

async function saveStoredImport(entry) {
  const db = await openImportsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readwrite");
    tx.objectStore(IMPORT_STORE).put({
      id: entry.id,
      name: entry.name,
      size: entry.size,
      lastModified: entry.lastModified,
      addedAt: entry.addedAt,
      sourceType: entry.sourceType,
      blob: entry.file,
    });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Could not save import."));
  });
  db.close();
}

async function replaceStoredImports(imports) {
  try {
    const db = await openImportsDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IMPORT_STORE, "readwrite");
      const store = tx.objectStore(IMPORT_STORE);
      for (const entry of imports) {
        if (!entry.file) continue;
        store.put({
          id: entry.id,
          name: entry.name,
          size: entry.size,
          lastModified: entry.lastModified,
          addedAt: entry.addedAt,
          sourceType: entry.sourceType,
          blob: entry.file,
        });
      }
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not replace imports."));
    });
    db.close();
  } catch {
    // Browser storage is optional once cloud state has been restored.
  }
}

async function listStoredImportRecords() {
  if (!("indexedDB" in window)) return [];
  const db = await openImportsDb();
  const records = await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readonly");
    const request = tx.objectStore(IMPORT_STORE).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error || new Error("Could not load imports."));
  });
  db.close();
  return records;
}

function recordsToFiles(records) {
  return records.map((record) => ({
    id: record.id,
    name: record.name,
    size: record.size,
    lastModified: record.lastModified,
    addedAt: record.addedAt,
    sourceType: record.sourceType || "pending",
    file: new File([record.blob], record.name, { type: record.blob?.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", lastModified: record.lastModified }),
  })).sort((left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || left.name.localeCompare(right.name));
}

async function loadStoredImportsByRefs(refs = []) {
  if (!refs.length) return [];
  const records = await listStoredImportRecords();
  const recordMap = new Map(records.map((record) => [record.id, record]));
  return recordsToFiles(
    refs
      .map((ref) => {
        const record = recordMap.get(ref.id);
        return record ? { ...record, addedAt: ref.addedAt || record.addedAt, sourceType: ref.sourceType || record.sourceType } : null;
      })
      .filter(Boolean),
  );
}

function allWorkspaceRefs() {
  const store = loadWorkspaceStore();
  return Object.values(store)
    .flatMap((workspace) => Array.isArray(workspace?.fileRefs) ? workspace.fileRefs : [])
    .map((ref) => ref.id);
}

async function garbageCollectStoredImports() {
  const referenced = new Set(allWorkspaceRefs());
  const records = await listStoredImportRecords();
  const unreferenced = records.filter((record) => !referenced.has(record.id)).map((record) => record.id);
  if (!unreferenced.length) return;
  await deleteStoredImportRecords(unreferenced);
}

async function deleteStoredImportRecords(ids = []) {
  const uniqueIds = [...new Set((ids || []).filter(Boolean))];
  if (!uniqueIds.length) return;
  const db = await openImportsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readwrite");
    const store = tx.objectStore(IMPORT_STORE);
    for (const id of uniqueIds) {
      store.delete(id);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Could not clean stored imports."));
  });
  db.close();
}

function removeImportRefsFromWorkspaceStore(id) {
  if (!id) return;
  const store = loadWorkspaceStore();
  let changed = false;
  for (const [key, workspace] of Object.entries(store)) {
    const refs = Array.isArray(workspace?.fileRefs) ? workspace.fileRefs : [];
    const nextRefs = refs.filter((ref) => ref.id !== id);
    if (nextRefs.length === refs.length) continue;
    store[key] = {
      ...workspace,
      fileRefs: nextRefs,
    };
    changed = true;
  }
  if (changed) saveWorkspaceStore(store);
}

function removeImportRefsFromCurrentSnapshot(id) {
  if (!id || !currentSnapshot) return;
  currentSnapshot = sanitizeWorkspaceSnapshot({
    ...currentSnapshot,
    fileRefs: Array.isArray(currentSnapshot.fileRefs)
      ? currentSnapshot.fileRefs.filter((ref) => ref.id !== id)
      : [],
  });
}

async function removeStoredImport(id) {
  cancelScheduledCloudStateSave();
  selectedFiles = selectedFiles.filter((entry) => entry.id !== id);
  removeImportRefsFromCurrentSnapshot(id);
  removeImportRefsFromWorkspaceStore(id);
  saveCurrentSessionState();
  try {
    await deleteStoredImportRecords([id]);
    await garbageCollectStoredImports();
  } catch {
    // Keep in-memory removal even if persistent storage is unavailable.
  }
  renderFileSurfaces();
  try {
    setStatus("Removing roster file...");
    await saveCloudState({
      ...snapshotCloudSavePayload(),
      imports: selectedFiles.map((entry) => ({ ...entry })),
      removedImportIds: [id],
    });
  } catch (error) {
    setStatus(error.message || "Could not save file removal.", true);
  }
  if (!selectedFiles.length) {
    resetDerivedState({ preserveSession: true });
    setStatus("Add a roster file to begin.");
    return;
  }
  await analyzeFiles();
  scheduleCloudStateSave();
  renderFileSurfaces();
}

function loadConflictSelections() {
  try {
    return JSON.parse(localStorage.getItem(CONFLICT_SELECTIONS_KEY) || "{}");
  } catch {
    return {};
  }
}

function saveConflictSelections() {
  localStorage.setItem(CONFLICT_SELECTIONS_KEY, JSON.stringify(conflictSelections));
}

async function bootstrapImports() {
  try {
    syncAccountsButton();
    if (!selectedFiles.length) {
      if (cloudAvailable) {
        selectedFiles = [];
      } else {
        const workspace = loadCurrentWorkspace();
        selectedFiles = await loadStoredImportsByRefs(workspace?.fileRefs || []);
        restoredSessionState = workspace?.session || restoredSessionState;
        currentSnapshot = sanitizeWorkspaceSnapshot(workspace?.snapshot);
        currentSnapshotStale = false;
      }
    }
    renderFilesList();
    if (currentSnapshot?.preview && currentSnapshot.doctorOptions?.length) {
      const snapshotInvalid = snapshotHasUnresolvablePreviewEvents(currentSnapshot);
      if ((currentSnapshotStale || snapshotInvalid) && selectedFiles.length) {
        setStatus("Refreshing calendar...");
        await ensureSelectedFilesLoaded();
        if (selectedFiles.length) {
          await analyzeFiles();
          scheduleCloudStateSave();
          return;
        }
      }
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      scheduleInsightWarmup();
      if (currentSnapshotStale || snapshotInvalid) {
        setStatus("Refreshing calendar...");
        void refreshSnapshotInBackground();
      } else {
        setStatus("Calendar loaded.");
      }
      return;
    }
    if (cloudAvailable && selectedFiles.length && selectedFiles.some((entry) => !entry.file)) {
      const loadedCalendar = await loadCloudCalendarEvents({ adminTargetEmail: adminViewingEmail ? viewedAccountEmail() : "" });
      if (loadedCalendar) {
        renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
        scheduleInsightWarmup();
        setStatus("Calendar loaded.");
        return;
      }
      renderClaimSection();
      syncActionState();
      setStatus("No D1 calendar events were found for this roster name. Check that the roster files have been indexed for the linked account.", true);
      return;
    }
    if (selectedFiles.length) {
      await analyzeFiles();
    } else {
      renderClaimSection();
      syncActionState();
      setStatus(availableRosterDoctors.length && !currentRosterClaims.length
        ? "Choose your roster name, or upload a roster if your name is not listed."
        : "Add a roster file to begin.");
    }
  } catch (error) {
    selectedFiles = [];
    renderFilesList();
    setStatus("Browser storage is unavailable. You can still import files for this session.", true);
  }
}

function snapshotHasUnresolvablePreviewEvents(snapshot) {
  const events = Array.isArray(snapshot?.preview?.events) ? snapshot.preview.events : [];
  if (!events.length) return false;
  const reviewIds = new Set(Array.isArray(snapshot?.preview?.review) ? snapshot.preview.review.map((item) => item?.id).filter(Boolean) : []);
  if (!reviewIds.size) return true;
  return events.some((event) => event?.id && !reviewIds.has(event.id));
}

function renderWorkspaceFromSnapshot(snapshot, session = {}) {
  currentSnapshot = sanitizeWorkspaceSnapshot(snapshot);
  if (!currentSnapshot) return;
  latestPreview = JSON.parse(JSON.stringify(currentSnapshot.preview));
  doctorOptions = Array.isArray(currentSnapshot.doctorOptions) ? JSON.parse(JSON.stringify(currentSnapshot.doctorOptions)) : [];
  detectedSources = JSON.parse(JSON.stringify(currentSnapshot.detectedSources || {}));
  parsedRosterSources = null;
  doctorRoleIndex = null;
  parsedImportDoctors = new Map();
  clearDoctorAnalysisCache();
  restoredSessionState = session && typeof session === "object" ? session : {};
  applySessionState(restoredSessionState, { inheritedSettings: rosterDefaultSettings() });
  reconcileMaterializedPreviewCustomEvents();
  hydrateInsightCacheFromSnapshot(currentSnapshot);
  pendingPreviewSnapToToday = true;
  renderSettings();
  renderFilesList();
  renderDoctorState();
  indexReviewItems(latestPreview.review || []);
  rebuildClientPreview();
  scheduleInsightWarmup();
  saveCurrentWorkspace();
}

async function ensureSelectedFilesLoaded() {
  if (!selectedFiles.some((entry) => !entry.file)) return true;
  const restored = await loadStoredImportsByRefs(selectedFiles.map(importRefForWorkspace));
  if (restored.length) {
    selectedFiles = restored;
    return true;
  }
  return false;
}

async function refreshSnapshotInBackground() {
  if (snapshotRefreshPromise) return snapshotRefreshPromise;
  snapshotRefreshPromise = (async () => {
    try {
      await analyzeFiles({ resetRange: false, preserveVisiblePreview: true });
      currentSnapshotStale = false;
      currentSnapshotBuiltAt = new Date().toISOString();
      setStatus("Calendar refreshed.");
    } catch (error) {
      setStatus(error.message || "Could not refresh the calendar.", true);
    } finally {
      snapshotRefreshPromise = null;
    }
  })();
  return snapshotRefreshPromise;
}

async function initVersionBar() {
  const el = document.getElementById("versionBar");
  if (!el) return;
  try {
    const res = await fetch("/api/version");
    const data = await res.json();
    const branch = escapeHtml(String(data.branch || ""));
    const commit = escapeHtml(String(data.commit || ""));
    el.dataset.branch = branch;
    el.dataset.commit = commit;
    updateVersionBar();
  } catch {
    const el2 = document.getElementById("versionBar");
    if (el2) el2.textContent = "";
  }
}

function updateVersionBar() {
  const el = document.getElementById("versionBar");
  if (!el) return;
  const branch = el.dataset.branch || "";
  const commit = el.dataset.commit || "";
  const gitPart = branch && commit ? `${branch} · ${commit}` : branch || commit || "";
  let cachePart = "";
  if (lastCacheDiagnostics) {
    if (lastCacheDiagnostics.cacheStoreError) cachePart = ` · store error: ${escapeHtml(lastCacheDiagnostics.cacheStoreError)}`;
    else if (lastCacheDiagnostics.cacheNotChecked === true) cachePart = " · no refresh";
    else if (!lastCacheDiagnostics.cacheEngine) cachePart = "";
    else if (lastCacheDiagnostics.cacheHit === true) cachePart = " · cache hit";
    else if (lastCacheDiagnostics.missReason) cachePart = ` · miss (${escapeHtml(lastCacheDiagnostics.missReason)})`;
    else cachePart = " · miss";
  }
  el.textContent = gitPart + cachePart;
}

async function bootstrapApp() {
  initVersionBar().catch(() => {});
  try {
    renderLoginState();
    if (!currentUserEmail || !currentUserPassword) {
      openLoginModal();
      setStatus("Log in with an email address to load your roster workspace.");
      return;
    }
    await restoreCloudState();
    renderLoginState();
    await bootstrapImports();
  } finally {
    hideLoadingScreen();
  }
}

function setStatus(message, isError = false) {
  const text = String(message || "").trim();
  if (!status || !text) return;

  removeSupersededStatusMessages(text);

  const entry = document.createElement("p");
  entry.className = "status";
  entry.textContent = text;
  entry.dataset.error = isError ? "true" : "false";
  entry.dataset.statusId = String(++statusMessageId);
  status.prepend(entry);
  status.dataset.error = isError ? "true" : "false";

  while (status.children.length > STATUS_MESSAGE_LIMIT) {
    status.lastElementChild?.remove();
  }

  window.setTimeout(() => {
    entry.classList.add("is-fading");
    window.setTimeout(() => {
      entry.remove();
      const newest = status.querySelector(".status");
      if (newest) {
        status.dataset.error = newest.dataset.error === "true" ? "true" : "false";
      } else {
        delete status.dataset.error;
      }
    }, STATUS_MESSAGE_FADE_MS);
  }, STATUS_MESSAGE_LIFETIME_MS);

  if (isError && message) {
    void reportAccountError(text);
  }
  void persistConsoleMessage(text, isError);
}

function removeSupersededStatusMessages(message) {
  const supersededMessages = STATUS_SUPERSEDED_MESSAGES.get(message);
  if (!supersededMessages?.length) return;

  const supersededSet = new Set(supersededMessages);
  status.querySelectorAll(".status").forEach((entry) => {
    if (supersededSet.has(entry.textContent.trim())) {
      entry.remove();
    }
  });
}

async function readJsonResponse(response, fallbackMessage = "Request failed.") {
  const text = await response.text().catch(() => "");
  if (!response.ok) {
    throw new Error(parseError(text, `${fallbackMessage} Server returned ${response.status}.`));
  }
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`${fallbackMessage} The server returned an invalid response.`);
  }
}

function parseError(text, fallbackMessage = "Request failed.") {
  try {
    return JSON.parse(text).error || fallbackMessage;
  } catch {
    if (/Worker exceeded resource limits|Error 1102/i.test(String(text || ""))) {
      return `${fallbackMessage} Cloudflare exceeded CPU or memory while handling this request.`;
    }
    const cleaned = String(text || "")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    return cleaned ? `${fallbackMessage} ${cleaned.slice(0, 220)}` : fallbackMessage;
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

bootstrapApp();
