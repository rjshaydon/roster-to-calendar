import {
  applyEventOverrides,
  buildRosterView,
  customEventsToEvents,
  defaultSettings as rosterDefaultSettings,
  doctorOptions as rosterDoctorOptions,
  findmyshiftProviderStaffOptions,
  applyRosterEventSeniorities,
  exportIcs,
  isIgnoredRosterIssueValue,
  parseUploadForm,
  parserRuleDefaults,
  parserRuleSeniorities,
  previewSummary,
  mergeMembershipDoctors,
  serializeConflict,
  serializeEvent,
  serializeReviewItem,
  setParserExtensions,
  sourceNames,
} from "./roster.js";
import { attachContactAllocations } from "./contact-allocations.js";

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
const rosterImportErrorModal = document.querySelector("#rosterImportErrorModal");
const rosterImportErrorMessage = document.querySelector("#rosterImportErrorMessage");
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
const appDialog = document.querySelector("#appDialog");
const appDialogTitle = document.querySelector("#appDialogTitle");
const appDialogMessage = document.querySelector("#appDialogMessage");
const appDialogCloseButton = document.querySelector("#appDialogCloseButton");
const appDialogCancelButton = document.querySelector("#appDialogCancelButton");
const appDialogConfirmButton = document.querySelector("#appDialogConfirmButton");
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
const shiftCodeReviewModal = document.querySelector("#shiftCodeReviewModal");
const shiftCodeReviewCloseButton = document.querySelector("#shiftCodeReviewCloseButton");
const shiftCodeReviewModalSubtitle = document.querySelector("#shiftCodeReviewModalSubtitle");
const shiftCodeReviewModalBody = document.querySelector("#shiftCodeReviewModalBody");
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
const inviteAccountForm = document.querySelector("#inviteAccountForm");
const invitePassword = document.querySelector("#invitePassword");
const currentDayPreview = document.querySelector("#currentDayPreview");
const exportButton = document.querySelector("#exportButton");
const facilityOverviewButton = document.querySelector("#facilityOverviewButton");
const facilityOverviewSection = document.querySelector("#facilityOverviewSection");
const facilityOverviewHeader = document.querySelector("#facilityOverviewHeader");
const facilityOverviewControls = document.querySelector("#facilityOverviewControls");
const facilityOverviewBody = document.querySelector("#facilityOverviewBody");
const facilityOverviewBackButton = document.querySelector("#facilityOverviewBackButton");
const mobileFacilityOverviewButton = document.querySelector("#mobileFacilityOverviewButton");
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
const switchOverlayCancelButton = document.querySelector("#switchOverlayCancelButton");
const rosterImportOverlay = document.querySelector("#rosterImportOverlay");
const rosterImportOverlayTitle = document.querySelector("#rosterImportOverlayTitle");
const ROSTER_IMPORT_OVERLAY_MAX_MS = 1500;

const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const OWNER_EMAIL = "rhaydon@gmail.com";
const OWNER_DOCTOR_KEY = "RICHARD HAYDON";
const MOBILE_RETURN_TO_CREATOR_VALUE = "__return_to_creator__";
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
// Versioned separately from the snapshot schema. Bumping this discards only
// browser-held rendering caches after a client cache correctness repair; it
// never affects D1 roster data, server snapshots, or subscription feeds.
const CALENDAR_SNAPSHOT_CACHE_KEY = "roster-calendar-snapshot-cache-v3";
const LEGACY_CALENDAR_SNAPSHOT_CACHE_KEYS = ["roster-calendar-snapshot-cache-v1", "roster-calendar-snapshot-cache-v2"];
const MAX_HOT_SNAPSHOT_CACHE_ENTRIES = 3;
const MAX_MEMORY_SNAPSHOT_CACHE_ENTRIES = 160;
const MAX_STORED_SNAPSHOT_CACHE_ENTRIES = 240;
const MAX_STORED_SNAPSHOT_CACHE_AGE_MS = 45 * 24 * 60 * 60 * 1000;
const ROSTER_OVERLAP_DOCTOR_CACHE_KEY = "roster-overlap-doctor-cache-v1";
const FACILITY_OVERVIEW_TAB_PREFERENCES_KEY = "roster-facility-overview-tabs-v1";
const FACILITY_OVERVIEW_SENIORITY_ORDER = ["SMS", "Senior Registrar", "CMO", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "NP", "Physio", "Intern", "Unknown"];
const CURRENT_EMAIL_KEY = "roster-current-email";
const CURRENT_PASSWORD_KEY = "roster-current-password";
const PERSISTENT_PASSWORD_KEY = "roster-persistent-password";
const STAY_LOGGED_IN_PREFERENCE_KEY = "roster-stay-logged-in-preference";
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
let viewedAccountId = normalizeEmail(currentUserEmail);
let viewedAccountType = currentUserRole === "creator" ? "creator" : "claimed-user";
let isImpersonating = false;
let impersonatedByCreator = false;
let returnToCreatorAvailable = false;
let currentDefaultDoctorKey = currentUserEmail === OWNER_EMAIL ? OWNER_DOCTOR_KEY : "";
let currentSnapshotOwnerType = currentUserEmail === OWNER_EMAIL ? "creator-account" : "user-account";
let currentSnapshotOwnerId = normalizeEmail(currentUserEmail);
let activeCalendarContext = initialCalendarContext();
let cloudAvailable = false;
let cloudSaveTimer = 0;
let backgroundCloudSaveTimer = 0;
let pendingCloudSaveSnapshot = null;
let cloudStateSaveQueue = Promise.resolve();
let serverUsers = [];
let currentRosterClaims = [];
let currentSuggestedClaims = [];
let latestNameMatches = [];
let availableRosterDoctors = [];
let calendarSnapshotMemoryCache = new Map();
let calendarImportPollPromise = null;
let calendarImportPollRunId = 0;
let currentSubscription = null;
let currentInsightsEnabled = currentUserRole === "creator";
let currentFacilityOverviewEnabled = currentUserRole === "creator";
let currentNonClinical = false;
let currentDirectorViewEnabled = currentUserRole === "creator";
let facilityOverviewCompactState = {
  latched: false,
  scroller: null,
  lastScrollTop: 0,
  userDirection: 0,
  touchScroller: null,
  touchY: 0,
};
let facilityOverviewState = {
  tab: "on-shift", date: formatDateKey(new Date()), facilityKey: "", includeClinicalSupport: false, requestId: 0, onShiftData: null, contactList: null,
  staffTermStart: formatDateKey(australianTermForDate(new Date()).start), staffTerms: [], staffContent: "", staffData: null, staffQuery: "", staffExpanded: new Set(), staffFocusSection: "", staffActionMenu: null, staffDesignationMenu: null, staffSeniorityMenu: null, staffMultiSelectSection: "", staffMultiSelectMembers: new Map(), staffBulkSeniorityMenu: null, staffMultiSelectSaving: false,
  preferredFacilityKey: "", preferredFacilityReason: "", preferredFacilityEvidenceDate: "", byStreamFrom: formatDateKey(new Date()), byStreamTo: formatDateKey(new Date()), byStreamRows: [], byStreamCatalog: [], byStreamCoverage: [], byStreamContent: "", byStreamData: null, byStreamLoading: false, byStreamMetadataLoading: false, byStreamMetadataKey: "", byStreamMetadataPromise: null, byStreamRequestId: 0, byStreamHideEmptyDates: true, byStreamRowId: 0,
  togetherStaffKeys: ["", ""], togetherRangeMode: "term",
  togetherTermStart: formatDateKey(australianTermForDate(new Date()).start),
  togetherFrom: formatDateKey(australianTermForDate(new Date()).start),
  togetherTo: formatDateKey(addDays(australianTermForDate(new Date()).end, -1)),
  togetherFacilityKey: "ALL", togetherContent: "", togetherHasSearched: false, togetherPinnedDoctors: [], togetherUserClearedAll: false,
};
let whoStaffActionMenu = null;
let whoStaffSeniorityMenu = null;
let whoStaffMenuContext = null;
const FACILITY_OVERVIEW_COMPACT_SCROLL_THRESHOLD = 28;
const FACILITY_OVERVIEW_SCROLL_TOLERANCE = 0;
let facilityOverviewNavigationLocked = false;
let facilityOverviewSessionNeedsInitialization = true;
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
let currentAdminTab = "users";
let adminUserSeniorityFilter = "";
let adminUserSearchQuery = "";
let createUserAccountExpanded = false;
let otherUsersExpanded = false;
let otherUsersExpandedBySearch = false;
const ROSTER_HOSPITAL_SORT_RANK = { mmc: 0, ddh: 1, casey: 2, mch: 3 };
let calendarStoreStatus = null;
let calendarStoreStatusError = "";
let rosterSyncStates = new Map();
const activeManualReparseIds = new Set();
const activeAutomatedSourceRefreshIds = new Set();
const pendingRemovedImportIds = new Set();
let rosterRemovalRetryRunId = 0;
let rosterSyncRefreshTimer = 0;
let lastRosterPersistence = null;
let adminConsoleOpen = false;
let adminConsoleLoading = false;
let adminConsoleMessages = [];
let creatorSwitcherAnnouncementBaseline = null;
let reportedIssueFingerprints = new Set();
let currentSnapshot = null;
let currentSnapshotStale = false;
let currentSnapshotBuiltAt = "";
let currentCalendarRevision = "";
let snapshotRefreshPromise = null;
let switchTargetPrefetchRunId = 0;
let switchTargetPrefetchPromise = null;
let switchOverlayRunId = 0;
let activeSwitchOverlayCancel = null;
let storedSnapshotMaintenanceQueued = false;
let storedSnapshotWritesSinceMaintenance = 0;
let deferredBootstrapRunId = 0;
let deferredBootstrapTimer = 0;
let deferredAccountContextRunId = 0;
let calendarTransitionRunId = 0;
let accountClaimResolutionTransition = null;
let pendingPreviewSnapToToday = false;
let insightWarmupTimer = 0;
let insightWarmupPromise = null;
let visibleInsightWarmCache = new Map();
let visibleInsightWarmKey = "";
let insightsRenderRunId = 0;
let parserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let globalParserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let localParserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
let parserRuleSuggestions = [];
let globalUnresolvedShiftCodes = [];
let globalUnresolvedShiftCodesLoading = false;
let globalUnresolvedShiftCodesLoaded = false;
let globalUnresolvedShiftCodesError = "";
let globalUnresolvedShiftCodeRunId = 0;
let shiftCodeReviewFilter = { query: "", source: "all" };
let previewIssueFocusTimer = 0;
let shiftCodeReviewReturnContext = null;
let pendingUnresolvedIssueFocusDate = "";
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
  shiftCodeReviewModal,
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
  if (!canUploadRosters()) {
    setStatus("Roster imports are managed by the Creator account.");
    return;
  }
  fileInput.click();
});

fileInput.addEventListener("change", async () => {
  if (!canUploadRosters()) {
    fileInput.value = "";
    return;
  }
  await importRosterFiles([...fileInput.files]);
  fileInput.value = "";
});

let rosterDragDepth = 0;
let rosterDragAborted = false;
let rosterImportErrorRunId = 0;
let rosterImportErrorTimer = 0;

function handleRosterDragOver(event) {
  if (!hasFileDrag(event.dataTransfer)) return;
  event.preventDefault();
  if (rosterDragAborted) {
    event.dataTransfer.dropEffect = "none";
    return;
  }
  syncRosterDragState(event.dataTransfer);
}

for (const eventName of ["dragenter", "dragover"]) {
  window.addEventListener(eventName, (event) => {
    if (!hasFileDrag(event.dataTransfer)) return;
    event.preventDefault();
    if (eventName === "dragenter" && !rosterDragAborted) rosterDragDepth += 1;
    handleRosterDragOver(event);
  });
}

window.addEventListener("dragleave", (event) => {
  if (!document.body.classList.contains("is-roster-dragging") && !rosterDragAborted) return;
  event.preventDefault();
  const related = event.relatedTarget;
  if (related && document.documentElement.contains(related)) return;
  clearRosterDragState();
});

window.addEventListener("dragend", clearRosterDragState);
window.addEventListener("blur", clearRosterDragState);
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState !== "visible") clearRosterDragState();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !rosterImportErrorModal?.classList.contains("hidden")) {
    event.preventDefault();
    dismissRosterImportError();
    return;
  }
  if (event.key !== "Escape") return;
  if (!document.body.classList.contains("is-roster-dragging") && !rosterDragAborted) return;
  event.preventDefault();
  abortRosterFileDrag();
}, true);

window.addEventListener("drop", async (event) => {
  if (!hasFileDrag(event.dataTransfer)) return;
  event.preventDefault();
  const wasAborted = rosterDragAborted;
  clearRosterDragState();
  if (wasAborted || !canUploadRosters()) return;
  await importRosterFiles([...event.dataTransfer.files]);
});

window.addEventListener("resize", queueFacilityOverviewMenuPositioning, { passive: true });

rosterImportErrorModal?.addEventListener("click", dismissRosterImportError);

filesButton?.addEventListener("click", openFilesModal);
filesCloseButton?.addEventListener("click", closeFilesModal);
filesModal?.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-files]")) closeFilesModal();
});
exportButton.addEventListener("click", openExportModal);
function toggleFacilityOverview() {
  if (facilityOverviewNavigationLocked) return;
  if (isFacilityOverviewOpen()) closeFacilityOverview();
  else {
    facilityOverviewNavigationLocked = true;
    void openFacilityOverview();
    window.setTimeout(() => { facilityOverviewNavigationLocked = false; }, 350);
  }
}
facilityOverviewButton?.addEventListener("click", toggleFacilityOverview);
mobileFacilityOverviewButton?.addEventListener("click", toggleFacilityOverview);
facilityOverviewBackButton?.addEventListener("click", closeFacilityOverview);
document.addEventListener("pointerdown", (event) => {
  const menu = facilityOverviewState.staffActionMenu;
  const target = event.target;
  if (!(target instanceof Element)) return;
  if (menu) {
    const insideMenu = target.closest("[data-facility-overview-staff-action-menu]");
    const trigger = target.closest("[data-facility-overview-staff-menu]");
    if (!insideMenu && trigger?.dataset.facilityOverviewStaffMenu !== menu.key) closeFacilityOverviewStaffActionMenu();
  }
  const designationMenu = facilityOverviewState.staffDesignationMenu;
  if (designationMenu) {
    const insideMenu = target.closest("[data-facility-overview-staff-designation-action-menu]");
    const trigger = target.closest("[data-facility-overview-staff-designation-menu]");
    if (!insideMenu && trigger?.dataset.facilityOverviewStaffDesignationMenu !== designationMenu.key) closeFacilityOverviewStaffDesignationMenu();
  }
  const seniorityMenu = facilityOverviewState.staffSeniorityMenu;
  if (seniorityMenu) {
    const insideMenu = target.closest("[data-facility-overview-staff-seniority-action-menu]");
    const trigger = target.closest("[data-facility-overview-staff-seniority-menu]");
    if (!insideMenu && trigger?.dataset.facilityOverviewStaffSeniorityMenu !== seniorityMenu.key) closeFacilityOverviewStaffSeniorityMenu();
  }
  const bulkMenu = facilityOverviewState.staffBulkSeniorityMenu;
  if (bulkMenu) {
    const insideMenu = target.closest("[data-facility-overview-staff-bulk-seniority-action-menu]");
    if (!insideMenu) closeFacilityOverviewStaffBulkSeniorityMenu();
  }
  if (whoStaffActionMenu || whoStaffSeniorityMenu) {
    const insideMenu = target.closest(".who-team-person [data-facility-overview-staff-action-menu], .who-team-person [data-facility-overview-staff-seniority-action-menu]");
    const trigger = target.closest("[data-who-staff-menu]");
    if (!insideMenu && !trigger) closeWhoStaffMenu();
  }
}, true);
document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape" || (!facilityOverviewState.staffActionMenu && !facilityOverviewState.staffDesignationMenu && !facilityOverviewState.staffSeniorityMenu && !facilityOverviewState.staffBulkSeniorityMenu && !whoStaffActionMenu && !whoStaffSeniorityMenu)) return;
  event.preventDefault();
  closeFacilityOverviewStaffActionMenu();
  closeFacilityOverviewStaffDesignationMenu();
  closeFacilityOverviewStaffSeniorityMenu();
  closeFacilityOverviewStaffBulkSeniorityMenu();
  closeWhoStaffMenu();
});
document.addEventListener("contextmenu", (event) => {
  if (!(event.target instanceof Element)) return;
  const trigger = event.target.closest("[data-who-staff-menu]");
  if (!trigger) return;
  const inlineContainer = trigger.closest(".event-inline-insight");
  openWhoStaffMenu(event, trigger, inlineContainer
    ? { kind: "inline", container: inlineContainer, date: inlineContainer.dataset.inlineWhoDate || "", source: inlineContainer.dataset.inlineWhoSource || "" }
    : { kind: "insights" });
});
document.addEventListener("click", (event) => {
  if (event.target instanceof Element) handleWhoStaffMenuAction(event);
});

function facilityOverviewIsScroller(element) {
  return element === facilityOverviewBody || Boolean(element?.matches?.(".facility-overview-together"));
}

function facilityOverviewScrollerForTarget(target) {
  if (!(target instanceof Element) || !facilityOverviewBody?.contains(target)) return null;
  return target.closest(".facility-overview-together") || facilityOverviewBody;
}

function facilityOverviewScrollerHasOverflow(scroller) {
  return scroller.scrollHeight > scroller.clientHeight + FACILITY_OVERVIEW_SCROLL_TOLERANCE;
}

function facilityOverviewShouldCompact(scroller, { userDirection = 0, movement = 0 } = {}) {
  return facilityOverviewScrollerHasOverflow(scroller)
    && scroller.scrollTop > FACILITY_OVERVIEW_COMPACT_SCROLL_THRESHOLD
    && (userDirection > 0 || movement > 0);
}

function facilityOverviewShouldReleaseCompact(scroller, userDirection) {
  return userDirection < 0 && scroller.scrollTop <= FACILITY_OVERVIEW_COMPACT_SCROLL_THRESHOLD;
}

function setFacilityOverviewCompactMode(scroller, compact) {
  facilityOverviewCompactState.latched = compact;
  facilityOverviewCompactState.scroller = compact ? scroller : null;
  facilityOverviewCompactState.lastScrollTop = compact ? scroller.scrollTop : 0;
  facilityOverviewSection?.classList.toggle("is-compact", compact);
}

function setFacilityOverviewScrollDirection(scroller, direction) {
  if (!direction || !facilityOverviewIsScroller(scroller)) return;
  facilityOverviewCompactState.userDirection = direction;
  if (
    direction < 0
    && facilityOverviewCompactState.latched
    && facilityOverviewCompactState.scroller === scroller
    && facilityOverviewShouldReleaseCompact(scroller, direction)
  ) {
    setFacilityOverviewCompactMode(scroller, false);
  }
}

function facilityOverviewKeyboardScrollDirection(event) {
  if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey) return 0;
  if (["ArrowUp", "PageUp", "Home"].includes(event.key)) return -1;
  if (["ArrowDown", "PageDown", "End"].includes(event.key)) return 1;
  if (event.key === " " || event.key === "Spacebar") return event.shiftKey ? -1 : 1;
  return 0;
}

facilityOverviewSection?.addEventListener("wheel", (event) => {
  const scroller = facilityOverviewScrollerForTarget(event.target);
  setFacilityOverviewScrollDirection(scroller, Math.sign(event.deltaY));
}, { capture: true, passive: true });

facilityOverviewSection?.addEventListener("touchstart", (event) => {
  const scroller = facilityOverviewScrollerForTarget(event.target);
  const touch = event.touches[0];
  if (!scroller || !touch) return;
  facilityOverviewCompactState.touchScroller = scroller;
  facilityOverviewCompactState.touchY = touch.clientY;
}, { capture: true, passive: true });

facilityOverviewSection?.addEventListener("touchmove", (event) => {
  const scroller = facilityOverviewCompactState.touchScroller;
  const touch = event.touches[0];
  if (!scroller || !touch) return;
  const movement = facilityOverviewCompactState.touchY - touch.clientY;
  if (Math.abs(movement) < 1) return;
  facilityOverviewCompactState.touchY = touch.clientY;
  setFacilityOverviewScrollDirection(scroller, Math.sign(movement));
}, { capture: true, passive: true });

facilityOverviewSection?.addEventListener("touchend", () => {
  facilityOverviewCompactState.touchScroller = null;
  facilityOverviewCompactState.touchY = 0;
}, { capture: true, passive: true });

facilityOverviewSection?.addEventListener("touchcancel", () => {
  facilityOverviewCompactState.touchScroller = null;
  facilityOverviewCompactState.touchY = 0;
}, { capture: true, passive: true });

facilityOverviewSection?.addEventListener("keydown", (event) => {
  if (event.target instanceof Element && event.target.closest("input, select, textarea, button, [contenteditable='true']")) return;
  const scroller = facilityOverviewScrollerForTarget(event.target) || facilityOverviewCompactState.scroller;
  setFacilityOverviewScrollDirection(scroller, facilityOverviewKeyboardScrollDirection(event));
}, { capture: true });

facilityOverviewSection?.addEventListener("scroll", (event) => {
  const scroller = event.target;
  if (!facilityOverviewIsScroller(scroller)) return;
  const previousScroller = facilityOverviewCompactState.scroller;
  const previousTop = previousScroller === scroller ? facilityOverviewCompactState.lastScrollTop : 0;
  const movement = Math.sign(scroller.scrollTop - previousTop);
  facilityOverviewCompactState.scroller = scroller;
  facilityOverviewCompactState.lastScrollTop = scroller.scrollTop;
  if (!facilityOverviewCompactState.latched) {
    if (facilityOverviewShouldCompact(scroller, { userDirection: facilityOverviewCompactState.userDirection, movement })) {
      setFacilityOverviewCompactMode(scroller, true);
    }
    return;
  }
  // A compact header makes this scrollport taller. The browser can therefore
  // clamp scrollTop back to zero even though the user is still scrolling down.
  // Never infer intent from that movement: only wheel, touch, or keyboard input
  // may set an upward direction and release the compact mode.
  if (
    facilityOverviewShouldReleaseCompact(scroller, facilityOverviewCompactState.userDirection)
  ) setFacilityOverviewCompactMode(scroller, false);
}, { capture: true, passive: true });
facilityOverviewSection?.addEventListener("click", (event) => {
  if (event.target.closest("[data-facility-overview-back-to-creator]")) {
    void returnToCreatorCalendar();
    return;
  }
  if (event.target.closest("[data-facility-overview-account]")) {
    void openAccountsSurface({ defaultAdminTab: "users" });
    return;
  }
  if (event.target.closest("[data-facility-overview-logout]")) {
    void logoutCurrentUser();
    return;
  }
  if (event.target.closest("[data-facility-overview-today]")) {
    if (facilityOverviewState.tab === "staff") return;
    if (facilityOverviewState.tab === "by-stream") {
      const today = formatDateKey(new Date());
      void setFacilityOverviewByStreamRange({ from: today, to: today });
      return;
    }
    clearFacilityOverviewStaffMultiSelect({ render: false });
    facilityOverviewState.tab = "on-shift";
    facilityOverviewState.date = formatDateKey(new Date());
    void loadFacilityOverviewOnShift();
    return;
  }
  const tab = event.target.closest("[data-facility-overview-tab]");
  if (tab) {
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = null;
    clearFacilityOverviewStaffMultiSelect({ render: false });
    facilityOverviewState.tab = tab.dataset.facilityOverviewTab || "on-shift";
    resetFacilityOverviewScroll();
    if (facilityOverviewState.tab === "staff") {
      void loadFacilityOverviewStaff();
    } else if (facilityOverviewState.tab === "by-stream") {
      void openFacilityOverviewByStream();
    } else {
      if (facilityOverviewState.tab === "together") initializeFacilityOverviewTogetherState();
      renderFacilityOverview();
    }
    return;
  }
  if (event.target.closest("[data-facility-overview-by-stream-add]")) {
    if (facilityOverviewState.byStreamRows.length < 6) {
      const previous = facilityOverviewState.byStreamRows.at(-1);
      facilityOverviewState.byStreamRows.push(newFacilityOverviewByStreamRow(previous));
      renderFacilityOverview();
    }
    return;
  }
  const removeByStream = event.target.closest("[data-facility-overview-by-stream-remove]");
  if (removeByStream) {
    const id = removeByStream.dataset.facilityOverviewByStreamRemove || "";
    if (facilityOverviewState.byStreamRows.length > 1) {
      facilityOverviewState.byStreamRows = facilityOverviewState.byStreamRows.filter((row) => row.id !== id);
      reconcileFacilityOverviewByStreamDuplicates();
      void loadFacilityOverviewByStream();
    }
    return;
  }
  if (event.target.closest("[data-facility-overview-together-add]")) {
    facilityOverviewState.togetherStaffKeys.push("");
    facilityOverviewState.togetherHasSearched = false;
    facilityOverviewState.togetherContent = "";
    renderFacilityOverview();
    return;
  }
  const removeTogetherStaff = event.target.closest("[data-facility-overview-together-remove]");
  if (removeTogetherStaff) {
    const index = Number(removeTogetherStaff.dataset.facilityOverviewTogetherRemove);
    if (Number.isInteger(index) && index >= 0 && index < facilityOverviewState.togetherStaffKeys.length) {
      facilityOverviewState.togetherStaffKeys.splice(index, 1);
      if (!facilityOverviewState.togetherStaffKeys.length) facilityOverviewState.togetherStaffKeys = [""];
      facilityOverviewState.togetherUserClearedAll = !facilityOverviewState.togetherStaffKeys.some(Boolean);
      facilityOverviewState.togetherHasSearched = false;
      facilityOverviewState.togetherContent = "";
      if (facilityOverviewState.togetherStaffKeys.some(Boolean)) void loadFacilityOverviewTogether();
      else renderFacilityOverview();
    }
    return;
  }
  const dateStep = event.target.closest("[data-facility-overview-date-step]");
  if (dateStep) {
    const step = Number(dateStep.dataset.facilityOverviewDateStep || 0);
    if (facilityOverviewState.tab === "by-stream") {
      const from = addDays(parseDateOnly(facilityOverviewState.byStreamFrom), step);
      const to = addDays(parseDateOnly(facilityOverviewState.byStreamTo), step);
      void setFacilityOverviewByStreamRange({ from: formatDateKey(from), to: formatDateKey(to) });
    } else {
      const date = parseDateOnly(facilityOverviewState.date);
      date.setDate(date.getDate() + step);
      facilityOverviewState.date = formatDateKey(date);
      void loadFacilityOverviewOnShift();
    }
    return;
  }
  const termStep = event.target.closest("[data-facility-overview-staff-term-step]");
  if (termStep) {
    const step = Number(termStep.dataset.facilityOverviewStaffTermStep || 0);
    const current = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart));
    const termNumber = current.termNumber + step;
    const next = termNumber < 1
      ? buildAustralianTerm(current.year - 1, 4, startMonthIndexForTerm(4))
      : termNumber > 4
        ? buildAustralianTerm(current.year + 1, 1, startMonthIndexForTerm(1))
        : buildAustralianTerm(current.year, termNumber, startMonthIndexForTerm(termNumber));
    const value = formatDateKey(next.start);
    if (facilityOverviewState.staffTerms.some((term) => term.value === value)) {
      clearFacilityOverviewStaffMultiSelect({ render: false });
      facilityOverviewState.staffTermStart = value;
      facilityOverviewState.staffExpanded.clear();
      facilityOverviewState.staffFocusSection = "";
      void loadFacilityOverviewStaff();
    }
  }
  const staffCalendar = event.target.closest("[data-facility-overview-open-staff-calendar]");
  if (staffCalendar) {
    void openFacilityOverviewStaffCalendar({
      doctorKey: staffCalendar.dataset.facilityOverviewOpenStaffCalendar || "",
      displayName: staffCalendar.dataset.facilityOverviewStaffDisplayName || "",
      sourceType: staffCalendar.dataset.facilityOverviewStaffSource || "",
    });
    return;
  }
  const multiSelectName = event.target.closest("[data-facility-overview-staff-multi-select-name]");
  if (multiSelectName && facilityOverviewStaffMultiSelectIsActive(multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "")) {
    event.preventDefault();
    toggleFacilityOverviewStaffMultiSelectMember({
      sectionKey: multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "",
      sourceType: multiSelectName.dataset.facilityOverviewStaffSource || "",
      doctorKey: multiSelectName.dataset.facilityOverviewStaffKey || "",
      displayName: multiSelectName.dataset.facilityOverviewStaffDisplayName || "",
    });
    return;
  }
  const mobileDesignationTrigger = event.target.closest("[data-facility-overview-staff-designation-menu], [data-facility-overview-staff-seniority-menu]");
  if (mobileDesignationTrigger && isMobileLayout() && isViewingCreatorAccount()) {
    event.preventDefault();
    const rect = mobileDesignationTrigger.getBoundingClientRect();
    const multiSelectSectionKey = facilityOverviewStaffMultiSelectSectionForControl(mobileDesignationTrigger);
    if (multiSelectSectionKey) {
      openFacilityOverviewStaffBulkSeniorityMenu({ sectionKey: multiSelectSectionKey, x: rect.left, y: rect.bottom });
      return;
    }
    facilityOverviewState.staffActionMenu = null;
    if (mobileDesignationTrigger.matches("[data-facility-overview-staff-seniority-menu]")) {
      facilityOverviewState.staffDesignationMenu = null;
      facilityOverviewState.staffSeniorityMenu = {
        key: mobileDesignationTrigger.dataset.facilityOverviewStaffSeniorityMenu || "",
        x: Math.max(8, Math.round(rect.left)), y: Math.max(8, Math.round(rect.bottom)),
      };
    } else {
      facilityOverviewState.staffSeniorityMenu = null;
      facilityOverviewState.staffDesignationMenu = {
        key: mobileDesignationTrigger.dataset.facilityOverviewStaffDesignationMenu || "",
        x: Math.max(8, Math.round(rect.left)), y: Math.max(8, Math.round(rect.bottom)),
      };
    }
    refreshFacilityOverviewStaffContent();
    renderFacilityOverviewStaffBodyPreservingViewport();
    return;
  }
  const setBulkSeniority = event.target.closest("[data-facility-overview-set-bulk-staff-seniority]");
  if (setBulkSeniority) {
    void setFacilityOverviewStaffSeniorityOverrides({
      seniority: setBulkSeniority.dataset.facilityOverviewSetBulkStaffSeniority || "",
      useRosterSeniority: setBulkSeniority.dataset.facilityOverviewUseRosterSeniority === "true",
    });
    return;
  }
  const setDesignation = event.target.closest("[data-facility-overview-set-staff-designation]");
  if (setDesignation) {
    void setFacilityOverviewStaffDesignation({
      designation: setDesignation.dataset.facilityOverviewSetStaffDesignation || "",
      sourceType: setDesignation.dataset.facilityOverviewStaffSource || "",
      doctorKey: setDesignation.dataset.facilityOverviewStaffKey || "",
      displayName: setDesignation.dataset.facilityOverviewStaffDisplayName || "",
      seniority: setDesignation.dataset.facilityOverviewStaffSeniority || "",
    });
    return;
  }
  const clearDesignation = event.target.closest("[data-facility-overview-clear-staff-designation]");
  if (clearDesignation) {
    void clearFacilityOverviewStaffDesignation(clearDesignation.dataset.facilityOverviewClearStaffDesignation || "");
    return;
  }
  const editSeniority = event.target.closest("[data-facility-overview-edit-staff-seniority]");
  if (editSeniority) {
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = {
      key: editSeniority.dataset.facilityOverviewStaffSeniorityMenu || "",
      x: Math.max(8, Math.round(Number(editSeniority.dataset.facilityOverviewMenuX) || 8)),
      y: Math.max(8, Math.round(Number(editSeniority.dataset.facilityOverviewMenuY) || 8)),
    };
    refreshFacilityOverviewStaffActionContent();
    renderFacilityOverview();
    return;
  }
  const setSeniority = event.target.closest("[data-facility-overview-set-staff-seniority]");
  if (setSeniority) {
    void setFacilityOverviewStaffSeniorityOverride({
      sourceType: setSeniority.dataset.facilityOverviewStaffSource || "",
      doctorKey: setSeniority.dataset.facilityOverviewStaffKey || "",
      displayName: setSeniority.dataset.facilityOverviewStaffDisplayName || "",
      seniority: setSeniority.dataset.facilityOverviewSetStaffSeniority || "",
      useRosterSeniority: setSeniority.dataset.facilityOverviewUseRosterSeniority === "true",
      termStart: setSeniority.dataset.facilityOverviewStaffTermStart || "",
    });
    return;
  }
  const workingTogether = event.target.closest("[data-facility-overview-open-working-together]");
  if (workingTogether) {
    openFacilityOverviewWorkingTogether({
      doctorKey: workingTogether.dataset.facilityOverviewOpenWorkingTogether || "",
      displayName: workingTogether.dataset.facilityOverviewStaffDisplayName || "",
      sourceType: workingTogether.dataset.facilityOverviewStaffSource || "",
    });
    return;
  }
  const staffSeniority = event.target.closest("[data-facility-overview-open-staff-section]");
  if (staffSeniority) {
    void openFacilityOverviewStaffSection({
      sourceType: staffSeniority.dataset.facilityOverviewStaffSource || "",
      seniority: staffSeniority.dataset.facilityOverviewOpenStaffSection || "",
      date: staffSeniority.dataset.facilityOverviewStaffDate || "",
    });
    return;
  }
  const staffSection = event.target.closest("[data-facility-overview-staff-section]");
  if (staffSection) {
    const key = staffSection.dataset.facilityOverviewStaffSection;
    if (facilityOverviewState.staffExpanded.has(key)) {
      facilityOverviewState.staffExpanded.delete(key);
      if (facilityOverviewStaffMultiSelectIsActive(key)) clearFacilityOverviewStaffMultiSelect({ render: false });
    } else facilityOverviewState.staffExpanded.add(key);
    refreshFacilityOverviewStaffContent();
    renderFacilityOverview();
  }
});
facilityOverviewSection?.addEventListener("contextmenu", (event) => {
  const multiSelectName = event.target.closest("[data-facility-overview-staff-multi-select-name]");
  if (multiSelectName && isViewingCreatorAccount() && facilityOverviewStaffMultiSelectIsSelected({
    sectionKey: multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "",
    sourceType: multiSelectName.dataset.facilityOverviewStaffSource || "",
    doctorKey: multiSelectName.dataset.facilityOverviewStaffKey || "",
  })) {
    event.preventDefault();
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = null;
    facilityOverviewState.staffBulkSeniorityMenu = {
      sectionKey: multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "",
      x: Math.max(8, Math.round(event.clientX || 0)),
      y: Math.max(8, Math.round(event.clientY || 0)),
    };
    refreshFacilityOverviewStaffContent();
    renderFacilityOverviewStaffBodyPreservingViewport();
    return;
  }
  const designationOrSeniorityTrigger = event.target.closest("[data-facility-overview-staff-designation-menu], [data-facility-overview-staff-seniority-menu]");
  const multiSelectSectionKey = isViewingCreatorAccount() ? facilityOverviewStaffMultiSelectSectionForControl(designationOrSeniorityTrigger) : "";
  if (multiSelectSectionKey) {
    event.preventDefault();
    openFacilityOverviewStaffBulkSeniorityMenu({ sectionKey: multiSelectSectionKey, x: event.clientX, y: event.clientY });
    return;
  }
  const designationTrigger = event.target.closest("[data-facility-overview-staff-designation-menu]");
  if (designationTrigger && isViewingCreatorAccount()) {
    event.preventDefault();
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = {
      key: designationTrigger.dataset.facilityOverviewStaffDesignationMenu || "",
      x: Math.max(8, Math.round(event.clientX || 0)),
      y: Math.max(8, Math.round(event.clientY || 0)),
    };
    refreshFacilityOverviewStaffActionContent();
    renderFacilityOverview();
    return;
  }
  const seniorityTrigger = event.target.closest("[data-facility-overview-staff-seniority-menu]");
  if (seniorityTrigger && isViewingCreatorAccount()) {
    event.preventDefault();
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = {
      key: seniorityTrigger.dataset.facilityOverviewStaffSeniorityMenu || "",
      x: Math.max(8, Math.round(event.clientX || 0)), y: Math.max(8, Math.round(event.clientY || 0)),
    };
    refreshFacilityOverviewStaffActionContent();
    renderFacilityOverview();
    return;
  }
  const staffMenu = event.target.closest("[data-facility-overview-staff-menu]");
  if (!staffMenu || !isViewingCreatorAccount()) return;
  event.preventDefault();
  facilityOverviewState.staffDesignationMenu = null;
  facilityOverviewState.staffSeniorityMenu = null;
  facilityOverviewState.staffActionMenu = {
    key: staffMenu.dataset.facilityOverviewStaffMenu || "",
    x: Math.max(8, Math.round(event.clientX || 0)),
    y: Math.max(8, Math.round(event.clientY || 0)),
  };
  refreshFacilityOverviewStaffActionContent();
  renderFacilityOverview();
});
facilityOverviewSection?.addEventListener("keydown", (event) => {
  if (event.key !== "ContextMenu" && !(event.shiftKey && event.key === "F10")) return;
  const multiSelectName = event.target.closest?.("[data-facility-overview-staff-multi-select-name]");
  if (multiSelectName && isViewingCreatorAccount() && facilityOverviewStaffMultiSelectIsSelected({
    sectionKey: multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "",
    sourceType: multiSelectName.dataset.facilityOverviewStaffSource || "",
    doctorKey: multiSelectName.dataset.facilityOverviewStaffKey || "",
  })) {
    event.preventDefault();
    const rect = multiSelectName.getBoundingClientRect();
    facilityOverviewState.staffActionMenu = null;
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = null;
    facilityOverviewState.staffBulkSeniorityMenu = {
      sectionKey: multiSelectName.dataset.facilityOverviewStaffMultiSelectSection || "",
      x: Math.max(8, Math.round(rect.left)),
      y: Math.max(8, Math.round(rect.bottom)),
    };
    refreshFacilityOverviewStaffContent();
    renderFacilityOverviewStaffBodyPreservingViewport();
    return;
  }
  const designationOrSeniorityTrigger = event.target.closest?.("[data-facility-overview-staff-designation-menu], [data-facility-overview-staff-seniority-menu]");
  const multiSelectSectionKey = isViewingCreatorAccount() ? facilityOverviewStaffMultiSelectSectionForControl(designationOrSeniorityTrigger) : "";
  if (multiSelectSectionKey) {
    event.preventDefault();
    const rect = designationOrSeniorityTrigger.getBoundingClientRect();
    openFacilityOverviewStaffBulkSeniorityMenu({ sectionKey: multiSelectSectionKey, x: rect.left, y: rect.bottom });
    return;
  }
  const trigger = event.target.closest?.("[data-facility-overview-staff-designation-menu]");
  const seniorityTrigger = event.target.closest?.("[data-facility-overview-staff-seniority-menu]");
  const target = seniorityTrigger || trigger;
  if (!target || !isViewingCreatorAccount()) return;
  event.preventDefault();
  const rect = target.getBoundingClientRect();
  facilityOverviewState.staffActionMenu = null;
  if (seniorityTrigger) {
    facilityOverviewState.staffDesignationMenu = null;
    facilityOverviewState.staffSeniorityMenu = { key: seniorityTrigger.dataset.facilityOverviewStaffSeniorityMenu || "", x: Math.max(8, Math.round(rect.left)), y: Math.max(8, Math.round(rect.bottom)) };
  } else {
    facilityOverviewState.staffSeniorityMenu = null;
    facilityOverviewState.staffDesignationMenu = { key: trigger.dataset.facilityOverviewStaffDesignationMenu || "", x: Math.max(8, Math.round(rect.left)), y: Math.max(8, Math.round(rect.bottom)) };
  }
  refreshFacilityOverviewStaffActionContent();
  renderFacilityOverview();
});
facilityOverviewSection?.addEventListener("change", (event) => {
  const multiSelectToggle = event.target.closest("[data-facility-overview-staff-multi-select]");
  if (multiSelectToggle) {
    const sectionKey = multiSelectToggle.dataset.facilityOverviewStaffMultiSelect || "";
    if (multiSelectToggle.checked) activateFacilityOverviewStaffMultiSelect(sectionKey);
    else clearFacilityOverviewStaffMultiSelect();
    return;
  }
  const togetherStaff = event.target.closest("[data-facility-overview-together-staff]");
  if (togetherStaff) {
    const index = Number(togetherStaff.dataset.facilityOverviewTogetherStaff);
    const value = String(togetherStaff.value || "");
    const duplicate = facilityOverviewState.togetherStaffKeys.some((key, keyIndex) => keyIndex !== index && key === value && value);
    if (duplicate) {
      togetherStaff.value = facilityOverviewState.togetherStaffKeys[index] || "";
      facilityOverviewState.togetherContent = `<article class="issue-card"><p>Choose each staff member only once.</p></article>`;
      facilityOverviewState.togetherHasSearched = false;
      renderFacilityOverview();
      return;
    }
    facilityOverviewState.togetherStaffKeys[index] = value;
    if (value) facilityOverviewState.togetherUserClearedAll = false;
    facilityOverviewState.togetherContent = "";
    facilityOverviewState.togetherHasSearched = false;
    const selectedCount = facilityOverviewState.togetherStaffKeys.filter(Boolean).length;
    if (value && selectedCount >= 1) {
      void loadFacilityOverviewTogether();
    } else {
      renderFacilityOverview();
    }
    return;
  }
  const togetherRangeMode = event.target.closest("[data-facility-overview-together-range-mode]");
  if (togetherRangeMode) {
    facilityOverviewState.togetherRangeMode = togetherRangeMode.value === "dates" ? "dates" : "term";
    refreshFacilityOverviewTogetherAfterFilterChange();
    return;
  }
  const togetherTerm = event.target.closest("[data-facility-overview-together-term]");
  if (togetherTerm) {
    facilityOverviewState.togetherTermStart = String(togetherTerm.value || "").slice(0, 10);
    refreshFacilityOverviewTogetherAfterFilterChange();
    return;
  }
  const togetherDate = event.target.closest("[data-facility-overview-together-date]");
  if (togetherDate) {
    const field = togetherDate.dataset.facilityOverviewTogetherDate;
    if (field === "from") facilityOverviewState.togetherFrom = String(togetherDate.value || "").slice(0, 10);
    if (field === "to") facilityOverviewState.togetherTo = String(togetherDate.value || "").slice(0, 10);
    refreshFacilityOverviewTogetherAfterFilterChange();
    return;
  }
  const togetherFacility = event.target.closest("[data-facility-overview-together-facility]");
  if (togetherFacility) {
    facilityOverviewState.togetherFacilityKey = String(togetherFacility.value || "ALL").toUpperCase();
    refreshFacilityOverviewTogetherAfterFilterChange();
    return;
  }
  const byStreamDate = event.target.closest("[data-facility-overview-by-stream-date]");
  if (byStreamDate) {
    const field = byStreamDate.dataset.facilityOverviewByStreamDate;
    const value = String(byStreamDate.value || "").slice(0, 10);
    void setFacilityOverviewByStreamRange(field === "from" ? { from: value } : { to: value });
    return;
  }
  const byStreamRow = event.target.closest("[data-facility-overview-by-stream-row]");
  if (byStreamRow) {
    const id = byStreamRow.dataset.facilityOverviewByStreamRow || "";
    const field = byStreamRow.dataset.facilityOverviewByStreamField || "";
    if (field === "hide-empty") {
      facilityOverviewState.byStreamHideEmptyDates = byStreamRow.checked;
      if (facilityOverviewState.byStreamData) facilityOverviewState.byStreamContent = facilityOverviewByStreamContentFromData(facilityOverviewState.byStreamData);
      renderFacilityOverview();
      return;
    }
    const row = facilityOverviewState.byStreamRows.find((candidate) => candidate.id === id);
    if (!row) return;
    if (field === "facility") {
      row.facilityKey = String(byStreamRow.value || "").toUpperCase();
      row.streamKey = facilityOverviewPreferredStreamKey(row.facilityKey);
      row.seniority = "ALL";
    } else if (field === "stream") {
      row.streamKey = String(byStreamRow.value || "");
      row.seniority = "ALL";
    } else if (field === "seniority") {
      row.seniority = String(byStreamRow.value || "ALL");
    }
    row.isPrefilled = false;
    reconcileFacilityOverviewByStreamDuplicates(row);
    if (facilityOverviewByStreamRowIsDuplicate(row)) {
      // A duplicate stays editable, but must not let an older request replace
      // the valid result lanes with a now-irrelevant response.
      facilityOverviewState.byStreamRequestId += 1;
      facilityOverviewState.byStreamLoading = false;
      if (facilityOverviewState.byStreamData) facilityOverviewState.byStreamContent = facilityOverviewByStreamContentFromData(facilityOverviewState.byStreamData);
      renderFacilityOverview();
      return;
    }
    void loadFacilityOverviewByStream();
    return;
  }
  const facility = event.target.closest("[data-facility-overview-facility]");
  if (facility) {
    clearFacilityOverviewStaffMultiSelect();
    facilityOverviewState.facilityKey = String(facility.value || "").toUpperCase();
    if (facilityOverviewState.tab === "staff") void loadFacilityOverviewStaff();
    else void loadFacilityOverviewOnShift();
    return;
  }
  const date = event.target.closest("[data-facility-overview-date]");
  if (date) {
    facilityOverviewState.date = String(date.value || "").slice(0, 10);
    void loadFacilityOverviewOnShift();
    return;
  }
  const includeCs = event.target.closest("[data-facility-overview-include-cs]");
  if (includeCs) {
    facilityOverviewState.includeClinicalSupport = includeCs.checked;
    void loadFacilityOverviewOnShift();
  }
  const term = event.target.closest("[data-facility-overview-staff-term]");
  if (term) {
    clearFacilityOverviewStaffMultiSelect();
    facilityOverviewState.staffTermStart = String(term.value || "").slice(0, 10);
    facilityOverviewState.staffExpanded.clear();
    facilityOverviewState.staffFocusSection = "";
    void loadFacilityOverviewStaff();
  }
});
facilityOverviewSection?.addEventListener("input", (event) => {
  const search = event.target.closest("[data-facility-overview-staff-search]");
  if (!search) return;
  facilityOverviewState.staffQuery = String(search.value || "");
  clearFacilityOverviewStaffMultiSelect({ render: false });
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBody();
});
mobileExportButton.addEventListener("click", openMobileExportModal);
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
  await openAccountsSurface({ defaultAdminTab: "users" });
});
mobileAccountButton?.addEventListener("click", async () => {
  await openAccountsSurface({ defaultAdminTab: "users" });
});
mobileAccountAccessButton?.addEventListener("click", async () => {
  await openAccountsSurface({ defaultAdminTab: "users" });
});
accountsCloseButton.addEventListener("click", closeAccountsModal);
accountsModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-accounts]")) closeAccountsModal();
});
insightsCloseButton.addEventListener("click", closeInsightsModal);
insightsModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-insights]")) closeInsightsModal();
});
shiftCodeReviewCloseButton?.addEventListener("click", closeShiftCodeReviewModal);
shiftCodeReviewModal?.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-shift-code-review]")) closeShiftCodeReviewModal();
});
shiftCodeReviewModalBody?.addEventListener("input", (event) => {
  const queryInput = event.target.closest("[data-shift-code-review-search]");
  if (!queryInput) return;
  shiftCodeReviewFilter = { ...shiftCodeReviewFilter, query: String(queryInput.value || "") };
  renderShiftCodeReviewResults();
});
shiftCodeReviewModalBody?.addEventListener("change", (event) => {
  const sourceFilter = event.target.closest("[data-shift-code-review-source]");
  if (!sourceFilter) return;
  shiftCodeReviewFilter = { ...shiftCodeReviewFilter, source: String(sourceFilter.value || "all") };
  renderShiftCodeReviewResults();
});
shiftCodeReviewModalBody?.addEventListener("click", (event) => {
  const goToEventButton = event.target.closest("[data-go-to-unresolved-event]");
  if (goToEventButton) {
    event.preventDefault();
    void openUnresolvedShiftIssueEvent(goToEventButton.dataset.goToUnresolvedEvent || "", {
      doctorKey: goToEventButton.dataset.unresolvedDoctorKey || "",
      displayName: goToEventButton.dataset.unresolvedDisplayName || "",
      date: goToEventButton.dataset.unresolvedDate || "",
      source: goToEventButton.dataset.unresolvedSource || "",
    });
    return;
  }
  const addRosterShiftCodeButton = event.target.closest("[data-add-roster-shift-code]");
  if (addRosterShiftCodeButton) {
    openRosterShiftCodeRuleModal(
      addRosterShiftCodeButton.dataset.addRosterShiftCode || "",
      splitShiftCodeSeniorities(addRosterShiftCodeButton.dataset.shiftCodeSeniorities || ""),
    );
    return;
  }
  const ignoreRosterShiftCodeRuleButton = event.target.closest("[data-ignore-roster-shift-code]");
  if (ignoreRosterShiftCodeRuleButton) {
    openRosterShiftCodeRuleModal(
      ignoreRosterShiftCodeRuleButton.dataset.ignoreRosterShiftCode || "",
      splitShiftCodeSeniorities(ignoreRosterShiftCodeRuleButton.dataset.shiftCodeSeniorities || ""),
      { ignore: true },
    );
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
  const ignoreShiftCodeButton = event.target.closest("[data-ignore-shift-code]");
  if (ignoreShiftCodeButton) {
    openParserRuleModal(
      ignoreShiftCodeButton.dataset.ignoreShiftCode || "",
      ignoreShiftCodeButton.dataset.errorId || "",
      splitShiftCodeSeniorities(ignoreShiftCodeButton.dataset.shiftCodeSeniorities || ""),
      { ignore: true },
    );
  }
});
let appDialogResolve = null;

function closeAppDialog(result = false) {
  if (!appDialog || appDialog.classList.contains("hidden")) return;
  appDialog.classList.add("hidden");
  appDialog.setAttribute("aria-hidden", "true");
  const resolve = appDialogResolve;
  appDialogResolve = null;
  resolve?.(result);
}

function showAppDialog({ title, message, confirmLabel = "" }) {
  if (!appDialog) return Promise.resolve(false);
  appDialogTitle.textContent = title;
  appDialogMessage.textContent = message;
  appDialogConfirmButton.textContent = confirmLabel || "Confirm";
  appDialogConfirmButton.classList.toggle("hidden", !confirmLabel);
  appDialog.classList.remove("hidden");
  appDialog.setAttribute("aria-hidden", "false");
  appDialogCancelButton.focus();
  return new Promise((resolve) => { appDialogResolve = resolve; });
}

appDialogCloseButton?.addEventListener("click", () => closeAppDialog(false));
appDialogCancelButton?.addEventListener("click", () => closeAppDialog(false));
appDialogConfirmButton?.addEventListener("click", () => closeAppDialog(true));
appDialog?.addEventListener("click", (event) => {
  if (event.target.closest("[data-close-app-dialog]")) closeAppDialog(false);
});

accountsBody.addEventListener("submit", (event) => {
  event.preventDefault();
  const createForm = event.target.closest("[data-create-account-form]");
  if (createForm) {
    createAccountFromOwner(createForm);
    return;
  }
  const adminUserForm = event.target.closest("[data-admin-user-form]");
  if (adminUserForm) {
    const email = adminUserForm.dataset.adminUserForm || "";
    const realName = adminUserForm.querySelector("[data-admin-user-real-name]")?.value.trim() || "";
    void saveAdminUserName(email, realName);
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
  const directorHospitalSelect = event.target.closest("[data-director-hospital-preference]");
  if (directorHospitalSelect) {
    updateDirectorHospitalPreference(directorHospitalSelect.value);
    return;
  }
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
  if (insightsToggle) {
    void setUserInsightsEnabled(insightsToggle.dataset.toggleUserInsights || "", insightsToggle.checked);
    return;
  }
  const overviewToggle = event.target.closest("[data-toggle-user-facility-overview]");
  if (overviewToggle) {
    void setUserFacilityOverviewEnabled(overviewToggle.dataset.toggleUserFacilityOverview || "", overviewToggle.checked);
    return;
  }
  const directorToggle = event.target.closest("[data-toggle-user-director-view]");
  if (directorToggle) void setUserDirectorViewEnabled(directorToggle.dataset.toggleUserDirectorView || "", directorToggle.checked);
});
accountsBody.addEventListener("input", (event) => {
  const searchInput = event.target.closest("[data-admin-user-search]");
  if (!searchInput) return;
  adminUserSearchQuery = String(searchInput.value || "");
  if (adminUserSearchQuery && !otherUsersExpanded) {
    otherUsersExpanded = true;
    otherUsersExpandedBySearch = true;
  } else if (!adminUserSearchQuery && otherUsersExpandedBySearch) {
    otherUsersExpanded = false;
    otherUsersExpandedBySearch = false;
  }
  renderAccountsModal();
  const replacement = accountsBody.querySelector("[data-admin-user-search]");
  if (!replacement) return;
  replacement.focus();
  replacement.setSelectionRange(adminUserSearchQuery.length, adminUserSearchQuery.length);
});
accountsBody.addEventListener("toggle", (event) => {
  const section = event.target;
  if (!(section instanceof HTMLDetailsElement)) return;
  if (section.matches("[data-create-user-account-section]")) createUserAccountExpanded = section.open;
  if (section.matches("[data-other-users-section]")) {
    otherUsersExpanded = section.open;
    if (!section.open) otherUsersExpandedBySearch = false;
  }
}, true);
accountsBody.addEventListener("click", (event) => {
  const sendInviteButton = event.target.closest("[data-send-account-invite]");
  if (sendInviteButton) {
    const inviteForm = sendInviteButton.closest("[data-create-account-form]");
    if (inviteForm) void sendAccountInvite(inviteForm);
    return;
  }
  const currentUsersSummary = event.target.closest("[data-other-users-section] > summary");
  const currentUsersControl = event.target.closest(".admin-user-search-filter, .admin-user-seniority-filter");
  if (currentUsersSummary && !currentUsersControl && adminUserSearchQuery.trim()) {
    event.preventDefault();
    adminUserSearchQuery = "";
    otherUsersExpanded = false;
    otherUsersExpandedBySearch = false;
    renderAccountsModal();
    return;
  }
  const adminTab = event.target.closest("[data-admin-tab]");
  if (adminTab) {
    const nextTab = adminTab.dataset.adminTab || "parser";
    if (nextTab === "system" || currentAdminTab === "system") {
      adminConsoleOpen = false;
    }
    currentAdminTab = nextTab;
    renderAccountsModal();
    queueGlobalUnresolvedShiftCodeLoad();
    return;
  }
  if (event.target.closest("[data-open-shift-code-review]")) {
    openShiftCodeReviewModal();
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
    void refreshCalendarStoreStatus({ silent: false, syncSwitcher: true });
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
  const addRosterShiftCodeButton = event.target.closest("[data-add-roster-shift-code]");
  if (addRosterShiftCodeButton) {
    openRosterShiftCodeRuleModal(
      addRosterShiftCodeButton.dataset.addRosterShiftCode || "",
      splitShiftCodeSeniorities(addRosterShiftCodeButton.dataset.shiftCodeSeniorities || ""),
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
  const ignoreRosterShiftCodeRuleButton = event.target.closest("[data-ignore-roster-shift-code]");
  if (ignoreRosterShiftCodeRuleButton) {
    openRosterShiftCodeRuleModal(
      ignoreRosterShiftCodeRuleButton.dataset.ignoreRosterShiftCode || "",
      splitShiftCodeSeniorities(ignoreRosterShiftCodeRuleButton.dataset.shiftCodeSeniorities || ""),
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
  const refreshAutomatedSourceButton = event.target.closest("[data-refresh-automated-source]");
  if (refreshAutomatedSourceButton) {
    if (!canRemoveImports()) return;
    void refreshAutomatedRosterSource(refreshAutomatedSourceButton.dataset.refreshAutomatedSource, {
      from: refreshAutomatedSourceButton.dataset.findmyshiftRangeFrom || "",
      to: refreshAutomatedSourceButton.dataset.findmyshiftRangeTo || "",
    });
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
    setEntranceStatus("Full name on roster, email address, and password are required.", true);
    return;
  }
  await loginWithEmail(email, password, { mode: "create", realName });
});
inviteAccountForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const inviteToken = new URLSearchParams(window.location.search).get("invite") || "";
  const password = invitePassword?.value || "";
  if (!inviteToken || !password) return;
  setEntranceStatus("Activating your account...");
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ action: "acceptInvite", inviteToken, newPassword: password }),
    });
    const data = await readJsonResponse(response, "Could not activate account.");
    window.history.replaceState({}, "", window.location.pathname);
    await loginWithEmail(data.email, password, { mode: "login", stayLoggedIn: true });
  } catch (error) {
    setEntranceStatus(error.message || "Could not activate account.", true);
  }
});
document.querySelectorAll("[data-password-visibility-toggle]").forEach((button) => {
  button.addEventListener("click", () => {
    const input = document.querySelector(`#${button.dataset.passwordVisibilityToggle}`);
    if (!(input instanceof HTMLInputElement)) return;
    const visible = input.type === "text";
    input.type = visible ? "password" : "text";
    button.setAttribute("aria-label", visible ? "Show password" : "Hide password");
    button.setAttribute("aria-pressed", visible ? "false" : "true");
  });
});
function hideVisiblePasswords() {
  document.querySelectorAll("[data-password-visibility-toggle]").forEach((button) => {
    const input = document.querySelector(`#${button.dataset.passwordVisibilityToggle}`);
    if (input instanceof HTMLInputElement && input.type === "text") input.type = "password";
    button.setAttribute("aria-label", "Show password");
    button.setAttribute("aria-pressed", "false");
  });
}

loginTabButton?.addEventListener("click", () => {
  hideVisiblePasswords();
  setEntranceTab("login");
});
createTabButton?.addEventListener("click", () => {
  hideVisiblePasswords();
  setEntranceTab("create");
});
document.addEventListener("visibilitychange", () => {
  if (document.hidden) hideVisiblePasswords();
});
window.addEventListener("blur", hideVisiblePasswords);
window.addEventListener("pagehide", hideVisiblePasswords);
logoutButton.addEventListener("click", () => {
  logoutCurrentUser();
});
backToCreatorButton.addEventListener("click", () => {
  void returnToCreatorCalendar();
});
doctorSelect.addEventListener("change", async () => {
  await switchDoctorSelection(doctorSelect.value, { resetRange: true });
});
doctorSelect.addEventListener("pointerdown", () => queueCreatorSwitchTargetPrefetch());
doctorSelect.addEventListener("focus", () => queueCreatorSwitchTargetPrefetch());
switchOverlayCancelButton?.addEventListener("click", () => {
  const cancel = activeSwitchOverlayCancel;
  if (!cancel || switchOverlayCancelButton.disabled) return;
  switchOverlayCancelButton.disabled = true;
  void Promise.resolve(cancel()).catch((error) => {
    setStatus(error.message || "Could not return to the Creator calendar.", true);
  });
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
  if (mobileDoctorSelect.value === MOBILE_RETURN_TO_CREATOR_VALUE && canReturnToCreator()) {
    closeSettingsPanel();
    try {
      showSwitchOverlay("Returning to creator...", "Restoring the creator calendar.");
      await returnToCreatorCalendar();
    } finally {
      hideSwitchOverlay();
    }
    return;
  }
  await switchDoctorSelection(mobileDoctorSelect.value, { resetRange: true });
});
mobileDoctorSelect?.addEventListener("pointerdown", () => queueCreatorSwitchTargetPrefetch());
mobileDoctorSelect?.addEventListener("focus", () => queueCreatorSwitchTargetPrefetch());
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
  const backToShiftCodesTrigger = event.target.closest("[data-preview-back-to-shift-codes]");
  if (backToShiftCodesTrigger) {
    void returnToShiftCodeReview();
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
  if (event.target.closest("[data-preview-doctor-select]")) {
    queueCreatorSwitchTargetPrefetch();
  }
  const chip = event.target.closest("[data-review-id]");
  if (!chip || event.button !== 0) return;
  if (isMobileLayout()) {
    startPendingPreviewGesture(event, chip);
    return;
  }
  startPreviewGesture(event, chip);
});
preview.addEventListener("focusin", (event) => {
  if (event.target.closest("[data-preview-doctor-select]")) {
    queueCreatorSwitchTargetPrefetch();
  }
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
    return;
  }
  const whenIncludeCsToggle = event.target.closest("[data-insights-when-include-cs]");
  if (whenIncludeCsToggle && insightsState?.mode === "when") {
    insightsState.includeCs = whenIncludeCsToggle.checked;
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
    if (parserRuleModal && !parserRuleModal.classList.contains("hidden")) {
      closeParserRuleModal();
      return;
    }
    if (shiftCodeReviewModal && !shiftCodeReviewModal.classList.contains("hidden")) {
      closeShiftCodeReviewModal();
      return;
    }
    closeReviewModal();
    closeCustomEventModal();
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
    && !event.target.closest("#mobileExportButton")
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
    directorHospitalPreference: "ALL",
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
      needsD1Resync: true,
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
    showRosterImportError("Please drop an Excel or PDF roster file (.xlsx, .xlsm, .xltx, .xltm, or .pdf).");
    return [];
  }
  return files;
}

function showRosterImportError(message) {
  if (!rosterImportErrorModal || !rosterImportErrorMessage) return;
  const runId = ++rosterImportErrorRunId;
  clearTimeout(rosterImportErrorTimer);
  rosterImportErrorMessage.textContent = String(message || "Please drop a valid roster file.").trim();
  rosterImportErrorModal.classList.remove("hidden", "is-closing");
  rosterImportErrorModal.setAttribute("aria-hidden", "false");
  rosterImportErrorTimer = window.setTimeout(() => {
    if (runId === rosterImportErrorRunId) dismissRosterImportError();
  }, 5000);
}

function dismissRosterImportError() {
  if (!rosterImportErrorModal || rosterImportErrorModal.classList.contains("hidden")) return;
  const runId = ++rosterImportErrorRunId;
  clearTimeout(rosterImportErrorTimer);
  rosterImportErrorTimer = 0;
  rosterImportErrorModal.classList.add("is-closing");
  rosterImportErrorModal.setAttribute("aria-hidden", "true");
  window.setTimeout(() => {
    if (runId !== rosterImportErrorRunId) return;
    rosterImportErrorModal.classList.add("hidden");
    rosterImportErrorModal.classList.remove("is-closing");
  }, 180);
}

function hasFileDrag(dataTransfer) {
  // Finder does not consistently expose item names, MIME types, or even the
  // standard `Files` token until a file is dropped.  Treat its file-url token
  // as a file drag too; validation remains at the drop boundary.
  const types = [...(dataTransfer?.types || [])].map((type) => String(type || "").trim().toLowerCase());
  return types.includes("files")
    || types.includes("public.file-url")
    || types.includes("application/x-moz-file");
}

function syncRosterDragState(dataTransfer) {
  if (rosterDragAborted) return;
  // During a drag, browser security deliberately withholds the file's name in
  // some macOS browsers. Show the import affordance for every file drag, then
  // validate Excel/PDF files after the user actually drops them.
  const active = hasFileDrag(dataTransfer);
  document.body.classList.toggle("is-roster-dragging", active);
  rosterDropOverlay.classList.toggle("hidden", !active);
  rosterDropOverlay.setAttribute("aria-hidden", active ? "false" : "true");
}

function abortRosterFileDrag() {
  rosterDragAborted = true;
  rosterDragDepth = 0;
  clearRosterDragVisualState();
}

function clearRosterDragVisualState() {
  document.body.classList.remove("is-roster-dragging");
  rosterDropOverlay.classList.add("hidden");
  rosterDropOverlay.setAttribute("aria-hidden", "true");
}

function clearRosterDragState() {
  rosterDragDepth = 0;
  rosterDragAborted = false;
  clearRosterDragVisualState();
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
    const reason = String(error?.message || "").trim();
    showRosterImportError(`Please drop an Excel or PDF roster file (.xlsx, .xlsm, .xltx, .xltm, or .pdf)${reason ? `: ${reason}` : "."}`);
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
    if (selectedDoctor() && !(isViewingCreatorAccount() && cloudAvailable)) {
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
    shiftCodeReviewModal,
    parserRuleModal,
    reviewModal,
    customEventModal,
  ].some((surface) => surface && !surface.classList.contains("hidden"));
  document.body.classList.toggle("has-active-popup", active);
  setMobileBodyScrollLock(active);
}

async function openAccountsSurface(options = {}) {
  closeSettingsPanel();
  adminConsoleOpen = false;
  if (isViewingCreatorAccount()) {
    if (options.defaultAdminTab) currentAdminTab = options.defaultAdminTab;
  }
  renderAccountsModal();
  accountsModal.classList.remove("hidden");
  accountsModal.setAttribute("aria-hidden", "false");
  focusAccountsModalChrome();
  queueGlobalUnresolvedShiftCodeLoad();
  if (activeCalendarMode() === "doctor-profile" && activeDoctorProfile && cloudAvailable) {
    void refreshDoctorProfileFileRefs().then(() => {
      if (!accountsModal.classList.contains("hidden")) renderAccountsModal();
    }).catch(() => null);
  }
  if (isCreatorAuthenticated()) {
    void loadServerUsers().then(() => {
      if (!accountsModal.classList.contains("hidden")) {
        renderAccountsModal();
        queueGlobalUnresolvedShiftCodeLoad();
      }
    });
    void refreshCalendarStoreStatus({ silent: true, syncSwitcher: false, lightweight: false });
  }
}

function focusAccountsModalChrome() {
  // Safari can focus the first input while a newly-visible modal is being
  // painted. Run after that paint and keep focus on a non-editable control.
  window.requestAnimationFrame(() => {
    if (accountsModal?.classList.contains("hidden")) return;
    const active = document.activeElement;
    if (active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement || active instanceof HTMLSelectElement) {
      active.blur();
    }
    accountsCloseButton?.focus({ preventScroll: true });
  });
}

async function refreshDoctorProfileFileRefs() {
  if (!activeDoctorProfile || !cloudAvailable) return;
  const data = await fetchDoctorProfileState(activeDoctorProfile, {
    cachedRevision: currentCalendarRevision || currentSnapshot?.calendarRevision || "",
    allowInlineBuild: false,
  });
  if (!Array.isArray(data.fileRefs) || !data.fileRefs.length) return;
  selectedFiles = importRefsToClientEntries(data.fileRefs);
  if (currentSnapshot) {
    currentSnapshot.fileRefs = data.fileRefs;
    saveCalendarSnapshotCacheForContext(currentSnapshot, {
      mode: "doctor-profile",
      ownerId: activeDoctorProfile.ownerId,
      doctorKey: activeDoctorProfile.doctorKey,
    });
  }
  renderFileSurfaces();
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
    const returnToCreatorOption = canReturnToCreator()
      ? [{
          key: MOBILE_RETURN_TO_CREATOR_VALUE,
          displayName: `Creator · ${formatRosterDisplayName(OWNER_DOCTOR_KEY)}`,
        }]
      : [];
    const mobilePickerOptions = [...returnToCreatorOption, ...pickerOptions];
    if (mobilePickerOptions.length > 1) {
      mobileDoctorSelect.innerHTML = mobilePickerOptions.map((doctor) => `
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

function normalizeAdminFilesSortOrder(value = "") {
  const normalized = String(value || "").trim().toLowerCase();
  return ADMIN_FILES_SORT_OPTIONS.some((option) => option.value === normalized) ? normalized : "hospital-term";
}

function rosterHospitalSortKey(sourceType = "") {
  const key = String(sourceType || "").trim().toLowerCase();
  const rank = Object.hasOwn(ROSTER_HOSPITAL_SORT_RANK, key) ? ROSTER_HOSPITAL_SORT_RANK[key] : 99;
  return `${String(rank).padStart(2, "0")}:${key}`;
}

function rosterFileTermSortKey(name = "") {
  const value = String(name || "");
  const termMatch = value.match(/term\s*([1-4])\D+(\d{4})/i);
  if (termMatch) {
    return `${termMatch[2]}-${String(Number(termMatch[1])).padStart(2, "0")}`;
  }
  const rangeMatch = value.match(/(\d{2})[-_](\d{2})[-_](\d{4}).*?(?:to|_to_).*?(\d{2})[-_](\d{2})[-_](\d{4})/i);
  if (rangeMatch) {
    const endDate = parseDateOnly(`${rangeMatch[6]}-${rangeMatch[5]}-${rangeMatch[4]}`);
    if (endDate) {
      const term = australianTermForDate(endDate);
      return `${term.year}-${String(term.termNumber).padStart(2, "0")}`;
    }
  }
  return "9999-99";
}

function enrichRosterFileEntry(entry, statusFile = null) {
  if (!entry) return entry;
  const addedAt = entry.addedAt || statusFile?.uploadedAt || statusFile?.addedAt || "";
  const lastModified = Number(entry.lastModified || statusFile?.lastModified || 0);
  return {
    ...statusFile,
    ...entry,
    addedAt,
    lastModified,
    sourceId: entry.sourceId || statusFile?.sourceId || "",
    startDate: entry.startDate || statusFile?.startDate || "",
    coverageEndDate: entry.coverageEndDate || statusFile?.coverageEndDate || "",
    endDate: entry.endDate || statusFile?.endDate || "",
  };
}

function sortRosterFileEntries(files = [], sortOrder = "hospital-term") {
  const order = normalizeAdminFilesSortOrder(sortOrder);
  const sorted = [...files];
  sorted.sort((left, right) => {
    const leftName = String(left?.name || "");
    const rightName = String(right?.name || "");
    const leftHospital = rosterHospitalSortKey(left?.sourceType);
    const rightHospital = rosterHospitalSortKey(right?.sourceType);
    const leftTerm = rosterFileTermSortKey(leftName);
    const rightTerm = rosterFileTermSortKey(rightName);
    const leftAdded = String(left?.addedAt || "");
    const rightAdded = String(right?.addedAt || "");
    const leftModified = Number(left?.lastModified || 0);
    const rightModified = Number(right?.lastModified || 0);
    let compare = 0;
    if (order === "hospital") {
      compare = leftHospital.localeCompare(rightHospital) || leftName.localeCompare(rightName);
    } else if (order === "term") {
      compare = leftTerm.localeCompare(rightTerm) || leftHospital.localeCompare(rightHospital) || leftName.localeCompare(rightName);
    } else if (order === "hospital-term") {
      compare = leftHospital.localeCompare(rightHospital) || leftTerm.localeCompare(rightTerm) || leftName.localeCompare(rightName);
    } else if (order === "alphabet") {
      compare = leftName.localeCompare(rightName, undefined, { sensitivity: "base" });
    } else if (order === "date-added") {
      compare = rightAdded.localeCompare(leftAdded) || leftName.localeCompare(rightName);
    } else if (order === "date-modified") {
      compare = rightModified - leftModified || leftName.localeCompare(rightName);
    }
    return compare;
  });
  return sorted;
}

function renderAdminFilesSortControl(sortOrder = "hospital-term") {
  const selected = normalizeAdminFilesSortOrder(sortOrder);
  return `
    <label class="field admin-files-sort">
      <span>Sort by</span>
      <select data-admin-files-sort>
        ${ADMIN_FILES_SORT_OPTIONS.map((option) => `
          <option value="${escapeHtml(option.value)}" ${option.value === selected ? "selected" : ""}>${escapeHtml(option.label)}</option>
        `).join("")}
      </select>
    </label>
  `;
}

function renderFilesMarkup({ canRemove = false, heading = "", description = "", canAdd = false, sortOrder = "", showSortControl = false, showSourceStatus = false } = {}) {
  const profileView = activeCalendarMode() === "doctor-profile";
  const hasUsableStatus = Boolean(calendarStoreStatus && calendarStoreStatus.unavailable !== true && !calendarStoreStatusError);
  const statusFiles = new Map((calendarStoreStatus?.files || []).map((file) => [file.id, file]));
  const populatedSelectedFileIds = new Set(calendarStoreStatus?.expectedFiles?.populatedFileIds || []);
  const statusOnlyEntries = hasUsableStatus && !profileView
    ? (calendarStoreStatus?.files || [])
      .filter((file) => file?.id)
      .map((file) => ({
        id: file.id,
        repoId: file.id,
        name: file.name,
        sourceType: file.sourceType,
        addedAt: file.uploadedAt || file.addedAt || "",
        lastModified: Number(file.lastModified || 0),
        fromRosterDatabase: true,
      }))
    : calendarFilesForActiveView();
  let displayFiles = rosterDisplayFiles(hasUsableStatus, statusOnlyEntries)
    .filter((entry) => !pendingRemovedImportIds.has(entry.id))
    .map((entry) => enrichRosterFileEntry(entry, statusFiles.get(entry.id)));
  if (sortOrder) {
    displayFiles = sortRosterFileEntries(displayFiles, sortOrder);
  }
  if (!displayFiles.length) {
    const emptyMessage = canRemove
      ? "Add rosters and they will stay here until removed."
      : "No files are currently linked to this calendar.";
    return `
      <article class="review-card">
        ${heading ? `<div class="review-top"><div><strong>${escapeHtml(heading)}</strong>${description ? `<span>${escapeHtml(description)}</span>` : ""}</div>${canAdd ? `<button type="button" class="button button-secondary" data-open-file-picker>Add files</button>` : ""}</div>` : ""}
        ${showSortControl ? renderAdminFilesSortControl(sortOrder) : ""}
        ${showSourceStatus ? renderRosterSourceStatusMarkup() : ""}
        <article class="issue-card"><strong>No files imported yet.</strong><p>${escapeHtml(emptyMessage)}</p></article>
      </article>
    `;
  }
  return `
    <article class="review-card">
      ${heading ? `<div class="review-top"><div><strong>${escapeHtml(heading)}</strong>${description ? `<span>${escapeHtml(description)}</span>` : ""}</div>${canAdd ? `<button type="button" class="button button-secondary" data-open-file-picker>Add files</button>` : ""}</div>` : ""}
      ${showSortControl ? renderAdminFilesSortControl(sortOrder) : ""}
      ${showSourceStatus ? renderRosterSourceStatusMarkup() : ""}
      <div class="file-summary">
        ${displayFiles.map((entry) => `
          <article class="file-pill" data-file-id="${entry.id}">
            <span>${escapeHtml(String(entry.sourceType || "").toUpperCase())}${entry.addedAt ? ` · Imported ${escapeHtml(formatTimestamp(entry.addedAt))}` : " · Roster database"}</span>
            <strong>${escapeHtml(entry.name)}</strong>
            ${Number(entry.lastModified || 0) > 0 ? `<span>Source modified ${escapeHtml(formatTimestamp(entry.lastModified))}</span>` : ""}
            ${rosterSyncLabel(entry) || (statusFiles.has(entry.id)
              ? statusFiles.get(entry.id)?.retainedSourceOnly
                ? `<span>Archived retained source · not used in calendar</span>`
                : renderRosterFileDoctorStatus(statusFiles.get(entry.id))
              : populatedSelectedFileIds.has(entry.id) ? `<span>Saved in D1 · inactive</span>`
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

function renderAdminFilesMarkup({ canRemove = false, canAdd = false } = {}) {
  const profileView = activeCalendarMode() === "doctor-profile";
  const hasUsableStatus = Boolean(calendarStoreStatus && calendarStoreStatus.unavailable !== true && !calendarStoreStatusError);
  const statusFiles = new Map((calendarStoreStatus?.files || []).map((file) => [file.id, file]));
  const populatedSelectedFileIds = new Set(calendarStoreStatus?.expectedFiles?.populatedFileIds || []);
  const statusOnlyEntries = hasUsableStatus && !profileView
    ? (calendarStoreStatus?.files || []).filter((file) => file?.id).map((file) => rosterStoreFileToClientEntry(file))
    : calendarFilesForActiveView();
  const displayFiles = rosterDisplayFiles(hasUsableStatus, statusOnlyEntries)
    .filter((entry) => !pendingRemovedImportIds.has(entry.id))
    .map((entry) => enrichRosterFileEntry(entry, statusFiles.get(entry.id)));
  const sourceStatuses = (calendarStoreStatus?.rosterSourceStatuses || []).filter((source) => source?.mode === "automated");
  const automatedSourceIds = new Set(sourceStatuses.map((source) => source.id));
  const isAutomatedFile = (file) => automatedSourceIds.has(String(file?.sourceId || "")) || String(file?.id || "").startsWith("automation:");
  const terms = adminRosterTerms();
  const manualFiles = displayFiles.filter((file) => !isAutomatedFile(file));
  const classifiedManualFiles = manualFiles.map((file) => ({ file, slot: adminRosterTermSlot(file, terms) }));
  const currentAndNextManualFiles = classifiedManualFiles
    .filter((item) => item.slot === "current" || item.slot === "next")
    .sort((left, right) => adminRosterFileCompare(left.file, right.file, terms));
  const previousManualFiles = classifiedManualFiles
    .filter((item) => item.slot !== "current" && item.slot !== "next")
    .sort((left, right) => adminPreviousRosterFileCompare(left.file, right.file));

  return `
    <article class="review-card admin-files-card">
      <div class="review-top"><div><strong>Files</strong><span>Roster files currently used to generate the creator calendar.</span></div>${canAdd ? `<button type="button" class="button button-secondary" data-open-file-picker>Add files</button>` : ""}</div>
      <section class="admin-file-section" aria-labelledby="admin-auto-sync-heading">
        <h3 id="admin-auto-sync-heading">Auto-sync</h3>
        <div class="admin-file-list admin-auto-sync-list">
          ${sourceStatuses.map((source) => renderAdminAutoSyncRow(source, displayFiles, terms)).join("") || `<article class="issue-card"><p>No automated roster sources are configured.</p></article>`}
        </div>
      </section>
      <section class="admin-file-section" aria-labelledby="admin-manual-imports-heading">
        <h3 id="admin-manual-imports-heading">Manual imports</h3>
        <div class="admin-file-list">
          ${currentAndNextManualFiles.length
            ? currentAndNextManualFiles.map(({ file, slot }) => renderAdminManualFileRow(file, slot, { canRemove, hasUsableStatus, statusFiles, populatedSelectedFileIds })).join("")
            : `<article class="issue-card"><p>No manual roster files for the current or next term.</p></article>`}
        </div>
      </section>
      <section class="admin-file-section" aria-labelledby="admin-previous-imports-heading">
        <h3 id="admin-previous-imports-heading">Previously imported files</h3>
        <div class="admin-file-list">
          ${previousManualFiles.length
            ? previousManualFiles.map(({ file, slot }) => renderAdminManualFileRow(file, slot, { canRemove, hasUsableStatus, statusFiles, populatedSelectedFileIds })).join("")
            : `<article class="issue-card"><p>No earlier manual roster files.</p></article>`}
        </div>
      </section>
    </article>
  `;
}

function adminRosterTerms() {
  const current = australianTermForDate(new Date());
  return { current, next: nextAustralianTerm(current) };
}

function adminRosterTermSlot(file, terms) {
  if (adminRosterFileOverlapsTerm(file, terms.current)) return "current";
  if (adminRosterFileOverlapsTerm(file, terms.next)) return "next";
  const fileTerm = rosterFileTermSortKey(file?.name || "");
  if (fileTerm === adminTermSortKey(terms.current)) return "current";
  if (fileTerm === adminTermSortKey(terms.next)) return "next";
  return "previous";
}

function adminRosterFileOverlapsTerm(file, term) {
  const start = String(file?.startDate || "").slice(0, 10);
  const end = String(file?.coverageEndDate || file?.endDate || start).slice(0, 10);
  if (!start || !end) return false;
  const termStart = formatDateKey(term.start);
  const termEnd = formatDateKey(addDays(term.end, -1));
  return start <= termEnd && end >= termStart;
}

function adminTermSortKey(term) {
  return `${term.year}-${String(term.termNumber).padStart(2, "0")}`;
}

function adminRosterFileCompare(left, right, terms) {
  const leftSlot = adminRosterTermSlot(left, terms);
  const rightSlot = adminRosterTermSlot(right, terms);
  const slotRank = { current: 0, next: 1, previous: 2 };
  return (slotRank[leftSlot] - slotRank[rightSlot])
    || rosterHospitalSortKey(left.sourceType).localeCompare(rosterHospitalSortKey(right.sourceType))
    || String(left.name || "").localeCompare(String(right.name || ""));
}

function adminPreviousRosterFileCompare(left, right) {
  return rosterFileTermSortKey(right.name).localeCompare(rosterFileTermSortKey(left.name))
    || String(right.addedAt || "").localeCompare(String(left.addedAt || ""))
    || String(left.name || "").localeCompare(String(right.name || ""));
}

function renderAdminAutoSyncRow(source, files, terms) {
  const sourceFiles = files.filter((file) => String(file?.sourceId || "") === String(source.id || ""));
  const fallbackNames = new Set((source.activeFileNames || [source.activeFileName]).filter(Boolean));
  const matchingFiles = sourceFiles.length ? sourceFiles : files.filter((file) => fallbackNames.has(file?.name));
  const currentFile = adminLatestTermFile(matchingFiles, terms.current);
  const nextFile = adminLatestTermFile(matchingFiles, terms.next);
  const operationalNote = ["received", "manual-current"].includes(source.state)
    ? ""
    : `<span class="admin-auto-sync-state${source.lastError ? " roster-source-error" : ""}">${escapeHtml(rosterSourceStateLabel(source))}${source.lastError ? ` · ${escapeHtml(source.lastError)}` : ""}</span>`;
  const refreshTitle = source.provider === "findmyshift"
    ? "Check FindMyShift for a newer roster, then reprocess"
    : "Reprocess the retained roster file";
  // This indicator belongs to the refresh the creator clicked. Source-level
  // status can remain stale while a background worker reports its final
  // result, so it must not keep the icon spinning after success or failure.
  const refreshInProgress = activeAutomatedSourceRefreshIds.has(String(source.id || ""));
  return `
    <article class="admin-file-row admin-auto-sync-row roster-source-${escapeHtml(source.state || "unknown")}">
      <div class="admin-auto-sync-heading"><strong>${escapeHtml(source.label)}</strong><button type="button" class="file-reparse file-reparse-visible${refreshInProgress ? " is-processing" : ""}" aria-label="${escapeHtml(refreshTitle)}" title="${escapeHtml(refreshInProgress ? "Refresh in progress" : refreshTitle)}" aria-busy="${refreshInProgress}" data-refresh-automated-source="${escapeHtml(source.id)}"${refreshInProgress ? " disabled" : ""}><span class="file-reparse-icon" aria-hidden="true">↻</span></button></div>
      <dl class="admin-file-details">
        <div><dt>Source modified</dt><dd>${source.providerModifiedAt ? escapeHtml(formatTimestamp(source.providerModifiedAt)) : "Not checked yet"}</dd></div>
        <div><dt>Successfully imported</dt><dd>${source.lastSuccessAt ? escapeHtml(formatTimestamp(source.lastSuccessAt)) : "Not yet imported"}</dd></div>
        ${renderAdminAutoTermDetail("Current term", currentFile)}
        ${nextFile ? renderAdminAutoTermDetail("Next term", nextFile) : ""}
      </dl>
      ${operationalNote}
    </article>
  `;
}

function adminLatestTermFile(files, term) {
  return files.filter((file) => adminRosterFileOverlapsTerm(file, term) || rosterFileTermSortKey(file?.name || "") === adminTermSortKey(term))
    .sort((left, right) => String(right.addedAt || "").localeCompare(String(left.addedAt || "")))[0] || null;
}

function renderAdminAutoTermDetail(label, file) {
  return `<div><dt>${escapeHtml(label)}</dt><dd>${file ? escapeHtml(file.name) : "Not available"}</dd></div>`;
}

function renderAdminManualFileRow(entry, slot, options = {}) {
  const sourceLabel = { mmc: "Monash Adults", mch: "Monash Paediatrics", ddh: "Dandenong", casey: "Casey" }[String(entry.sourceType || "").toLowerCase()] || String(entry.sourceType || "Roster").toUpperCase();
  const termLabel = slot === "current" ? "Current term" : slot === "next" ? "Next term" : adminRosterFileTermLabel(entry);
  const statusFile = options.statusFiles?.get(entry.id);
  const status = rosterSyncLabel(entry) || (statusFile?.retainedSourceOnly
    ? "Archived retained source · not used in calendar"
    : statusFile ? rosterAdminFileStoreStatus(statusFile)
      : options.populatedSelectedFileIds?.has(entry.id) ? "Saved in D1 · inactive"
        : entry.file && options.hasUsableStatus ? "Not yet confirmed in D1"
        : entry.file ? "Roster database status not checked" : "");
  const canReparse = options.canRemove && (!entry.fromRosterDatabase || statusFile?.rawSourceAvailable === true);
  const reparseInProgress = activeManualReparseIds.has(String(entry.id || ""));
  return `
    <article class="admin-file-row admin-manual-file-row" data-file-id="${escapeHtml(entry.id)}">
      <div class="admin-file-row-main">
        <span class="admin-file-kicker">${escapeHtml(sourceLabel)} · ${escapeHtml(termLabel)}</span>
        <strong>${escapeHtml(entry.name)}</strong>
        <span class="admin-file-meta">${entry.addedAt ? `Imported ${escapeHtml(formatTimestamp(entry.addedAt))}` : "Imported roster"}${Number(entry.lastModified || 0) > 0 ? ` · Source modified ${escapeHtml(formatTimestamp(entry.lastModified))}` : ""}</span>
        ${status ? `<span class="admin-file-status">${status}</span>` : ""}
        ${statusFile?.rawSourceAvailable === false ? `<span class="admin-file-status">Source file not retained · re-upload once to enable reparse</span>` : ""}
      </div>
      ${options.canRemove ? `<div class="admin-file-actions">${canReparse ? `<button type="button" class="file-reparse file-reparse-visible${reparseInProgress ? " is-processing" : ""}" aria-label="Reparse roster file" title="${reparseInProgress ? "Reparse in progress" : "Reparse roster file"}" aria-busy="${reparseInProgress}" data-reparse-import="${escapeHtml(entry.id)}"${reparseInProgress ? " disabled" : ""}>↻</button>` : ""}<button type="button" class="file-remove file-remove-visible" aria-label="Remove file" title="Remove file" data-remove-import="${escapeHtml(entry.id)}">🗑</button></div>` : ""}
    </article>
  `;
}

function rosterAdminFileStoreStatus(file) {
  if (file.status === "populated") return "Saved in D1";
  if (file.status === "partial") return "Partially saved in D1";
  if (file.status === "retained") return "Archived retained source · not used in calendar";
  return "No parsed shifts saved in D1";
}

function adminRosterFileTermLabel(file) {
  const start = String(file?.startDate || "").slice(0, 10);
  if (start) return formatAustralianTermLabel(australianTermForDate(parseDateOnly(start)));
  const term = rosterFileTermSortKey(file?.name || "");
  return term === "9999-99" ? "Earlier import" : `Term ${Number(term.slice(5))} ${term.slice(0, 4)}`;
}

function rosterStatusDoctorLabel() {
  return activeDoctorProfile?.displayName
    || selectedDoctor()?.displayName
    || selectedDoctor()?.key
    || "selected doctor";
}

function rosterStatusDoctorKey() {
  if (activeDoctorProfile?.doctorKey) return normalizeRosterName(activeDoctorProfile.doctorKey);
  if (activeCalendarMode() === "claimed-account") {
    return normalizeRosterName(currentDefaultDoctorKey || currentRosterClaims[0]?.key || selectedDoctor()?.key || "");
  }
  return normalizeRosterName(selectedDoctor()?.key || OWNER_DOCTOR_KEY);
}

function renderRosterFileDoctorStatus(file = {}) {
  if (file.selectedDoctorEventCount === null || file.selectedDoctor === null && calendarStoreStatus?.selectedDoctorEventCount === null) {
    return `<span>Checking shifts for ${escapeHtml(rosterStatusDoctorLabel())}…</span>`;
  }
  const doctor = file.selectedDoctor;
  if (!doctor) return `<span>No roster match for ${escapeHtml(rosterStatusDoctorLabel())}</span>`;
  const count = Number(doctor.eventCount || 0);
  if (!count) return `<span>${escapeHtml(doctor.displayName)} · no shifts in this roster</span>`;
  return `
    <span>${escapeHtml(doctor.displayName)} · ${count} shift${count === 1 ? "" : "s"}</span>
    <details class="file-doctor-shifts">
      <summary>Show shifts</summary>
      <ul>${(doctor.shifts || []).map((shift) => `<li>${escapeHtml(formatRosterFileShift(shift))}</li>`).join("")}</ul>
    </details>
  `;
}

function formatRosterFileShift(shift = {}) {
  const date = String(shift.start || "").slice(0, 10);
  const time = shift.allDay ? "All day" : shift.timeLabel || rosterTimestampTime(shift.start, shift.end);
  const location = shift.location ? ` · ${shift.location}` : "";
  return `${formatDate(date)} · ${time} · ${shift.title || "Shift"}${location}`;
}

function rosterTimestampTime(start, end) {
  const startTime = String(start || "").slice(11, 16);
  const endTime = String(end || "").slice(11, 16);
  return startTime && endTime ? `${startTime}–${endTime}` : startTime || endTime || "Time not recorded";
}

function renderRosterSourceStatusMarkup() {
  const sources = calendarStoreStatus?.rosterSourceStatuses || [];
  if (!sources.length) return "";
  return `
    <section class="roster-source-status" aria-label="Roster update status">
      ${sources.map((source) => `
        <article class="roster-source-card roster-source-${escapeHtml(source.state || "unknown")}">
          <strong>${escapeHtml(source.label)}</strong>
          <span>${escapeHtml(rosterSourceStateLabel(source))}</span>
          ${source.providerModifiedAt ? `<small>Source modified ${escapeHtml(formatTimestamp(source.providerModifiedAt))}</small>` : ""}
          ${source.lastSuccessAt ? `<small>Imported ${escapeHtml(formatTimestamp(source.lastSuccessAt))}</small>` : ""}
          ${renderRosterProcessorDispatch(source.processorDispatch)}
          ${renderRosterSourceFileNames(source)}
          ${source.lastError ? `<small class="roster-source-error">${escapeHtml(source.lastError)}</small>` : ""}
        </article>
      `).join("")}
    </section>
  `;
}

function renderRosterSourceFileNames(source = {}) {
  const names = [...new Set((Array.isArray(source.activeFileNames) ? source.activeFileNames : [source.activeFileName])
    .map((name) => String(name || "").trim())
    .filter(Boolean))];
  if (!names.length) return "";
  const label = names.length === 1 ? "Current file" : `${names.length} current files`;
  return `<small>${escapeHtml(label)}: ${names.map(escapeHtml).join(", ")}</small>`;
}

function rosterSourceStateLabel(source = {}) {
  if (source.mode === "manual") return source.state === "manual-current" ? "Manual roster uploaded" : "Manual roster needed";
  if (source.state === "not-configured") return "Not connected";
  if (source.state === "queued") return "Update queued for background processing";
  if (source.state === "processing") return "Update is being imported";
  if (source.state === "failed") return "Latest update failed";
  if (source.state === "received") return "Latest source update imported";
  return "Waiting for first source update";
}

function renderRosterProcessorDispatch(dispatch = null) {
  if (!dispatch?.status) return "";
  const status = String(dispatch.status || "");
  if (status === "accepted") return dispatch.acceptedAt ? `<small>Processor requested ${escapeHtml(formatTimestamp(dispatch.acceptedAt))}</small>` : "<small>Processor requested</small>";
  if (status === "running") return dispatch.startedAt ? `<small>Processor running since ${escapeHtml(formatTimestamp(dispatch.startedAt))}</small>` : "<small>Processor running</small>";
  if (status === "failed") return `<small class="roster-source-error">Processor dispatch failed${dispatch.lastError ? `: ${escapeHtml(dispatch.lastError)}` : ""}</small>`;
  if (status === "requested") return dispatch.requestedAt ? `<small>Processor dispatch requested ${escapeHtml(formatTimestamp(dispatch.requestedAt))}</small>` : "<small>Processor dispatch requested</small>";
  return "";
}

function renderFilesList() {
  if (!filesList) return;
  filesList.innerHTML = renderFilesMarkup({ canRemove: canRemoveImports(), canAdd: canUploadRosters() });
}

function renderFileSurfaces() {
  renderFilesList();
  if (accountsModal && !accountsModal.classList.contains("hidden")) {
    renderAccountsModal();
  }
}

function creatorDoctorSwitcherSignature(options = doctorPickerOptions()) {
  if (!canUseCreatorDoctorSwitcher()) return "";
  return (options || [])
    .map((doctor) => `${doctorIdentityKey(doctor)}:${doctor.displayName}`)
    .sort()
    .join("|");
}

function creatorSwitcherOptionSignature(selectElement) {
  if (!selectElement?.options?.length) return "";
  return [...selectElement.options]
    .map((option) => {
      const displayName = option.dataset.displayName || option.textContent || "";
      return `${doctorIdentityKey({ key: option.value, displayName })}:${displayName}`;
    })
    .sort()
    .join("|");
}

function visibleCreatorSwitcherSignature() {
  if (!canUseCreatorDoctorSwitcher()) return null;
  const pickerOptions = doctorPickerOptions();
  if (isMobileLayout() && mobileDoctorSelect && !mobileDoctorSelect.disabled && mobileDoctorSelect.options.length) {
    return creatorSwitcherOptionSignature(mobileDoctorSelect);
  }
  if (doctorSelect && !doctorSelect.classList.contains("hidden") && doctorSelect.options.length) {
    return creatorSwitcherOptionSignature(doctorSelect);
  }
  if (!pickerOptions.length) return "";
  return null;
}

function captureCreatorSwitcherVisibleBaseline() {
  if (!canUseCreatorDoctorSwitcher()) {
    creatorSwitcherAnnouncementBaseline = null;
    return;
  }
  if (creatorSwitcherAnnouncementBaseline !== null) return;
  creatorSwitcherAnnouncementBaseline = visibleCreatorSwitcherSignature();
}

function isCreatorSwitcherVisibleAndAligned() {
  const visibleSignature = visibleCreatorSwitcherSignature();
  if (visibleSignature === null) return false;
  return visibleSignature === creatorDoctorSwitcherSignature();
}

function isCreatorSwitcherRepositorySettled() {
  if (!isViewingCreatorAccount() || !cloudAvailable) return true;
  if (selectedFilesHavePendingD1Uploads()) return false;
  const activeSync = [...rosterSyncStates.values()].some((state) => (
    ["pending", "parsing", "saving", "uploading-source"].includes(state.status)
  ));
  if (activeSync) return false;
  const storeIds = new Set((calendarStoreStatus?.files || []).map((file) => file.id).filter(Boolean));
  for (const entry of selectedFiles) {
    if (!entry?.id) continue;
    if (!storeIds.has(entry.id)) return false;
  }
  return true;
}

function tryAnnounceCreatorSwitcherRosterUpdate() {
  return new Promise((resolve) => {
    if (!canUseCreatorDoctorSwitcher() || creatorSwitcherAnnouncementBaseline === null) {
      resolve(false);
      return;
    }
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (!isCreatorSwitcherRepositorySettled() || !isCreatorSwitcherVisibleAndAligned()) {
          resolve(false);
          return;
        }
        const baseline = creatorSwitcherAnnouncementBaseline;
        const visibleSignature = visibleCreatorSwitcherSignature();
        if (visibleSignature === null || visibleSignature === baseline) {
          resolve(false);
          return;
        }
        setStatus("Switcher menu updated.");
        creatorSwitcherAnnouncementBaseline = null;
        resolve(true);
      });
    });
  });
}

function remainingSelectedSourceTypes() {
  return new Set(selectedFiles.map((entry) => String(entry?.sourceType || "").toLowerCase()).filter(Boolean));
}

function removedSourceTypesForEntry(removedEntry) {
  const removedType = String(removedEntry?.sourceType || "").toLowerCase();
  if (!removedType) return new Set();
  const remaining = remainingSelectedSourceTypes();
  if (remaining.has(removedType)) return new Set();
  return new Set([removedType]);
}

function pickerHasRemovedSourceDoctors(removedSourceTypes) {
  if (!removedSourceTypes?.size) return false;
  return doctorPickerOptions().some((doctor) => (
    normalizedDoctorSourceTypes(doctor).some((type) => removedSourceTypes.has(type))
  ));
}

function filterSnapshotDoctorsAfterRemoval(snapshot, removedEntry) {
  if (!snapshot?.doctorOptions?.length) return snapshot;
  const removedSourceTypes = removedSourceTypesForEntry(removedEntry);
  if (!removedSourceTypes.size) return snapshot;
  const filtered = snapshot.doctorOptions.filter((doctor) => (
    !normalizedDoctorSourceTypes(doctor).some((type) => removedSourceTypes.has(type))
  ));
  if (filtered.length === snapshot.doctorOptions.length) return snapshot;
  return { ...snapshot, doctorOptions: filtered };
}

function applyAuthoritativeAvailableDoctors(doctors = []) {
  availableRosterDoctors = sanitizeAvailableRosterDoctors(doctors);
}

async function waitForAnimationFrame() {
  await new Promise((resolve) => requestAnimationFrame(resolve));
}

async function isCreatorSwitcherRemovalStable(removedSourceTypes) {
  for (let frame = 0; frame < 2; frame += 1) {
    await waitForAnimationFrame();
    if (!isCreatorSwitcherRepositorySettled()) return false;
    if (!isCreatorSwitcherVisibleAndAligned()) return false;
    if (pickerHasRemovedSourceDoctors(removedSourceTypes)) return false;
  }
  return true;
}

async function waitForCreatorSwitcherRemovalSettled(removedEntry, options = {}) {
  if (!canUseCreatorDoctorSwitcher()) return { settled: true };
  const removedSourceTypes = removedSourceTypesForEntry(removedEntry);
  if (!selectedFiles.length) {
    for (const delay of options.delays || [0, 500, 1500, 4000, 10000]) {
      if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
      await waitForAnimationFrame();
      if (!doctorPickerOptions().length && isCreatorSwitcherRepositorySettled()) {
        return { settled: true };
      }
      availableRosterDoctors = [];
      renderDoctorState();
    }
    throw new Error("Switcher did not refresh after roster removal.");
  }
  if (!removedSourceTypes.size) return { settled: true };
  for (const delay of options.delays || [0, 500, 1500, 4000, 10000]) {
    if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
    try {
      await syncCreatorDoctorPickerWithRemainingRosters({ localOnly: true });
    } catch {
      // Keep polling with the last merged doctor list.
    }
    renderDoctorState();
    if (await isCreatorSwitcherRemovalStable(removedSourceTypes)) {
      return { settled: true };
    }
  }
  throw new Error("Switcher did not refresh after roster removal.");
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
    syncActionState();
    syncMobileChrome();
    return;
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
  const shouldShow = !currentNonClinical && !canUseDoctorPicker() && !doctorOptions.length && availableRosterDoctors.length;
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
    return false;
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
    return true;
  } catch (error) {
    const preserveVisiblePreview = Boolean(latestPreview && isViewingCreatorAccount() && cloudAvailable);
    if (!preserveVisiblePreview) {
      clearPreviewData();
    }
    setStatus(error.message, true);
    return false;
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
  if (!(doctor.key === OWNER_DOCTOR_KEY && canUseCreatorDoctorSwitcher())
    && ![...requestedKeys].some((key) => validDoctors.has(key))) {
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
    .filter((event) => matchesPreviewHospitalFilter(event, "all"));
  visibleEvents.sort(comparePreviewEvents);
  return visibleEvents;
}

function buildResolvedPreviewEvents(baseData) {
  const activeCustomEventIds = new Set(customEventsForActiveCalendar().map((event) => event.id));
  const rosterEvents = (baseData.events || []).filter((event) => !isCustomPreviewEvent(event));
  const visibleCustomEventIds = new Set(
    customEventsToEvents(customEventsForActiveCalendar(), settings, rosterEvents).map((event) => event.id),
  );
  const baseEvents = new Map(
    (baseData.events || [])
      .filter((event) => !(
        baseData.customEventsMaterialized === true
        && isCustomPreviewEvent(event)
        && (!activeCustomEventIds.has(event.id) || !visibleCustomEventIds.has(event.id))
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
    if (!visibleCustomEventIds.has(event.id)) continue;
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
  if (latestPreview && previewMatchesDoctor(latestPreview, doctor)) {
    return latestPreview;
  }
  if (currentSnapshot?.preview && snapshotMatchesDoctor(currentSnapshot, doctor)) {
    latestPreview = JSON.parse(JSON.stringify(currentSnapshot.preview));
    indexReviewItems(latestPreview.review || []);
    return latestPreview;
  }
  if (cloudAvailable && selectedFiles.some((entry) => !entry.file)) {
    const loaded = await loadCloudCalendarEvents({
      adminTargetEmail: adminViewingEmail ? viewedAccountEmail() : "",
      doctorKey: doctor.key,
      preserveExistingSnapshot: true,
    });
    if (loaded && currentSnapshot?.preview && snapshotMatchesDoctor(currentSnapshot, doctor)) {
      latestPreview = JSON.parse(JSON.stringify(currentSnapshot.preview));
      indexReviewItems(latestPreview.review || []);
      return latestPreview;
    }
  }
  if (!selectedFiles.length || selectedFiles.some((entry) => !entry.file)) {
    throw new Error("Calendar data is not loaded yet. Refresh the calendar, then try exporting again.");
  }

  const data = await buildBrowserPreviewData(doctor);
  latestPreview = data;
  indexReviewItems(data.review || []);
  return latestPreview;
}

function previewMatchesDoctor(previewData, doctor) {
  if (!previewData || !doctor?.key) return false;
  if (currentSnapshot?.preview && samePreviewData(previewData, currentSnapshot.preview)) {
    return snapshotMatchesDoctor(currentSnapshot, doctor);
  }
  if (previewData.derivedFromD1) return false;
  return selectedDoctor()?.key === doctor.key;
}

function samePreviewData(left, right) {
  if (!left || !right) return false;
  if (left === right) return true;
  const leftIds = (left.events || []).map((event) => event.id).join("|");
  const rightIds = (right.events || []).map((event) => event.id).join("|");
  return String(left.lastParsed || "") === String(right.lastParsed || "")
    && leftIds === rightIds;
}

function snapshotMatchesDoctor(snapshot, doctor) {
  if (!snapshot?.preview || !doctor?.key) return false;
  const snapshotDoctorKey = normalizeRosterName(snapshot.session?.doctorKey || "");
  if (!snapshotDoctorKey) return selectedDoctor()?.key === doctor.key;
  return doctorOptionMatchesKey(doctor, snapshotDoctorKey);
}

function doctorOptionMatchesKey(doctor, key) {
  const normalizedKey = normalizeRosterName(key);
  if (!normalizedKey || !doctor?.key) return false;
  return [doctor.key, ...(doctor.aliases || []).map((alias) => alias.key)]
    .some((item) => normalizeRosterName(item) === normalizedKey);
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
    // Normal account/profile switches should still open at the current month.
    // A shift-code review jump deliberately keeps its historical roster date in view.
    if (!pendingUnresolvedIssueFocusDate) {
      requestAnimationFrame(() => snapPreviewToCurrentMonth(false));
    }
  }
  queueGlobalUnresolvedShiftCodeLoad();
}

function renderPreviewHeader(doctor, data) {
  return `
    <div class="preview-head">
      ${renderPreviewDoctorControl(doctor)}
      <div class="preview-toolbar">
        ${renderPreviewRangeControls(data.previewStart, data.previewEnd)}
        <span class="preview-event-count">${data.count} events</span>
        ${canReturnToCreator()
          ? `<button type="button" class="button button-secondary preview-back-button" data-preview-back-to-creator>Back to creator</button>`
          : ""}
        ${shiftCodeReviewReturnContext && canReturnToCreator()
          ? `<button type="button" class="button button-secondary preview-back-button" data-preview-back-to-shift-codes>Back to shift codes</button>`
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

function isClinicalSupportEvent(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  return text.includes("clinical support") || /\bcs\b/.test(text) || /\bcso\b/.test(text);
}

function eventTone(event) {
  const text = `${event.title || ""} ${event.rawValue || ""}`.toLowerCase();
  if (text.includes("annual") || text.includes("conference") || text.includes("leave")) return "leave";
  if (text.includes("phnw")) return "phnw";
  if (isClinicalSupportEvent(event)) return "cs";
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
    includeCs: false,
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
    includeCs: false,
    comparisonDoctorKey: normalizedKey,
  };
  await renderInsightsModal();
}

function closeInsightsModal() {
  insightsRenderRunId += 1;
  insightsState = null;
  insightsModal.classList.add("hidden");
  insightsModal.setAttribute("aria-hidden", "true");
  insightsModalBody.innerHTML = "";
}

async function renderInsightsModal() {
  if (!insightsState) return;
  const renderRunId = insightsRenderRunId + 1;
  insightsRenderRunId = renderRunId;
  insightsState.renderRunId = renderRunId;
  if (insightsState.mode === "who") {
    await renderWhoInsight();
  } else if (insightsState.mode === "when") {
    renderWhenInsightLoading();
    showInsightsModal();
    await renderWhenInsight({ renderRunId });
    return;
  }
  showInsightsModal();
}

function showInsightsModal() {
  insightsModal.classList.remove("hidden");
  insightsModal.setAttribute("aria-hidden", "false");
}

function isCurrentInsightRender(renderRunId, mode) {
  return Boolean(insightsState && insightsState.mode === mode && insightsState.renderRunId === renderRunId);
}

function renderWhenInsightLoading() {
  insightsModalTitle.textContent = "When am I working with…?";
  insightsModalSubtitle.textContent = "Find future dates where both doctors are working from the selected date.";
  insightsModalBody.innerHTML = `<article class="issue-card"><p>Loading doctors and shared shifts...</p></article>`;
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
        allowFallback: true,
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

async function fetchRosterInsightRows({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], doctorKeys = [], overlapDoctorKeys = [], allowFallback = true } = {}) {
  if (!cloudAvailable || !startDate) return { ok: false, unavailable: true, rows: [] };
  const cacheKey = `${rosterInsightCacheKey({ startDate, endDate, sourceTypes, excludeDoctorKeys, doctorKeys, overlapDoctorKeys })}|fallback:${allowFallback ? "1" : "0"}`;
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
        allowFallback,
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

async function fetchRosterOverlapDoctors({ startDate, endDate = startDate, sourceTypes = [], excludeDoctorKeys = [], overlapDoctorKeys = [], allowFallback = true } = {}) {
  if (!cloudAvailable || !startDate || !overlapDoctorKeys.length) return { ok: false, unavailable: true, doctors: [] };
  const cacheKey = `${rosterOverlapDoctorCacheKey({ startDate, endDate, sourceTypes, excludeDoctorKeys, overlapDoctorKeys })}|fallback:${allowFallback ? "1" : "0"}`;
  if (visibleInsightWarmCache.has(cacheKey)) {
    return { ok: true, doctors: visibleInsightWarmCache.get(cacheKey), elapsedMs: 0, cached: true };
  }
  const persistentDoctors = readPersistentRosterOverlapDoctors(cacheKey);
  if (persistentDoctors) {
    visibleInsightWarmCache.set(cacheKey, persistentDoctors);
    return { ok: true, doctors: persistentDoctors, elapsedMs: 0, cached: true, persistentCached: true };
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
        allowFallback,
      }),
    });
    const data = await readJsonResponse(response, "Could not load roster overlap doctors.");
    const elapsedMs = Math.round(performance.now() - startedAt);
    if (!data.ok || data.unavailable || !Array.isArray(data.doctors)) return { ok: false, unavailable: true, doctors: [], elapsedMs };
    if (elapsedMs > 1000) console.warn("Roster overlap doctor SQL lookup was slow", { elapsedMs, queryMs: data.queryMs, startDate, endDate, sourceTypes, overlapDoctorKeys });
    visibleInsightWarmCache.set(cacheKey, data.doctors);
    writePersistentRosterOverlapDoctors(cacheKey, data.doctors);
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

function loadPersistentRosterOverlapDoctorCache() {
  try {
    const parsed = JSON.parse(localStorage.getItem(ROSTER_OVERLAP_DOCTOR_CACHE_KEY) || "{}");
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function savePersistentRosterOverlapDoctorCache(store) {
  try {
    const entries = Object.entries(store || {})
      .filter(([, entry]) => Array.isArray(entry?.doctors))
      .sort((left, right) => Number(right[1]?.savedAt || 0) - Number(left[1]?.savedAt || 0))
      .slice(0, 80);
    localStorage.setItem(ROSTER_OVERLAP_DOCTOR_CACHE_KEY, JSON.stringify(Object.fromEntries(entries)));
  } catch {
    // Persistent insight cache must never affect calendar interaction.
  }
}

function readPersistentRosterOverlapDoctors(cacheKey) {
  if (!cacheKey || !currentCalendarRevision) return null;
  const entry = loadPersistentRosterOverlapDoctorCache()[cacheKey];
  if (!entry || entry.revision !== currentCalendarRevision || !Array.isArray(entry.doctors)) return null;
  return entry.doctors;
}

function writePersistentRosterOverlapDoctors(cacheKey, doctors = []) {
  if (!cacheKey || !currentCalendarRevision || !Array.isArray(doctors)) return;
  const store = loadPersistentRosterOverlapDoctorCache();
  store[cacheKey] = {
    revision: currentCalendarRevision,
    savedAt: Date.now(),
    doctors,
  };
  savePersistentRosterOverlapDoctorCache(store);
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

async function renderWhenInsight({ renderRunId = insightsState?.renderRunId } = {}) {
  const hospitalFilters = Array.isArray(insightsState.hospitalFilters) ? insightsState.hospitalFilters : [];
  const includeCs = Boolean(insightsState.includeCs);
  const fromDate = insightsState.fromDate || formatDateKey(new Date());
  const toDate = insightsState.termEnd || currentCalendarInsightDateRange().end || fromDate;
  const doctorResult = await fetchRosterOverlapDoctors({
    startDate: fromDate,
    endDate: toDate,
    sourceTypes: hospitalFilters.map((item) => item.toLowerCase()),
    excludeDoctorKeys: selectedInsightDoctorKeys(),
    overlapDoctorKeys: selectedInsightDoctorKeys(),
  });
  if (!isCurrentInsightRender(renderRunId, "when")) return;
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
      mine: filterWhenInsightEvents(selectedDoctorEventsForInsights(fromDate, toDate, hospitalFilters), includeCs),
      theirs: [],
      fromDate,
      toDate,
      hospitalFilters,
      includeCs,
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
  if (!isCurrentInsightRender(renderRunId, "when")) return;
  if (serverResult.ok) {
    const serverRows = serverResult.rows;
    const selectedComparison = options.find((doctor) => doctor.key === selectedKey) || null;
    const serverEvents = insightRowsToEventsByDoctor(serverRows);
    const mine = filterWhenInsightEvents(selectedDoctorEventsForInsights(fromDate, toDate, hospitalFilters), includeCs);
    const theirs = selectedComparison
      ? filterWhenInsightEvents(serverEvents.get(selectedComparison.key) || [], includeCs)
      : [];
    const hospitalOptions = availableHospitalsFromInsightEvents([...mine, ...[...serverEvents.values()].flat()]);
    renderWhenInsightResult({ options, selectedComparison, mine, theirs, fromDate, toDate, hospitalFilters, includeCs, hospitalOptions });
    return;
  }
  insightsModalTitle.textContent = "When am I working with…?";
  insightsModalSubtitle.textContent = "Find future dates where both doctors are working from the selected date.";
  insightsModalBody.innerHTML = renderRosterInsightUnavailable();
}

function renderWhenInsightResult({ options, selectedComparison, mine, theirs, fromDate, toDate, hospitalFilters, includeCs = false, hospitalOptions = null }) {
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
        <label class="toggle">
          <input type="checkbox" data-insights-when-include-cs ${includeCs ? "checked" : ""}>
          Include CS
        </label>
        <p class="status">Unticked hides Clinical Support overlaps.</p>
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
  // Some provider shifts omit the staff grade but retain it in the title (for
  // example "Physiotherapist"). Use the same recognised-grade detection as
  // At a glance before falling back to stored doctor metadata.
  const eventRole = eventSeniorityRoleCode(event)
    || normalizeWhoRole(facilityOverviewDetectedSeniority(event, event?.seniority || ""));
  const role = eventRole || metadata[source]?.role || metadata.any?.role || inferredWhoRoleForDoctor(doctor) || "";
  const activeRule = parserRuleForWhoEvent(event, source, role);
  const ruleTitle = activeRule ? parserRulePreviewTitle(activeRule) : "";
  const eventForGrouping = ruleTitle ? { ...event, title: ruleTitle } : event;
  const period = activeRule?.period ? whoPeriodLabel({ ...eventForGrouping, rawValue: activeRule.period }) : whoPeriodLabel(eventForGrouping);
  const rawTeam = activeRule?.base ? activeRule.base : whoTeamLabel(eventForGrouping);
  const isNightSsu = period === "Night" && rawTeam === "SSU";
  const isNightIc = isWhoNightIcShift({ event, period, rawTeam, rule: activeRule, ruleTitle });
  const team = ["NP", "Physio"].includes(whoRoleDisplayLabel(role)) ? "Fast Track" : whoDisplayTeamLabel({ period, rawTeam, isNightIc });
  return {
    doctorKey: doctor.key,
    doctorName: doctor.displayName,
    role,
    roleLabel: whoRoleDisplayLabel(role),
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
  for (const assignment of coalesceWhoAssignments(assignments)) {
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

function coalesceWhoAssignments(assignments = []) {
  const byShift = new Map();
  for (const assignment of assignments) {
    const key = [
      normalizeRosterName(assignment.doctorKey || assignment.doctorName),
      assignment.date,
      assignment.source,
      assignment.period,
      assignment.team,
      assignment.timeLabel,
    ].join("|");
    const existing = byShift.get(key);
    if (!existing) {
      byShift.set(key, assignment);
      continue;
    }
    const existingHasRole = Boolean(normalizeWhoRole(existing.role));
    const incomingHasRole = Boolean(normalizeWhoRole(assignment.role));
    if (!existingHasRole && incomingHasRole) byShift.set(key, assignment);
  }
  return [...byShift.values()];
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
  const target = {
    doctorKey: item.doctorKey || "",
    displayName: item.doctorName || item.doctorKey || "Staff member",
    sourceType: String(item.source || "").toLowerCase(),
    seniority: facilityOverviewNormalizeSeniority(item.role || item.event?.seniority || "Unknown"),
    termStart: formatDateKey(australianTermForDate(parseDateOnly(item.date || formatDateKey(new Date()))).start),
  };
  const menuKey = facilityOverviewStaffActionMenuKey(target);
  const actionMenuOpen = isViewingCreatorAccount() && whoStaffActionMenu?.key === menuKey;
  const seniorityMenuOpen = isViewingCreatorAccount() && whoStaffSeniorityMenu?.key === facilityOverviewStaffSeniorityMenuKey(target);
  return `
    <div class="who-team-person">
      <div class="facility-overview-staff-action"><button type="button" class="who-team-name" ${doctorAttribute}="${escapeHtml(item.doctorKey || "")}" data-who-staff-menu="${escapeHtml(menuKey)}" data-facility-overview-staff-source="${escapeHtml(target.sourceType)}" data-facility-overview-staff-key="${escapeHtml(target.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(target.displayName)}" data-facility-overview-staff-seniority="${escapeHtml(target.seniority)}" data-facility-overview-staff-term-start="${escapeHtml(target.termStart)}" title="Show future shifts with ${escapeHtml(item.doctorName)}">${escapeHtml(item.doctorName)}</button>${actionMenuOpen ? renderFacilityOverviewStaffActionMenu(target, whoStaffActionMenu) : ""}${seniorityMenuOpen ? renderFacilityOverviewStaffSeniorityMenu(target, whoStaffSeniorityMenu) : ""}</div>
      <span class="who-team-meta">
        ${item.specialTime ? `<span class="who-team-time">${escapeHtml(item.specialTime)}</span>` : ""}
        ${roleParts.length ? `<span class="who-team-role">${escapeHtml(roleParts.join(" · "))}</span>` : ""}
      </span>
    </div>
  `;
}

function closeWhoStaffMenu() {
  if (!whoStaffActionMenu && !whoStaffSeniorityMenu) return;
  whoStaffActionMenu = null;
  whoStaffSeniorityMenu = null;
  refreshWhoStaffMenuContext();
}

function refreshWhoStaffMenuContext() {
  const context = whoStaffMenuContext;
  if (!context) return;
  if (context.kind === "inline" && context.container?.isConnected) {
    void renderInlineWhoInsight(context.container, context.date, { source: context.source });
    return;
  }
  if (context.kind === "insights") void renderInsightsModal();
}

function openWhoStaffMenu(event, trigger, context) {
  if (!isViewingCreatorAccount()) return;
  event.preventDefault();
  const rect = trigger.getBoundingClientRect();
  const target = {
    doctorKey: trigger.dataset.facilityOverviewStaffKey || "",
    displayName: trigger.dataset.facilityOverviewStaffDisplayName || "",
    sourceType: trigger.dataset.facilityOverviewStaffSource || "",
    seniority: trigger.dataset.facilityOverviewStaffSeniority || "Unknown",
    termStart: trigger.dataset.facilityOverviewStaffTermStart || "",
  };
  whoStaffMenuContext = context;
  whoStaffSeniorityMenu = null;
  whoStaffActionMenu = {
    key: trigger.dataset.whoStaffMenu || facilityOverviewStaffActionMenuKey(target),
    x: Math.max(8, Math.round(event.clientX || rect.left)),
    y: Math.max(8, Math.round(event.clientY || rect.bottom)),
  };
  refreshWhoStaffMenuContext();
}

function handleWhoStaffMenuAction(event) {
  const action = event.target.closest(".who-team-person [data-facility-overview-open-staff-calendar], .who-team-person [data-facility-overview-open-working-together], .who-team-person [data-facility-overview-edit-staff-seniority], .who-team-person [data-facility-overview-set-staff-seniority]");
  if (!action) return false;
  event.preventDefault();
  if (action.matches("[data-facility-overview-open-staff-calendar]")) {
    void openFacilityOverviewStaffCalendar({ doctorKey: action.dataset.facilityOverviewOpenStaffCalendar || "", displayName: action.dataset.facilityOverviewStaffDisplayName || "", sourceType: action.dataset.facilityOverviewStaffSource || "" });
    closeWhoStaffMenu();
    return true;
  }
  if (action.matches("[data-facility-overview-open-working-together]")) {
    openFacilityOverviewWorkingTogether({ doctorKey: action.dataset.facilityOverviewOpenWorkingTogether || "", displayName: action.dataset.facilityOverviewStaffDisplayName || "", sourceType: action.dataset.facilityOverviewStaffSource || "" });
    closeWhoStaffMenu();
    return true;
  }
  if (action.matches("[data-facility-overview-edit-staff-seniority]")) {
    whoStaffActionMenu = null;
    whoStaffSeniorityMenu = {
      key: action.dataset.facilityOverviewStaffSeniorityMenu || "",
      x: Math.max(8, Math.round(Number(action.dataset.facilityOverviewMenuX) || 8)),
      y: Math.max(8, Math.round(Number(action.dataset.facilityOverviewMenuY) || 8)),
    };
    refreshWhoStaffMenuContext();
    return true;
  }
  whoStaffActionMenu = null;
  whoStaffSeniorityMenu = null;
  void setFacilityOverviewStaffSeniorityOverride({
    sourceType: action.dataset.facilityOverviewStaffSource || "",
    doctorKey: action.dataset.facilityOverviewStaffKey || "",
    displayName: action.dataset.facilityOverviewStaffDisplayName || "",
    seniority: action.dataset.facilityOverviewSetStaffSeniority || "",
    useRosterSeniority: action.dataset.facilityOverviewUseRosterSeniority === "true",
    termStart: action.dataset.facilityOverviewStaffTermStart || "",
  }).then(() => refreshWhoStaffMenuContext());
  return true;
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
    const mine = filterWhenInsightEvents(selectedDoctorEventsForInsights(fromDate, toDate, []), false);
    const theirs = selectedComparison
      ? filterWhenInsightEvents(serverEvents.get(selectedComparison.key) || [], false)
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
  if (text.includes("physiotherapist") || /\bphysio\b/.test(text) || text.includes("nurse practitioner") || /\b(?:enp|np|d1)\b/.test(text)) return "Fast Track";
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
    NP: 7,
    Physio: 8,
  };
  return Object.prototype.hasOwnProperty.call(ranks, normalized) ? ranks[normalized] : 99;
}

function normalizeWhoRole(role) {
  const upper = String(role || "").trim().toUpperCase();
  if (!upper) return "";
  if (upper === "UNKNOWN") return "";
  if (upper === "SMS" || upper.includes("SENIOR MEDICAL STAFF") || upper.includes("CONSULTANT") || upper.includes("STAFF SPECIALIST")) return "SMS";
  if (upper === "CMO" || upper.includes("CMO")) return "CMO";
  if (upper === "SR" || upper.includes("SENIOR REGISTRAR") || upper.includes("SENIOR REG")) return "SR";
  if (upper === "IR" || upper === "TR" || upper.includes("TRANSITIONAL") || upper.includes("INTERMEDIATE")) return "IR";
  if (upper === "JR" || upper.includes("JUNIOR REGISTRAR") || upper.includes("JUNIOR REG")) return "JR";
  if (upper === "H" || upper === "HMO" || upper.includes("HMO")) return "HMO";
  if (upper === "I" || upper.includes("INTERN")) return "I";
  if (upper === "ENP" || upper === "NP" || upper.includes("NURSE PRACTITIONER")) return "NP";
  if (upper === "AMP" || upper === "PHYSIO" || upper.includes("PHYSIOTHERAPIST")) return "Physio";
  return upper;
}

function eventSeniorityRoleCode(event) {
  return normalizeWhoRole(event?.seniority || event?.role || "");
}

function inferredWhoRoleForDoctor(doctor) {
  const key = normalizeRosterName(doctor?.key || "");
  if (key && key === OWNER_DOCTOR_KEY) return "SMS";
  return "";
}

function whoRoleDisplayLabel(role) {
  const normalized = normalizeWhoRole(role);
  if (normalized === "I") return "Intern";
  return normalized;
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

function filterWhenInsightEvents(events, includeCs = false) {
  return events.filter(isRosterShiftEvent).filter((event) => includeCs || !isClinicalSupportEvent(event));
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
  await fetchRosterOverlapDoctors({
    startDate,
    endDate,
    excludeDoctorKeys: selectedKeys,
    overlapDoctorKeys: selectedKeys,
    allowFallback: true,
  });
}

function syncActionState() {
  syncControlBarVisibility();
  const ready = canUseExportControls();
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
  if (activeCalendarMode() === "claimed-account") return normalizeRosterName(currentDefaultDoctorKey || currentRosterClaims[0]?.key || "");
  if (currentUserEmail === OWNER_EMAIL && !adminViewingEmail && !activeDoctorProfile) return OWNER_DOCTOR_KEY;
  return "";
}

function preferredDoctorKeyForAccountEmail(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail) return "";
  if (targetEmail === OWNER_EMAIL) return OWNER_DOCTOR_KEY;
  const serverUser = serverUsers.map(normalizeServerUser).find((user) => user.email === targetEmail);
  if (serverUser?.defaultDoctorKey) return normalizeRosterName(serverUser.defaultDoctorKey);
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
  return Boolean(isCreatorAuthenticated() && (canUseDoctorPicker() || activeCalendarMode() === "doctor-profile"));
}

function canReturnToCreator() {
  return Boolean(isCreatorAuthenticated() && (returnToCreatorAvailable || adminViewingEmail || activeDoctorProfile));
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

function showSwitchOverlay(title, message, options = {}) {
  const runId = ++switchOverlayRunId;
  activeSwitchOverlayCancel = typeof options.onCancel === "function" ? options.onCancel : null;
  switchOverlayTitle.textContent = title || "Switching…";
  switchOverlayMessage.textContent = message || "Loading calendar…";
  if (switchOverlayCancelButton) {
    switchOverlayCancelButton.classList.toggle("hidden", !activeSwitchOverlayCancel);
    switchOverlayCancelButton.disabled = false;
  }
  if (!switchOverlay) return runId;
  switchOverlay.classList.remove("hidden");
  switchOverlay.setAttribute("aria-hidden", "false");
  return runId;
}

function hideSwitchOverlay(runId = null) {
  if (runId !== null && runId !== switchOverlayRunId) return;
  activeSwitchOverlayCancel = null;
  if (switchOverlayCancelButton) {
    switchOverlayCancelButton.classList.add("hidden");
    switchOverlayCancelButton.disabled = false;
  }
  if (!switchOverlay) return;
  switchOverlay.classList.add("hidden");
  switchOverlay.setAttribute("aria-hidden", "true");
}

async function cancelCreatorCalendarSwitch() {
  const runId = showSwitchOverlay("Returning to Creator…", "Restoring the Creator calendar.");
  try {
    await returnToCreatorCalendar({ skipOutgoingSave: true, restoreOnFailure: false });
    setStatus("Returned to the Creator calendar.");
  } finally {
    hideSwitchOverlay(runId);
  }
}

function showRosterImportOverlay(fileCount = 1) {
  if (!rosterImportOverlay) return;
  if (rosterImportOverlayTitle) {
    rosterImportOverlayTitle.textContent = fileCount === 1 ? "Adding Roster File" : "Adding Roster Files";
  }
  rosterImportOverlay.classList.remove("hidden");
  rosterImportOverlay.setAttribute("aria-hidden", "false");
}

function hideRosterImportOverlay() {
  if (!rosterImportOverlay) return;
  rosterImportOverlay.classList.add("hidden");
  rosterImportOverlay.setAttribute("aria-hidden", "true");
}

async function finishRosterImportOverlay(startedAt = Date.now()) {
  const remaining = ROSTER_IMPORT_OVERLAY_MAX_MS - (Date.now() - startedAt);
  if (remaining > 0) await new Promise((resolve) => setTimeout(resolve, remaining));
  hideRosterImportOverlay();
}

async function importRosterFiles(files) {
  const accepted = validateIncomingFiles(files);
  if (!accepted.length) return;
  if (!await validateFreshRosterUploads(accepted)) return;
  const startedAt = Date.now();
  showRosterImportOverlay(accepted.length);
  try {
    await mergeFiles(accepted);
    await refreshCreatorCalendarAfterFileChange();
  } finally {
    await finishRosterImportOverlay(startedAt);
  }
}

async function switchDoctorSelection(selectedKey, options = {}) {
  const resetRange = options.resetRange !== false;
  const normalizedSelectedKey = normalizeRosterName(selectedKey);
  doctorSelect.value = selectedKey;
  const canSwitchAsCreator = canUseCreatorDoctorSwitcher();
  if (canSwitchAsCreator && cloudAvailable && (!serverUsers.length || !availableRosterDoctors.length)) {
    void loadServerUsers();
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
    resolvedAccount = locallyResolvedDoctorAccountForSwitch(selectedOption, selectedKey);
    if (!resolvedAccount) {
      setStatus(`Opening ${selectedOption.displayName}...`);
      try {
        resolvedAccount = await resolveDoctorAccountForSwitch(selectedOption);
      } catch (error) {
        setStatus(error.message || "Could not check whether that calendar is claimed.", true);
        return;
      }
    }
  }
  const claimedEmail = normalizeEmail(resolvedAccount?.email || selectedOption?.accountEmail || claimedEmailForDoctorKey(selectedKey, selectedOption?.displayName || ""));
  let switchOverlayId = null;
  if (canSwitchAsCreator && selectedOption) {
    const switchProfile = doctorProfileForDoctor(selectedOption);
    const targetContext = resolvedAccount?.mode === "claimed-account" && claimedEmail
      ? accountCalendarContextForEmail(claimedEmail)
      : switchProfile
        ? calendarSnapshotContext({
            mode: "doctor-profile",
            ownerId: switchProfile.ownerId,
            doctorKey: switchProfile.doctorKey,
          })
        : null;
    const targetSnapshotReady = targetContext
      ? Boolean(await loadCachedCalendarSnapshotForContextAsync(targetContext))
      : false;
    if (!targetContext || !targetSnapshotReady) {
      switchOverlayId = showSwitchOverlay(
        `Switching to ${selectedOption.displayName}…`,
        resolvedAccount?.mode === "claimed-account" ? "Opening the linked account calendar." : "Opening the roster calendar and loading saved doctor-profile edits.",
        { onCancel: cancelCreatorCalendarSwitch },
      );
    }
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
      hideSwitchOverlay(switchOverlayId);
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

function locallyResolvedDoctorAccountForSwitch(doctor, selectedKey = "") {
  const accountEmail = currentClaimedAccountEmail(doctor?.accountEmail || doctor?.claimedBy || "")
    || claimedEmailForDoctorKey(selectedKey || doctor?.key || "", doctor?.displayName || "");
  if (accountEmail) {
    return {
      mode: "claimed-account",
      email: accountEmail,
    };
  }
  if (
    doctor?.targetMode === "doctor-profile"
    && availableRosterDoctors.length
    && serverUsers.length
  ) {
    return {
      mode: "doctor-profile",
      email: "",
    };
  }
  return null;
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
  const fromDoctor = normalizedDoctorSourceTypes(doctor);
  if (fromDoctor.length) return fromDoctor;
  const repositoryDoctor = (availableRosterDoctors || []).find(
    (entry) => doctorIdentityKey(entry) === doctorIdentityKey(doctor),
  );
  return normalizedDoctorSourceTypes(repositoryDoctor);
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
  const [startAddress, endAddress] = String(sheet?.["!ref"] || "A1:A1").split(":");
  const start = decodeSheetCellAddress(startAddress);
  return { s: start, e: decodeSheetCellAddress(endAddress || startAddress) };
}

function cleanSheetCell(sheet, row, col) {
  const address = encodeSheetCellAddress({ r: row - 1, c: col - 1 });
  return String(sheet?.[address]?.v ?? "").trim();
}

function decodeSheetCellAddress(value) {
  const match = String(value || "").toUpperCase().match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!match) return { r: 0, c: 0 };
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return { r: Math.max(0, Number(match[2]) - 1), c: Math.max(0, column - 1) };
}

function encodeSheetCellAddress({ r, c }) {
  let column = Math.max(0, Number(c) || 0) + 1;
  let letters = "";
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return `${letters}${Math.max(0, Number(r) || 0) + 1}`;
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
  creatorSwitcherAnnouncementBaseline = null;
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

function clearActiveViewedAccountState() {
  cancelDeferredBootstrapImports();
  cancelDeferredAccountContextLoad();
  currentSnapshot = null;
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = "";
  currentCalendarRevision = "";
  restoredSessionState = null;
  selectedFiles = [];
  doctorOptions = [];
  activeDoctorProfile = null;
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

async function openUnresolvedShiftIssueEvent(issueId, focus = {}) {
  const issue = globalUnresolvedShiftCodes.find((item) => item.id === String(issueId || ""));
  const doctorKey = normalizeRosterName(focus.doctorKey || issue?.doctorKey || "");
  const displayName = String(focus.displayName || issue?.displayName || doctorKey).trim();
  const date = String(focus.date || issue?.sampleDate || "").slice(0, 10);
  if (!doctorKey || !date) {
    setStatus("This unresolved code does not have a specific roster person and date to open.", true);
    return;
  }
  const source = String(focus.source || issue?.source || "").toLowerCase();
  const candidates = dedupeDoctorOptions([
    ...(availableRosterDoctors || []),
    ...(doctorOptions || []),
  ]);
  const doctor = candidates.find((item) => normalizeRosterName(item.key) === doctorKey && (!source || normalizedDoctorSourceTypes(item).includes(source)))
    || candidates.find((item) => normalizeRosterName(item.key) === doctorKey);
  if (!doctor) {
    setStatus(`Could not find ${displayName || doctorKey} in the available roster calendars.`, true);
    return;
  }

  closeShiftCodeReviewModal();
  shiftCodeReviewReturnContext = { id: String(issue.id || ""), code: String(issue.code || ""), source: String(issue.source || "") };
  pendingUnresolvedIssueFocusDate = date;
  setStatus(`Opening ${doctor.displayName || displayName || doctorKey} on ${formatDate(date)}…`);
  await switchDoctorSelection(doctor.key, { resetRange: false });
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (latestPreview?.events?.length && normalizeRosterName(selectedDoctor()?.key) === doctorKey) break;
    await new Promise((resolve) => window.setTimeout(resolve, 100));
  }
  if (!latestPreview?.events?.length || normalizeRosterName(selectedDoctor()?.key) !== doctorKey) {
    pendingUnresolvedIssueFocusDate = "";
    setStatus("The calendar is still loading. Select the same review item again in a moment.", true);
    return;
  }

  const term = australianTermForDate(parseDateOnly(date));
  settings.dateFrom = formatDateKey(term.start);
  settings.dateTo = formatDateKey(addDays(term.end, -1));
  if (settingsInputs.dateFrom) settingsInputs.dateFrom.value = settings.dateFrom;
  if (settingsInputs.dateTo) settingsInputs.dateTo.value = settings.dateTo;
  rebuildClientPreview();
  saveCurrentSessionState();
  requestAnimationFrame(() => focusPreviewIssueDate(date));
  setStatus(`Opened ${doctor.displayName || displayName || doctorKey} on ${formatDate(date)}.`);
}

function focusPreviewIssueDate(date) {
  const cell = preview.querySelector(`[data-add-date="${CSS.escape(date)}"]`);
  if (!cell) return;
  pendingUnresolvedIssueFocusDate = "";
  if (previewIssueFocusTimer) window.clearTimeout(previewIssueFocusTimer);
  preview.querySelectorAll(".is-unresolved-issue-focus").forEach((item) => item.classList.remove("is-unresolved-issue-focus"));
  cell.classList.add("is-unresolved-issue-focus");
  cell.scrollIntoView({ block: "center", behavior: "smooth" });
  previewIssueFocusTimer = window.setTimeout(() => {
    cell.classList.remove("is-unresolved-issue-focus");
    previewIssueFocusTimer = 0;
  }, 6000);
}

async function returnToShiftCodeReview() {
  const context = shiftCodeReviewReturnContext;
  if (!context) {
    await returnToCreatorCalendar();
    return;
  }
  await returnToCreatorCalendar();
  await openAccountsSurface({ defaultAdminTab: "parser" });
  shiftCodeReviewFilter = { query: context.code, source: context.source || "all" };
  openShiftCodeReviewModal();
  requestAnimationFrame(() => {
    const row = shiftCodeReviewModalBody?.querySelector(`[data-shift-code-review-id="${CSS.escape(context.id)}"]`);
    if (!row) return;
    row.open = true;
    row.scrollIntoView({ block: "center", behavior: "smooth" });
  });
  shiftCodeReviewReturnContext = null;
  setStatus("Returned to the selected shift-code review.");
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
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  }).formatToParts(date);
  const valueFor = (type) => parts.find((part) => part.type === type)?.value || "";
  const dayPeriod = valueFor("dayPeriod").toLowerCase();
  return `${valueFor("day")} ${valueFor("month")} ${valueFor("year")} at ${valueFor("hour")}:${valueFor("minute")}${dayPeriod ? ` ${dayPeriod}` : ""}`;
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
          <span>Display as…</span>
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

function hasLoadedExportPreview() {
  return Boolean(latestPreview || currentSnapshot?.preview);
}

function exportSourceEntries() {
  if (selectedFiles.length) return selectedFiles;
  return importRefsToClientEntries(currentSnapshot?.fileRefs || []);
}

function canOpenExportModal() {
  return Boolean(selectedFiles.length || hasLoadedExportPreview());
}

function canUseExportControls() {
  return Boolean(selectedDoctor() && canOpenExportModal());
}

function openMobileExportModal(event) {
  event?.preventDefault();
  event?.stopPropagation();
  if (mobileExportButton?.disabled) return;
  if (!settingsPanel.classList.contains("hidden")) closeSettingsPanel();
  syncMobileViewportInsets();
  openExportModal();
}

function openExportModal() {
  if (!canOpenExportModal()) {
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
      const snapshot = snapshotCloudSavePayload();
      snapshot.session = {
        ...snapshot.session,
        exportRange: normalizeExportRangeState(exportConfig.mode === "range" ? exportConfig : defaultExportRangeState()),
      };
      snapshot.imports = exportSourceEntries().map(importRefForWorkspace);
      const url = subscriptionUrl("webcal", exportConfig.mode === "range" ? "range" : "full");
      if (!url) throw new Error("No subscription link is available for this account yet.");
      closeExportModal();
      setStatus("Opening Apple Calendar subscription...");
      queueBackgroundCloudStateSave(snapshot, {
        reportErrors: false,
        onError: () => setStatus("Opening Apple Calendar subscription. Feed settings could not be saved in the background.", true),
      });
      window.location.href = url;
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
    snapshot.imports = exportSourceEntries().map(importRefForWorkspace);
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
  adminConsoleOpen = false;
  closeShiftCodeReviewModal();
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

function rosterClaimsForDoctorProfile(profile) {
  if (!profile) return [];
  const aliases = Array.isArray(profile.aliases) && profile.aliases.length
    ? profile.aliases
    : (profile.sourceTypes || []).map((sourceType) => ({
      sourceType,
      key: profile.doctorKey,
      displayName: profile.displayName,
    }));
  return sanitizeRosterClaims(aliases.map((alias) => ({
    sourceType: alias.sourceType,
    key: alias.key || profile.doctorKey,
    displayName: alias.displayName || profile.displayName,
  })));
}

function currentAccount() {
  if (activeCalendarMode() === "doctor-profile" && activeDoctorProfile) {
    const linkedEmail = linkedAccountEmailForDoctorProfile(activeDoctorProfile);
    if (linkedEmail) {
      const serverAccount = serverUsers.map(normalizeServerUser).find((user) => user.email === linkedEmail);
      if (serverAccount) return serverAccount;
      const localAccount = accountState.users.find((user) => user.email === linkedEmail);
      if (localAccount) return localAccount;
    }
    return {
      email: "",
      realName: activeDoctorProfile.displayName || "",
      password: "",
      role: "user",
      claims: rosterClaimsForDoctorProfile(activeDoctorProfile),
    };
  }
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

function linkedAccountEmailForDoctorProfile(profile) {
  if (!profile?.doctorKey) return "";
  const identity = doctorIdentityKey({ key: profile.doctorKey, displayName: profile.displayName });
  const doctor = (availableRosterDoctors || []).find((entry) => doctorIdentityKey(entry) === identity)
    || (availableRosterDoctors || []).find((entry) => normalizeRosterName(entry.key) === normalizeRosterName(profile.doctorKey));
  return normalizeEmail(doctor?.accountEmail || doctor?.claimedBy || "");
}

function viewedAccountEmail() {
  if (activeCalendarMode() === "doctor-profile") {
    return linkedAccountEmailForDoctorProfile(activeDoctorProfile) || "";
  }
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

function canUseFacilityOverview() {
  if (activeCalendarMode() === "doctor-profile") return isCreatorAuthenticated();
  if (currentUserRole === "creator" && !adminViewingEmail) return true;
  return currentFacilityOverviewEnabled === true;
}

function syncFacilityOverviewAccess() {
  const enabled = canUseFacilityOverview();
  const calendarAvailable = !(currentNonClinical && currentDirectorViewEnabled);
  facilityOverviewButton?.classList.toggle("hidden", !enabled || !calendarAvailable);
  mobileFacilityOverviewButton?.classList.toggle("hidden", !enabled || !calendarAvailable);
  facilityOverviewBackButton?.classList.toggle("hidden", !calendarAvailable);
  mobileActionBar?.classList.toggle("has-facility-overview", enabled);
  if (!enabled && facilityOverviewSection && !facilityOverviewSection.classList.contains("hidden")) closeFacilityOverview();
  syncFacilityOverviewNavigationState();
}

function facilityOverviewMelbourneClock(now = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
  }).formatToParts(now);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return {
    today: `${values.year}-${values.month}-${values.day}`,
    time: `${values.hour || "00"}:${values.minute || "00"}`,
  };
}

function facilityOverviewPreferredFacilityFromEvents(events = [], options = {}) {
  const clock = facilityOverviewMelbourneClock(options.now || new Date());
  const today = String(options.today || clock.today).slice(0, 10);
  const nowKey = `${today}T${clock.time}`;
  const sourceOrder = ["MMC", "DDH", "CASEY", "MCH"];
  const sourceRank = (source) => {
    const index = sourceOrder.indexOf(source);
    return index >= 0 ? index : 99;
  };
  const rows = (events || []).map((event) => ({ event, sourceType: eventSourceCode(event), date: eventRosterDateKey(event) }))
    .filter((row) => row.sourceType && row.date && isRosterShiftEvent(row.event) && String(row.event?.status || "").toLowerCase() !== "unknown" && String(row.event?.kind || "").toLowerCase() !== "unknown")
    .filter((row) => !(row.sourceType === "DDH" && /\b(?:hith|vhh)\b/i.test(`${row.event?.title || ""} ${row.event?.rawValue || ""}`)))
    .sort((left, right) => String(left.event?.start || "").localeCompare(String(right.event?.start || "")) || sourceRank(left.sourceType) - sourceRank(right.sourceType));
  const weekday = parseDateOnly(today).getDay();
  const weekStart = formatDateKey(addDays(parseDateOnly(today), weekday === 0 ? -6 : 1 - weekday));
  const weekEnd = formatDateKey(addDays(parseDateOnly(weekStart), 6));
  const weekRows = rows.filter((row) => row.date >= weekStart && row.date <= weekEnd);
  const result = (row, reason) => row ? ({ facilityKey: row.sourceType, reason, evidenceDate: row.date }) : null;
  const uniqueSources = [...new Set(weekRows.map((row) => row.sourceType))];
  if (uniqueSources.length === 1) return result(weekRows[0], "sole-current-week-facility");
  if (uniqueSources.length > 1) {
    const active = weekRows.find((row) => {
      const start = String(row.event?.start || "").slice(0, 16);
      const end = String(row.event?.end || "").slice(0, 16);
      return start && end && start <= nowKey && nowKey < end;
    });
    if (active) return result(active, "active-shift");
    const todayRows = weekRows.filter((row) => row.date === today);
    const nextToday = todayRows.find((row) => String(row.event?.start || "").slice(0, 16) > nowKey);
    if (nextToday) return result(nextToday, "today-next-shift");
    const completedToday = [...todayRows].filter((row) => String(row.event?.end || "").slice(0, 16) <= nowKey)
      .sort((left, right) => String(right.event?.end || "").localeCompare(String(left.event?.end || "")) || sourceRank(left.sourceType) - sourceRank(right.sourceType))[0];
    if (completedToday) return result(completedToday, "today-last-shift");
  }
  const next = rows.find((row) => String(row.event?.start || "").slice(0, 16) > nowKey);
  if (next) return result(next, "next-shift");
  const linkedSources = [...new Set((options.linkedSourceTypes || []).map(normalizeEventSourceCode).filter(Boolean))]
    .sort((left, right) => sourceRank(left) - sourceRank(right));
  return linkedSources[0] ? { facilityKey: linkedSources[0], reason: "sole-or-first-linked-facility", evidenceDate: today } : null;
}

function refreshFacilityOverviewPreferredFacility() {
  if (!canUseFacilityOverview()) return;
  const directorPreference = currentNonClinical && currentDirectorViewEnabled ? directorHospitalPreference() : "";
  if (directorPreference && directorPreference !== "ALL") {
    facilityOverviewState.preferredFacilityKey = directorPreference;
    facilityOverviewState.preferredFacilityReason = "director-hospital-preference";
    facilityOverviewState.preferredFacilityEvidenceDate = "";
    return;
  }
  const doctor = selectedDoctor();
  const clock = facilityOverviewMelbourneClock();
  const preferred = facilityOverviewPreferredFacilityFromEvents(currentSnapshot?.preview?.events || latestPreview?.events || [], {
    today: clock.today,
    linkedSourceTypes: normalizedDoctorSourceTypes(doctor),
  });
  facilityOverviewState.preferredFacilityKey = preferred?.facilityKey || "";
  facilityOverviewState.preferredFacilityReason = preferred?.reason || "";
  facilityOverviewState.preferredFacilityEvidenceDate = preferred?.evidenceDate || "";
}

function isFacilityOverviewOpen() {
  return Boolean(facilityOverviewSection && !facilityOverviewSection.classList.contains("hidden"));
}

function facilityOverviewLabel() {
  // Creators can use the same tools, but this is a Director-facing label only
  // when the viewed account has explicitly been granted Director access.
  return currentDirectorViewEnabled && currentUserRole !== "creator"
    ? "Director overview"
    : "At a glance";
}

function syncFacilityOverviewNavigationState() {
  const open = isFacilityOverviewOpen();
  const label = facilityOverviewLabel();
  const heading = facilityOverviewSection?.querySelector(".facility-overview-head h2");
  const tabList = facilityOverviewSection?.querySelector(".facility-overview-tabs");
  if (heading) heading.textContent = label;
  if (facilityOverviewSection) facilityOverviewSection.setAttribute("aria-label", `${label} ED overview`);
  if (tabList) tabList.setAttribute("aria-label", `${label} views`);
  if (facilityOverviewButton) {
    facilityOverviewButton.textContent = open ? "My calendar" : label;
    facilityOverviewButton.setAttribute("aria-label", open ? "Return to my calendar" : `Open ${label}`);
  }
  if (mobileFacilityOverviewButton) {
    mobileFacilityOverviewButton.setAttribute("aria-label", open ? "Return to my calendar" : `${label} ED overview`);
    const shortLabel = mobileFacilityOverviewButton.querySelector("[aria-hidden='true']");
    const accessibleLabel = mobileFacilityOverviewButton.querySelector(".sr-only");
    if (shortLabel) shortLabel.textContent = open ? "Cal" : "ED";
    if (accessibleLabel) accessibleLabel.textContent = open ? "My calendar" : label;
  }
}

function resetFacilityOverviewScroll() {
  facilityOverviewCompactState.latched = false;
  facilityOverviewCompactState.scroller = null;
  facilityOverviewCompactState.lastScrollTop = 0;
  facilityOverviewCompactState.userDirection = 0;
  facilityOverviewCompactState.touchScroller = null;
  facilityOverviewCompactState.touchY = 0;
  facilityOverviewSection?.classList.remove("is-compact");
  if (facilityOverviewBody) facilityOverviewBody.scrollTop = 0;
  const togetherScroller = facilityOverviewBody?.querySelector(".facility-overview-together");
  if (togetherScroller) togetherScroller.scrollTop = 0;
}

function facilityOverviewFacilityOptions() {
  const values = new Set();
  for (const source of ["mmc", "ddh", "casey", "mch"]) {
    if (Array.isArray(latestPreview?.sources?.[source]) && latestPreview.sources[source].length) values.add(source);
  }
  for (const code of availablePreviewHospitals) values.add(String(code || "").toUpperCase());
  for (const code of latestPreview?.hospitals || []) values.add(String(code || "").toUpperCase());
  for (const code of availableHospitalsForPreview(latestPreview?.events || [])) values.add(String(code || "").toUpperCase());
  for (const file of selectedFiles) values.add(String(file?.sourceType || "").toUpperCase());
  for (const doctor of availableRosterDoctors || []) {
    for (const source of doctor?.sourceTypes || [doctor?.sourceType]) values.add(String(source || "").toUpperCase());
  }
  for (const coverage of facilityOverviewState.byStreamCoverage || []) values.add(String(coverage?.sourceType || "").toUpperCase());
  for (const stream of facilityOverviewState.byStreamCatalog || []) values.add(String(stream?.facilityKey || "").toUpperCase());
  const facilities = [...values].filter((value) => ["MMC", "DDH", "CASEY", "MCH"].includes(value));
  if (!facilities.length && ["MMC", "DDH", "CASEY", "MCH"].includes(facilityOverviewState.facilityKey)) {
    facilities.push(facilityOverviewState.facilityKey);
  }
  return facilities.sort((left, right) => {
    const order = { MMC: 0, DDH: 1, CASEY: 2, MCH: 3 };
    return (order[left] ?? 99) - (order[right] ?? 99);
  });
}

async function loadFacilityOverviewMetadata() {
  if (!canUseFacilityOverview()) return null;
  const metadataKey = [currentSnapshot?.calendarRevision || currentCalendarRevision || "", formatDateKey(australianTermForDate(new Date()).start), normalizedDoctorSourceTypes(selectedDoctor()).sort().join(",")].join("|");
  if (facilityOverviewState.byStreamMetadataKey === metadataKey) return { ok: true };
  if (facilityOverviewState.byStreamMetadataPromise) return facilityOverviewState.byStreamMetadataPromise;
  facilityOverviewState.byStreamMetadataLoading = true;
  facilityOverviewState.byStreamMetadataPromise = (async () => {
    try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryFacilityOverviewMetadata",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        sourceTypes: normalizedDoctorSourceTypes(selectedDoctor()),
      }),
    });
    const data = await readJsonResponse(response, "Could not load available streams.");
    facilityOverviewState.byStreamCoverage = data?.facilities || [];
    facilityOverviewState.byStreamCatalog = facilityOverviewBuildStreamCatalog(data?.catalogEvents || []);
    facilityOverviewState.byStreamMetadataKey = metadataKey;
    return data;
    } catch (error) {
      console.warn("Could not load At a glance stream metadata", error);
      return null;
    } finally {
      facilityOverviewState.byStreamMetadataLoading = false;
      facilityOverviewState.byStreamMetadataPromise = null;
    }
  })();
  return facilityOverviewState.byStreamMetadataPromise;
}

function facilityOverviewStreamKey(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function facilityOverviewAssignmentForRangeRow(row) {
  const event = row?.event;
  if (!event || !isRosterShiftEvent(event)) return null;
  const person = { key: String(row.doctorKey || ""), displayName: String(row.displayName || row.doctorKey || "") };
  const assignment = buildWhoAssignment(person, {}, { ...event, seniority: row.seniority || event.seniority || "Unknown" });
  if (!assignment || !facilityOverviewIsMeaningfulStream(assignment.team, row.sourceType || assignment.source)) return null;
  const streamKey = facilityOverviewStreamKey(assignment.team);
  if (!streamKey) return null;
  return {
    ...assignment,
    sourceType: String(row.sourceType || assignment.source || "").toUpperCase(),
    streamKey,
    streamLabel: assignment.team,
    seniority: facilityOverviewDetectedSeniority(event, row.seniority || event.seniority || "Unknown"),
    date: String(row.date || eventRosterDateKey(event)).slice(0, 10),
    event,
    doctorKey: String(row.doctorKey || person.key),
    displayName: String(row.displayName || person.displayName),
  };
}

function facilityOverviewBuildStreamCatalog(rows = []) {
  const entries = new Map();
  for (const row of rows || []) {
    const assignment = facilityOverviewAssignmentForRangeRow(row);
    if (!assignment) continue;
    const key = `${assignment.sourceType}|${assignment.streamKey}`;
    const existing = entries.get(key) || {
      facilityKey: assignment.sourceType,
      streamKey: assignment.streamKey,
      label: assignment.streamLabel,
      seniorities: new Set(),
      firstSeenDate: assignment.date,
      lastSeenDate: assignment.date,
      rank: whoTeamRank(assignment.streamLabel, assignment.sourceType),
    };
    existing.seniorities.add(assignment.seniority);
    if (assignment.date < existing.firstSeenDate) existing.firstSeenDate = assignment.date;
    if (assignment.date > existing.lastSeenDate) existing.lastSeenDate = assignment.date;
    entries.set(key, existing);
  }
  return [...entries.values()].map((entry) => ({
    ...entry,
    seniorities: [...entry.seniorities].sort(compareFacilityOverviewSeniorities),
  })).sort((left, right) => left.facilityKey.localeCompare(right.facilityKey) || left.rank - right.rank || left.label.localeCompare(right.label));
}

function facilityOverviewStreamsForFacility(facilityKey) {
  const key = String(facilityKey || "").toUpperCase();
  return (facilityOverviewState.byStreamCatalog || []).filter((entry) => entry.facilityKey === key);
}

function facilityOverviewFirstStreamKey(facilityKey) {
  return facilityOverviewStreamsForFacility(facilityKey)[0]?.streamKey || "";
}

function facilityOverviewPreferredStreamKey(facilityKey) {
  const facility = String(facilityKey || "").toUpperCase();
  const now = new Date();
  const assignments = buildWhoAssignments(selectedDoctor(), latestPreview?.events || [])
    .filter((assignment) => String(assignment.source || "").toUpperCase() === facility)
    .filter((assignment) => facilityOverviewIsMeaningfulStream(assignment.team, assignment.source))
    .sort((left, right) => String(left.event?.start || "").localeCompare(String(right.event?.start || "")));
  const active = assignments.find((assignment) => {
    const start = new Date(assignment.event?.start || "");
    const end = new Date(assignment.event?.end || "");
    return !assignment.event?.allDay && start <= now && now < end;
  });
  const next = assignments.find((assignment) => new Date(assignment.event?.start || "") > now);
  return facilityOverviewStreamKey((active || next)?.team || "") || facilityOverviewFirstStreamKey(facility);
}

function newFacilityOverviewByStreamRow(options = {}) {
  const currentFacility = String(facilityOverviewState.facilityKey || "").toUpperCase();
  const facilityKey = String(options.facilityKey || facilityOverviewState.preferredFacilityKey || (currentFacility === "ALL" ? "" : currentFacility) || facilityOverviewFacilityOptions()[0] || "MMC").toUpperCase();
  const streamKey = String(options.streamKey || facilityOverviewPreferredStreamKey(facilityKey));
  facilityOverviewState.byStreamRowId += 1;
  return {
    id: `stream-${facilityOverviewState.byStreamRowId}`,
    facilityKey,
    streamKey,
    seniority: options.seniority || "ALL",
    isPrefilled: Boolean(options.id),
    duplicateOfId: "",
  };
}

function initializeFacilityOverviewByStreamState() {
  if (!Array.isArray(facilityOverviewState.byStreamRows) || !facilityOverviewState.byStreamRows.length) {
    facilityOverviewState.byStreamRows = [newFacilityOverviewByStreamRow()];
  }
  for (const row of facilityOverviewState.byStreamRows) {
    const streams = facilityOverviewStreamsForFacility(row.facilityKey);
    if (streams.length && !streams.some((stream) => stream.streamKey === row.streamKey)) row.streamKey = streams[0].streamKey;
    if (!row.seniority) row.seniority = "ALL";
    if (typeof row.isPrefilled !== "boolean") row.isPrefilled = false;
    if (typeof row.duplicateOfId !== "string") row.duplicateOfId = "";
  }
  reconcileFacilityOverviewByStreamDuplicates();
}

function facilityOverviewByStreamRowIsDuplicate(row) {
  return facilityOverviewByStreamDuplicateRows().some((duplicate) => duplicate.row.id === row?.id);
}

function facilityOverviewByStreamSelectionKey(row) {
  return row?.facilityKey && row?.streamKey && row?.seniority
    ? `${row.facilityKey}|${row.streamKey}|${row.seniority}`
    : "";
}

function reconcileFacilityOverviewByStreamDuplicates(changedRow = null) {
  const rows = facilityOverviewState.byStreamRows || [];
  const rowsById = new Map(rows.map((row) => [row.id, row]));
  for (const row of rows) {
    const original = rowsById.get(row.duplicateOfId);
    if (!original || facilityOverviewByStreamSelectionKey(row) !== facilityOverviewByStreamSelectionKey(original)) {
      row.duplicateOfId = "";
    }
  }
  if (!changedRow) return;
  const key = facilityOverviewByStreamSelectionKey(changedRow);
  const original = rows.find((row) => row.id !== changedRow.id && !row.isPrefilled && facilityOverviewByStreamSelectionKey(row) === key);
  changedRow.duplicateOfId = original?.id || "";
}

function facilityOverviewByStreamDuplicateRows(rows = facilityOverviewState.byStreamRows) {
  const rowsById = new Map((rows || []).map((row, index) => [row.id, { row, index }]));
  const duplicates = [];
  for (const [index, row] of (rows || []).entries()) {
    const original = rowsById.get(row?.duplicateOfId);
    if (original && facilityOverviewByStreamSelectionKey(row) === facilityOverviewByStreamSelectionKey(original.row)) {
      duplicates.push({ row, index, first: original });
    }
  }
  return duplicates;
}

function facilityOverviewByStreamDistinctRows(rows = facilityOverviewState.byStreamRows) {
  const seen = new Set();
  return (rows || []).filter((row) => {
    const key = facilityOverviewByStreamSelectionKey(row);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function facilityOverviewAccountKey() {
  return normalizeEmail(viewedAccountEmail() || currentUserEmail);
}

function savedFacilityOverviewTabs() {
  try {
    const stored = JSON.parse(localStorage.getItem(FACILITY_OVERVIEW_TAB_PREFERENCES_KEY) || "{}");
    return stored && typeof stored === "object" && !Array.isArray(stored) ? stored : {};
  } catch {
    return {};
  }
}

function savedFacilityOverviewTabForCurrentAccount() {
  const tab = String(savedFacilityOverviewTabs()[facilityOverviewAccountKey()] || "");
  return ["on-shift", "staff", "together", "by-stream"].includes(tab) ? tab : "";
}

function rememberFacilityOverviewTabForCurrentAccount() {
  const accountKey = facilityOverviewAccountKey();
  const tab = String(facilityOverviewState.tab || "");
  if (!accountKey || !["on-shift", "staff", "together", "by-stream"].includes(tab)) return;
  try {
    const stored = savedFacilityOverviewTabs();
    if (stored[accountKey] === tab) return;
    localStorage.setItem(FACILITY_OVERVIEW_TAB_PREFERENCES_KEY, JSON.stringify({ ...stored, [accountKey]: tab }));
  } catch {
    // The overview still works if browser storage is unavailable.
  }
}

function beginFacilityOverviewAccountSession() {
  rememberFacilityOverviewTabForCurrentAccount();
  facilityOverviewSessionNeedsInitialization = true;
}

function resetFacilityOverviewSessionState() {
  const today = formatDateKey(new Date());
  const currentTerm = australianTermForDate(new Date());
  const currentTermStart = formatDateKey(currentTerm.start);
  const currentTermEnd = formatDateKey(addDays(currentTerm.end, -1));
  const directorPreference = currentNonClinical && currentDirectorViewEnabled ? directorHospitalPreference() : "";
  const defaultTab = directorPreference === "ALL" ? "staff" : "on-shift";
  facilityOverviewState.tab = savedFacilityOverviewTabForCurrentAccount() || defaultTab;
  facilityOverviewState.date = today;
  facilityOverviewState.facilityKey = directorPreference || "";
  facilityOverviewState.includeClinicalSupport = false;
  facilityOverviewState.onShiftData = null;
  facilityOverviewState.content = "";
  facilityOverviewState.staffTermStart = currentTermStart;
  facilityOverviewState.staffTerms = [];
  facilityOverviewState.staffContent = "";
  facilityOverviewState.staffData = null;
  facilityOverviewState.staffQuery = "";
  facilityOverviewState.staffExpanded = new Set();
  facilityOverviewState.staffFocusSection = "";
  facilityOverviewState.staffActionMenu = null;
  facilityOverviewState.staffDesignationMenu = null;
  facilityOverviewState.staffSeniorityMenu = null;
  facilityOverviewState.staffMultiSelectSection = "";
  facilityOverviewState.staffMultiSelectMembers = new Map();
  facilityOverviewState.staffBulkSeniorityMenu = null;
  facilityOverviewState.staffMultiSelectSaving = false;
  facilityOverviewState.byStreamFrom = today;
  facilityOverviewState.byStreamTo = today;
  facilityOverviewState.byStreamRows = [];
  facilityOverviewState.byStreamCatalog = [];
  facilityOverviewState.byStreamCoverage = [];
  facilityOverviewState.byStreamContent = "";
  facilityOverviewState.byStreamData = null;
  facilityOverviewState.byStreamLoading = false;
  facilityOverviewState.byStreamMetadataLoading = false;
  facilityOverviewState.byStreamMetadataKey = "";
  facilityOverviewState.byStreamMetadataPromise = null;
  facilityOverviewState.byStreamHideEmptyDates = true;
  facilityOverviewState.togetherStaffKeys = [""];
  facilityOverviewState.togetherRangeMode = "term";
  facilityOverviewState.togetherTermStart = currentTermStart;
  facilityOverviewState.togetherFrom = currentTermStart;
  facilityOverviewState.togetherTo = currentTermEnd;
  facilityOverviewState.togetherFacilityKey = "ALL";
  facilityOverviewState.togetherContent = "";
  facilityOverviewState.togetherHasSearched = false;
  facilityOverviewState.togetherPinnedDoctors = [];
  facilityOverviewState.togetherUserClearedAll = true;
  facilityOverviewSessionNeedsInitialization = false;
}

async function openFacilityOverview(options = {}) {
  if (!canUseFacilityOverview()) return;
  refreshFacilityOverviewPreferredFacility();
  if (facilityOverviewSessionNeedsInitialization) resetFacilityOverviewSessionState();
  if (options.preserveFacility !== true && options.preserveStaffTerm !== true) {
    const directorPreference = currentNonClinical && currentDirectorViewEnabled ? directorHospitalPreference() : "";
    const preferred = String(facilityOverviewState.preferredFacilityKey || "").toUpperCase();
    if (directorPreference) {
      facilityOverviewState.facilityKey = directorPreference;
      facilityOverviewState.togetherFacilityKey = directorPreference;
    } else if (preferred) {
      facilityOverviewState.facilityKey = preferred;
    }
    if (!options.preserveDate) facilityOverviewState.date = formatDateKey(new Date());
    if (!options.preserveByStreamRange) {
      const today = formatDateKey(new Date());
      facilityOverviewState.byStreamFrom = today;
      facilityOverviewState.byStreamTo = today;
      facilityOverviewState.byStreamRows = [];
      facilityOverviewState.byStreamData = null;
      facilityOverviewState.byStreamContent = "";
    }
  }
  const facilities = facilityOverviewFacilityOptions();
  if (!facilityOverviewState.facilityKey || (facilityOverviewState.facilityKey !== "ALL" && !facilities.includes(facilityOverviewState.facilityKey))) {
    facilityOverviewState.facilityKey = facilities[0] || "MMC";
  }
  form?.classList.add("is-facility-overview-active");
  previewSection?.classList.add("hidden");
  facilityOverviewSection?.classList.remove("hidden");
  resetFacilityOverviewScroll();
  syncFacilityOverviewNavigationState();
  if (facilityOverviewState.tab === "staff" && !options.preserveStaffTerm) {
    facilityOverviewState.staffTermStart = formatDateKey(australianTermForDate(new Date()).start);
  }
  renderFacilityOverview();
  if (facilityOverviewState.tab === "staff") await loadFacilityOverviewStaff();
  else if (facilityOverviewState.tab === "by-stream") await openFacilityOverviewByStream();
  else if (facilityOverviewState.tab === "together") initializeFacilityOverviewTogetherState();
  else await loadFacilityOverviewOnShift();
}

async function openFacilityOverviewByStream() {
  if (!canUseFacilityOverview() || facilityOverviewState.tab !== "by-stream") return;
  facilityOverviewState.byStreamContent = `<article class="issue-card"><p>Loading available streams…</p></article>`;
  renderFacilityOverview();
  await loadFacilityOverviewMetadata();
  if (facilityOverviewState.tab !== "by-stream") return;
  initializeFacilityOverviewByStreamState();
  renderFacilityOverview();
  await loadFacilityOverviewByStream();
}

function closeFacilityOverview() {
  facilityOverviewNavigationLocked = false;
  facilityOverviewState.requestId += 1;
  facilityOverviewState.byStreamRequestId += 1;
  facilityOverviewState.staffActionMenu = null;
  facilityOverviewState.staffDesignationMenu = null;
  facilityOverviewState.staffSeniorityMenu = null;
  clearFacilityOverviewStaffMultiSelect({ render: false });
  form?.classList.remove("is-facility-overview-active");
  resetFacilityOverviewScroll();
  facilityOverviewSection?.classList.add("hidden");
  if (latestPreview) previewSection?.classList.remove("hidden");
  syncFacilityOverviewNavigationState();
}

function renderFacilityOverview() {
  if (!facilityOverviewBody || !facilityOverviewControls) return;
  if (!facilityOverviewSessionNeedsInitialization) rememberFacilityOverviewTabForCurrentAccount();
  if (facilityOverviewHeader) facilityOverviewHeader.innerHTML = renderFacilityOverviewHeader();
  syncFacilityOverviewTabOrder();
  const activeTab = facilityOverviewState.tab || "on-shift";
  facilityOverviewBody.classList.toggle("is-working-together", activeTab === "together");
  facilityOverviewSection?.querySelectorAll("[data-facility-overview-tab]").forEach((button) => {
    const selected = button.dataset.facilityOverviewTab === activeTab;
    button.classList.toggle("is-active", selected);
    button.setAttribute("aria-selected", String(selected));
  });
  if (activeTab === "staff") {
    const facilities = facilityOverviewFacilityOptions();
    const selected = facilityOverviewState.facilityKey === "ALL" || facilities.includes(facilityOverviewState.facilityKey)
      ? facilityOverviewState.facilityKey : "ALL";
    facilityOverviewState.facilityKey = selected || "ALL";
    const selectedTerm = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date())));
    const terms = facilityOverviewState.staffTerms.length ? facilityOverviewState.staffTerms : [{ value: formatDateKey(selectedTerm.start), label: formatAustralianTermLabel(selectedTerm) }];
    facilityOverviewControls.innerHTML = `
      <label class="field"><span>ED</span><select data-facility-overview-facility><option value="ALL" ${facilityOverviewState.facilityKey === "ALL" ? "selected" : ""}>All EDs</option>${facilities.map((facility) => `<option value="${escapeHtml(facility)}" ${facility === facilityOverviewState.facilityKey ? "selected" : ""}>${escapeHtml(displaySourceCode(facility))}</option>`).join("")}</select></label>
      <label class="field facility-overview-staff-search"><span>Find staff</span><input type="search" value="${escapeHtml(facilityOverviewState.staffQuery)}" placeholder="Name" data-facility-overview-staff-search></label>
      <label class="field facility-overview-staff-term-control"><span>Term</span><select data-facility-overview-staff-term>${terms.map((term) => `<option value="${escapeHtml(term.value)}" ${term.value === facilityOverviewState.staffTermStart ? "selected" : ""}>${escapeHtml(term.label)}</option>`).join("")}</select></label>
    `;
    renderFacilityOverviewStaffBody();
    queueFacilityOverviewMenuPositioning();
    return;
  }
  if (activeTab === "together") {
    initializeFacilityOverviewTogetherState();
    facilityOverviewControls.innerHTML = "";
    facilityOverviewBody.innerHTML = renderFacilityOverviewTogetherProposal();
    queueFacilityOverviewMenuPositioning();
    return;
  }
  if (activeTab === "by-stream") {
    initializeFacilityOverviewByStreamState();
    facilityOverviewControls.innerHTML = `
      <label class="field"><span>From</span><input type="date" value="${escapeHtml(facilityOverviewState.byStreamFrom)}" data-facility-overview-by-stream-date="from"></label>
      <label class="field"><span>To</span><input type="date" value="${escapeHtml(facilityOverviewState.byStreamTo)}" data-facility-overview-by-stream-date="to"></label>
      <button type="button" class="button button-secondary facility-overview-by-stream-today" data-facility-overview-today>Today</button>
      ${renderFacilityOverviewDateNavigation("range")}
    `;
    facilityOverviewBody.innerHTML = renderFacilityOverviewByStream();
    queueFacilityOverviewMenuPositioning();
    return;
  }
  const facilities = facilityOverviewFacilityOptions();
  const selected = facilityOverviewState.facilityKey === "ALL" || facilities.includes(facilityOverviewState.facilityKey)
    ? facilityOverviewState.facilityKey : facilities[0] || "MMC";
  facilityOverviewState.facilityKey = selected;
  const content = facilityOverviewState.content || `<article class="issue-card"><p>Choose an ED and date to load rostered staff.</p></article>`;
  facilityOverviewControls.innerHTML = `
    <label class="field"><span>ED</span><select data-facility-overview-facility><option value="ALL" ${selected === "ALL" ? "selected" : ""}>All EDs</option>${facilities.map((facility) => `<option value="${escapeHtml(facility)}" ${facility === selected ? "selected" : ""}>${escapeHtml(displaySourceCode(facility))}</option>`).join("")}</select></label>
    <label class="field"><span>Date</span><input type="date" value="${escapeHtml(facilityOverviewState.date)}" data-facility-overview-date></label>
    ${renderFacilityOverviewDateNavigation("day")}
    <label class="toggle facility-overview-cs-toggle">CS <input type="checkbox" data-facility-overview-include-cs ${facilityOverviewState.includeClinicalSupport ? "checked" : ""}></label>
  `;
  facilityOverviewBody.innerHTML = `<div class="facility-overview-results">${content}</div>`;
  queueFacilityOverviewMenuPositioning();
}

function syncFacilityOverviewTabOrder() {
  const tabs = facilityOverviewSection?.querySelector(".facility-overview-tabs");
  if (!tabs) return;
  const order = currentNonClinical && currentDirectorViewEnabled && directorHospitalPreference() === "ALL"
    ? ["staff", "on-shift", "together", "by-stream"]
    : ["on-shift", "staff", "together", "by-stream"];
  for (const tabName of order) {
    const button = tabs.querySelector(`[data-facility-overview-tab="${tabName}"]`);
    if (button) tabs.append(button);
  }
}

function renderFacilityOverviewStaffBody() {
  if (!facilityOverviewBody || facilityOverviewState.tab !== "staff") return;
  facilityOverviewBody.innerHTML = `<div class="facility-overview-results">${facilityOverviewState.staffContent || `<article class="issue-card"><p>Loading ED staff…</p></article>`}</div>`;
  queueFacilityOverviewMenuPositioning();
}

function queueFacilityOverviewMenuPositioning() {
  window.requestAnimationFrame(() => {
    const margin = 8;
    for (const menu of document.querySelectorAll(".facility-overview-staff-action-menu")) {
      const requestedLeft = Number.parseFloat(menu.style.getPropertyValue("--facility-overview-menu-left")) || margin;
      const requestedTop = Number.parseFloat(menu.style.getPropertyValue("--facility-overview-menu-top")) || margin;
      menu.style.maxHeight = `${Math.max(120, window.innerHeight - margin * 2)}px`;
      menu.style.overflowY = "auto";
      const rect = menu.getBoundingClientRect();
      const left = Math.min(Math.max(margin, requestedLeft), Math.max(margin, window.innerWidth - rect.width - margin));
      const top = Math.min(Math.max(margin, requestedTop), Math.max(margin, window.innerHeight - rect.height - margin));
      menu.style.left = `${left}px`;
      menu.style.top = `${top}px`;
    }
  });
}

function renderFacilityOverviewHeader() {
  const doctor = selectedDoctor();
  const displayName = doctor?.displayName || currentAccount().realName || "Selected doctor";
  const start = latestPreview?.previewStart || settings.dateFrom || "";
  const end = latestPreview?.previewEnd || settings.dateTo || "";
  const staffView = facilityOverviewState.tab === "staff";
  const togetherView = facilityOverviewState.tab === "together";
  const byStreamView = facilityOverviewState.tab === "by-stream";
  const selectedTerm = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date())));
  const terms = facilityOverviewState.staffTerms.length ? facilityOverviewState.staffTerms : [{ value: formatDateKey(selectedTerm.start), label: formatAustralianTermLabel(selectedTerm) }];
  const previousTerm = selectedTerm.termNumber === 1
    ? buildAustralianTerm(selectedTerm.year - 1, 4, startMonthIndexForTerm(4))
    : buildAustralianTerm(selectedTerm.year, selectedTerm.termNumber - 1, startMonthIndexForTerm(selectedTerm.termNumber - 1));
  const nextTerm = nextAustralianTerm(selectedTerm);
  const hasPrevious = terms.some((term) => term.value === formatDateKey(previousTerm.start));
  const hasNext = terms.some((term) => term.value === formatDateKey(nextTerm.start));
  return `
    <div class="preview-head facility-overview-preview-head">
      <div class="preview-doctor-control">
        <span>Doctor</span>
        <strong>${escapeHtml(displayName)}</strong>
      </div>
      <div class="preview-toolbar">
        <div class="preview-range-controls facility-overview-range-controls" aria-label="${byStreamView ? "By stream date range" : "Calendar range; unavailable in At a glance"}">
          ${staffView ? `
            <button type="button" class="button button-secondary preview-today-button" data-facility-overview-today disabled aria-disabled="true">Today</button>
            ${hasPrevious ? `<button type="button" class="button button-secondary" data-facility-overview-staff-term-step="-1" aria-label="Previous term">‹</button>` : ""}
            <label class="preview-term-start-control"><span class="preview-range-label">Term</span>
              <select class="preview-range-button preview-term-start-select" data-facility-overview-staff-term>
                ${terms.map((term) => `<option value="${escapeHtml(term.value)}" ${term.value === facilityOverviewState.staffTermStart ? "selected" : ""}>${escapeHtml(term.label)}</option>`).join("")}
              </select>
            </label>
            ${hasNext ? `<button type="button" class="button button-secondary" data-facility-overview-staff-term-step="1" aria-label="Next term">›</button>` : ""}
          ` : togetherView ? `<span class="facility-overview-header-note">Compare staff rosters</span>` : byStreamView ? `
            <label class="preview-range-input-control"><span class="preview-range-label">From</span><input type="date" class="preview-range-button preview-range-date-input" value="${escapeHtml(facilityOverviewState.byStreamFrom)}" data-facility-overview-by-stream-date="from"></label>
            <label class="preview-range-input-control"><span class="preview-range-label">To</span><input type="date" class="preview-range-button preview-range-date-input" value="${escapeHtml(facilityOverviewState.byStreamTo)}" data-facility-overview-by-stream-date="to"></label>
            <button type="button" class="button button-secondary preview-today-button" data-facility-overview-today>Today</button>
            ${renderFacilityOverviewDateNavigation("range", { header: true })}
          ` : `
            <span class="preview-range-label">From</span><button type="button" class="preview-range-button" disabled>${escapeHtml(start ? formatDate(start) : "Set date")}</button>
            <span class="preview-range-label">To</span><button type="button" class="preview-range-button" disabled>${escapeHtml(end ? formatDate(end) : "Set date")}</button>
            <button type="button" class="button button-secondary preview-today-button" data-facility-overview-today>Today</button>
            ${renderFacilityOverviewDateNavigation("day", { header: true })}
          `}
        </div>
        ${currentDirectorViewEnabled
          ? `<button type="button" class="button button-secondary" data-facility-overview-account>Account</button>`
          : ""}
        ${canReturnToCreator()
          ? `<button type="button" class="button button-secondary" data-facility-overview-back-to-creator>Back to creator</button>`
          : ""}
        <button type="button" class="button button-secondary preview-logout-button" data-facility-overview-logout>Log out</button>
      </div>
    </div>
  `;
}

function renderFacilityOverviewDateNavigation(kind, { header = false } = {}) {
  const range = kind === "range";
  const previous = range ? "Previous date range" : "Previous day";
  const next = range ? "Next date range" : "Next day";
  const className = header ? "facility-overview-header-date-navigation" : "facility-overview-date-actions";
  return `<div class="${className}" aria-label="${range ? "Move date range" : "Move date"}">
    <button type="button" class="button button-secondary" data-facility-overview-date-step="-1" aria-label="${previous}"><span class="facility-overview-date-navigation-label">Previous</span><span class="facility-overview-date-navigation-chevron" aria-hidden="true">‹</span></button>
    <button type="button" class="button button-secondary" data-facility-overview-date-step="1" aria-label="${next}"><span class="facility-overview-date-navigation-label">Next</span><span class="facility-overview-date-navigation-chevron" aria-hidden="true">›</span></button>
  </div>`;
}

function facilityOverviewMergeStreamCatalog(next = []) {
  const merged = new Map();
  for (const entry of [...(facilityOverviewState.byStreamCatalog || []), ...(next || [])]) {
    const key = `${entry.facilityKey}|${entry.streamKey}`;
    const existing = merged.get(key) || { ...entry, seniorities: new Set(), firstSeenDate: entry.firstSeenDate, lastSeenDate: entry.lastSeenDate };
    for (const seniority of entry.seniorities || []) existing.seniorities.add(seniority);
    if (entry.firstSeenDate && (!existing.firstSeenDate || entry.firstSeenDate < existing.firstSeenDate)) existing.firstSeenDate = entry.firstSeenDate;
    if (entry.lastSeenDate && (!existing.lastSeenDate || entry.lastSeenDate > existing.lastSeenDate)) existing.lastSeenDate = entry.lastSeenDate;
    merged.set(key, existing);
  }
  facilityOverviewState.byStreamCatalog = [...merged.values()].map((entry) => ({
    ...entry,
    seniorities: [...entry.seniorities].sort(compareFacilityOverviewSeniorities),
  })).sort((left, right) => left.facilityKey.localeCompare(right.facilityKey) || left.rank - right.rank || left.label.localeCompare(right.label));
}

function facilityOverviewByStreamSeniorityOptions(row) {
  const stream = facilityOverviewStreamsForFacility(row.facilityKey).find((entry) => entry.streamKey === row.streamKey);
  const values = new Set(["SMS", ...(stream?.seniorities || [])]);
  return [...values].sort(compareFacilityOverviewSeniorities);
}

function renderFacilityOverviewByStream() {
  const facilities = facilityOverviewFacilityOptions();
  const rows = facilityOverviewState.byStreamRows || [];
  const resultRows = facilityOverviewByStreamDistinctRows(rows);
  const duplicateRows = facilityOverviewByStreamDuplicateRows(rows);
  const duplicatesById = new Map(duplicateRows.map((duplicate) => [duplicate.row.id, duplicate]));
  return `
    <section class="facility-overview-by-stream" aria-label="By stream comparison">
      <aside class="facility-overview-by-stream-selectors">
        <div class="facility-overview-by-stream-selector-head"><h3>Streams to compare</h3><p>Each row is one result lane.</p>${duplicateRows.length ? `<p class="facility-overview-by-stream-duplicate-warning" role="status">Choose each ED, stream, and seniority combination only once. Duplicate selections are shown once.</p>` : ""}</div>
        <div class="facility-overview-by-stream-row-list">
          ${rows.map((row, index) => {
            const streams = facilityOverviewStreamsForFacility(row.facilityKey);
            const seniorities = facilityOverviewByStreamSeniorityOptions(row);
            const duplicate = duplicatesById.get(row.id);
            return `<fieldset class="facility-overview-by-stream-row${duplicate ? " is-duplicate" : ""}"${duplicate ? ` aria-describedby="facility-overview-by-stream-duplicate-${escapeHtml(row.id)}"` : ""}><legend>Stream selection ${index + 1}</legend>
              <label class="field facility-overview-by-stream-field-ed"><span>ED</span><select data-facility-overview-by-stream-row="${escapeHtml(row.id)}" data-facility-overview-by-stream-field="facility">${facilities.map((facility) => `<option value="${escapeHtml(facility)}" ${facility === row.facilityKey ? "selected" : ""}>${escapeHtml(displaySourceCode(facility))}</option>`).join("")}</select></label>
              <label class="field facility-overview-by-stream-field-stream"><span>Stream</span><select data-facility-overview-by-stream-row="${escapeHtml(row.id)}" data-facility-overview-by-stream-field="stream" ${streams.length ? "" : "disabled"}><option value="">Choose stream…</option>${streams.map((stream) => `<option value="${escapeHtml(stream.streamKey)}" ${stream.streamKey === row.streamKey ? "selected" : ""}>${escapeHtml(stream.label)}</option>`).join("")}</select></label>
              <label class="field facility-overview-by-stream-field-seniority"><span>Seniority</span><select data-facility-overview-by-stream-row="${escapeHtml(row.id)}" data-facility-overview-by-stream-field="seniority"><option value="ALL" ${row.seniority === "ALL" ? "selected" : ""}>All team</option>${seniorities.map((seniority) => `<option value="${escapeHtml(seniority)}" ${seniority === row.seniority ? "selected" : ""}>${escapeHtml(seniority)}</option>`).join("")}</select></label>
              ${duplicate ? `<p id="facility-overview-by-stream-duplicate-${escapeHtml(row.id)}" class="facility-overview-by-stream-row-duplicate">Duplicates stream selection ${duplicate.first.index + 1}; it is shown once in the results.</p>` : ""}
              ${rows.length > 1 ? `<button type="button" class="facility-overview-by-stream-remove" data-facility-overview-by-stream-remove="${escapeHtml(row.id)}" aria-label="Remove stream selection ${index + 1}">Remove</button>` : ""}
            </fieldset>`;
          }).join("")}
        </div>
        <button type="button" class="facility-overview-add-staff" data-facility-overview-by-stream-add ${rows.length >= 6 ? "disabled" : ""}>${rows.length >= 6 ? "Maximum of six streams" : "+ Add another stream"}</button>
        <label class="toggle facility-overview-by-stream-hide-empty"><input type="checkbox" data-facility-overview-by-stream-row="options" data-facility-overview-by-stream-field="hide-empty" ${facilityOverviewState.byStreamHideEmptyDates ? "checked" : ""}> Hide dates without assignments</label>
      </aside>
      <div class="facility-overview-by-stream-results ${resultRows.length > 1 ? "is-multi-lane" : "is-single-lane"}" aria-live="polite">${facilityOverviewState.byStreamContent || `<article class="issue-card"><p>Choose an ED and stream to view coverage.</p></article>`}</div>
    </section>
  `;
}

function facilityOverviewByStreamCoverageFor(facilityKey, date) {
  const source = String(facilityKey || "").toLowerCase();
  return (facilityOverviewState.byStreamCoverage || []).some((coverage) => String(coverage.sourceType || "").toLowerCase() === source && coverage.startDate <= date && coverage.endDate >= date);
}

function facilityOverviewByStreamDates() {
  const start = facilityOverviewState.byStreamFrom;
  const end = facilityOverviewState.byStreamTo;
  const dates = [];
  if (!start || !end || end < start) return dates;
  for (let date = parseDateOnly(start); formatDateKey(date) <= end; date = addDays(date, 1)) dates.push(formatDateKey(date));
  return dates;
}

function setFacilityOverviewByStreamRange({ from, to } = {}) {
  let nextFrom = String(from ?? facilityOverviewState.byStreamFrom ?? "").slice(0, 10);
  let nextTo = String(to ?? facilityOverviewState.byStreamTo ?? "").slice(0, 10);
  if (nextFrom && nextTo && nextFrom > nextTo) {
    if (from != null) nextTo = nextFrom;
    else nextFrom = nextTo;
  }
  facilityOverviewState.byStreamFrom = nextFrom;
  facilityOverviewState.byStreamTo = nextTo;
  return loadFacilityOverviewByStream();
}

function facilityOverviewByStreamRangeLabel() {
  const start = facilityOverviewState.byStreamFrom;
  const end = facilityOverviewState.byStreamTo;
  if (!start || !end) return "";
  const startDate = parseDateOnly(start);
  const endDate = parseDateOnly(end);
  const day = new Intl.DateTimeFormat("en-AU", { day: "numeric" });
  const month = new Intl.DateTimeFormat("en-AU", { month: "short" });
  if (start === end) return `${day.format(startDate)} ${month.format(startDate)}`;
  if (startDate.getFullYear() === endDate.getFullYear() && startDate.getMonth() === endDate.getMonth()) {
    return `${day.format(startDate)}–${day.format(endDate)} ${month.format(endDate)}`;
  }
  return `${day.format(startDate)} ${month.format(startDate)} – ${day.format(endDate)} ${month.format(endDate)}`;
}

function facilityOverviewByStreamContentFromData(data) {
  const selected = facilityOverviewByStreamDistinctRows();
  const assignments = (data?.events || []).map(facilityOverviewAssignmentForRangeRow).filter(Boolean);
  const bySelection = new Map(selected.map((row) => [row.id, []]));
  const observedByStreamDate = new Set();
  for (const assignment of assignments) {
    observedByStreamDate.add(`${assignment.sourceType}|${assignment.streamKey}|${assignment.date}`);
    for (const row of selected) {
      if (assignment.sourceType !== row.facilityKey || assignment.streamKey !== row.streamKey) continue;
      if (row.seniority !== "ALL" && facilityOverviewNormalizeSeniority(assignment.seniority) !== facilityOverviewNormalizeSeniority(row.seniority)) continue;
      bySelection.get(row.id).push(assignment);
    }
  }
  const dates = facilityOverviewByStreamDates();
  let hiddenCount = 0;
  const visibleDates = dates.filter((date) => {
    const hasMatch = selected.some((row) => (bySelection.get(row.id) || []).some((assignment) => assignment.date === date));
    const hasUncovered = selected.some((row) => !facilityOverviewByStreamCoverageFor(row.facilityKey, date));
    if (facilityOverviewState.byStreamHideEmptyDates && !hasMatch && !hasUncovered) {
      hiddenCount += 1;
      return false;
    }
    return true;
  });
  const laneHeading = (row, className = "facility-overview-by-stream-lane-head") => {
    const stream = facilityOverviewStreamsForFacility(row.facilityKey).find((entry) => entry.streamKey === row.streamKey);
    return `<header class="${className}"><strong>${escapeHtml(stream?.label || row.streamKey || "Stream")} · ${escapeHtml(row.seniority === "ALL" ? "All team" : row.seniority)}</strong><span>${escapeHtml(displaySourceCode(row.facilityKey))} · ${escapeHtml(facilityOverviewByStreamRangeLabel())}</span></header>`;
  };
  const dayMarkup = (row, date, className = "facility-overview-by-stream-day") => {
    const matching = (bySelection.get(row.id) || []).filter((assignment) => assignment.date === date)
      .sort(compareFacilityOverviewAssignmentsByStart);
    const covered = facilityOverviewByStreamCoverageFor(row.facilityKey, date);
    const streamObserved = observedByStreamDate.has(`${row.facilityKey}|${row.streamKey}|${date}`);
    let empty = "";
    if (!covered) empty = `<p class="facility-overview-by-stream-empty">No active roster covers this ED on this date.</p>`;
    else if (!streamObserved) empty = `<p class="facility-overview-by-stream-empty">This stream was not observed on this date.</p>`;
    else empty = `<p class="facility-overview-by-stream-empty">No ${escapeHtml(row.seniority === "ALL" ? "team" : row.seniority)} assignment was found.</p>`;
    const dayLabel = parseDateOnly(date).toLocaleDateString("en-AU", { weekday: "short", day: "numeric", month: "short" });
    const shiftBlocks = facilityOverviewByStreamShiftBlocks(matching);
    return `<section class="${className}">
      <h4><time datetime="${escapeHtml(date)}">${escapeHtml(dayLabel)}</time></h4>
      ${matching.length ? `<div class="facility-overview-by-stream-shifts">${shiftBlocks.map((block) => `<section class="facility-overview-by-stream-shift"><h5>${escapeHtml(block.label)}</h5><div class="facility-overview-by-stream-people">${block.assignments.map((assignment) => `<article class="facility-overview-by-stream-person"><div class="facility-overview-by-stream-person-name"><strong>${escapeHtml(assignment.displayName)}</strong><span class="facility-overview-by-stream-grade">${escapeHtml(facilityOverviewSeniorityAbbreviation(assignment.seniority))}</span></div>${assignment.roleNote ? `<small>${escapeHtml(assignment.roleNote)}</small>` : ""}</article>`).join("")}</div></section>`).join("")}</div>` : `<div class="facility-overview-by-stream-day-empty">${empty}</div>`}
    </section>`;
  };
  if (!visibleDates.length) return `<article class="issue-card"><p>No selected stream assignments were found in this date range.${hiddenCount ? " Turn off “Hide dates without assignments” to inspect covered dates." : ""}</p></article>`;
  if (selected.length > 1) {
    const headers = selected.map((row) => laneHeading(row, "facility-overview-by-stream-comparison-head")).join("");
    const dateRows = visibleDates.map((date) => `<section class="facility-overview-by-stream-comparison-date"><div class="facility-overview-by-stream-comparison-day-grid">${selected.map((row) => dayMarkup(row, date, "facility-overview-by-stream-comparison-day")).join("")}</div></section>`).join("");
    return `${hiddenCount ? `<p class="facility-overview-by-stream-summary">${escapeHtml(`${hiddenCount} date${hiddenCount === 1 ? "" : "s"} without an assignment hidden`)}</p>` : ""}<section class="facility-overview-by-stream-comparison"><div class="facility-overview-by-stream-comparison-head-grid">${headers}</div>${dateRows}</section>`;
  }
  const lanes = selected.map((row) => {
    const days = visibleDates.map((date) => dayMarkup(row, date)).join("");
    return `<section class="facility-overview-by-stream-lane">
      ${laneHeading(row)}
      <div class="facility-overview-by-stream-day-grid">${days}</div>
    </section>`;
  }).filter(Boolean);
  if (!lanes.length) return `<article class="issue-card"><p>No selected stream assignments were found in this date range.${hiddenCount ? " Turn off “Hide dates without assignments” to inspect covered dates." : ""}</p></article>`;
  return `${hiddenCount ? `<p class="facility-overview-by-stream-summary">${escapeHtml(`${hiddenCount} date cell${hiddenCount === 1 ? "" : "s"} without an assignment hidden`)}</p>` : ""}${lanes.join("")}`;
}

async function loadFacilityOverviewByStream() {
  if (!canUseFacilityOverview() || facilityOverviewState.tab !== "by-stream") return;
  initializeFacilityOverviewByStreamState();
  const startDate = facilityOverviewState.byStreamFrom;
  const endDate = facilityOverviewState.byStreamTo;
  const rows = facilityOverviewByStreamDistinctRows().filter((row) => row.facilityKey && row.streamKey && row.seniority);
  if (!startDate || !endDate || endDate < startDate || !rows.length) {
    facilityOverviewState.byStreamContent = `<article class="issue-card"><p>Choose a valid date range and stream selection.</p></article>`;
    renderFacilityOverview();
    return;
  }
  const requestId = facilityOverviewState.byStreamRequestId + 1;
  facilityOverviewState.byStreamRequestId = requestId;
  facilityOverviewState.byStreamLoading = true;
  facilityOverviewState.byStreamContent = `<article class="issue-card"><p>Loading stream coverage…</p></article>`;
  renderFacilityOverview();
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryFacilityOverviewByStream", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        startDate, endDate, selections: rows,
      }),
    });
    const data = await readJsonResponse(response, "Could not load stream coverage.");
    if (facilityOverviewState.byStreamRequestId !== requestId || facilityOverviewState.tab !== "by-stream") return;
    facilityOverviewState.byStreamData = data;
    facilityOverviewState.byStreamCoverage = data.coverage || [];
    facilityOverviewMergeStreamCatalog(facilityOverviewBuildStreamCatalog(data.events || []));
    facilityOverviewState.byStreamContent = facilityOverviewByStreamContentFromData(data);
  } catch (error) {
    if (facilityOverviewState.byStreamRequestId !== requestId) return;
    facilityOverviewState.byStreamContent = `<article class="issue-card"><p>${escapeHtml(error.message || "Stream coverage is unavailable right now.")}</p></article>`;
  }
  facilityOverviewState.byStreamLoading = false;
  renderFacilityOverview();
}

function facilityOverviewTogetherStaffOptions() {
  const activeDoctor = selectedDoctor();
  const activeClaims = currentRosterClaims.map((claim) => ({
    key: claim.key,
    displayName: claim.displayName,
    sourceType: claim.sourceType,
  }));
  const activeViewer = activeDoctor?.key ? [{
    ...activeDoctor,
    displayName: activeDoctor.displayName || currentAccount().realName || activeDoctor.key,
    aliases: Array.isArray(activeDoctor.aliases) && activeDoctor.aliases.length ? activeDoctor.aliases : activeClaims,
    sourceTypes: normalizedDoctorSourceTypes(activeDoctor).length
      ? normalizedDoctorSourceTypes(activeDoctor)
      : [...new Set(activeClaims.map((claim) => claim.sourceType))],
  }] : [];
  return dedupeDoctorOptions([...(availableRosterDoctors || []), ...doctorPickerOptions(), ...activeViewer, ...(facilityOverviewState.togetherPinnedDoctors || [])])
    .map((doctor) => ({ ...doctor, identity: doctorIdentityKey(doctor) }))
    .filter((doctor) => doctor.identity)
    .sort((left, right) => left.displayName.localeCompare(right.displayName));
}

function initializeFacilityOverviewTogetherState() {
  if (!Array.isArray(facilityOverviewState.togetherStaffKeys)) facilityOverviewState.togetherStaffKeys = ["", ""];
  if (!facilityOverviewState.togetherStaffKeys.length) facilityOverviewState.togetherStaffKeys = [""];
  const options = facilityOverviewTogetherStaffOptions();
  const identities = new Set(options.map((doctor) => doctor.identity));
  facilityOverviewState.togetherStaffKeys = facilityOverviewState.togetherStaffKeys.map((key) => identities.has(key) ? key : "");
  if (!facilityOverviewState.togetherStaffKeys[0] && !facilityOverviewState.togetherUserClearedAll) {
    const currentIdentity = doctorIdentityKey(selectedDoctor());
    if (identities.has(currentIdentity)) facilityOverviewState.togetherStaffKeys[0] = currentIdentity;
  }
}

function facilityOverviewTogetherTermOptions() {
  const current = australianTermForDate(new Date());
  const terms = new Map();
  const add = (term) => terms.set(formatDateKey(term.start), { value: formatDateKey(term.start), label: formatAustralianTermLabel(term) });
  for (let year = current.year - 1; year <= current.year + 1; year += 1) {
    for (let termNumber = 1; termNumber <= 4; termNumber += 1) add(buildAustralianTerm(year, termNumber, startMonthIndexForTerm(termNumber)));
  }
  for (const term of facilityOverviewState.staffTerms || []) terms.set(term.value, term);
  return [...terms.values()].sort((left, right) => right.value.localeCompare(left.value));
}

function renderFacilityOverviewTogetherProposal() {
  const options = facilityOverviewTogetherStaffOptions();
  const selected = new Set(facilityOverviewState.togetherStaffKeys.filter(Boolean));
  const facilities = facilityOverviewFacilityOptions();
  const rangeMode = facilityOverviewState.togetherRangeMode === "dates" ? "dates" : "term";
  const terms = facilityOverviewTogetherTermOptions();
  return `
    <section class="facility-overview-together" aria-label="Working together search">
      <div class="facility-overview-together-intro">
        <h3>Select one person to view shifts, or two or more people to find shared shifts. Site can be combined with either option.</h3>
      </div>
      <div class="facility-overview-together-builder">
        <fieldset class="facility-overview-together-section">
          <legend><span>1</span> Staff</legend>
          <div class="facility-overview-together-staff-list">
            ${facilityOverviewState.togetherStaffKeys.map((key, index) => `
              <div class="facility-overview-together-staff-row">
                <label class="field"><span>Staff member ${index + 1}</span>
                  <select data-facility-overview-together-staff="${index}">
                    <option value="">Choose staff…</option>
                    ${options.map((doctor) => {
                      const sourceLabel = normalizedDoctorSourceTypes(doctor).map((source) => displaySourceCode(source)).join(", ");
                      const disabled = selected.has(doctor.identity) && doctor.identity !== key;
                      return `<option value="${escapeHtml(doctor.identity)}" ${doctor.identity === key ? "selected" : ""} ${disabled ? "disabled" : ""}>${escapeHtml(doctor.displayName)}${sourceLabel ? ` — ${escapeHtml(sourceLabel)}` : ""}</option>`;
                    }).join("")}
                  </select>
                </label>
                <button type="button" class="facility-overview-remove-staff" data-facility-overview-together-remove="${index}" aria-label="Remove staff member ${index + 1}"><span aria-hidden="true">🗑</span><span class="sr-only">Remove staff member ${index + 1}</span></button>
              </div>
            `).join("")}
          </div>
          <button type="button" class="facility-overview-add-staff" data-facility-overview-together-add><span aria-hidden="true">+</span> Add another staff member</button>
        </fieldset>
        <fieldset class="facility-overview-together-section">
          <legend><span>2</span> When</legend>
          <div class="facility-overview-segmented" role="radiogroup" aria-label="Search period">
            <label><input type="radio" name="facility-overview-together-range" value="term" data-facility-overview-together-range-mode ${rangeMode === "term" ? "checked" : ""}><span>Term</span></label>
            <label><input type="radio" name="facility-overview-together-range" value="dates" data-facility-overview-together-range-mode ${rangeMode === "dates" ? "checked" : ""}><span>Date range</span></label>
          </div>
          ${rangeMode === "term" ? `
            <label class="field"><span>Term</span><select data-facility-overview-together-term>${terms.map((term) => `<option value="${escapeHtml(term.value)}" ${term.value === facilityOverviewState.togetherTermStart ? "selected" : ""}>${escapeHtml(term.label)}</option>`).join("")}</select></label>
          ` : `
            <div class="facility-overview-together-dates">
              <label class="field"><span>From</span><input type="date" value="${escapeHtml(facilityOverviewState.togetherFrom)}" data-facility-overview-together-date="from"></label>
              <label class="field"><span>To</span><input type="date" value="${escapeHtml(facilityOverviewState.togetherTo)}" data-facility-overview-together-date="to"></label>
            </div>
          `}
        </fieldset>
        <fieldset class="facility-overview-together-section">
          <legend><span>3</span> Site <small>Optional</small></legend>
          <label class="field"><span>ED site</span><select data-facility-overview-together-facility><option value="ALL">All EDs</option>${facilities.map((facility) => `<option value="${escapeHtml(facility)}" ${facility === facilityOverviewState.togetherFacilityKey ? "selected" : ""}>${escapeHtml(displaySourceCode(facility))}</option>`).join("")}</select></label>
          <p class="facility-overview-filter-hint">This narrows either the selected term or date range.</p>
        </fieldset>
      </div>
      <div class="facility-overview-together-output" aria-live="polite">
        ${facilityOverviewState.togetherContent || (facilityOverviewState.togetherHasSearched ? "" : renderFacilityOverviewTogetherEmptyState())}
      </div>
    </section>
  `;
}

function renderFacilityOverviewTogetherEmptyState() {
  const count = facilityOverviewState.togetherStaffKeys.filter(Boolean).length;
  if (count === 1) {
    return `<div class="facility-overview-empty-state"><span aria-hidden="true">↗</span><strong>Loading rostered shifts</strong><p>Add another staff member at any time to find shifts they share.</p></div>`;
  }
  return `<div class="facility-overview-empty-state"><span aria-hidden="true">↗</span><strong>Rostered shifts will appear here</strong><p>Choose a staff member and set the search scope above.</p></div>`;
}

function facilityOverviewTogetherDateRange() {
  if (facilityOverviewState.togetherRangeMode === "dates") {
    return { startDate: facilityOverviewState.togetherFrom, endDate: facilityOverviewState.togetherTo };
  }
  const term = australianTermForDate(parseDateOnly(facilityOverviewState.togetherTermStart));
  return { startDate: formatDateKey(term.start), endDate: formatDateKey(addDays(term.end, -1)) };
}

function refreshFacilityOverviewTogetherAfterFilterChange() {
  facilityOverviewState.togetherHasSearched = false;
  facilityOverviewState.togetherContent = "";
  if (facilityOverviewState.togetherStaffKeys.some(Boolean)) void loadFacilityOverviewTogether();
  else renderFacilityOverview();
}

function facilityOverviewTogetherDoctorKeys(doctor) {
  return [...new Set([
    doctor?.key,
    ...(doctor?.aliases || []).map((alias) => alias.key),
  ].map(normalizeRosterName).filter(Boolean))];
}

async function loadFacilityOverviewTogether() {
  if (!canUseFacilityOverview() || facilityOverviewState.tab !== "together") return;
  const options = facilityOverviewTogetherStaffOptions();
  const selectedDoctors = facilityOverviewState.togetherStaffKeys
    .map((identity) => options.find((doctor) => doctor.identity === identity))
    .filter(Boolean);
  if (!selectedDoctors.length) {
    facilityOverviewState.togetherHasSearched = false;
    facilityOverviewState.togetherContent = `<article class="issue-card"><p>Choose at least one staff member to view rostered shifts.</p></article>`;
    renderFacilityOverview();
    return;
  }
  const { startDate, endDate } = facilityOverviewTogetherDateRange();
  if (!startDate || !endDate || endDate < startDate) {
    facilityOverviewState.togetherHasSearched = false;
    facilityOverviewState.togetherContent = `<article class="issue-card"><p>Choose a valid date range with the end date on or after the start date.</p></article>`;
    renderFacilityOverview();
    return;
  }
  const requestId = facilityOverviewState.requestId + 1;
  facilityOverviewState.requestId = requestId;
  facilityOverviewState.togetherHasSearched = true;
  facilityOverviewState.togetherContent = `<article class="issue-card"><p>${selectedDoctors.length === 1 ? "Finding rostered shifts…" : "Finding shared roster days…"}</p></article>`;
  renderFacilityOverview();
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryFacilityOverviewWorkingTogether",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        startDate,
        endDate,
        doctorKeys: [...new Set(selectedDoctors.flatMap(facilityOverviewTogetherDoctorKeys))],
        sourceTypes: facilityOverviewState.togetherFacilityKey === "ALL" ? [] : [facilityOverviewState.togetherFacilityKey],
      }),
    });
    const data = await readJsonResponse(response, "Could not search shared roster days.");
    if (facilityOverviewState.requestId !== requestId || facilityOverviewState.tab !== "together") return;
    facilityOverviewState.togetherContent = renderFacilityOverviewTogetherResults(data.events || [], selectedDoctors, { startDate, endDate });
  } catch (error) {
    if (facilityOverviewState.requestId !== requestId) return;
    facilityOverviewState.togetherContent = `<article class="issue-card"><p>${escapeHtml(error.message || "Shared roster days are unavailable right now.")}</p></article>`;
  }
  renderFacilityOverview();
}

function renderFacilityOverviewTogetherResults(rows, selectedDoctors, range) {
  const ownerByRosterKey = new Map();
  for (const doctor of selectedDoctors) {
    for (const key of facilityOverviewTogetherDoctorKeys(doctor)) ownerByRosterKey.set(key, doctor.identity);
  }
  const eventsByDoctor = new Map(selectedDoctors.map((doctor) => [doctor.identity, []]));
  for (const row of rows || []) {
    const identity = ownerByRosterKey.get(normalizeRosterName(row?.doctorKey || ""));
    if (!identity || !row?.event || !isRosterShiftEvent(row.event)) continue;
    eventsByDoctor.get(identity).push(serializeEvent(row.event));
  }
  const indexed = selectedDoctors.map((doctor) => ({
    doctor,
    intervals: facilityOverviewWorkingIntervals(eventsByDoctor.get(doctor.identity) || [], range),
  }));
  if (selectedDoctors.length === 1) {
    const matches = facilityOverviewDedupeMatches(facilityOverviewIntervalMatches(indexed, range));
    if (!matches.length) {
      return `<article class="facility-overview-no-match"><span aria-hidden="true">○</span><strong>No rostered shifts found</strong><p>${escapeHtml(selectedDoctors[0].displayName)} has no rostered shifts during this period. Try a wider range or all EDs.</p></article>`;
    }
    return `<section class="facility-overview-match-group"><div class="facility-overview-match-group-head"><h4>${escapeHtml(selectedDoctors[0].displayName)}’s shifts</h4></div>${renderFacilityOverviewTogetherMatchCards(matches, { singlePerson: true })}</section>`;
  }
  const allMatches = facilityOverviewIntervalMatches(indexed, range);
  const allIntervalsByFacility = facilityOverviewUnionIntervalsByFacility(allMatches);
  const pairMatches = facilityOverviewTogetherPairs(indexed)
    .flatMap((pair) => facilityOverviewIntervalMatches(pair, range).flatMap((match) => (
      facilityOverviewSubtractIntervals(match, allIntervalsByFacility.get(match.facility) || [])
        .map((interval) => ({ ...match, ...interval, pairLabel: pair.map((entry) => entry.doctor.displayName).join(" + ") }))
    )));
  const uniqueAllMatches = facilityOverviewDedupeMatches(allMatches);
  const uniquePairMatches = facilityOverviewDedupeMatches(pairMatches);
  if (!uniqueAllMatches.length && !uniquePairMatches.length) {
    return `<article class="facility-overview-no-match"><span aria-hidden="true">○</span><strong>No roster overlaps found</strong><p>${escapeHtml(selectedDoctors.map((doctor) => doctor.displayName).join(", "))} were not working at the same hospital at the same time during this period. Try a wider range or all EDs.</p></article>`;
  }
  if (selectedDoctors.length === 2) return renderFacilityOverviewTogetherMatchCards(uniqueAllMatches);
  const showGroups = uniqueAllMatches.length && uniquePairMatches.length;
  return `
    ${uniqueAllMatches.length ? `${showGroups ? `<section class="facility-overview-match-group facility-overview-match-group-all"><div class="facility-overview-match-group-head"><h4>All selected staff</h4></div>` : ""}${renderFacilityOverviewTogetherMatchCards(uniqueAllMatches)}${showGroups ? "</section>" : ""}` : ""}
    ${uniquePairMatches.length ? `${showGroups ? `<section class="facility-overview-match-group"><div class="facility-overview-match-group-head"><h4>Two-person overlaps</h4></div>` : ""}${renderFacilityOverviewTogetherMatchCards(uniquePairMatches, { showPair: true })}${showGroups ? "</section>" : ""}` : ""}
  `;
}

function facilityOverviewWorkingIntervals(events, range) {
  const rangeStart = new Date(`${range.startDate}T00:00:00`);
  const rangeEnd = new Date(`${range.endDate}T00:00:00`);
  rangeEnd.setDate(rangeEnd.getDate() + 1);
  return (events || []).flatMap((event) => {
    const facility = eventSourceCode(event);
    const start = new Date(event.start);
    let end = new Date(event.end);
    if (!facility || event.allDay || Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return [];
    if (end <= start && extractTimePortion(event.end) === "00:00") end = addDays(end, 1);
    if (end <= start || end <= rangeStart || start >= rangeEnd) return [];
    return [{ facility, start, end, event }];
  });
}

function facilityOverviewTogetherPairs(indexed) {
  const pairs = [];
  for (let left = 0; left < indexed.length - 1; left += 1) {
    for (let right = left + 1; right < indexed.length; right += 1) pairs.push([indexed[left], indexed[right]]);
  }
  return pairs;
}

function facilityOverviewIntervalMatches(indexed, range) {
  if (!indexed.length || indexed.some((entry) => !entry.intervals.length)) return [];
  const rangeStart = new Date(`${range.startDate}T00:00:00`);
  const rangeEnd = new Date(`${range.endDate}T00:00:00`);
  rangeEnd.setDate(rangeEnd.getDate() + 1);
  const matches = [];
  const visit = (index, facility, start, end, people) => {
    if (index === indexed.length) {
      if (start < end) matches.push({ facility, start, end, people });
      return;
    }
    for (const interval of indexed[index].intervals) {
      if (facility && interval.facility !== facility) continue;
      const nextStart = start && start > interval.start ? start : interval.start;
      const nextEnd = end && end < interval.end ? end : interval.end;
      if (nextStart >= nextEnd || nextEnd <= rangeStart || nextStart >= rangeEnd) continue;
      visit(index + 1, facility || interval.facility, nextStart, nextEnd, [...people, { doctor: indexed[index].doctor, event: interval.event }]);
    }
  };
  visit(0, "", null, null, []);
  return matches.sort(facilityOverviewCompareMatches);
}

function facilityOverviewUnionIntervalsByFacility(matches) {
  const grouped = new Map();
  for (const match of matches || []) {
    if (!grouped.has(match.facility)) grouped.set(match.facility, []);
    grouped.get(match.facility).push({ start: match.start, end: match.end });
  }
  for (const [facility, intervals] of grouped) {
    const merged = [];
    for (const interval of intervals.sort((left, right) => left.start - right.start)) {
      const previous = merged[merged.length - 1];
      if (previous && interval.start <= previous.end) previous.end = new Date(Math.max(previous.end, interval.end));
      else merged.push({ ...interval });
    }
    grouped.set(facility, merged);
  }
  return grouped;
}

function facilityOverviewSubtractIntervals(match, exclusions) {
  let segments = [{ start: match.start, end: match.end }];
  for (const exclusion of exclusions || []) {
    segments = segments.flatMap((segment) => {
      if (exclusion.end <= segment.start || exclusion.start >= segment.end) return [segment];
      const remaining = [];
      if (exclusion.start > segment.start) remaining.push({ start: segment.start, end: new Date(Math.min(segment.end, exclusion.start)) });
      if (exclusion.end < segment.end) remaining.push({ start: new Date(Math.max(segment.start, exclusion.end)), end: segment.end });
      return remaining.filter((item) => item.start < item.end);
    });
  }
  return segments;
}

function facilityOverviewDedupeMatches(matches) {
  const unique = new Map();
  for (const match of matches || []) {
    const key = [match.facility, match.start.toISOString(), match.end.toISOString(), match.people.map((person) => person.doctor.identity).join("|"), match.pairLabel || ""].join("~");
    if (!unique.has(key)) unique.set(key, match);
  }
  return [...unique.values()].sort(facilityOverviewCompareMatches);
}

function facilityOverviewCompareMatches(left, right) {
  return left.start - right.start || left.facility.localeCompare(right.facility) || left.end - right.end;
}

function renderFacilityOverviewTogetherMatchCards(matches, options = {}) {
  const canOpenStaffCalendars = canUseCreatorDoctorSwitcher();
  return `<div class="facility-overview-together-match-list">${matches.map((match) => `
    <article class="facility-overview-together-match">
      ${options.showPair ? `<div class="facility-overview-match-pair">${match.people.map((person, index) => `${index ? `<span aria-hidden="true"> + </span>` : ""}${renderFacilityOverviewStaffName({ ...person.doctor, sourceType: match.facility }, { directCalendar: canOpenStaffCalendars })}`).join("")}</div>` : ""}
      <div class="facility-overview-together-match-date"><time datetime="${escapeHtml(formatDateKey(match.start))}"><strong>${escapeHtml(new Intl.DateTimeFormat("en-AU", { weekday: "short" }).format(match.start))}</strong><span>${escapeHtml(formatDate(formatDateKey(match.start)))}</span></time><span class="facility-overview-site-pill">${escapeHtml(displaySourceCode(match.facility))}</span></div>
      <span class="facility-overview-overlap-time">${options.singlePerson ? "Shift" : "Overlap"} ${escapeHtml(facilityOverviewFormatOverlap(match.start, match.end))}</span>
      <div class="facility-overview-together-match-people">${match.people.map((person) => `<div>${renderFacilityOverviewStaffName({ ...person.doctor, sourceType: match.facility }, { directCalendar: canOpenStaffCalendars })}<span>${escapeHtml(renderFacilityOverviewStream([person.event], { includeTimes: true }) || "Rostered shift")}${renderFacilityOverviewSeniorityLink(person.event?.seniority, { sourceType: match.facility, date: formatDateKey(match.start) })}</span></div>`).join("")}</div>
    </article>
  `).join("")}</div>`;
}

function facilityOverviewFormatOverlap(start, end) {
  const format = (value) => new Intl.DateTimeFormat("en-AU", { hour: "2-digit", minute: "2-digit", hourCycle: "h23" }).format(value);
  const startsAndEndsSameDay = formatDateKey(start) === formatDateKey(end);
  return `${format(start)} – ${startsAndEndsSameDay ? format(end) : `${formatDate(formatDateKey(end))} ${format(end)}`}`;
}

async function loadFacilityOverviewOnShift() {
  if (!canUseFacilityOverview() || facilityOverviewState.tab !== "on-shift") return;
  const requestId = facilityOverviewState.requestId + 1;
  facilityOverviewState.requestId = requestId;
  facilityOverviewState.staffData = null;
  facilityOverviewState.onShiftData = null;
  facilityOverviewState.contactList = null;
  facilityOverviewState.content = `<article class="issue-card"><p>Loading rostered staff…</p></article>`;
  renderFacilityOverview();
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryFacilityOverviewOnShift",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        facilityKey: facilityOverviewState.facilityKey,
        date: facilityOverviewState.date,
        includeClinicalSupport: facilityOverviewState.includeClinicalSupport === true,
      }),
    });
    const data = await readJsonResponse(response, "Could not load the ED overview.");
    if (facilityOverviewState.requestId !== requestId || facilityOverviewState.tab !== "on-shift") return;
    facilityOverviewState.onShiftData = data.events || [];
    facilityOverviewState.contactList = data.contactList || null;
    facilityOverviewState.content = renderFacilityOverviewOnShiftResults(facilityOverviewState.onShiftData);
  } catch (error) {
    if (facilityOverviewState.requestId !== requestId) return;
    facilityOverviewState.content = `<article class="issue-card"><p>${escapeHtml(error.message || "The ED overview is unavailable right now.")}</p></article>`;
  }
  renderFacilityOverview();
}

function renderFacilityOverviewOnShiftResults(rows) {
  const canUseStaffActions = canUseFacilityOverview();
  const termStart = formatDateKey(australianTermForDate(parseDateOnly(facilityOverviewState.date)).start);
  const people = new Map();
  for (const row of rows) {
    const event = row?.event;
    if (!event || !isRosterShiftEvent(event)) continue;
    const doctorKey = String(row.doctorKey || "").trim();
    const groupKey = `${String(row.sourceType || facilityOverviewState.facilityKey || "").toUpperCase()}|${doctorKey}`;
    if (!doctorKey) continue;
    const seniority = facilityOverviewDetectedSeniority(event, row.seniority || "Unknown");
    const entry = people.get(groupKey) || {
      doctorKey, displayName: String(row.displayName || doctorKey), sourceType: String(row.sourceType || facilityOverviewState.facilityKey || "").toLowerCase(),
      seniority, events: [], markers: new Set(),
    };
    if (entry.seniority === "Unknown" && seniority !== "Unknown") entry.seniority = seniority;
    const marker = `${event.title}|${event.start}|${event.end}|${event.rawValue || ""}`;
    if (!entry.markers.has(marker)) {
      entry.markers.add(marker);
      entry.events.push({ ...event, seniority: seniority === "Unknown" ? entry.seniority : seniority });
    }
    people.set(groupKey, entry);
  }
  const assignments = [...people.values()].flatMap((person) => person.events.map((event) => {
    const base = buildWhoAssignment({ key: person.doctorKey, displayName: person.displayName }, {}, event);
    return base ? { ...base, person, event } : null;
  })).filter(Boolean);
  if (!assignments.length) return `<article class="issue-card"><p>No recognised working shifts were found for this ED and date.</p></article>`;
  const contactMatches = attachContactAllocations(assignments, facilityOverviewState.contactList?.contacts || []);
  const periods = new Map();
  for (const assignment of contactMatches.assignments) {
    if (!periods.has(assignment.period)) periods.set(assignment.period, []);
    periods.get(assignment.period).push(assignment);
  }
  return `${renderFacilityOverviewContactListStatus(contactMatches)}${["AM", "PM", "Night"].filter((period) => periods.has(period)).map((period) => `
    <section class="facility-overview-period"><h3>${period}</h3><div class="facility-overview-staff-grid">${renderFacilityOverviewOnShiftPeriod(periods.get(period), { canUseStaffActions, termStart, period })}</div></section>
  `).join("")}`;
}

function renderFacilityOverviewOnShiftPeriod(assignments, options = {}) {
  const active = (assignments || []).filter((assignment) => !assignment.contactDisplacedBy?.length);
  const rosterOnly = (assignments || []).filter((assignment) => assignment.contactDisplacedBy?.length);
  const content = facilityOverviewIsDdhPeriod(active)
    ? renderFacilityOverviewDdhOnShiftPeriod(active, options)
    : facilityOverviewIsMmcNightPeriod(active, options)
      ? renderFacilityOverviewMmcNightPeriod(active, options)
      : renderFacilityOverviewGenericOnShiftPeriod(active, options);
  return `${content}${rosterOnly.length ? renderFacilityOverviewRosterOnlyCard(rosterOnly, options) : ""}`;
}

function facilityOverviewIsDdhPeriod(assignments) {
  return (assignments || []).some((assignment) => String(assignment.source || assignment.person?.sourceType || "").trim().toUpperCase() === "DDH");
}

function facilityOverviewIsMmcNightPeriod(assignments, options = {}) {
  return options.period === "Night"
    && (assignments || []).some((assignment) => String(assignment.source || assignment.person?.sourceType || "").trim().toUpperCase() === "MMC");
}

function renderFacilityOverviewDdhOnShiftPeriod(assignments, options = {}) {
  if (options.period === "Night") return renderFacilityOverviewDdhNightPeriod(assignments, options);
  const handled = new Set();
  const teamAssignments = (team) => (assignments || []).filter((assignment) => String(assignment.team || "").trim().toLowerCase() === team);
  const orange = teamAssignments("orange");
  const silver = teamAssignments("silver");
  const fastTrack = teamAssignments("fast track");
  const avao = teamAssignments("avao");
  const ssu = teamAssignments("ssu");
  const ssuSms = ssu.filter((assignment) => String(assignment.person?.seniority || assignment.role || "").trim().toUpperCase() === "SMS");
  const ssuInternHmo = ssu.filter((assignment) => !ssuSms.includes(assignment));
  for (const assignment of [...orange, ...silver, ...fastTrack, ...avao, ...ssu]) handled.add(assignment);

  const renderPlacedCard = (stream, items, column) => items.length
    ? renderFacilityOverviewStreamCard(stream, items, { ...options, cardClass: `facility-overview-ddh-column-${column}` })
    : "";
  const mainRow = [
    renderPlacedCard("Orange", orange, 1),
    renderPlacedCard("Silver", silver, 2),
    renderPlacedCard("Fast Track", fastTrack, 3),
  ].join("");
  const supportRow = options.period === "AM"
    ? [
      renderPlacedCard("AVAO", avao, 1),
      renderPlacedCard("SSU", ssu, 2),
    ].join("")
    : [
      renderPlacedCard("AVAO", avao, 1),
      renderPlacedCard("SSU", ssuSms, 2),
      renderPlacedCard("SSU", ssuInternHmo, ssuSms.length ? 3 : 2),
    ].join("");
  const remaining = (assignments || []).filter((assignment) => !handled.has(assignment));
  return `${mainRow ? `<div class="facility-overview-ddh-row facility-overview-ddh-main-row">${mainRow}</div>` : ""}
    ${supportRow ? `<div class="facility-overview-ddh-row facility-overview-ddh-support-row">${supportRow}</div>` : ""}
    ${remaining.length ? renderFacilityOverviewGenericOnShiftPeriod(remaining, options) : ""}`;
}

function renderFacilityOverviewDdhNightPeriod(assignments, options = {}) {
  const isSsu = (assignment) => ["ssu", "night ssu"].includes(String(assignment.team || "").trim().toLowerCase());
  const ssu = (assignments || []).filter(isSsu);
  const nonSsu = (assignments || []).filter((assignment) => !isSsu(assignment));
  const seniorRegistrars = nonSsu.filter((assignment) => normalizeWhoRole(assignment.role) === "SR");
  const mainTeam = nonSsu.filter((assignment) => !seniorRegistrars.includes(assignment));
  return [
    ["Night SR", seniorRegistrars],
    ["Main team", mainTeam],
    ["SSU team", ssu],
  ].filter(([, items]) => items.length)
    .map(([label, items]) => renderFacilityOverviewStreamCard(label, items, options))
    .join("");
}

function renderFacilityOverviewMmcNightPeriod(assignments, options = {}) {
  const isTeam = (assignment, labels) => labels.includes(String(facilityOverviewEffectiveTeam(assignment) || "").trim().toLowerCase());
  const hub = (assignments || []).filter((assignment) => isTeam(assignment, ["hub", "night hub"]));
  const ssu = (assignments || []).filter((assignment) => isTeam(assignment, ["ssu", "night ssu"]));
  const assignedToDedicatedTeam = new Set([...hub, ...ssu]);
  const remaining = (assignments || []).filter((assignment) => !assignedToDedicatedTeam.has(assignment));
  const seniorRegistrars = remaining.filter((assignment) => normalizeWhoRole(assignment.role) === "SR");
  const mainTeam = remaining.filter((assignment) => !seniorRegistrars.includes(assignment));
  const cardOptions = { ...options, showSpecialTimes: false };
  return [
    ["Night SR", seniorRegistrars],
    ["Hub", hub],
    ["SSU", ssu],
    ["Main team", mainTeam],
  ].filter(([, items]) => items.length)
    .map(([label, items]) => renderFacilityOverviewStreamCard(label, items, cardOptions))
    .join("");
}

function renderFacilityOverviewGenericOnShiftPeriod(assignments, options = {}) {
  const streamed = new Map();
  const unstreamed = new Map();
  for (const assignment of assignments || []) {
    const team = facilityOverviewEffectiveTeam(assignment);
    const isStreamed = facilityOverviewIsMeaningfulStream(team, assignment.source || assignment.person?.sourceType);
    const target = isStreamed ? streamed : unstreamed;
    const key = isStreamed ? team : String(assignment.person?.seniority || assignment.role || "Unknown");
    if (!target.has(key)) target.set(key, []);
    target.get(key).push(assignment);
  }
  const streamCards = [...streamed.entries()]
    .sort(([left, itemsLeft], [right, itemsRight]) => whoTeamRank(left, itemsLeft[0]?.source || "") - whoTeamRank(right, itemsRight[0]?.source || "") || left.localeCompare(right))
    .map(([stream, items]) => renderFacilityOverviewStreamCard(stream, items, options));
  const seniorityCards = [...unstreamed.entries()]
    .sort(([left], [right]) => compareFacilityOverviewSeniorities(left, right))
    .map(([seniority, items]) => renderFacilityOverviewUnstreamedCard(seniority, items, options));
  return [...streamCards, ...seniorityCards].join("");
}

function facilityOverviewEffectiveTeam(assignment) {
  return String(assignment?.contactAllocation?.streamLabel || assignment?.team || "").trim();
}

function renderFacilityOverviewRosterOnlyCard(assignments, options = {}) {
  return `<article class="issue-card facility-overview-staff-card facility-overview-roster-only-card">
    <strong>Roster only</strong>
    ${renderFacilityOverviewOnShiftNames(assignments, { ...options, showContactDiscrepancy: true })}
  </article>`;
}

function facilityOverviewIsMeaningfulStream(team, source = "") {
  const value = String(team || "").trim().toLowerCase();
  const sourceCode = String(source || "").trim().toUpperCase();
  if (sourceCode === "MMC" && ["am", "pm", "am shift", "pm shift", "night", "night shift", "shift"].includes(value)) return false;
  return Boolean(value && ![
    "other", "float", "rover", "shift", "clinical support", "cs", "cso", "sms", "cmo", "senior registrar",
    "transitional/intermediate registrar", "junior registrar", "hmo", "intern", "unknown",
  ].includes(value));
}

function renderFacilityOverviewStreamCard(stream, assignments, options = {}) {
  return `<article class="issue-card facility-overview-staff-card facility-overview-stream-card${options.cardClass ? ` ${options.cardClass}` : ""}">
    <strong class="facility-overview-stream-card-title">${escapeHtml(stream)}</strong>
    ${renderFacilityOverviewOnShiftNames(assignments, options)}
  </article>`;
}

function renderFacilityOverviewUnstreamedCard(seniority, assignments, options = {}) {
  return `<article class="issue-card facility-overview-staff-card facility-overview-unstreamed-card">
    <strong>${renderFacilityOverviewSeniorityLink(seniority, { sourceType: assignments[0]?.person?.sourceType, date: facilityOverviewState.date }) || escapeHtml(seniority)}</strong>
    ${renderFacilityOverviewOnShiftNames(assignments, options)}
  </article>`;
}

function renderFacilityOverviewOnShiftNames(assignments, options = {}) {
  const byPerson = new Map();
  for (const assignment of assignments || []) {
    const person = assignment.person;
    if (!person) continue;
    const existing = byPerson.get(person.doctorKey) || { person, specialTimes: new Set() };
    if (assignment.specialTime) existing.specialTimes.add(assignment.specialTime);
    byPerson.set(person.doctorKey, existing);
  }
  return `<div class="facility-overview-on-shift-names">${[...byPerson.values()].sort((left, right) => compareFacilityOverviewPeople(left.person, right.person)).map(({ person, specialTimes }) => {
    const sourceAssignment = (assignments || []).find((assignment) => assignment.person?.doctorKey === person.doctorKey);
    const allocation = sourceAssignment?.contactAllocation;
    const discrepancy = options.showContactDiscrepancy && sourceAssignment?.contactDisplacedBy?.length
      ? `<small class="facility-overview-contact-discrepancy">Not on live allocation</small>` : "";
    return `<div>${renderFacilityOverviewStaffName(person, { ...options, seniority: person.seniority })}${renderFacilityOverviewOnShiftSeniority(person, options)}${allocation ? renderFacilityOverviewContactAllocation(allocation) : ""}${discrepancy}${options.showSpecialTimes !== false && specialTimes.size ? `<small>${escapeHtml([...specialTimes].join(" · "))}</small>` : ""}</div>`;
  }).join("")}</div>`;
}

function renderFacilityOverviewContactAllocation(allocation) {
  const phone = String(allocation?.phone || "").trim();
  if (!phone) return `<span class="facility-overview-contact-number is-empty">No phone recorded</span>`;
  const dial = phone.replace(/[^0-9+]/g, "");
  return `<a class="facility-overview-contact-number" href="tel:${escapeHtml(dial)}" title="Allocated contact number">${escapeHtml(phone)}</a>`;
}

function renderFacilityOverviewContactListStatus(matches) {
  const contactList = facilityOverviewState.contactList;
  if (!contactList?.status || contactList.status === "unavailable") return "";
  if (contactList.status !== "available") return `<p class="facility-overview-contact-status">Live contact allocation is not available for this date.</p>`;
  const received = contactList.providerModifiedAt || contactList.receivedAt || "";
  const freshness = received ? ` · updated ${formatFacilityOverviewContactTime(received)}` : "";
  const unresolved = matches.unmatched || [];
  const review = unresolved.length
    ? isViewingCreatorAccount()
      ? `<details class="facility-overview-contact-review"><summary>${unresolved.length} allocation${unresolved.length === 1 ? "" : "s"} need review</summary>${unresolved.map((contact) => `<div>${escapeHtml(contact.shift)} · ${escapeHtml(contact.role)} · ${escapeHtml(contact.name)}${contact.phone ? ` · ${escapeHtml(contact.phone)}` : ""}${contact.reviewReason ? ` · ${escapeHtml(contact.reviewReason)}` : ""}</div>`).join("")}</details>`
      : `<span> · ${unresolved.length} allocation${unresolved.length === 1 ? "" : "s"} need review</span>`
    : "";
  return `<div class="facility-overview-contact-status"><span>Live contact allocations for ${escapeHtml(contactList.sourceDate)}${freshness} · ${matches.matchedCount} matched</span>${review}</div>`;
}

function formatFacilityOverviewContactTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.valueOf())) return String(value);
  return new Intl.DateTimeFormat("en-AU", { hour: "2-digit", minute: "2-digit", hourCycle: "h23" }).format(date);
}

function renderFacilityOverviewOnShiftSeniority(person, options = {}) {
  const label = facilityOverviewCompactSeniorityLabel(person?.seniority);
  if (!label) return "";
  const target = { ...person, seniority: person.seniority, termStart: options.termStart };
  if (!isViewingCreatorAccount()) return `<span class="facility-overview-on-shift-seniority">${escapeHtml(label)}</span>`;
  const key = facilityOverviewStaffActionMenuKey(target);
  return `<button type="button" class="facility-overview-on-shift-seniority facility-overview-on-shift-seniority-trigger" data-facility-overview-staff-menu="${escapeHtml(key)}" data-facility-overview-staff-source="${escapeHtml(person.sourceType)}" data-facility-overview-staff-key="${escapeHtml(person.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(person.displayName)}" data-facility-overview-staff-seniority="${escapeHtml(person.seniority)}" data-facility-overview-staff-term-start="${escapeHtml(options.termStart || "")}" aria-label="Edit ${escapeHtml(person.displayName)}'s designation">${escapeHtml(label)}</button>`;
}

function refreshFacilityOverviewStaffActionContent() {
  if (facilityOverviewState.tab === "on-shift" && Array.isArray(facilityOverviewState.onShiftData)) {
    facilityOverviewState.content = renderFacilityOverviewOnShiftResults(facilityOverviewState.onShiftData);
  } else if (facilityOverviewState.tab === "staff") {
    refreshFacilityOverviewStaffContent();
  }
}

async function loadFacilityOverviewStaff() {
  if (!canUseFacilityOverview() || facilityOverviewState.tab !== "staff") return;
  const term = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date())));
  facilityOverviewState.staffTermStart = formatDateKey(term.start);
  const requestId = facilityOverviewState.requestId + 1;
  facilityOverviewState.requestId = requestId;
  facilityOverviewState.staffContent = `<article class="issue-card"><p>Loading ED staff…</p></article>`;
  renderFacilityOverview();
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "queryFacilityOverviewStaff", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        facilityKey: facilityOverviewState.facilityKey === "ALL" ? "all" : facilityOverviewState.facilityKey,
        termStart: facilityOverviewState.staffTermStart, termEnd: formatDateKey(addDays(term.end, -1)),
      }),
    });
    const data = await readJsonResponse(response, "Could not load ED staff.");
    if (facilityOverviewState.requestId !== requestId || facilityOverviewState.tab !== "staff") return;
    facilityOverviewState.staffTerms = facilityOverviewTermsFromCoverage(data.coverage || []);
    facilityOverviewState.staffData = data;
    facilityOverviewState.staffContent = renderFacilityOverviewStaffResults(data, term);
  } catch (error) {
    if (facilityOverviewState.requestId !== requestId) return;
    facilityOverviewState.staffContent = `<article class="issue-card"><p>${escapeHtml(error.message || "The ED staff list is unavailable right now.")}</p></article>`;
  }
  renderFacilityOverview();
  focusFacilityOverviewStaffSection();
}

function refreshFacilityOverviewStaffContent() {
  if (!facilityOverviewState.staffData) return;
  const term = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date())));
  facilityOverviewState.staffContent = renderFacilityOverviewStaffResults(facilityOverviewState.staffData, term);
}

function facilityOverviewStaffMultiSelectMemberKey({ sourceType = "", doctorKey = "" } = {}) {
  return `${String(sourceType || "").toLowerCase()}|${normalizeRosterName(doctorKey)}`;
}

function facilityOverviewStaffMultiSelectIsActive(sectionKey) {
  return Boolean(sectionKey) && facilityOverviewState.staffMultiSelectSection === sectionKey;
}

function facilityOverviewStaffMultiSelectIsSelected({ sectionKey = "", sourceType = "", doctorKey = "" } = {}) {
  return facilityOverviewStaffMultiSelectIsActive(sectionKey)
    && facilityOverviewState.staffMultiSelectMembers.has(facilityOverviewStaffMultiSelectMemberKey({ sourceType, doctorKey }));
}

function facilityOverviewStaffMultiSelectSectionForControl(control) {
  const sourceType = String(control?.dataset?.facilityOverviewStaffSource || "").toLowerCase();
  const doctorKey = control?.dataset?.facilityOverviewStaffKey || "";
  const sectionKey = sourceType ? `${sourceType}:Unknown` : "";
  return facilityOverviewStaffMultiSelectIsSelected({ sectionKey, sourceType, doctorKey }) ? sectionKey : "";
}

function openFacilityOverviewStaffBulkSeniorityMenu({ sectionKey = "", x = 8, y = 8 } = {}) {
  if (!sectionKey) return;
  facilityOverviewState.staffActionMenu = null;
  facilityOverviewState.staffDesignationMenu = null;
  facilityOverviewState.staffSeniorityMenu = null;
  facilityOverviewState.staffBulkSeniorityMenu = { sectionKey, x: Math.max(8, Math.round(x || 0)), y: Math.max(8, Math.round(y || 0)) };
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport(sectionKey);
}

function activateFacilityOverviewStaffMultiSelect(sectionKey) {
  if (!sectionKey || !isViewingCreatorAccount()) return;
  facilityOverviewState.staffMultiSelectSection = sectionKey;
  facilityOverviewState.staffMultiSelectMembers = new Map();
  facilityOverviewState.staffBulkSeniorityMenu = null;
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport(sectionKey);
}

function clearFacilityOverviewStaffMultiSelect(options = {}) {
  const wasActive = Boolean(facilityOverviewState.staffMultiSelectSection || facilityOverviewState.staffMultiSelectMembers.size || facilityOverviewState.staffBulkSeniorityMenu);
  facilityOverviewState.staffMultiSelectSection = "";
  facilityOverviewState.staffMultiSelectMembers = new Map();
  facilityOverviewState.staffBulkSeniorityMenu = null;
  facilityOverviewState.staffMultiSelectSaving = false;
  if (wasActive && options.render !== false && facilityOverviewState.tab === "staff") {
    refreshFacilityOverviewStaffContent();
    renderFacilityOverviewStaffBodyPreservingViewport(options.sectionKey || "");
  }
}

function toggleFacilityOverviewStaffMultiSelectMember({ sectionKey = "", sourceType = "", doctorKey = "", displayName = "" } = {}) {
  if (!facilityOverviewStaffMultiSelectIsActive(sectionKey) || !sourceType || !doctorKey || facilityOverviewState.staffMultiSelectSaving) return;
  const key = facilityOverviewStaffMultiSelectMemberKey({ sourceType, doctorKey });
  if (facilityOverviewState.staffMultiSelectMembers.has(key)) facilityOverviewState.staffMultiSelectMembers.delete(key);
  else facilityOverviewState.staffMultiSelectMembers.set(key, { sourceType, doctorKey, displayName });
  facilityOverviewState.staffBulkSeniorityMenu = null;
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport(sectionKey);
}

function renderFacilityOverviewStaffBodyPreservingViewport(sectionKey = "") {
  if (!facilityOverviewBody || facilityOverviewState.tab !== "staff") return;
  const scrollTop = facilityOverviewBody.scrollTop;
  const previousAnchor = sectionKey
    ? [...facilityOverviewBody.querySelectorAll("[data-facility-overview-staff-section]")].find((element) => element.dataset.facilityOverviewStaffSection === sectionKey)
    : null;
  const bodyTop = facilityOverviewBody.getBoundingClientRect().top;
  const anchorOffset = previousAnchor ? previousAnchor.getBoundingClientRect().top - bodyTop : null;
  renderFacilityOverviewStaffBody();
  window.requestAnimationFrame(() => {
    if (!facilityOverviewBody) return;
    const nextAnchor = sectionKey
      ? [...facilityOverviewBody.querySelectorAll("[data-facility-overview-staff-section]")].find((element) => element.dataset.facilityOverviewStaffSection === sectionKey)
      : null;
    if (nextAnchor && anchorOffset !== null) {
      facilityOverviewBody.scrollTop += nextAnchor.getBoundingClientRect().top - facilityOverviewBody.getBoundingClientRect().top - anchorOffset;
    } else {
      facilityOverviewBody.scrollTop = scrollTop;
    }
  });
}

function applyFacilityOverviewStaffSeniorityOverrides(overrides) {
  if (!Array.isArray(overrides) || !overrides.length) return;
  const byPerson = new Map((facilityOverviewState.staffData?.seniorityOverrides || []).map((override) => [`${override.sourceType}|${override.doctorKey}`, override]));
  for (const override of overrides) {
    if (!override?.sourceType || !override?.doctorKey) continue;
    byPerson.set(`${override.sourceType}|${override.doctorKey}`, override);
  }
  if (facilityOverviewState.staffData) {
    facilityOverviewState.staffData.seniorityOverrides = [...byPerson.values()];
    refreshFacilityOverviewStaffContent();
    if (facilityOverviewState.tab === "staff") renderFacilityOverviewStaffBodyPreservingViewport(facilityOverviewState.staffMultiSelectSection);
  }
  if (Array.isArray(facilityOverviewState.onShiftData)) {
    const byKey = new Map(overrides.map((override) => [`${override.sourceType}|${override.doctorKey}`, override]));
    facilityOverviewState.onShiftData = facilityOverviewState.onShiftData.map((row) => {
      const override = byKey.get(`${row.sourceType}|${row.doctorKey}`);
      if (!override || override.useRosterSeniority) return row;
      return { ...row, seniority: override.seniority, seniorityOverride: override, event: { ...row.event, seniority: override.seniority, facilitySeniorityOverride: true } };
    });
    if (facilityOverviewState.tab === "on-shift") {
      facilityOverviewState.content = renderFacilityOverviewOnShiftResults(facilityOverviewState.onShiftData);
      renderFacilityOverview();
    }
  }
}

function applyFacilityOverviewStaffDesignation(designation, { clear = false } = {}) {
  if (!facilityOverviewState.staffData || !designation?.sourceType || !designation?.doctorKey) return;
  const current = facilityOverviewState.staffData.designations || [];
  const key = `${designation.sourceType}|${designation.doctorKey}`;
  facilityOverviewState.staffData.designations = clear
    ? current.filter((entry) => `${entry.sourceType}|${entry.doctorKey}` !== key)
    : [...current.filter((entry) => `${entry.sourceType}|${entry.doctorKey}` !== key), designation];
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport();
}

function facilityOverviewTermsFromCoverage(coverage) {
  const terms = new Map();
  const add = (term) => terms.set(formatDateKey(term.start), { value: formatDateKey(term.start), label: formatAustralianTermLabel(term) });
  add(australianTermForDate(new Date()));
  for (const item of coverage || []) {
    const start = String(item?.startDate || "").slice(0, 10);
    const end = String(item?.endDate || "").slice(0, 10);
    if (!start || !end) continue;
    let term = australianTermForDate(parseDateOnly(start));
    const last = australianTermForDate(parseDateOnly(end));
    while (term.year < last.year || (term.year === last.year && term.termNumber <= last.termNumber)) {
      add(term);
      term = nextAustralianTerm(term);
    }
  }
  return [...terms.values()].sort((left, right) => right.value.localeCompare(left.value));
}

function renderFacilityOverviewStaffResults(data, term) {
  const canUseStaffActions = canUseFacilityOverview();
  const designations = new Map((data.designations || []).map((designation) => [`${designation.sourceType}|${designation.doctorKey}`, designation]));
  const seniorityOverrides = new Map((data.seniorityOverrides || []).map((override) => [`${override.sourceType}|${override.doctorKey}`, override]));
  const byPerson = new Map();
  for (const designation of data.designations || []) {
    const key = `${designation.sourceType}|${designation.doctorKey}`;
    if (!designation?.doctorKey || byPerson.has(key)) continue;
    byPerson.set(key, {
      sourceType: designation.sourceType, doctorKey: designation.doctorKey, displayName: designation.displayName || designation.doctorKey,
      membershipGrades: new Set([facilityOverviewNormalizeSeniority(designation.seniority || "Unknown")]), eventGrades: [], events: [], eventMarkers: new Set(), coverageStarts: [], coverageEnds: [], membershipSources: new Set(["designation"]),
    });
  }
  for (const member of data.members || []) {
    const key = `${member.sourceType}|${member.doctorKey}`;
    const entry = byPerson.get(key) || {
      sourceType: member.sourceType, doctorKey: member.doctorKey, displayName: member.displayName,
      membershipGrades: new Set(), eventGrades: [], events: [], eventMarkers: new Set(), coverageStarts: [], coverageEnds: [], membershipSources: new Set(),
    };
    if (member.seniority) entry.membershipGrades.add(facilityOverviewNormalizeSeniority(member.seniority));
    if (member.coverageStart) entry.coverageStarts.push(member.coverageStart);
    if (member.coverageEnd && member.membershipSource !== "sms-continuity") entry.coverageEnds.push(member.coverageEnd);
    entry.membershipSources.add(member.membershipSource || "roster");
    byPerson.set(key, entry);
  }
  for (const row of data.events || []) {
    const key = `${row.sourceType}|${row.doctorKey}`;
    const entry = byPerson.get(key) || { sourceType: row.sourceType, doctorKey: row.doctorKey, displayName: row.displayName, membershipGrades: new Set(), eventGrades: [], events: [], eventMarkers: new Set(), coverageStarts: [], coverageEnds: [], membershipSources: new Set(["event"]) };
    if (row.seniority) entry.eventGrades.push({ grade: facilityOverviewDetectedSeniority(row.event, row.seniority), date: String(row.event?.start || "").slice(0, 10) });
    const marker = `${row.event?.title || ""}|${row.event?.start || ""}|${row.event?.end || ""}|${row.event?.rawValue || ""}`;
    if (row.event && isRosterShiftEvent(row.event) && !isClinicalSupportEvent(row.event) && !entry.eventMarkers.has(marker)) {
      entry.eventMarkers.add(marker);
      entry.events.push(row.event);
    }
    byPerson.set(key, entry);
  }
  const query = facilityOverviewState.staffQuery.trim().toLocaleLowerCase();
  const panels = new Map();
  for (const person of byPerson.values()) {
    if (query && !person.displayName.toLocaleLowerCase().includes(query)) continue;
    const latestKnownEventGrade = person.eventGrades
      .map((entry) => ({ ...entry, grade: facilityOverviewNormalizeSeniority(entry.grade) }))
      .filter((entry) => entry.grade && entry.grade !== "Unknown")
      .sort((left, right) => right.date.localeCompare(left.date))[0]?.grade;
    const knownMembershipGrade = [...person.membershipGrades]
      .map(facilityOverviewNormalizeSeniority)
      .find((grade) => grade && grade !== "Unknown");
    // FindMyShift staff-list groups describe the clinician's current grade.
    // Use that term membership before a dated shift label so ED Staff, On
    // shift and Who all report the same current designation.
    person.seniority = knownMembershipGrade || latestKnownEventGrade || "Unknown";
    person.multipleGrades = false;
    person.designation = designations.get(`${person.sourceType}|${person.doctorKey}`) || null;
    person.seniorityOverride = seniorityOverrides.get(`${person.sourceType}|${person.doctorKey}`) || null;
    if (person.seniorityOverride && !person.seniorityOverride.useRosterSeniority) {
      person.seniority = facilityOverviewNormalizeSeniority(person.seniorityOverride.seniority);
      person.multipleGrades = false;
    }
    if (!panels.has(person.sourceType)) panels.set(person.sourceType, []);
    panels.get(person.sourceType).push(person);
  }
  if (!panels.size) return `<article class="issue-card"><p>No ED staff are recorded for ${escapeHtml(formatAustralianTermLabel(term))}${query ? " matching that name" : ""}.</p></article>`;
  const selectedFacilities = (facilityOverviewState.facilityKey === "ALL" ? facilityOverviewFacilityOptions() : [facilityOverviewState.facilityKey])
    .map((facility) => String(facility || "").toLowerCase());
  return selectedFacilities.filter((source) => panels.has(source)).map((source) => {
    const people = panels.get(source);
    const partial = people.some((person) => person.coverageStarts.some((date) => date > formatDateKey(term.start)) || person.coverageEnds.some((date) => date < formatDateKey(addDays(term.end, -1))));
    const groups = new Map();
    const previousStaff = [];
    for (const person of people) {
      if (person.designation?.designation === "previous_staff") {
        previousStaff.push(person);
        continue;
      }
      if (!groups.has(person.seniority)) groups.set(person.seniority, []);
      groups.get(person.seniority).push(person);
    }
    const countColumnLabelWidth = Math.max(...[...groups.keys(), ...(previousStaff.length ? ["Previous staff"] : [])].map((label) => String(label).length), 1) + 1;
    return `<section class="facility-overview-ed-panel">
      <div class="facility-overview-ed-heading"><h3>${escapeHtml(displaySourceCode(source))}</h3>${partial ? `<span class="facility-overview-coverage-warning">Partial roster coverage</span>` : ""}</div>
      <div class="facility-overview-staff-sections" style="--facility-overview-seniority-label-width: ${countColumnLabelWidth}ch">${[...groups.entries()].sort(([left], [right]) => facilityOverviewSeniorityRank(left) - facilityOverviewSeniorityRank(right)).map(([grade, staff]) => {
        const key = `${source}:${grade}`;
        const expanded = query || facilityOverviewState.staffExpanded.has(key);
        const canMultiSelect = grade === "Unknown" && expanded && isViewingCreatorAccount();
        const multiSelectActive = facilityOverviewStaffMultiSelectIsActive(key);
        const selectedCount = multiSelectActive ? facilityOverviewState.staffMultiSelectMembers.size : 0;
        return `<section class="facility-overview-staff-section">
          <div class="facility-overview-staff-section-banner"><button type="button" class="facility-overview-staff-section-toggle" aria-expanded="${expanded}" data-facility-overview-staff-section="${escapeHtml(key)}"><span class="facility-overview-staff-section-title">${escapeHtml(grade)}</span><span class="facility-overview-staff-section-count">(${staff.length})</span><span class="facility-overview-staff-section-chevron" aria-hidden="true"></span></button>${canMultiSelect ? `<label class="facility-overview-staff-multi-select"><input type="checkbox" data-facility-overview-staff-multi-select="${escapeHtml(key)}" ${multiSelectActive ? "checked" : ""} ${facilityOverviewState.staffMultiSelectSaving ? "disabled" : ""}><span>Multi-select${multiSelectActive ? ` (${selectedCount})` : ""}</span></label>` : ""}</div>
          ${expanded ? `<div class="facility-overview-staff-list">${staff.sort((left, right) => left.displayName.localeCompare(right.displayName)).map((person) => {
            const dates = person.events.map((event) => String(event.start || "").slice(0, 10)).filter(Boolean).sort();
            const activity = dates.length ? `${formatDate(dates[0])}${dates.length > 1 ? ` – ${formatDate(dates[dates.length - 1])}` : ""} · ${dates.length} shift${dates.length === 1 ? "" : "s"}` : "No working shifts recorded this term";
            const selected = facilityOverviewStaffMultiSelectIsSelected({ sectionKey: key, sourceType: person.sourceType, doctorKey: person.doctorKey });
            const name = renderFacilityOverviewStaffName(person, { canUseStaffActions, designation: person.designation, seniority: person.seniority, termStart: formatDateKey(term.start), multiSelectSectionKey: grade === "Unknown" ? key : "", multiSelected: selected });
            return `<article class="facility-overview-staff-member${selected ? " is-multi-selected" : ""}">${name}${renderFacilityOverviewStaffActivity(person, activity)}${renderFacilityOverviewStaffSeniorityStatus(person, term)}</article>`;
          }).join("")}</div>${multiSelectActive && facilityOverviewState.staffBulkSeniorityMenu?.sectionKey === key ? renderFacilityOverviewStaffBulkSeniorityMenu(facilityOverviewState.staffBulkSeniorityMenu, selectedCount) : ""}` : ""}
        </section>`;
      }).join("")}${previousStaff.length ? renderFacilityOverviewPreviousStaffSection(source, previousStaff, query, canUseStaffActions) : ""}</div>
    </section>`;
  }).join("") || `<article class="issue-card"><p>No ED staff are recorded for this selection.</p></article>`;
}

function renderFacilityOverviewStaffActivity(person, fallback) {
  const designation = person?.designation;
  const label = designation?.label || fallback;
  if (!isViewingCreatorAccount() || fallback !== "No working shifts recorded this term" && !designation?.label) return `<span>${escapeHtml(label)}</span>`;
  const key = facilityOverviewStaffDesignationMenuKey(person, designation);
  const menu = facilityOverviewState.staffDesignationMenu;
  const expanded = menu?.key === key;
  return `<div class="facility-overview-staff-designation-action"><button type="button" class="facility-overview-staff-designation-trigger" data-facility-overview-staff-designation-menu="${escapeHtml(key)}" data-facility-overview-staff-source="${escapeHtml(person.sourceType)}" data-facility-overview-staff-key="${escapeHtml(person.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(person.displayName)}" data-facility-overview-staff-seniority="${escapeHtml(person.seniority)}" aria-haspopup="menu" aria-expanded="${expanded}">${escapeHtml(label)}</button>${expanded ? renderFacilityOverviewStaffDesignationMenu(person, designation, menu) : ""}</div>`;
}

function renderFacilityOverviewPreviousStaffSection(source, staff, query, canUseStaffActions) {
  const key = `${source}:previous-staff`;
  const expanded = query || facilityOverviewState.staffExpanded.has(key);
  return `<section class="facility-overview-staff-section facility-overview-previous-staff-section">
    <button type="button" class="facility-overview-staff-section-toggle" aria-expanded="${expanded}" data-facility-overview-staff-section="${escapeHtml(key)}"><span class="facility-overview-staff-section-title">Previous staff</span><span class="facility-overview-staff-section-count">(${staff.length})</span><span class="facility-overview-staff-section-chevron" aria-hidden="true"></span></button>
    ${expanded ? `<div class="facility-overview-staff-list">${staff.sort((left, right) => left.displayName.localeCompare(right.displayName)).map((person) => `<article class="facility-overview-staff-member">${renderFacilityOverviewStaffName(person, { canUseStaffActions, designation: person.designation, seniority: person.seniority, termStart: facilityOverviewState.staffTermStart })}<span>No longer works for this ED</span></article>`).join("")}</div>` : ""}
  </section>`;
}

function renderFacilityOverviewStaffName(person, options = {}) {
  const displayName = String(person?.displayName || person?.doctorKey || "Staff member");
  const target = { doctorKey: person?.doctorKey || person?.key || "", displayName, sourceType: person?.sourceType || "", seniority: facilityOverviewNormalizeSeniority(options.seniority || person?.seniority || "Unknown"), termStart: options.termStart || "" };
  if (options.directCalendar && canUseCreatorDoctorSwitcher()) {
    return `<button type="button" class="facility-overview-staff-calendar-link" data-facility-overview-open-staff-calendar="${escapeHtml(target.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(displayName)}" data-facility-overview-staff-source="${escapeHtml(target.sourceType)}" aria-label="Open ${escapeHtml(displayName)}'s calendar">${escapeHtml(displayName)}</button>`;
  }
  if (!options.canUseStaffActions) return `<strong>${escapeHtml(displayName)}</strong>`;
  const menuKey = facilityOverviewStaffActionMenuKey(target);
  const menu = facilityOverviewState.staffActionMenu;
  const expanded = isViewingCreatorAccount() && menu?.key === menuKey;
  const multiSelectSectionKey = String(options.multiSelectSectionKey || "");
  const multiSelectActive = facilityOverviewStaffMultiSelectIsActive(multiSelectSectionKey);
  const multiSelectAttributes = multiSelectSectionKey ? ` data-facility-overview-staff-multi-select-name data-facility-overview-staff-multi-select-section="${escapeHtml(multiSelectSectionKey)}"${multiSelectActive ? ` aria-pressed="${options.multiSelected ? "true" : "false"}"` : ""}` : "";
  const seniorityMenuKey = facilityOverviewStaffSeniorityMenuKey(target);
  const seniorityMenu = facilityOverviewState.staffSeniorityMenu;
  const seniorityExpanded = isViewingCreatorAccount() && seniorityMenu?.key === seniorityMenuKey;
  return `<div class="facility-overview-staff-action"><button type="button" class="facility-overview-staff-calendar-link" data-facility-overview-open-working-together="${escapeHtml(target.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(displayName)}" data-facility-overview-staff-source="${escapeHtml(target.sourceType)}" data-facility-overview-staff-menu="${escapeHtml(menuKey)}" data-facility-overview-staff-key="${escapeHtml(target.doctorKey)}" data-facility-overview-staff-seniority="${escapeHtml(target.seniority)}" data-facility-overview-staff-term-start="${escapeHtml(target.termStart)}" data-facility-overview-staff-designation-id="${escapeHtml(options.designation?.id || "")}" aria-expanded="${expanded}" aria-haspopup="${isViewingCreatorAccount() ? "menu" : "false"}" aria-label="${multiSelectActive ? `${options.multiSelected ? "Deselect" : "Select"} ${escapeHtml(displayName)}` : `Find times working with ${escapeHtml(displayName)}`}"${multiSelectAttributes}>${escapeHtml(displayName)}</button>${expanded ? renderFacilityOverviewStaffActionMenu({ ...target, designation: options.designation }, menu) : ""}${seniorityExpanded ? renderFacilityOverviewStaffSeniorityMenu(target, seniorityMenu) : ""}</div>`;
}

function facilityOverviewStaffActionMenuKey({ doctorKey = "", displayName = "", sourceType = "" } = {}) {
  return `${String(sourceType || "").toLowerCase()}|${normalizeRosterName(doctorKey)}|${rosterIdentityKey(displayName || doctorKey)}`;
}

function renderFacilityOverviewStaffActionMenu(target, menu) {
  const attributes = `data-facility-overview-staff-display-name="${escapeHtml(target.displayName)}" data-facility-overview-staff-source="${escapeHtml(target.sourceType)}" data-facility-overview-staff-key="${escapeHtml(target.doctorKey)}" data-facility-overview-staff-seniority="${escapeHtml(target.seniority || "Unknown")}" data-facility-overview-staff-term-start="${escapeHtml(target.termStart || "")}"`;
  const left = Math.max(8, Number(menu?.x) || 8);
  const top = Math.max(8, Number(menu?.y) || 8);
  return `<div class="facility-overview-staff-action-menu" role="menu" data-facility-overview-staff-action-menu="${escapeHtml(menu?.key || "")}" style="--facility-overview-menu-left: ${left}px; --facility-overview-menu-top: ${top}px">
    <button type="button" role="menuitem" data-facility-overview-open-staff-calendar="${escapeHtml(target.doctorKey)}" ${attributes}>Person's calendar</button>
    <button type="button" role="menuitem" data-facility-overview-open-working-together="${escapeHtml(target.doctorKey)}" ${attributes}>When working together</button>
    ${isViewingCreatorAccount() ? `<button type="button" role="menuitem" data-facility-overview-edit-staff-seniority data-facility-overview-staff-seniority-menu="${escapeHtml(facilityOverviewStaffSeniorityMenuKey(target))}" data-facility-overview-menu-x="${left}" data-facility-overview-menu-y="${top}" ${attributes}>Edit designation</button>` : ""}
    ${target.designation?.designation === "previous_staff" ? `<button type="button" role="menuitem" data-facility-overview-clear-staff-designation="${escapeHtml(target.designation.id)}">Restore to current staff</button>` : ""}
  </div>`;
}

function facilityOverviewStaffSeniorityMenuKey(person) {
  return `${facilityOverviewStaffActionMenuKey(person)}|seniority|${String(person?.termStart || facilityOverviewState.staffTermStart || "")}`;
}

function renderFacilityOverviewStaffSeniorityMenu(person, menu) {
  const attributes = `data-facility-overview-staff-source="${escapeHtml(person.sourceType)}" data-facility-overview-staff-key="${escapeHtml(person.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(person.displayName)}" data-facility-overview-staff-term-start="${escapeHtml(person.termStart || facilityOverviewState.staffTermStart || "")}"`;
  const left = Math.max(8, Number(menu?.x) || 8);
  const top = Math.max(8, Number(menu?.y) || 8);
  const choices = ["SMS", "CMO", "Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "Intern", "NP", "Physio", "Unknown"];
  return `<div class="facility-overview-staff-action-menu" role="menu" data-facility-overview-staff-seniority-action-menu="${escapeHtml(menu?.key || "")}" style="--facility-overview-menu-left: ${left}px; --facility-overview-menu-top: ${top}px">${choices.map((seniority) => `<button type="button" role="menuitem" data-facility-overview-set-staff-seniority="${escapeHtml(seniority)}" ${attributes}>${escapeHtml(seniority)}</button>`).join("")}<button type="button" role="menuitem" data-facility-overview-set-staff-seniority="" data-facility-overview-use-roster-seniority="true" ${attributes}>Use roster designation</button></div>`;
}

function renderFacilityOverviewStaffBulkSeniorityMenu(menu, selectedCount) {
  const left = Math.max(8, Number(menu?.x) || 8);
  const top = Math.max(8, Number(menu?.y) || 8);
  const choices = ["SMS", "CMO", "Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "Intern", "NP", "Physio", "Unknown"];
  const countLabel = `${selectedCount} selected staff member${selectedCount === 1 ? "" : "s"}`;
  return `<div class="facility-overview-staff-action-menu" role="menu" aria-label="Change designation for ${escapeHtml(countLabel)}" data-facility-overview-staff-bulk-seniority-action-menu style="--facility-overview-menu-left: ${left}px; --facility-overview-menu-top: ${top}px"><span class="facility-overview-staff-bulk-menu-label">Change designation for ${escapeHtml(countLabel)}</span>${choices.map((seniority) => `<button type="button" role="menuitem" data-facility-overview-set-bulk-staff-seniority="${escapeHtml(seniority)}" ${facilityOverviewState.staffMultiSelectSaving ? "disabled" : ""}>${escapeHtml(seniority)}</button>`).join("")}<button type="button" role="menuitem" data-facility-overview-set-bulk-staff-seniority="" data-facility-overview-use-roster-seniority="true" ${facilityOverviewState.staffMultiSelectSaving ? "disabled" : ""}>Use roster designation</button></div>`;
}

function renderFacilityOverviewStaffSeniorityStatus(person, term) {
  const label = person.seniorityOverride && !person.seniorityOverride.useRosterSeniority
    ? `Creator-set ${facilityOverviewCompactSeniorityLabel(person.seniority)}`
    : facilityOverviewCompactSeniorityLabel(person.seniority);
  if (!isViewingCreatorAccount()) return person.seniorityOverride ? `<small>${escapeHtml(label)}</small>` : "";
  const target = { ...person, termStart: formatDateKey(term.start) };
  const key = facilityOverviewStaffSeniorityMenuKey(target);
  const menu = facilityOverviewState.staffSeniorityMenu;
  const expanded = menu?.key === key;
  return `<div class="facility-overview-staff-seniority-action"><button type="button" class="facility-overview-staff-seniority-trigger" data-facility-overview-staff-seniority-menu="${escapeHtml(key)}" data-facility-overview-staff-source="${escapeHtml(person.sourceType)}" data-facility-overview-staff-key="${escapeHtml(person.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(person.displayName)}" data-facility-overview-staff-seniority="${escapeHtml(person.seniority)}" data-facility-overview-staff-term-start="${escapeHtml(formatDateKey(term.start))}" aria-haspopup="menu" aria-expanded="${expanded}">${escapeHtml(label)}</button>${expanded ? renderFacilityOverviewStaffSeniorityMenu(target, menu) : ""}</div>`;
}

function facilityOverviewStaffDesignationMenuKey(person, designation = null) {
  return `${facilityOverviewStaffActionMenuKey(person)}|${designation?.id || "none"}`;
}

function renderFacilityOverviewStaffDesignationMenu(person, designation, menu) {
  const attributes = `data-facility-overview-staff-source="${escapeHtml(person.sourceType)}" data-facility-overview-staff-key="${escapeHtml(person.doctorKey)}" data-facility-overview-staff-display-name="${escapeHtml(person.displayName)}" data-facility-overview-staff-seniority="${escapeHtml(person.seniority)}"`;
  const left = Math.max(8, Number(menu?.x) || 8);
  const top = Math.max(8, Number(menu?.y) || 8);
  if (designation?.id) {
    return `<div class="facility-overview-staff-action-menu" role="menu" data-facility-overview-staff-designation-action-menu="${escapeHtml(menu?.key || "")}" style="--facility-overview-menu-left: ${left}px; --facility-overview-menu-top: ${top}px"><button type="button" role="menuitem" data-facility-overview-clear-staff-designation="${escapeHtml(designation.id)}">Remove ${escapeHtml(designation.label || "leave")}</button></div>`;
  }
  const choices = [
    ["long_service_leave", "Long Service Leave"],
    ["sabbatical_leave", "Sabbatical Leave"],
    ["sick_leave", "Sick Leave"],
    ["personal_leave", "Personal Leave"],
    ["previous_staff", "No longer works for this ED"],
  ].filter(([value]) => value !== "previous_staff" || person.seniority === "SMS");
  return `<div class="facility-overview-staff-action-menu" role="menu" data-facility-overview-staff-designation-action-menu="${escapeHtml(menu?.key || "")}" style="--facility-overview-menu-left: ${left}px; --facility-overview-menu-top: ${top}px">${choices.map(([value, label]) => `<button type="button" role="menuitem" data-facility-overview-set-staff-designation="${value}" ${attributes}>${label}</button>`).join("")}</div>`;
}

function renderFacilityOverviewSeniorityLink(seniority, options = {}) {
  const grade = String(seniority || "").trim();
  const sourceType = String(options.sourceType || "").trim().toLowerCase();
  if (!grade || grade === "Unknown" || !sourceType || !canUseFacilityOverview()) return "";
  return ` <button type="button" class="facility-overview-seniority facility-overview-seniority-link" data-facility-overview-open-staff-section="${escapeHtml(grade)}" data-facility-overview-staff-source="${escapeHtml(sourceType)}" data-facility-overview-staff-date="${escapeHtml(options.date || "")}" aria-label="View ${escapeHtml(grade)} staff at ${escapeHtml(displaySourceCode(sourceType))}">${escapeHtml(grade)}</button>`;
}

function facilityOverviewDoctorOptionFor({ doctorKey = "", displayName = "", sourceType = "" } = {}) {
  const normalizedKey = normalizeRosterName(doctorKey);
  const identity = rosterIdentityKey(displayName || doctorKey);
  const source = String(sourceType || "").trim().toLowerCase();
  const hasKey = (doctor) => [doctor?.key, ...(doctor?.aliases || []).map((alias) => alias?.key)]
    .map(normalizeRosterName)
    .includes(normalizedKey);
  const hasSource = (doctor) => normalizedDoctorSourceTypes(doctor).includes(source)
    || (doctor?.aliases || []).some((alias) => String(alias?.sourceType || "").toLowerCase() === source);
  const options = doctorPickerOptions();
  return options.find((doctor) => normalizedKey && hasKey(doctor) && (!source || hasSource(doctor)))
    || options.find((doctor) => normalizedKey && hasKey(doctor))
    || options.find((doctor) => identity && rosterIdentityKey(doctor?.displayName || doctor?.key) === identity && (!source || hasSource(doctor)))
    || options.find((doctor) => identity && rosterIdentityKey(doctor?.displayName || doctor?.key) === identity)
    || null;
}

async function openFacilityOverviewStaffCalendar(target) {
  if (!canUseCreatorDoctorSwitcher()) return;
  const doctor = facilityOverviewDoctorOptionFor(target);
  if (!doctor) {
    setStatus("That staff member's calendar profile is not available yet.", true);
    return;
  }
  closeFacilityOverview();
  await switchDoctorSelection(doctor.key, { resetRange: true });
}

function facilityOverviewTogetherOptionFor({ doctorKey = "", displayName = "", sourceType = "" } = {}) {
  const normalizedKey = normalizeRosterName(doctorKey);
  const identity = rosterIdentityKey(displayName || doctorKey);
  const source = String(sourceType || "").trim().toLowerCase();
  const hasKey = (doctor) => facilityOverviewTogetherDoctorKeys(doctor).includes(normalizedKey);
  const hasSource = (doctor) => normalizedDoctorSourceTypes(doctor).includes(source)
    || (doctor?.aliases || []).some((alias) => String(alias?.sourceType || "").toLowerCase() === source);
  const options = facilityOverviewTogetherStaffOptions();
  return options.find((doctor) => normalizedKey && hasKey(doctor) && (!source || hasSource(doctor)))
    || options.find((doctor) => normalizedKey && hasKey(doctor))
    || options.find((doctor) => identity && doctorIdentityKey(doctor) === identity && (!source || hasSource(doctor)))
    || options.find((doctor) => identity && doctorIdentityKey(doctor) === identity)
    || null;
}

function facilityOverviewTogetherFallbackOption({ doctorKey = "", displayName = "", sourceType = "", aliases = [], sourceTypes = [] } = {}) {
  const key = normalizeRosterName(doctorKey);
  if (!key) return null;
  const normalizedSource = String(sourceType || "").trim().toLowerCase();
  const normalizedAliases = Array.isArray(aliases) ? aliases.filter((alias) => alias?.key) : [];
  const option = {
    key,
    displayName: String(displayName || key),
    sourceType: normalizedSource,
    sourceTypes: [...new Set([normalizedSource, ...(sourceTypes || [])].map((item) => String(item || "").toLowerCase()).filter(Boolean))],
    aliases: normalizedAliases,
  };
  return { ...option, identity: doctorIdentityKey(option) };
}

function openFacilityOverviewWorkingTogether(target) {
  if (!canUseFacilityOverview()) return;
  const selectedPerson = facilityOverviewTogetherOptionFor(target) || facilityOverviewTogetherFallbackOption(target);
  if (!selectedPerson) return;
  const activeDoctor = currentNonClinical ? null : selectedDoctor();
  const viewer = activeDoctor && (facilityOverviewTogetherOptionFor(activeDoctor)
    || facilityOverviewTogetherFallbackOption({
      doctorKey: activeDoctor.key || currentRosterClaims[0]?.key || "",
      displayName: activeDoctor.displayName || currentAccount().realName || currentRosterClaims[0]?.displayName || "",
      sourceType: activeDoctor.sourceType || currentRosterClaims[0]?.sourceType || "",
      sourceTypes: normalizedDoctorSourceTypes(activeDoctor),
      aliases: Array.isArray(activeDoctor.aliases) && activeDoctor.aliases.length ? activeDoctor.aliases : currentRosterClaims,
    }));
  facilityOverviewState.tab = "together";
  facilityOverviewState.staffActionMenu = null;
  facilityOverviewState.togetherPinnedDoctors = viewer ? [selectedPerson, viewer] : [selectedPerson];
  facilityOverviewState.togetherStaffKeys = !viewer || selectedPerson.identity === viewer.identity
    ? [selectedPerson.identity, ""]
    : [selectedPerson.identity, viewer.identity];
  facilityOverviewState.togetherUserClearedAll = false;
  facilityOverviewState.togetherContent = "";
  facilityOverviewState.togetherHasSearched = false;
  resetFacilityOverviewScroll();
  renderFacilityOverview();
  void loadFacilityOverviewTogether();
}

function closeFacilityOverviewStaffActionMenu() {
  if (!facilityOverviewState.staffActionMenu) return;
  facilityOverviewState.staffActionMenu = null;
  refreshFacilityOverviewStaffMenuContent();
}

function closeFacilityOverviewStaffDesignationMenu() {
  if (!facilityOverviewState.staffDesignationMenu) return;
  facilityOverviewState.staffDesignationMenu = null;
  refreshFacilityOverviewStaffMenuContent();
}

function closeFacilityOverviewStaffSeniorityMenu() {
  if (!facilityOverviewState.staffSeniorityMenu) return;
  facilityOverviewState.staffSeniorityMenu = null;
  refreshFacilityOverviewStaffMenuContent();
}

function refreshFacilityOverviewStaffMenuContent() {
  refreshFacilityOverviewStaffActionContent();
  if (facilityOverviewState.tab === "staff") renderFacilityOverviewStaffBodyPreservingViewport();
  else renderFacilityOverview();
}

function closeFacilityOverviewStaffBulkSeniorityMenu() {
  if (!facilityOverviewState.staffBulkSeniorityMenu) return;
  facilityOverviewState.staffBulkSeniorityMenu = null;
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport(facilityOverviewState.staffMultiSelectSection);
}

async function setFacilityOverviewStaffDesignation({ designation = "", sourceType = "", doctorKey = "", displayName = "", seniority = "" } = {}) {
  if (!isViewingCreatorAccount() || !designation || !sourceType || !doctorKey) return;
  const term = australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date())));
  closeFacilityOverviewStaffDesignationMenu();
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setFacilityStaffDesignation", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        facilityKey: sourceType, doctorKey, displayName, seniority, designation,
        termStart: formatDateKey(term.start), termEnd: formatDateKey(addDays(term.end, -1)),
      }),
    });
    const data = await readJsonResponse(response, "Could not save the staff designation.");
    setStatus(designation === "previous_staff" ? `${displayName} moved to Previous staff.` : `${displayName} marked as ${designation.replaceAll("_", " ")}.`);
    facilityOverviewState.onShiftData = null;
    applyFacilityOverviewStaffDesignation(data.designation);
  } catch (error) {
    setStatus(error.message || "Could not save the staff designation.", true);
  }
}

async function clearFacilityOverviewStaffDesignation(designationId) {
  if (!isViewingCreatorAccount() || !designationId) return;
  closeFacilityOverviewStaffActionMenu();
  closeFacilityOverviewStaffDesignationMenu();
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "clearFacilityStaffDesignation", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        designationId,
      }),
    });
    const data = await readJsonResponse(response, "Could not remove the staff designation.");
    setStatus("Staff designation removed.");
    facilityOverviewState.onShiftData = null;
    applyFacilityOverviewStaffDesignation(data.designation, { clear: true });
  } catch (error) {
    setStatus(error.message || "Could not remove the staff designation.", true);
  }
}

async function setFacilityOverviewStaffSeniorityOverride({ sourceType = "", doctorKey = "", displayName = "", seniority = "", useRosterSeniority = false, termStart = "" } = {}) {
  if (!isViewingCreatorAccount() || !sourceType || !doctorKey) return;
  const effectiveTermStart = termStart || formatDateKey(australianTermForDate(parseDateOnly(facilityOverviewState.tab === "on-shift" ? facilityOverviewState.date : facilityOverviewState.staffTermStart)).start);
  closeFacilityOverviewStaffActionMenu();
  closeFacilityOverviewStaffSeniorityMenu();
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setFacilityStaffSeniorityOverride", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        facilityKey: sourceType, doctorKey, displayName, seniority, useRosterSeniority, termStart: effectiveTermStart,
      }),
    });
    const data = await readJsonResponse(response, "Could not save the staff designation.");
    setStatus(useRosterSeniority ? `${displayName} will use the roster designation from this term.` : `${displayName} is now designated ${seniority} from this term.`);
    visibleInsightWarmCache.clear();
    if (useRosterSeniority && facilityOverviewState.tab === "on-shift") void loadFacilityOverviewOnShift();
    else applyFacilityOverviewStaffSeniorityOverrides([data.override]);
  } catch (error) {
    setStatus(error.message || "Could not save the staff designation.", true);
  }
}

async function setFacilityOverviewStaffSeniorityOverrides({ seniority = "", useRosterSeniority = false } = {}) {
  const sectionKey = facilityOverviewState.staffMultiSelectSection;
  const selectedStaff = [...facilityOverviewState.staffMultiSelectMembers.values()];
  const sourceType = String(sectionKey || "").split(":")[0];
  if (!isViewingCreatorAccount() || !sectionKey || !sourceType || !selectedStaff.length || facilityOverviewState.staffMultiSelectSaving) return;
  const termStart = formatDateKey(australianTermForDate(parseDateOnly(facilityOverviewState.staffTermStart || formatDateKey(new Date()))).start);
  facilityOverviewState.staffMultiSelectSaving = true;
  facilityOverviewState.staffBulkSeniorityMenu = null;
  refreshFacilityOverviewStaffContent();
  renderFacilityOverviewStaffBodyPreservingViewport(sectionKey);
  try {
    const response = await fetch("/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setFacilityStaffSeniorityOverrides", email: authUserEmail || currentUserEmail, password: authUserPassword || currentUserPassword,
        facilityKey: sourceType, staff: selectedStaff, seniority, useRosterSeniority, termStart,
      }),
    });
    const data = await readJsonResponse(response, "Could not save the staff designations.");
    visibleInsightWarmCache.clear();
    facilityOverviewState.staffMultiSelectSaving = false;
    clearFacilityOverviewStaffMultiSelect({ render: false });
    if (useRosterSeniority && facilityOverviewState.tab === "on-shift") void loadFacilityOverviewOnShift();
    else applyFacilityOverviewStaffSeniorityOverrides(data.overrides || []);
    const verb = useRosterSeniority ? "will use the roster designation" : `are now designated ${seniority}`;
    setStatus(`${selectedStaff.length} staff member${selectedStaff.length === 1 ? "" : "s"} ${verb}.`);
  } catch (error) {
    facilityOverviewState.staffMultiSelectSaving = false;
    refreshFacilityOverviewStaffContent();
    renderFacilityOverviewStaffBodyPreservingViewport(sectionKey);
    setStatus(error.message || "Could not save the staff designations.", true);
  }
}

async function openFacilityOverviewStaffSection({ sourceType = "", seniority = "", date = "" } = {}) {
  if (!canUseFacilityOverview()) return;
  const source = String(sourceType || "").trim().toUpperCase();
  const grade = String(seniority || "").trim();
  const dateKey = String(date || facilityOverviewState.date || "").slice(0, 10);
  if (!source || !grade || grade === "Unknown" || !/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return;
  const term = australianTermForDate(parseDateOnly(dateKey));
  const sectionKey = `${source.toLowerCase()}:${grade}`;
  clearFacilityOverviewStaffMultiSelect({ render: false });
  facilityOverviewState.tab = "staff";
  facilityOverviewState.facilityKey = source;
  facilityOverviewState.staffTermStart = formatDateKey(term.start);
  facilityOverviewState.staffQuery = "";
  facilityOverviewState.staffExpanded = new Set([sectionKey]);
  facilityOverviewState.staffFocusSection = sectionKey;
  resetFacilityOverviewScroll();
  await openFacilityOverview({ preserveStaffTerm: true });
}

function focusFacilityOverviewStaffSection() {
  const key = facilityOverviewState.staffFocusSection;
  if (!key || facilityOverviewState.tab !== "staff") return;
  const button = [...(facilityOverviewBody?.querySelectorAll("[data-facility-overview-staff-section]") || [])]
    .find((candidate) => candidate.dataset.facilityOverviewStaffSection === key);
  if (!button) return;
  facilityOverviewState.staffFocusSection = "";
  window.requestAnimationFrame(() => {
    button.scrollIntoView({ block: "nearest" });
    button.focus({ preventScroll: true });
  });
}

function facilityOverviewEventStartMinutes(event) {
  const match = String(event?.timeLabel || "").match(/^(\d{1,2}):(\d{2})/);
  if (!match) return Number.NaN;
  return Number(match[1]) * 60 + Number(match[2]);
}

function renderFacilityOverviewStream(events, options = {}) {
  return [...events]
    .sort(compareInsightEvents)
    .map((event) => {
      const stream = String(event?.title || "").replace(/^[^:]+:\s*/, "").trim();
      return options.includeTimes && event?.timeLabel ? `${stream} (${event.timeLabel})` : stream;
    })
    .filter(Boolean)
    .join(" · ");
}

function facilityOverviewSeniorityRank(value) {
  return FACILITY_OVERVIEW_SENIORITY_ORDER.indexOf(facilityOverviewNormalizeSeniority(value));
}

function compareFacilityOverviewSeniorities(left, right) {
  const rankDifference = facilityOverviewSeniorityRank(left) - facilityOverviewSeniorityRank(right);
  if (rankDifference) return rankDifference;
  return facilityOverviewNormalizeSeniority(left).localeCompare(facilityOverviewNormalizeSeniority(right));
}

function compareFacilityOverviewPeople(left, right) {
  return compareFacilityOverviewSeniorities(left?.seniority, right?.seniority)
    || String(left?.displayName || "").localeCompare(String(right?.displayName || ""));
}

function compareFacilityOverviewAssignmentsByStart(left, right) {
  return String(left?.event?.start || "").localeCompare(String(right?.event?.start || ""))
    || String(left?.displayName || "").localeCompare(String(right?.displayName || ""));
}

function compareFacilityOverviewAssignmentsBySeniority(left, right) {
  return compareFacilityOverviewSeniorities(left?.seniority, right?.seniority)
    || String(left?.displayName || "").localeCompare(String(right?.displayName || ""));
}

function facilityOverviewSeniorityAbbreviation(value) {
  const seniority = facilityOverviewNormalizeSeniority(value);
  const abbreviations = {
    "Senior Registrar": "SR",
    "Transitional/Intermediate Registrar": "TR",
    "Junior Registrar": "JR",
    Intern: "I",
  };
  return abbreviations[seniority] || seniority;
}

function facilityOverviewByStreamShiftBlocks(assignments = []) {
  const normalStartMinutes = { AM: 8 * 60, PM: 14 * 60 + 30, Night: 23 * 60 };
  const blocks = new Map();
  for (const assignment of assignments) {
    const start = extractTimePortion(assignment?.event?.start || "");
    const [hours = "0", minutes = "0"] = start.split(":");
    let startMinutes = Number(hours) * 60 + Number(minutes);
    if (assignment.period === "Night" && startMinutes < 6 * 60) startMinutes += 24 * 60;
    const isExceptional = Boolean(assignment.specialTime && start);
    const key = isExceptional ? `time:${start}` : `period:${assignment.period}`;
    const block = blocks.get(key) || {
      key,
      label: isExceptional ? start : assignment.period,
      sortMinutes: isExceptional ? startMinutes : (normalStartMinutes[assignment.period] ?? startMinutes),
      assignments: [],
    };
    block.assignments.push(assignment);
    blocks.set(key, block);
  }
  return [...blocks.values()]
    .sort((left, right) => left.sortMinutes - right.sortMinutes || left.label.localeCompare(right.label))
    .map((block) => ({ ...block, assignments: block.assignments.sort(compareFacilityOverviewAssignmentsBySeniority) }));
}

function facilityOverviewNormalizeSeniority(value) {
  const seniority = String(value || "").trim();
  const normalized = seniority.toLowerCase();
  if (normalized === "sr" || normalized.includes("senior registrar")) return "Senior Registrar";
  if (normalized === "tr" || normalized === "ir" || normalized.includes("transitional") || normalized.includes("intermediate")) return "Transitional/Intermediate Registrar";
  if (normalized === "jr" || normalized.includes("junior registrar")) return "Junior Registrar";
  if (normalized === "enp" || normalized === "np" || normalized.includes("nurse practitioner")) return "NP";
  if (normalized === "amp" || normalized === "physio" || normalized.includes("physiotherapist")) return "Physio";
  if (normalized === "sms" || normalized === "cmo" || normalized === "hmo") return normalized.toUpperCase();
  if (normalized === "intern" || normalized === "i") return "Intern";
  if (!seniority || normalized === "unknown") return "Unknown";
  return "Unknown";
}

function facilityOverviewDetectedSeniority(event, seniority = "") {
  const text = `${event?.title || ""} ${event?.rawValue || ""} ${event?.seniority || ""} ${seniority || ""}`.toLowerCase();
  if (text.includes("physiotherapist") || /\bphysio\b/.test(text)) return "Physio";
  if (text.includes("nurse practitioner") || /\b(?:enp|np|d1)\b/.test(text)) return "NP";
  return facilityOverviewNormalizeSeniority(seniority || event?.seniority || "Unknown");
}

function facilityOverviewCompactSeniorityLabel(value) {
  const seniority = facilityOverviewNormalizeSeniority(value);
  if (seniority === "Senior Registrar") return "SR";
  if (seniority === "Transitional/Intermediate Registrar") return "TR";
  if (seniority === "Junior Registrar") return "JR";
  return seniority;
}

function knownInsightsAccessForEmail(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail) return false;
  if (targetEmail === OWNER_EMAIL) return true;
  const serverUser = serverUsers.map(normalizeServerUser).find((user) => user.email === targetEmail);
  return serverUser?.insightsEnabled === true;
}

function primeInsightsAccessForCurrentView(options = {}) {
  if (activeCalendarMode() === "doctor-profile") {
    currentInsightsEnabled = isCreatorAuthenticated();
    return;
  }
  if (currentUserRole === "creator" && !adminViewingEmail) {
    currentInsightsEnabled = true;
    return;
  }
  if (typeof options.insightsEnabled === "boolean") {
    currentInsightsEnabled = options.insightsEnabled === true;
    return;
  }
  currentInsightsEnabled = knownInsightsAccessForEmail(viewedAccountEmail());
}

function canRemoveImports() {
  return isViewingCreatorAccount();
}

function canUploadRosters() {
  return isViewingCreatorAccount();
}

function isCreatorAuthenticated() {
  return normalizeEmail(authUserEmail || currentUserEmail) === OWNER_EMAIL && Boolean(authUserPassword || currentUserPassword);
}

function syncAccountsButton() {
  const ownerView = isViewingCreatorAccount();
  if (addRosterFilesButton) addRosterFilesButton.classList.toggle("hidden", !canUploadRosters());
  const issueCount = ownerView ? adminIssueCount() : 0;
  accountsButton.innerHTML = ownerView
    ? `Admin${issueCount ? `<span class="notification-badge">${issueCount}</span>` : ""}`
    : "Account";
  syncFacilityOverviewAccess();
}

function renderAccountsModal() {
  const me = currentAccount();
  const ownerView = isViewingCreatorAccount();
  const doctorProfileView = activeCalendarMode() === "doctor-profile" && activeDoctorProfile;
  accountsModalTitle.textContent = ownerView ? "Admin" : "Account";
  accountsModalSubtitle.textContent = ownerView
    ? "Review user issues, manage accounts, and update the owner account."
    : doctorProfileView
      ? "Review the roster account details used for this calendar."
      : "Manage your account details.";

  const serverOtherUsers = serverUsers
    .map(normalizeServerUser)
    .filter((user) => user.email !== me.email);
  const localOtherUsers = accountState.users.filter((user) => user.email !== me.email);
  const otherUsers = serverOtherUsers.length ? serverOtherUsers : localOtherUsers;
  const availableUserSeniorities = [...new Set(otherUsers.flatMap((user) => normalizeServerUser(user).seniorities || []))].sort();
  if (adminUserSeniorityFilter && !availableUserSeniorities.includes(adminUserSeniorityFilter)) adminUserSeniorityFilter = "";
  const seniorityFilteredUsers = adminUserSeniorityFilter
    ? otherUsers.filter((user) => normalizeServerUser(user).seniorities.includes(adminUserSeniorityFilter))
    : otherUsers;
  const normalizedUserSearchQuery = adminUserSearchQuery.trim().toLocaleLowerCase();
  const filteredOtherUsers = normalizedUserSearchQuery
    ? seniorityFilteredUsers.filter((user) => [
        user.realName,
        user.email,
        ...sanitizeRosterClaims(user.claims || []).flatMap((claim) => [claim.displayName, claim.sourceType]),
      ].some((value) => String(value || "").toLocaleLowerCase().includes(normalizedUserSearchQuery)))
    : seniorityFilteredUsers;
  const linkedNames = renderLinkedRosterNames(currentRosterClaims, currentSuggestedClaims);
  if (ownerView && !["parser", "system", "users", "files", "owner"].includes(currentAdminTab)) currentAdminTab = "users";
  const issueCount = adminIssueCount();
  const adminTabs = ownerView ? `
    <div class="admin-tabs" role="tablist" aria-label="Admin sections">
      <button type="button" class="entrance-tab ${currentAdminTab === "users" ? "is-active" : ""}" data-admin-tab="users">Users</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "files" ? "is-active" : ""}" data-admin-tab="files">Files</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "owner" ? "is-active" : ""}" data-admin-tab="owner">Account</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "parser" ? "is-active" : ""}" data-admin-tab="parser">Parser${issueCount ? `<span class="notification-badge">${issueCount}</span>` : ""}</button>
      <button type="button" class="entrance-tab ${currentAdminTab === "system" ? "is-active" : ""}" data-admin-tab="system">System</button>
    </div>
  ` : "";
  const ownerCard = `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>${ownerView ? "Account" : doctorProfileView ? "Roster account" : "Your account"}</strong>
          <span>${escapeHtml(me.realName || "Name not set")}${me.email ? ` · ${escapeHtml(me.email)}` : doctorProfileView ? " · No linked account" : ""}</span>
        </div>
      </div>
      ${doctorProfileView ? `
        <div class="review-body">
          <label class="field">
            <span>Email address</span>
            <input type="email" value="" readonly placeholder="No linked account">
          </label>
        </div>
      ` : `
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
      `}
      ${renderAccountHospitalLocationsCard({ directorPreference: currentNonClinical && currentDirectorViewEnabled && !doctorProfileView })}
      ${currentNonClinical && currentDirectorViewEnabled ? "" : linkedNames}
      ${ownerView || (currentNonClinical && currentDirectorViewEnabled) ? "" : renderFilesMarkup({
        canRemove: false,
        canAdd: !doctorProfileView,
        heading: "Files used to generate your calendar...",
        description: "These roster files currently feed your calendar.",
      })}
    </article>
  `;
  const usersCard = ownerView ? `
      <details class="review-card create-user-card" data-create-user-account-section ${createUserAccountExpanded ? "open" : ""}>
        <summary class="review-top">
          <div>
            <strong>Create user account</strong>
            <span>Create an account and enter it immediately for setup or testing.</span>
          </div>
          <span class="collapsible-chevron" aria-hidden="true">⌄</span>
        </summary>
        <form class="review-body" data-create-account-form novalidate>
          <label class="field">
            <span>Full name on roster</span>
            <input type="text" data-create-real-name placeholder="Name shown to the user" autocomplete="name">
          </label>
          <label class="field">
            <span>Email address</span>
            <input type="text" inputmode="email" data-create-email placeholder="doctor@example.com" autocomplete="email">
          </label>
          <label class="field">
            <span>Temporary password (manual setup only)</span>
            <input type="password" data-create-password placeholder="Not needed when sending an invite" autocomplete="new-password">
          </label>
          <div class="create-user-options" aria-label="Account access options">
            <label class="toggle review-toggle">
              <input type="checkbox" data-create-non-clinical>
              Non-clinical
            </label>
            <label class="toggle review-toggle">
              <input type="checkbox" data-create-director-view>
              Director view
            </label>
          </div>
          <div class="modal-actions">
            <button type="submit" class="button button-primary">Create and enter account</button>
            <button type="button" class="button button-secondary" data-send-account-invite>Send invite</button>
          </div>
        </form>
      </details>
      <article class="review-card creator-qr-card" aria-labelledby="creator-qr-title">
        <img class="creator-qr-code" src="/static/rtc-curiousmind-qr.svg" alt="QR code that opens rtc.curiousmind.app">
        <div class="creator-qr-copy">
          <strong id="creator-qr-title">Open Roster to Calendar</strong>
          <span>Scan this code with a phone to open the app.</span>
          <a href="https://rtc.curiousmind.app" target="_blank" rel="noopener noreferrer">rtc.curiousmind.app</a>
        </div>
      </article>
      <details class="review-card other-users-card" data-other-users-section ${otherUsersExpanded ? "open" : ""}>
        <summary class="review-top admin-users-header">
          <div class="admin-users-summary">
            <strong>Current users</strong>
            <span>${filteredOtherUsers.length ? `${filteredOtherUsers.length} account${filteredOtherUsers.length === 1 ? "" : "s"}` : otherUsers.length ? "No matching users." : "No other users have logged in yet."}</span>
          </div>
          <label class="field admin-user-filter admin-user-search-filter">
            <span>Search users</span>
            <input type="search" value="${escapeHtml(adminUserSearchQuery)}" data-admin-user-search placeholder="Name or email">
          </label>
          <label class="field admin-user-filter admin-user-seniority-filter">
            <span>Filter by seniority</span>
            <select data-admin-user-seniority-filter>
              <option value="">All seniorities</option>
              ${availableUserSeniorities.map((seniority) => `<option value="${escapeHtml(seniority)}" ${seniority === adminUserSeniorityFilter ? "selected" : ""}>${escapeHtml(seniority)}</option>`).join("")}
            </select>
          </label>
          <span class="collapsible-chevron" aria-hidden="true">⌄</span>
        </summary>
        <div class="issues-list">
          ${filteredOtherUsers.length ? filteredOtherUsers.map((user) => `
            <article class="issue-card account-user-card">
              <div class="account-user-summary">
                <strong class="account-user-name">${escapeHtml(user.realName || "Name not set")}</strong>
                <p class="account-user-email">${escapeHtml(user.email)}</p>
                ${renderAdminUserClaims(user)}
              </div>
              <div class="account-user-controls">
                ${user.role === "owner" ? "" : `
                <div class="account-user-permissions">
                  <label class="toggle review-toggle">
                    <input type="checkbox" ${user.insightsEnabled ? "checked" : ""} data-toggle-user-insights="${escapeHtml(user.email)}">
                    Who/When?
                  </label>
                  <label class="toggle review-toggle">
                    <input type="checkbox" ${user.facilityOverviewEnabled ? "checked" : ""} data-toggle-user-facility-overview="${escapeHtml(user.email)}">
                    At a glance
                  </label>
                  <label class="toggle review-toggle">
                    <input type="checkbox" ${user.directorViewEnabled ? "checked" : ""} data-toggle-user-director-view="${escapeHtml(user.email)}">
                    Director
                  </label>
                </div>
              `}
                <div class="account-actions">
                  <button type="button" class="button button-secondary button-small" data-enter-account="${escapeHtml(user.email)}">Enter account</button>
                  ${user.role === "owner" ? "" : `<button type="button" class="button button-secondary button-small" data-edit-roster-claims="${escapeHtml(user.email)}">Edit</button>`}
                  ${user.email !== OWNER_EMAIL ? `<button type="button" class="button button-danger button-small" data-delete-account="${escapeHtml(user.email)}">Delete</button>` : ""}
                </div>
              </div>
              ${user.role === "owner" ? "" : `
                <div class="account-claim-editor hidden" data-claim-editor="${escapeHtml(user.email)}">
                  <form class="admin-user-name-form" data-admin-user-form="${escapeHtml(user.email)}">
                    <label class="field">
                      <span>User name</span>
                      <input type="text" value="${escapeHtml(user.realName || "")}" data-admin-user-real-name required autocomplete="name" placeholder="Name shown in the user's profile">
                    </label>
                    <button type="submit" class="button button-primary">Save name</button>
                  </form>
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
      </details>
    ` : "";
  const parserCard = ownerView ? renderParserAdminCard(serverOtherUsers) : "";
  const systemCard = ownerView ? renderSystemAdminCard() : "";
  const filesCard = ownerView ? renderAdminFilesMarkup({
    canRemove: canRemoveImports(),
    canAdd: true,
  }) : "";
  const adminBody = ownerView
    ? (currentAdminTab === "parser"
        ? parserCard
        : currentAdminTab === "system"
          ? systemCard
          : currentAdminTab === "users"
            ? usersCard
            : currentAdminTab === "files"
              ? filesCard
              : ownerCard)
    : ownerCard;
  accountsBody.innerHTML = `${adminTabs}${adminBody}`;
  if (ownerView && currentAdminTab === "users") {
    const currentUsersCard = accountsBody.querySelector(".other-users-card");
    const createUserCard = accountsBody.querySelector(".create-user-card");
    if (currentUsersCard && createUserCard) accountsBody.insertBefore(currentUsersCard, createUserCard);
  }
}

function renderSystemAdminCard() {
  return `
    <div class="issues-list">
      ${renderLoginPerformanceCard()}
      ${renderCalendarStoreCard()}
    </div>
  `;
}

function renderLoginPerformanceCard() {
  const client = lastLoginTimings || {};
  const server = client.server || {};
  const value = (milliseconds) => Number.isFinite(Number(milliseconds)) ? `${Math.round(Number(milliseconds))} ms` : "Not measured";
  const source = String(server.snapshotSource || "unknown").replaceAll("-", " ");
  return `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>Latest login performance</strong>
          <span>Calendar shown: ${escapeHtml(value(client.firstCalendarPaintCommitted || client.firstCalendarPaint))} · Server response: ${escapeHtml(value(server.serverTotalMs))}</span>
        </div>
      </div>
      <div class="review-body system-admin-body">
        <p>Authentication: ${escapeHtml(value(server.authMs))} · Registry: ${escapeHtml(value(server.registryLookupMs))} · R2: ${escapeHtml(value(server.r2ReadMs))}</p>
        <p>Snapshot: ${escapeHtml(source)}${server.validationDeferred === true ? " · revision checked after display" : ""}</p>
      </div>
    </article>
  `;
}

function renderAccountHospitalLocationsCard({ directorPreference = false } = {}) {
  if (directorPreference) {
    const selected = directorHospitalPreference();
    return `
      <section class="review-body account-hospital-locations">
        <div class="section-head">
          <h4>Your hospital(s)</h4>
          <p>Choose the ED to show first throughout At a glance. Select All EDs for the Program Director view.</p>
        </div>
        <label class="field">
          <span>Director preference</span>
          <select data-director-hospital-preference>
            <option value="ALL" ${selected === "ALL" ? "selected" : ""}>All EDs</option>
            ${DIRECTOR_HOSPITAL_OPTIONS.map((facility) => `<option value="${facility.key}" ${selected === facility.key ? "selected" : ""}>${escapeHtml(facility.label)}</option>`).join("")}
          </select>
        </label>
      </section>
    `;
  }
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

const DIRECTOR_HOSPITAL_OPTIONS = [
  { key: "MMC", label: "Monash Medical Centre (MMC)" },
  { key: "DDH", label: "Dandenong Hospital (DDH)" },
  { key: "CASEY", label: "Casey Hospital" },
  { key: "MCH", label: "Monash Children's Hospital (MCH)" },
];

function directorHospitalPreference() {
  const preference = String(settings.directorHospitalPreference || "ALL").toUpperCase();
  return preference === "ALL" || DIRECTOR_HOSPITAL_OPTIONS.some((facility) => facility.key === preference)
    ? preference : "ALL";
}

function updateDirectorHospitalPreference(value) {
  const preference = String(value || "ALL").toUpperCase();
  settings.directorHospitalPreference = DIRECTOR_HOSPITAL_OPTIONS.some((facility) => facility.key === preference)
    ? preference : "ALL";
  facilityOverviewState.facilityKey = settings.directorHospitalPreference;
  facilityOverviewState.togetherFacilityKey = settings.directorHospitalPreference;
  facilityOverviewState.tab = settings.directorHospitalPreference === "ALL" ? "staff" : "on-shift";
  facilityOverviewState.byStreamRows = [];
  refreshFacilityOverviewPreferredFacility();
  saveCurrentSessionState();
  renderAccountsModal();
  if (isFacilityOverviewOpen()) void openFacilityOverview({ preserveFacility: true, preserveStaffTerm: true, preserveDate: true });
  setStatus("Director hospital preference updated.");
}

const ACCOUNT_HOSPITAL_LOCATION_ORDER = ["mmc", "ddh", "mch", "casey"];

function recognizedHospitalTypesForActiveAccount() {
  if (isViewingCreatorAccount()) {
    return ACCOUNT_HOSPITAL_LOCATION_ORDER.filter((sourceType) => detectedSources[sourceType]?.length);
  }
  if (activeCalendarMode() === "doctor-profile" && activeDoctorProfile?.sourceTypes?.length) {
    const linkedTypes = new Set(activeDoctorProfile.sourceTypes.map((sourceType) => String(sourceType || "").toLowerCase()));
    return ACCOUNT_HOSPITAL_LOCATION_ORDER.filter((sourceType) => linkedTypes.has(sourceType));
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
  const statusMatchesSelection = selectedPersistence.statusMatchesEntries;
  const hasUsableStatus = Boolean(status && !unavailable && !calendarStoreStatusError);
  const missingSelectedFiles = hasUsableStatus ? selectedPersistence.missingEntries.slice(0, 3) : [];
  const retainedSourceCount = (status?.files || []).filter((file) => file.rawSourceAvailable === true).length;
  const retainedSourceTotal = (status?.files || []).length;
  const retainedSourceDetail = status && !unavailable
    ? `${retainedSourceCount}/${retainedSourceTotal} source file${retainedSourceTotal === 1 ? "" : "s"} retained.`
    : "";
  const checkedAt = status?.checkedAt ? `Last checked ${formatTimestamp(status.checkedAt)}.` : "";
  const syncSummary = rosterSyncSummary();
  const serverSyncedCount = Number(status?.populated || 0);
  const serverExpectedCount = Number(status?.total || 0);
  const serverStatusComplete = Boolean(status && !unavailable && serverExpectedCount > 0 && serverSyncedCount === serverExpectedCount);
  let statusErrorDetail = "";
  if (calendarStoreStatusError) {
    const errorSuffix = `Status check failed: ${calendarStoreStatusError}.${checkedAt ? ` ${checkedAt}` : ""}`;
    if (status && statusMatchesSelection) {
      statusErrorDetail = `Last good check: ${selectedPersistence.persistedCount}/${selectedPersistence.expectedCount || "?"} roster files synced. ${errorSuffix}`;
    } else if (status) {
      statusErrorDetail = `Status check failed — last result is outdated. ${errorSuffix}`;
    } else {
      statusErrorDetail = errorSuffix;
    }
  }
  const detail = syncSummary
    ? `${syncSummary}.${checkedAt ? ` ${checkedAt}` : ""}`
    : calendarStoreStatusError
    ? statusErrorDetail
    : unavailable
      ? "Roster database is unavailable to this deployment."
      : serverStatusComplete
        ? `${serverSyncedCount} roster file${serverSyncedCount === 1 ? "" : "s"} synced. ${retainedSourceDetail}${checkedAt ? ` ${checkedAt}` : ""}`
        : status && selectedPersistence.expectedCount > 0 && selectedPersistence.complete
          ? `${selectedPersistence.persistedCount} roster file${selectedPersistence.persistedCount === 1 ? "" : "s"} synced. ${retainedSourceDetail}${checkedAt ? ` ${checkedAt}` : ""}`
        : status
          ? `Sync issue detected: ${serverSyncedCount}/${serverExpectedCount} roster files confirmed. ${retainedSourceDetail}${checkedAt ? ` ${checkedAt}` : ""}`
          : "Status check pending — roster files may already be synced.";
  return `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>Roster database</strong>
          <span>${escapeHtml(detail)}</span>
        </div>
      </div>
      <div class="review-body system-admin-body">
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
          <button type="button" class="button button-secondary" data-view-console>${adminConsoleOpen ? "Hide console" : "View console"}</button>
        </div>
        <details class="advanced-roster-recovery">
          <summary>Advanced recovery</summary>
          <p>Only rebuild when retained source files are known to be correct but the derived roster database is corrupted. Normal roster updates are automatic.</p>
          ${hasPendingRosterAutomation()
            ? `<p>Recovery is unavailable while roster updates are queued or processing.</p>`
            : `<button type="button" class="button button-secondary" data-replace-active-rosters>Rebuild all retained rosters</button>`}
        </details>
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
          <button type="button" class="button button-secondary button-small" data-open-shift-code-review ${unknownIssues.length ? "" : "disabled"}>Review</button>
          <button type="button" class="button button-secondary button-small" data-add-manual-shift-code>Add</button>
        </div>
        <div class="issues-list">
          ${unknownIssues.length ? `<article class="issue-card"><p>Open the review list to triage grouped unrecognised shift codes.</p></article>` : `<article class="issue-card"><p>No missing or unresolved shift codes need review.</p></article>`}
          ${globalUnresolvedShiftCodesLoading ? `<article class="issue-card"><p>Checking roster database for unresolved shift codes...</p></article>` : ""}
          ${globalUnresolvedShiftCodesError ? `<article class="issue-card"><p>${escapeHtml(globalUnresolvedShiftCodesError)}</p></article>` : ""}
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
              ${renderUnknownShiftCodeHierarchy(allUnknownIssues.filter((item) => item.source === group.source), { compact: true })}
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

function renderParserAdminCard(users) {
  return `
    <div class="issues-list">
      ${renderParserRulesCard()}
      ${renderAdminErrorsCard(users)}
    </div>
  `;
}

function openShiftCodeReviewModal() {
  if (!shiftCodeReviewModal || !shiftCodeReviewModalBody) return;
  renderShiftCodeReviewModal();
  shiftCodeReviewModal.classList.remove("hidden");
  shiftCodeReviewModal.setAttribute("aria-hidden", "false");
  queueGlobalUnresolvedShiftCodeLoad();
}

function closeShiftCodeReviewModal() {
  if (!shiftCodeReviewModal) return;
  shiftCodeReviewModal.classList.add("hidden");
  shiftCodeReviewModal.setAttribute("aria-hidden", "true");
  if (shiftCodeReviewModalBody) shiftCodeReviewModalBody.innerHTML = "";
  // Returning from a calendar briefly filters to the code just reviewed.
  // That focus must not leak into a later, ordinary Creator review session.
  shiftCodeReviewFilter = { query: "", source: "all" };
}

function renderShiftCodeReviewModal() {
  if (!shiftCodeReviewModalBody) return;
  const allItems = collectUnknownShiftIssues();
  const sources = [...new Set(allItems.map((item) => item.source).filter(Boolean))]
    .sort((left, right) => sourceSortRank(left) - sourceSortRank(right) || left.localeCompare(right));
  shiftCodeReviewModalBody.innerHTML = `
    <div class="shift-code-review-controls">
      <label class="field">
        <span>Search</span>
        <input type="search" value="${escapeHtml(shiftCodeReviewFilter.query)}" data-shift-code-review-search placeholder="Code, doctor, seniority, raw value">
      </label>
      <label class="field">
        <span>Hospital</span>
        <select data-shift-code-review-source>
          <option value="all" ${shiftCodeReviewFilter.source === "all" ? "selected" : ""}>All hospitals</option>
          ${sources.map((source) => `<option value="${escapeHtml(source)}" ${shiftCodeReviewFilter.source === source ? "selected" : ""}>${escapeHtml(source)}</option>`).join("")}
        </select>
      </label>
    </div>
    <div id="shiftCodeReviewResults"></div>
  `;
  renderShiftCodeReviewResults();
}

function renderShiftCodeReviewResults() {
  const results = document.querySelector("#shiftCodeReviewResults");
  if (!results) return;
  const allItems = collectUnknownShiftIssues();
  const filteredItems = filterShiftCodeReviewItems(allItems);
  if (shiftCodeReviewModalSubtitle) {
    shiftCodeReviewModalSubtitle.textContent = `${filteredItems.length} of ${allItems.length} grouped code${allItems.length === 1 ? "" : "s"} shown.`;
  }
  results.innerHTML = renderShiftCodeReviewResultsMarkup(filteredItems);
}

function filterShiftCodeReviewItems(items) {
  const query = String(shiftCodeReviewFilter.query || "").trim().toLowerCase();
  const source = sanitizeIssueSource(shiftCodeReviewFilter.source);
  return (items || []).filter((item) => {
    if (source && item.source !== source) return false;
    if (!query) return true;
    return [
      item.source,
      item.code,
      item.seniorityLabel,
      item.message,
      item.sample,
      item.rawValue,
    ].some((value) => String(value || "").toLowerCase().includes(query));
  });
}

function renderShiftCodeReviewResultsMarkup(items) {
  if (globalUnresolvedShiftCodesLoading) {
    return `<article class="issue-card"><p>Checking roster database for unresolved shift codes...</p></article>`;
  }
  if (globalUnresolvedShiftCodesError) {
    return `<article class="issue-card"><p>${escapeHtml(globalUnresolvedShiftCodesError)}</p></article>`;
  }
  if (!items.length) {
    return `<article class="issue-card"><p>No unresolved shift codes match this filter.</p></article>`;
  }
  return renderUnknownShiftCodeHierarchy(items);
}

function renderUnknownShiftCodeHierarchy(items, options = {}) {
  const bySource = new Map();
  for (const item of items || []) {
    if (!bySource.has(item.source)) bySource.set(item.source, new Map());
    const bySeniority = bySource.get(item.source);
    if (!bySeniority.has(item.seniority)) bySeniority.set(item.seniority, []);
    bySeniority.get(item.seniority).push(item);
  }
  return [...bySource.entries()]
    .sort(([left], [right]) => sourceSortRank(left) - sourceSortRank(right) || left.localeCompare(right))
    .map(([source, bySeniority]) => {
      const count = [...bySeniority.values()].flat().length;
      return `<details class="issue-card shift-code-review-source" ${options.compact ? "" : "open"}>
        <summary><strong>${escapeHtml(source)}</strong> · ${count} unrecognised code${count === 1 ? "" : "s"}</summary>
        <div class="issues-list">
          ${[...bySeniority.entries()]
            .sort(([left], [right]) => senioritySortRank(left) - senioritySortRank(right) || left.localeCompare(right))
            .map(([seniority, codes]) => `
              <details class="issue-card shift-code-review-seniority" ${options.compact ? "" : "open"}>
                <summary><strong>${escapeHtml(seniority || "Unknown seniority")}</strong> · ${codes.length} code${codes.length === 1 ? "" : "s"}</summary>
                <div class="issues-list">
                  ${codes.sort((left, right) => left.code.localeCompare(right.code)).map((item) => `
                    <details class="issue-card shift-code-review-row" data-shift-code-review-id="${escapeHtml(item.id)}">
                      <summary><strong>${escapeHtml(item.code)}</strong>${item.count > 1 ? ` · seen ${item.count} times` : ""}</summary>
                      <div>
                        <p>${escapeHtml(item.message || "Shift code not recognised.")}</p>
                        <p>${escapeHtml(item.sample)}</p>
                        ${renderShiftCodeReviewExamples(item)}
                      </div>
                      <div class="account-actions">${renderShiftCodeReviewIssueActions(item)}</div>
                    </details>
                  `).join("")}
                </div>
              </details>
            `).join("")}
        </div>
      </details>`;
    }).join("");
}

function renderShiftCodeReviewExamples(item) {
  const examples = (Array.isArray(item?.examples) ? item.examples : [])
    .filter((example) => example && (example.displayName || example.doctorKey || example.date));
  if (examples.length < 2) return "";
  const byPerson = new Map();
  for (const example of examples) {
    const doctorKey = normalizeRosterName(example.doctorKey || example.displayName || "");
    const key = doctorKey || `${example.displayName || "Roster"}|${example.date || ""}`;
    const person = byPerson.get(key) || {
      displayName: String(example.displayName || example.doctorKey || "Roster").trim(),
      doctorKey: String(example.doctorKey || example.displayName || "").trim(),
      dates: [],
    };
    const date = String(example.date || "").slice(0, 10);
    if (date && !person.dates.some((entry) => entry.date === date && entry.rawValue === example.rawValue)) {
      person.dates.push({ date, rawValue: String(example.rawValue || "").trim(), timeLabel: String(example.timeLabel || "").trim() });
    }
    byPerson.set(key, person);
  }
  const people = [...byPerson.values()]
    .sort((left, right) => left.displayName.localeCompare(right.displayName));
  return `
    <details class="shift-code-review-examples">
      <summary>More examples · ${people.length} ${people.length === 1 ? "person" : "people"} · ${examples.length} occurrences</summary>
      <div class="issues-list">
        ${people.map((person) => `
          <details class="shift-code-review-example-person">
            <summary>${escapeHtml(person.displayName || "Roster")} · ${person.dates.length} ${person.dates.length === 1 ? "date" : "dates"}</summary>
            <ul class="shift-code-review-example-dates">
              ${person.dates.sort((left, right) => left.date.localeCompare(right.date)).map((entry) => `
                <li><button type="button" class="button button-secondary" data-go-to-unresolved-event="${escapeHtml(item.id)}" data-unresolved-doctor-key="${escapeHtml(person.doctorKey)}" data-unresolved-display-name="${escapeHtml(person.displayName)}" data-unresolved-date="${escapeHtml(entry.date)}" data-unresolved-source="${escapeHtml(item.source)}">${escapeHtml(formatDate(entry.date))}${entry.timeLabel ? ` · ${escapeHtml(entry.timeLabel)}` : ""}${entry.rawValue && entry.rawValue !== item.rawValue ? ` · ${escapeHtml(entry.rawValue)}` : ""}</button></li>
              `).join("")}
            </ul>
          </details>
        `).join("")}
      </div>
    </details>
  `;
}

function renderShiftCodeReviewIssueActions(item) {
  const seniorities = escapeHtml((item.seniorities || []).join("|"));
  const goToEvent = item.doctorKey && item.sampleDate
    ? `<button type="button" class="button button-secondary" data-go-to-unresolved-event="${escapeHtml(item.id)}">Go to event</button>`
    : "";
  if (item.email) {
    return `
      ${goToEvent}
      <button type="button" class="button button-secondary" data-add-shift-code="${escapeHtml(item.email)}" data-error-id="${escapeHtml(item.id)}" data-shift-code-seniorities="${seniorities}">Edit shift code</button>
      <button type="button" class="button button-secondary" data-ignore-shift-code="${escapeHtml(item.email)}" data-error-id="${escapeHtml(item.id)}" data-shift-code-seniorities="${seniorities}">Ignore</button>
    `;
  }
  return `
    ${goToEvent}
    <button type="button" class="button button-secondary" data-add-roster-shift-code="${escapeHtml(item.id)}" data-shift-code-seniorities="${seniorities}">Edit shift code</button>
    <button type="button" class="button button-secondary" data-ignore-roster-shift-code="${escapeHtml(item.id)}" data-shift-code-seniorities="${seniorities}">Ignore</button>
  `;
}

function sourceSortRank(source) {
  const order = ["MMC", "DDH", "Casey", "MCH"];
  const index = order.indexOf(sanitizeIssueSource(source));
  return index >= 0 ? index : order.length;
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
      if (isSystemRosterReviewNotice(issue)) continue;
      const source = sanitizeIssueSource(issue.source);
      const seniority = sanitizeRuleSeniority(issue.seniority);
      const code = parserRuleCodeForIssue(issue);
      if (!source || !code) continue;
      if (isKnownResolvedShiftCodeValue(source, issue.rawValue, issue.suggestedTitle)) continue;
      if (isShiftCodeResolvedByActiveRules({ source, seniority, code, rawValue: code })) continue;
      addUnknownShiftIssueToMap(byKey, {
        origin: "admin",
        source,
        seniority,
        code,
        id: issue.id || issue.fingerprint || "",
        email: user.email,
        message: issue.message || "",
        rawValue: issue.rawValue || code,
        sample: `${user.realName || user.email} · ${formatDate(issue.startDay || issue.date || "")} · ${issue.rawValue || code}`,
        count: issue.count || 1,
        lastSeenAt: issue.lastSeenAt || "",
      });
    }
  }
  for (const item of globalUnresolvedShiftCodes) {
    if (isSystemRosterReviewNotice(item)) continue;
    const source = sanitizeIssueSource(item.source);
    const seniority = sanitizeRuleSeniority(item.seniority);
    const code = parserRuleCodeForIssue(item);
    if (!source || !code) continue;
    if (isKnownResolvedShiftCodeValue(source, item.rawValue)) continue;
    if (isShiftCodeResolvedByGlobalRules({ source, seniority, code, rawValue: item.rawValue || code })) continue;
    addUnknownShiftIssueToMap(byKey, {
      origin: "roster",
      source,
      seniority,
      code,
      id: item.id,
      email: "",
      message: item.message || "",
      rawValue: item.rawValue || code,
      sample: `${item.sampleName || "Roster"} · ${formatDate(item.sampleDate || "")} · ${item.rawValue || code}`,
      doctorKey: item.doctorKey || "",
      displayName: item.displayName || item.sampleName || "",
      sampleDate: item.sampleDate || "",
      count: item.count || 1,
      lastSeenAt: item.lastSeenAt || item.sampleDate || "",
      examples: item.examples || [],
    });
  }
  return [...byKey.values()].sort((left, right) => {
    if (left.source !== right.source) return left.source.localeCompare(right.source);
    return left.code.localeCompare(right.code);
  });
}

function addUnknownShiftIssueToMap(byKey, item) {
  const key = `${item.source}|${sanitizeRuleSeniority(item.seniority)}|${item.code}`;
  const existing = byKey.get(key);
  if (existing) {
    existing.count += item.count || 1;
    existing.seniorities = addUniqueSeniority(existing.seniorities, item.seniority);
    existing.seniorityLabel = formatShiftCodeSeniorities(existing.seniorities);
    if (!existing.email && item.email) {
      existing.email = item.email;
      existing.id = item.id || existing.id;
    }
    if (!existing.rawValue && item.rawValue) existing.rawValue = item.rawValue;
    if (!existing.sample && item.sample) existing.sample = item.sample;
    if (!existing.doctorKey && item.doctorKey) existing.doctorKey = item.doctorKey;
    if (!existing.displayName && item.displayName) existing.displayName = item.displayName;
    if (!existing.sampleDate && item.sampleDate) existing.sampleDate = item.sampleDate;
    if ((item.lastSeenAt || "") > (existing.lastSeenAt || "")) existing.lastSeenAt = item.lastSeenAt || "";
    existing.examples = [...(existing.examples || []), ...(item.examples || [])];
    return;
  }
  const seniorities = addUniqueSeniority([], item.seniority);
  byKey.set(key, {
    origin: item.origin || "",
    source: item.source,
    seniority: sanitizeRuleSeniority(item.seniority),
    seniorities,
    seniorityLabel: formatShiftCodeSeniorities(seniorities),
    code: item.code,
    id: item.id || "",
    email: normalizeEmail(item.email || ""),
    message: item.message || "",
    rawValue: item.rawValue || item.code,
    sample: item.sample || "",
    doctorKey: normalizeRosterName(item.doctorKey || ""),
    displayName: String(item.displayName || "").trim(),
    sampleDate: String(item.sampleDate || "").slice(0, 10),
    count: item.count || 1,
    lastSeenAt: item.lastSeenAt || "",
    examples: Array.isArray(item.examples) ? item.examples : [],
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

function renderAdminUserClaims(user) {
  const correctName = String(user?.realName || "").trim();
  const normalizedCorrectName = correctName.replace(/\s+/g, " ").toLocaleLowerCase();
  const claims = sanitizeRosterClaims(user?.claims || []);
  if (!claims.length) return `<p class="status account-user-claims-empty">No roster names linked.</p>`;
  return `
    <div class="account-user-claims">
      ${claims.map((claim) => {
        const normalizedClaimName = claim.displayName.replace(/\s+/g, " ").toLocaleLowerCase();
        const showVariation = !normalizedCorrectName || normalizedClaimName !== normalizedCorrectName;
        return `<div class="account-user-claim"><strong>${escapeHtml(claim.sourceType.toUpperCase())}</strong>${showVariation ? ` <span>"${escapeHtml(claim.displayName)}"</span>` : ""}</div>`;
      }).join("")}
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
    if (decision === "approveGlobal") refreshGlobalUnresolvedShiftCodesAfterRuleChange();
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

function findGlobalUnresolvedShiftCodeIssue(issueId = "") {
  const id = String(issueId || "").trim();
  if (!id) return null;
  const item = globalUnresolvedShiftCodes.find((issue) => issue.id === id);
  if (!item) return null;
  return {
    id: item.id,
    fingerprint: issueFingerprint(item.source, item.rawValue || item.code, item.seniority),
    source: item.source,
    seniority: item.seniority,
    code: item.code,
    rawValue: item.rawValue || item.code,
    startDay: item.sampleDate || "",
    message: item.message || "Shift code not recognised.",
    suggestedTitle: item.suggestedTitle || "",
    timeLabel: item.timeLabel || "",
  };
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

function isShiftCodeResolvedByGlobalRules(issue) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = parserRuleCodeForIssue(issue);
  if (!source || !code) return false;
  const sourceKey = source.toLowerCase();
  const rules = sanitizeParserExtensionRuleList(globalParserExtensions?.[sourceKey], source);
  if (seniority !== "Unknown") {
    return rules.some((rule) => rule.code === code && rule.seniority === seniority);
  }
  return rules.some((rule) => rule.code === code);
}

function isKnownResolvedShiftCodeValue(sourceValue, rawValue, normalizedTitle = "") {
  const source = sanitizeIssueSource(sourceValue);
  const code = parserRuleCodeFromRawValue(source, rawValue);
  if (!source || !code) return false;
  if (isIgnoredRosterIssueValue(source, rawValue || code)) return true;
  if (["AM", "PM", "NIGHT"].includes(code)) return true;
  if (code === "PHNW") return true;
  if (source === "MCH" && ["CS", "OCS", "0CS", "CSOS"].includes(code)) return true;
  if (source === "DDH") {
    if (["CS", "CS ONSITE", "SSU", "DAY OFF IN LIEU"].includes(code)) return true;
    if (/^(ORANGE|SILVER|FAST|AVAO|ROVER)\s+(AM|PM)$/.test(code)) return true;
  }
  const titleCode = incompleteShiftCodeFromTitle(source, normalizedTitle);
  return Boolean(titleCode && isShiftCodeResolvedByActiveRules({ source, seniority: "Unknown", code: titleCode }));
}

function isSystemRosterReviewNotice(issue) {
  return /^roster\s+supersession\s+review\s*:/i.test(String(issue?.rawValue || "").trim());
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
  const selectedSeniorities = Array.isArray(options.selectedSeniorities)
    ? options.selectedSeniorities.map(sanitizeRuleSeniority).filter(Boolean)
    : [];
  setParserRuleModalIssueFields({ ...issue, source, seniority, code }, selectedSeniorities.length ? selectedSeniorities : [seniority], {
    allowMultipleSeniorities: options.allowMultipleSeniorities,
    readonlySourceRaw: options.readonlySourceRaw,
    ignore: options.ignore,
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

function openRosterShiftCodeRuleModal(issueId = "", selectedSeniorities = [], options = {}) {
  const issue = findGlobalUnresolvedShiftCodeIssue(issueId);
  if (!issue) {
    setStatus("Could not find that parser warning.", true);
    return;
  }
  openParserRuleModalFromSyntheticIssue(issue, {
    mode: "global",
    targetEmail: "",
    allowMultipleSeniorities: true,
    selectedSeniorities,
    ignore: options.ignore === true,
    title: options.ignore === true ? "Ignore shift code" : "Edit shift code",
  });
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
    globalParserExtensions = sanitizeParserExtensions(data.parserExtensions);
    refreshGlobalUnresolvedShiftCodesAfterRuleChange();
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
    globalParserExtensions = sanitizeParserExtensions(data.parserExtensions);
    refreshGlobalUnresolvedShiftCodesAfterRuleChange();
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
      if (!accountsModal.classList.contains("hidden") && currentAdminTab === "parser") renderAccountsModal();
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
      if (data.user) {
        const updatedEmail = normalizeServerUser(data.user).email;
        serverUsers = [
          ...serverUsers.filter((user) => normalizeServerUser(user).email !== updatedEmail),
          data.user,
        ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
      }
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
  const nonClinical = formElement.querySelector("[data-create-non-clinical]")?.checked === true;
  const directorViewEnabled = formElement.querySelector("[data-create-director-view]")?.checked === true;
  if (!isValidEmailAddress(email)) {
    await showAppDialog({ title: "Invalid email address", message: "Please enter a valid email address" });
    return;
  }
  if (!realName || !password) {
    await showAppDialog({
      title: "Account details required",
      message: !realName ? "Please enter a real name" : "Please enter a temporary password",
    });
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
        nonClinical,
        directorViewEnabled,
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

async function sendAccountInvite(formElement) {
  if (!isCreatorAuthenticated()) return;
  const realName = formElement.querySelector("[data-create-real-name]")?.value.trim() || "";
  const email = normalizeEmail(formElement.querySelector("[data-create-email]")?.value || "");
  const nonClinical = formElement.querySelector("[data-create-non-clinical]")?.checked === true;
  const directorViewEnabled = formElement.querySelector("[data-create-director-view]")?.checked === true;
  if (!isValidEmailAddress(email)) {
    await showAppDialog({ title: "Invalid email address", message: "Please enter a valid email address" });
    return;
  }
  if (!realName) {
    await showAppDialog({ title: "Account details required", message: "Please enter a real name" });
    return;
  }
  const confirmed = await showAppDialog({
    title: "Confirm invitation email",
    message: `Send the account invitation to ${email}?`,
    confirmLabel: "Send invite",
  });
  if (!confirmed) return;
  setStatus(`Sending invitation to ${email}...`);
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "adminSendInvite",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail: email,
        targetRealName: realName,
        nonClinical,
        directorViewEnabled,
      }),
    });
    const data = await readJsonResponse(response, "Could not send invitation.");
    if (data.user) {
      serverUsers = [...serverUsers.filter((user) => normalizeServerUser(user).email !== email), data.user]
        .sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
    }
    formElement.reset();
    await loadServerUsers();
    renderAccountsModal();
    setStatus(`Invitation sent to ${email}. It expires in seven days.`);
  } catch (error) {
    setStatus(error.message || "Could not send invitation.", true);
  }
}

function isValidEmailAddress(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || ""));
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

async function setUserFacilityOverviewEnabled(email, enabled) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || !isCreatorAuthenticated()) return;
  const previousUsers = serverUsers.map((user) => ({ ...normalizeServerUser(user) }));
  serverUsers = serverUsers.map((user) => {
    const normalized = normalizeServerUser(user);
    return normalized.email === targetEmail ? { ...normalized, facilityOverviewEnabled: enabled === true } : normalized;
  });
  renderAccountsModal();
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setUserFacilityOverviewEnabled",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail,
        facilityOverviewEnabled: enabled === true,
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
      currentFacilityOverviewEnabled = enabled === true;
      syncFacilityOverviewAccess();
    }
    setStatus(enabled ? "At a glance ED overview enabled for that user." : "At a glance ED overview disabled for that user.");
  } catch (error) {
    serverUsers = previousUsers;
    renderAccountsModal();
    setStatus(error.message || "Could not update user feature access.", true);
  }
}

async function setUserDirectorViewEnabled(email, enabled) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || !isCreatorAuthenticated()) return;
  const previousUsers = serverUsers.map((user) => ({ ...normalizeServerUser(user) }));
  serverUsers = serverUsers.map((user) => {
    const normalized = normalizeServerUser(user);
    return normalized.email === targetEmail ? { ...normalized, directorViewEnabled: enabled === true } : normalized;
  });
  renderAccountsModal();
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "setUserDirectorViewEnabled",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail,
        directorViewEnabled: enabled === true,
      }),
    });
    const data = await readJsonResponse(response, "Could not update Director access.");
    if (data.user) {
      serverUsers = [
        ...serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail),
        data.user,
      ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
      renderAccountsModal();
    }
    if (targetEmail === currentUserEmail) currentDirectorViewEnabled = enabled === true;
    setStatus(enabled ? "Director access enabled for that user." : "Director access disabled for that user.");
  } catch (error) {
    serverUsers = previousUsers;
    renderAccountsModal();
    setStatus(error.message || "Could not update Director access.", true);
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
    const nameInput = editor.querySelector("[data-admin-user-real-name]");
    if (!isHidden && nameInput) nameInput.focus();
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

async function saveAdminUserName(email, realName) {
  const targetEmail = normalizeEmail(email);
  const nextRealName = String(realName || "").trim();
  if (!targetEmail || !isCreatorAuthenticated()) return;
  if (!nextRealName) {
    setStatus("Enter a user name before saving.", true);
    return;
  }
  try {
    const response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "updateAccount",
        email: authUserEmail || currentUserEmail,
        password: authUserPassword || currentUserPassword,
        targetEmail,
        realName: nextRealName,
      }),
    });
    const data = await readJsonResponse(response, "Could not update the user name.");
    if (data.user) {
      serverUsers = [
        ...serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail),
        data.user,
      ].sort((left, right) => normalizeServerUser(left).email.localeCompare(normalizeServerUser(right).email));
    }
    accountState.users = accountState.users.map((user) => normalizeEmail(user.email) === targetEmail
      ? { ...user, realName: nextRealName }
      : user);
    saveAccountState();
    renderAccountsModal();
    setStatus(`Updated ${nextRealName}.`);
  } catch (error) {
    setStatus(error.message || "Could not update the user name.", true);
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
  if (options.preserveRenderedSnapshot && visibleSnapshotIsCurrent({ requireNotStale: true })) {
    return;
  }
  const targetEmail = normalizeEmail(context.ownerEmail || context.ownerId);
  const cachedSnapshot = options.preserveRenderedSnapshot ? currentSnapshot : null;
  const cachedRevision = cachedSnapshot?.calendarRevision || "";
  await restoreCloudState({
    adminTargetEmail: targetEmail === OWNER_EMAIL ? "" : targetEmail,
    preserveSessionOnFailure: true,
    deferHydration: true,
    responseMode: "fast",
    cachedRevision,
    allowInlineBuild: !cachedSnapshot?.preview,
    preserveExistingSnapshot: Boolean(cachedSnapshot?.preview),
    accountSwitchStartedAt: options.accountSwitchStartedAt,
    transition: options.transition,
  });
  if (!calendarTransitionStillCurrent(options.transition)) return;
  if (cachedSnapshot?.preview && (currentSnapshotStale || !currentSnapshot?.preview || currentSnapshot.cacheKey !== cachedSnapshot.cacheKey)) {
    currentSnapshot = cachedSnapshot;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = cachedSnapshot.cachedAt || cachedSnapshot.preview?.lastParsed || "";
  }
  if (!cachedSnapshot?.preview && currentSnapshot?.preview) {
    renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
    setStatus(currentSnapshotStale ? "Refreshing calendar..." : "Calendar loaded.");
    renderLoginState();
  }
  await hydrateAuthenticatedWorkspace({
    adminTargetEmail: targetEmail === OWNER_EMAIL ? "" : targetEmail,
    includeBootstrap: true,
    accountSwitchStartedAt: options.accountSwitchStartedAt,
    doctorKey: context.doctorKey || "",
    cachedRevision,
    allowInlineBuild: !cachedSnapshot?.preview,
    transition: options.transition,
  }, 0);
  if (!calendarTransitionStillCurrent(options.transition)) return;
  renderLoginState();
}

async function validateDoctorProfileCalendarInBackground(doctor, previousState, options = {}) {
  // A browser profile cache can be complete enough to render immediately, but
  // it is never authoritative. In particular, its revision may have been
  // copied from a different calendar context before this profile's snapshot
  // was refreshed. Always obtain this profile's full server snapshot before
  // declaring the visible calendar current. Do not send cachedRevision here:
  // the profile endpoint intentionally treats it as a validation-only request
  // and may therefore omit the snapshot body we need to replace stale events.
  let result = await loadUnclaimedDoctorCalendar(doctor, previousState, {
    profile: options.profile,
    cachedRevision: "",
    allowInlineBuild: false,
    transition: options.transition,
  });
  if (!result && !calendarSnapshotMatchesActiveContext(currentSnapshot)) {
    result = await waitForDoctorProfileCalendarBuild(doctor, previousState, {
      profile: options.profile,
      transition: options.transition,
    });
  }
  if (!calendarTransitionStillCurrent(options.transition)) return;
  if (result) {
    await commitCalendarLoad(result, { saveInBackground: true, transition: options.transition });
  } else if (calendarSnapshotMatchesActiveContext(currentSnapshot)) {
    renderLoginState();
    setStatus("Calendar is up to date.");
  } else {
    throw new Error(`${doctor.displayName} calendar is not ready yet. Try again in a moment.`);
  }
}

async function waitForDoctorProfileCalendarBuild(doctor, previousState, options = {}) {
  const retryDelays = [250, 500, 1000, 1500, 2500, 4000, 5000];
  let attempt = 0;
  while (calendarTransitionStillCurrent(options.transition)) {
    const delayMs = retryDelays[Math.min(attempt, retryDelays.length - 1)];
    attempt += 1;
    await new Promise((resolve) => window.setTimeout(resolve, delayMs));
    if (!calendarTransitionStillCurrent(options.transition)) return null;
    const result = await loadUnclaimedDoctorCalendar(doctor, previousState, {
      profile: options.profile,
      allowInlineBuild: false,
      transition: options.transition,
    });
    if (result || calendarSnapshotMatchesActiveContext(currentSnapshot)) return result;
  }
  return null;
}

async function enterUserAccount(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || (!isOwnerAccount() && !isCreatorAuthenticated())) return;
  const previousState = captureCalendarViewState();
  beginFacilityOverviewAccountSession();
  const accountSwitchStartedAt = performance.now();
  const creatorEmail = authUserEmail || currentUserEmail;
  const creatorPassword = authUserPassword || currentUserPassword;
  if (normalizeEmail(creatorEmail) !== OWNER_EMAIL || !creatorPassword) {
    setStatus("Creator authentication is required to enter another account.", true);
    return;
  }
  cancelScheduledCloudStateSave();
  const outgoingSave = capturePendingCloudStateSave() || outgoingSnapshotSavePayload(previousState);
  if (outgoingSave) queueBackgroundCloudStateSave(outgoingSave, { delayMs: 1500 });

  closeAccountsModal();
  authUserEmail = creatorEmail;
  authUserPassword = creatorPassword;
  adminViewingEmail = targetEmail;
  viewedAccountId = targetEmail;
  viewedAccountType = "claimed-user";
  isImpersonating = true;
  impersonatedByCreator = true;
  returnToCreatorAvailable = true;
  currentDefaultDoctorKey = preferredDoctorKeyForAccountEmail(targetEmail);
  activeDoctorProfile = null;
  setActiveCalendarContext(targetEmail === OWNER_EMAIL ? "creator-account" : "claimed-account", { email: targetEmail });
  currentUserEmail = targetEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = targetEmail === OWNER_EMAIL ? "creator" : "user";
  primeInsightsAccessForCurrentView();
  const transition = beginCalendarTransition();
  forceConsoleSkin();
  setStatus(`Entering ${targetEmail}...`);
  const targetContext = accountCalendarContextForEmail(targetEmail);
  const renderedCachedSnapshot = await renderCachedCalendarSnapshotForContextAsync(targetContext, { accountSwitchStartedAt, transition });
  if (!renderedCachedSnapshot) {
    clearPreviewData();
    doctorOptions = [];
    selectedFiles = [];
  }
  renderLoginState();
  try {
    const validation = validateClaimedAccountCalendarInBackground(targetContext, {
      accountSwitchStartedAt,
      preserveRenderedSnapshot: renderedCachedSnapshot,
      transition,
    });
    if (renderedCachedSnapshot) {
      void validation.catch((error) => {
        if (!calendarTransitionStillCurrent(transition)) return;
        reportBackgroundValidationError(error, {
          preserveRenderedSnapshot: true,
          fallbackMessage: `Could not update ${targetEmail}.`,
        });
      });
    } else {
      await validation;
    }
  } catch (error) {
    if (!calendarTransitionStillCurrent(transition)) return;
    restoreCalendarViewState(previousState);
    renderLoginState();
    setStatus(normalizeAuthMessage(error.message || `Could not enter ${targetEmail}.`), true);
  }
}

async function enterDoctorProfileView(doctor) {
  if (!isOwnerAccount() && !isCreatorAuthenticated()) return;
  rememberCreatorCalendarSourceRefs();
  const previousState = captureCalendarViewState();
  beginFacilityOverviewAccountSession();
  const creatorEmail = authUserEmail || currentUserEmail;
  const creatorPassword = authUserPassword || currentUserPassword;
  cancelScheduledCloudStateSave();
  const outgoingSave = capturePendingCloudStateSave() || outgoingSnapshotSavePayload(previousState);
  if (outgoingSave) queueBackgroundCloudStateSave(outgoingSave, { delayMs: 1500 });
  const profile = doctorProfileForDoctor(doctor);
  const targetContext = profile ? calendarSnapshotContext({
    mode: "doctor-profile",
    ownerId: profile.ownerId,
    doctorKey: profile.doctorKey,
  }) : null;
  adminViewingEmail = "";
  viewedAccountId = "";
  viewedAccountType = "unclaimed-user";
  isImpersonating = false;
  impersonatedByCreator = false;
  returnToCreatorAvailable = true;
  currentDefaultDoctorKey = normalizeRosterName(profile?.doctorKey || "");
  currentUserEmail = creatorEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = "creator";
  activeDoctorProfile = profile;
  if (profile) {
    currentRosterClaims = rosterClaimsForDoctorProfile(profile);
    currentSuggestedClaims = [];
  }
  if (profile) setActiveCalendarContext("doctor-profile", { email: currentUserEmail, profile });
  primeInsightsAccessForCurrentView();
  const accountSwitchStartedAt = performance.now();
  const transition = beginCalendarTransition();
  localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  setStatus(`Opening ${doctor.displayName}...`);
  const renderedCachedSnapshot = targetContext
    ? await renderCachedCalendarSnapshotForContextAsync(targetContext, { accountSwitchStartedAt, transition })
    : false;
  if (!renderedCachedSnapshot) {
    currentSnapshot = null;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = "";
    currentCalendarRevision = "";
    resetTransientCalendarData();
    doctorOptions = [];
  }
  renderLoginState();
  try {
    const validation = validateDoctorProfileCalendarInBackground(doctor, previousState, {
      profile,
      cachedRevision: renderedCachedSnapshot ? currentSnapshot?.calendarRevision || "" : "",
      renderedCachedSnapshot,
      transition,
    });
    if (renderedCachedSnapshot) {
      void validation.catch((error) => {
        if (!calendarTransitionStillCurrent(transition)) return;
        reportBackgroundValidationError(error, {
          preserveRenderedSnapshot: true,
          fallbackMessage: `Could not update ${doctor.displayName}.`,
        });
      });
    } else {
      await validation;
    }
  } catch (error) {
    if (!calendarTransitionStillCurrent(transition)) return;
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
  const profileData = await fetchDoctorProfileState(profile, {
    cachedRevision: options.cachedRevision || "",
    allowInlineBuild: options.allowInlineBuild !== false,
  });
  if (!calendarTransitionStillCurrent(options.transition)) return null;
  if (profileData.snapshotCurrent === true && calendarSnapshotMatchesActiveContext(currentSnapshot)) {
    currentCalendarRevision = String(profileData.calendarRevision || currentCalendarRevision || "");
    if (Array.isArray(profileData.fileRefs) && profileData.fileRefs.length) {
      selectedFiles = importRefsToClientEntries(profileData.fileRefs);
      if (currentSnapshot) {
        currentSnapshot.fileRefs = profileData.fileRefs;
        currentSnapshot.calendarRevision = currentCalendarRevision;
        saveCalendarSnapshotCacheForContext(currentSnapshot, {
          mode: "doctor-profile",
          ownerId: profile.ownerId,
          doctorKey: profile.doctorKey,
        });
      }
    } else if (currentSnapshot) {
      currentSnapshot.calendarRevision = currentCalendarRevision;
      saveCalendarSnapshotCacheForContext(currentSnapshot, {
        mode: "doctor-profile",
        ownerId: profile.ownerId,
        doctorKey: profile.doctorKey,
      });
    }
    return null;
  }
  const snapshot = sanitizeWorkspaceSnapshot(profileData.snapshot);
  if (snapshot && profileData.calendarRevision) snapshot.calendarRevision = String(profileData.calendarRevision);
  const snapshotUsable = snapshot?.preview && snapshot?.doctorOptions?.length;
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
  if (options.allowInlineBuild === false && calendarSnapshotMatchesActiveContext(currentSnapshot)) return null;
  if (options.allowInlineBuild === false) {
    const building = profileData.snapshotSource === "server-cache-building" || profileData.snapshotStatus === "building";
    setStatus(`${doctor.displayName} calendar is ${building ? "building" : "preparing"}...`);
    return null;
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
      allowInlineBuild: options.allowInlineBuild !== false,
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
  if (!calendarTransitionStillCurrent(options.transition)) return;
  if (!result || result.mode !== "doctor-profile") {
    throw new Error("Unsupported calendar load result.");
  }
  activeDoctorProfile = result.profile;
  currentRosterClaims = rosterClaimsForDoctorProfile(result.profile);
  currentSuggestedClaims = [];
  setActiveCalendarContext("doctor-profile", { email: currentUserEmail, profile: activeDoctorProfile });
  resetTransientCalendarData();
  currentDefaultDoctorKey = normalizeRosterName(result.profile?.doctorKey || result.doctor?.key || "");
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
  refreshFacilityOverviewPreferredFacility();
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
  if (!calendarTransitionStillCurrent(options.transition)) return;
  renderLoginState();
  setStatus("Calendar loaded.");
}

function captureCalendarViewState() {
  return {
    adminViewingEmail,
    activeDoctorProfile,
    viewedAccountId,
    viewedAccountType,
    isImpersonating,
    impersonatedByCreator,
    returnToCreatorAvailable,
    activeCalendarContext: activeCalendarContext ? { ...activeCalendarContext } : null,
    currentUserEmail,
    currentUserPassword,
    currentUserRole,
    selectedFiles: selectedFiles.map((entry) => ({ ...entry })),
    creatorCalendarSourceFileRefs: creatorCalendarSourceFileRefs.map((entry) => ({ ...entry })),
    currentSnapshot: currentSnapshot ? JSON.parse(JSON.stringify(currentSnapshot)) : null,
    currentSnapshotStale,
    currentSnapshotBuiltAt,
    facilityOverviewSessionNeedsInitialization,
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
  viewedAccountId = state.viewedAccountId;
  viewedAccountType = state.viewedAccountType;
  isImpersonating = state.isImpersonating;
  impersonatedByCreator = state.impersonatedByCreator;
  returnToCreatorAvailable = state.returnToCreatorAvailable;
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
  facilityOverviewSessionNeedsInitialization = state.facilityOverviewSessionNeedsInitialization === true;
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
  viewedAccountId = normalizeEmail(currentUserEmail);
  viewedAccountType = "creator";
  isImpersonating = false;
  impersonatedByCreator = false;
  returnToCreatorAvailable = false;
  currentDefaultDoctorKey = OWNER_DOCTOR_KEY;
  setActiveCalendarContext("creator-account", { email: currentUserEmail });
  clearPreviewData();
  restoredSessionState = loadCurrentSessionState();
  currentSnapshot = sanitizeWorkspaceSnapshot(loadCurrentWorkspace()?.snapshot);
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = "";
  await bootstrapImports();
  renderLoginState();
}

async function returnToCreatorCalendar(options = {}) {
  await returnToCreatorAccount(options);
}

async function returnToCreatorAccount(options = {}) {
  const previousState = captureCalendarViewState();
  beginFacilityOverviewAccountSession();
  const accountSwitchStartedAt = performance.now();
  const creatorEmail = authUserEmail || OWNER_EMAIL;
  const creatorPassword = authUserPassword || currentUserPassword;
  cancelScheduledCloudStateSave();
  if (options.skipOutgoingSave !== true) {
    const outgoingSave = capturePendingCloudStateSave() || outgoingSnapshotSavePayload(previousState);
    if (outgoingSave) queueBackgroundCloudStateSave(outgoingSave, { delayMs: 1500 });
  }
  adminViewingEmail = "";
  viewedAccountId = normalizeEmail(creatorEmail);
  viewedAccountType = "creator";
  isImpersonating = false;
  impersonatedByCreator = false;
  returnToCreatorAvailable = false;
  currentDefaultDoctorKey = OWNER_DOCTOR_KEY;
  activeDoctorProfile = null;
  currentUserEmail = creatorEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = "creator";
  setActiveCalendarContext("creator-account", { email: currentUserEmail });
  primeInsightsAccessForCurrentView();
  const transition = beginCalendarTransition();
  localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  forceConsoleSkin();
  setStatus("Returning to creator account...");
  forceCreatorDoctorSession();
  restoreCreatorImportFilesIfNeeded();
  const targetContext = accountCalendarContextForEmail(OWNER_EMAIL);
  const renderedCachedSnapshot = await renderCachedCalendarSnapshotForContextAsync(targetContext, { accountSwitchStartedAt, transition });
  restoreCreatorImportFilesIfNeeded();
  renderFileSurfaces();
  renderLoginState();
  const validateCreator = async () => {
    await validateClaimedAccountCalendarInBackground(targetContext, {
      accountSwitchStartedAt,
      preserveRenderedSnapshot: renderedCachedSnapshot,
      transition,
    });
    if (!calendarTransitionStillCurrent(transition)) return;
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
    if (selectedDoctor()?.key !== OWNER_DOCTOR_KEY || cloudAvailable) {
      if (cloudAvailable) {
        if (!visibleSnapshotIsCurrent({ requireNotStale: true })) {
          const loaded = await loadCloudCalendarEvents({
            doctorKey: OWNER_DOCTOR_KEY,
            cachedRevision: currentSnapshot?.calendarRevision || currentCalendarRevision || "",
            allowInlineBuild: false,
            preserveExistingSnapshot: true,
            transition,
          });
          if (loaded && currentSnapshot) {
            renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
          }
        }
      } else {
        clearPreviewData();
        await updatePreview({ resetRange: false });
      }
    }
    void syncCreatorFileListFromStore().catch(() => null);
    renderLoginState();
  };
  if (renderedCachedSnapshot) {
    void validateCreator().catch((error) => {
      if (!calendarTransitionStillCurrent(transition)) return;
      reportBackgroundValidationError(error, {
        preserveRenderedSnapshot: true,
        fallbackMessage: "Could not update the creator calendar.",
      });
    });
  } else {
    try {
      await validateCreator();
    } catch (error) {
      if (!calendarTransitionStillCurrent(transition)) return;
      if (options.restoreOnFailure !== false) restoreCalendarViewState(previousState);
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
    ownerType: String(value.ownerType || "").trim(),
    ownerId: String(value.ownerId || "").trim(),
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

function restoreCreatorImportFilesIfNeeded() {
  if (!isViewingCreatorAccount() || selectedFiles.length) return false;
  const skipIds = pendingRemovedImportIds;
  const refsFromMemory = creatorCalendarSourceFileRefs.filter((entry) => entry?.id && !skipIds.has(entry.id));
  if (refsFromMemory.length) {
    selectedFiles = importRefsToClientEntries(refsFromMemory);
    return selectedFiles.length > 0;
  }
  const refsFromSnapshot = (currentSnapshot?.fileRefs || []).filter((entry) => entry?.id && !skipIds.has(entry.id));
  if (refsFromSnapshot.length) {
    selectedFiles = importRefsToClientEntries(refsFromSnapshot);
    if (selectedFiles.length) rememberCreatorCalendarSourceRefs();
    return selectedFiles.length > 0;
  }
  return false;
}

async function syncCreatorFileListFromStore(options = {}) {
  if (!isViewingCreatorAccount() || !cloudAvailable) return;
  if (!selectedFiles.length) restoreCreatorImportFilesIfNeeded();
  await refreshCalendarStoreStatus({ silent: true, includeAvailableDoctors: options.includeAvailableDoctors === true }).catch(() => null);
  if ((calendarStoreStatus?.files || []).length) {
    mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, {
      force: true,
      removeMissingFromStore: options.removeMissingFromStore === true,
      removedIds: options.removedIds || [],
    });
  } else if (!selectedFiles.length) {
    restoreCreatorImportFilesIfNeeded();
  }
  renderFileSurfaces();
}

function applyLoadedCalendarFileRefs(snapshot) {
  const snapshotRefs = (Array.isArray(snapshot?.fileRefs) ? snapshot.fileRefs : [])
    .filter((entry) => entry?.id && !pendingRemovedImportIds.has(entry.id));
  if (snapshotRefs.length) {
    const fromSnapshot = importRefsToClientEntries(snapshotRefs);
    selectedFiles = isViewingCreatorAccount() && Array.isArray(calendarStoreStatus?.files) && calendarStoreStatus.files.length
      ? mergeRosterFileEntries(fromSnapshot, calendarStoreStatus)
      : fromSnapshot;
    rememberCreatorCalendarSourceRefs();
    return;
  }
  if (isViewingCreatorAccount() && !selectedFiles.length) {
    restoreCreatorImportFilesIfNeeded();
  }
}

function rosterStoreFileToClientEntry(file) {
  if (!file?.id) return null;
  return {
    id: file.id,
    repoId: file.id,
    name: file.name || "roster.xlsx",
    sourceType: file.sourceType || "",
    size: Number(file.size || 0),
    lastModified: Number(file.lastModified || 0),
    addedAt: file.uploadedAt || "",
    sourceId: String(file.sourceId || ""),
    startDate: String(file.startDate || ""),
    coverageEndDate: String(file.coverageEndDate || ""),
    endDate: String(file.endDate || ""),
    fromRosterDatabase: true,
  };
}

function mergeRosterFileEntries(baseEntries, status = calendarStoreStatus, options = {}) {
  if (!isViewingCreatorAccount() || !status || !Array.isArray(status.files)) {
    return Array.isArray(baseEntries) ? [...baseEntries] : [];
  }
  const storeIds = new Set(status.files.map((file) => file.id).filter(Boolean));
  const removedIds = new Set([
    ...(Array.isArray(options.removedIds) ? options.removedIds.filter(Boolean) : []),
    ...pendingRemovedImportIds,
  ]);
  const byId = new Map();
  const includeRetainedSourceEntries = options.includeRetainedSourceEntries === true;
  const retainedSourceIsBeingImported = (file) => {
    const syncState = rosterSyncStates.get(file?.id);
    return Boolean(syncState && ["pending", "uploading-source", "parsing", "saving"].includes(syncState.status));
  };

  for (const entry of baseEntries || []) {
    if (!entry?.id || removedIds.has(entry.id)) continue;
    const storeFile = status.files.find((file) => file?.id === entry.id);
    if (storeFile?.retainedSourceOnly && !includeRetainedSourceEntries && !retainedSourceIsBeingImported(storeFile)) continue;
    if (String(entry.id).startsWith("automation:") && !storeIds.has(entry.id)) continue;
    if (options.removeMissingFromStore && storeIds.size && !storeIds.has(entry.id)) continue;
    byId.set(entry.id, entry);
  }

  for (const file of status.files) {
    if (!file?.id || removedIds.has(file.id)) continue;
    if (file.retainedSourceOnly && !includeRetainedSourceEntries && !retainedSourceIsBeingImported(file)) continue;
    const storeEntry = rosterStoreFileToClientEntry(file);
    if (!storeEntry) continue;
    const existing = byId.get(file.id);
    if (existing) {
      byId.set(file.id, {
        ...storeEntry,
        ...existing,
        sourceType: existing.sourceType === "pending"
          ? (storeEntry.sourceType || existing.sourceType)
          : (existing.sourceType || storeEntry.sourceType),
        file: existing.file || null,
        addedAt: existing.addedAt || storeEntry.addedAt,
        fromRosterDatabase: !existing.file,
      });
    } else {
      byId.set(file.id, storeEntry);
    }
  }

  return [...byId.values()].sort(
    (left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || String(left.name || "").localeCompare(String(right.name || "")),
  );
}

function mergeSelectedFilesWithRosterStoreStatus(status = calendarStoreStatus, options = {}) {
  const next = mergeRosterFileEntries(selectedFiles, status, options);
  const prevKey = selectedFiles.map((entry) => entry.id).join("|");
  const nextKey = next.map((entry) => entry.id).join("|");
  selectedFiles = next;
  if (prevKey !== nextKey || options.force === true) rememberCreatorCalendarSourceRefs();
  return selectedFiles;
}

function calendarFilesForActiveView() {
  const fromSelected = (selectedFiles || []).filter((entry) => entry?.id);
  if (fromSelected.length) return fromSelected;
  let fromSnapshot = importRefsToClientEntries(currentSnapshot?.fileRefs || []);
  if (activeCalendarMode() === "doctor-profile" && activeDoctorProfile?.sourceTypes?.length) {
    const allowedSources = new Set(activeDoctorProfile.sourceTypes.map((item) => String(item || "").toLowerCase()));
    fromSnapshot = fromSnapshot.filter((entry) => allowedSources.has(String(entry.sourceType || "").toLowerCase()));
  }
  return fromSnapshot;
}

function rosterDisplayFiles(hasUsableStatus, statusOnlyEntries = []) {
  if (activeCalendarMode() === "doctor-profile") {
    return calendarFilesForActiveView();
  }
  if (isViewingCreatorAccount() && hasUsableStatus && (calendarStoreStatus?.files || []).length) {
    return mergeRosterFileEntries(selectedFiles, calendarStoreStatus, { includeRetainedSourceEntries: true })
      .filter((entry) => !pendingRemovedImportIds.has(entry.id));
  }
  return (selectedFiles.length ? selectedFiles : statusOnlyEntries)
    .filter((entry) => !pendingRemovedImportIds.has(entry.id));
}

function selectedFilesNeedD1CalendarReload() {
  if (!isViewingCreatorAccount() || !cloudAvailable || !selectedFiles.length) return false;
  return selectedFiles.some((entry) => !entry.file);
}

function selectedFilesHavePendingD1Uploads(status = calendarStoreStatus) {
  if (!selectedFiles.length) return false;
  const failedIds = new Set(
    [...rosterSyncStates.entries()]
      .filter(([, state]) => state.status === "failed")
      .map(([id]) => id),
  );
  return selectedFiles.some(
    (entry) => entry?.file && (!isLocalRosterFileSyncedToD1(entry, status) || failedIds.has(entry.id)),
  );
}

async function refreshCreatorCalendarAfterFileChange(options = {}) {
  if (!selectedFiles.length) return;
  const afterRosterRemoval = options.removeMissingFromStore === true;
  if (isViewingCreatorAccount() && cloudAvailable) {
    captureCreatorSwitcherVisibleBaseline();
  }
  if (options.refreshStatus !== false && isCreatorAuthenticated()) {
    await refreshCalendarStoreStatus({ silent: true }).catch(() => null);
  }
  if (isViewingCreatorAccount()) {
    mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, {
      removeMissingFromStore: afterRosterRemoval,
      removedIds: options.removedIds || [],
      force: options.force === true,
    });
  }
  let performedLocalUpload = false;
  const unsyncedLocalEntries = selectedFiles.filter((entry) => entry?.file && (
    entry.needsD1Resync === true || !isLocalRosterFileSyncedToD1(entry)
  ));
  if (unsyncedLocalEntries.length && cloudAvailable && isCreatorAuthenticated()) {
    performedLocalUpload = true;
    try {
      await saveSelectedRosterFilesToD1(unsyncedLocalEntries, {
        force: unsyncedLocalEntries.some((entry) => entry.needsD1Resync === true),
      });
      for (const entry of unsyncedLocalEntries) {
        if (entry) entry.needsD1Resync = false;
      }
    } catch (error) {
      setStatus(error.message || "Could not save roster file to D1.", true);
    }
    if (isCreatorAuthenticated()) {
      await refreshCalendarStoreStatus({ silent: true }).catch(() => null);
      mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
    }
    try {
      await syncCreatorDoctorPickerWithRemainingRosters({ localOnly: afterRosterRemoval });
      renderDoctorState();
    } catch {
      // Keep the repository-backed doctor list until the next status refresh.
    }
    renderFileSurfaces();
  }
  if (isViewingCreatorAccount() && cloudAvailable) {
    let deferSwitcherAnnouncementToPoll = false;
    try {
      await refreshAvailableDoctorsAfterRosterChange({ localOnly: afterRosterRemoval });
      mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
      renderFileSurfaces();
      setStatus(performedLocalUpload ? "Updating calendar..." : "Loading calendar...");
      try {
        const loaded = await loadCloudCalendarEvents({
          preserveExistingSnapshot: !afterRosterRemoval,
          allowInlineBuild: afterRosterRemoval,
          cachedRevision: "",
        });
        if (loaded && currentSnapshot && !currentSnapshotStale) {
          renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
          mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
          try {
            await syncCreatorDoctorPickerWithRemainingRosters({ localOnly: afterRosterRemoval });
          } catch {
            // Keep the last merged doctor list.
          }
          renderDoctorState();
          setStatus(performedLocalUpload ? "Calendar refreshed." : "Calendar loaded.");
        } else if (performedLocalUpload || currentSnapshotStale) {
          setStatus("Roster saved. Calendar snapshot is building...");
          deferSwitcherAnnouncementToPoll = true;
          void pollCalendarAfterRosterChange();
        } else if (!loaded) {
          setStatus("Could not reload calendar from roster database.", true);
        }
      } catch (error) {
        const overload = /503|CPU|memory|overload/i.test(String(error?.message || ""));
        setStatus(overload
          ? "Roster saved. Calendar snapshot is building..."
          : (error.message || "Could not reload calendar from roster database."), overload);
        deferSwitcherAnnouncementToPoll = true;
        void pollCalendarAfterRosterChange();
      }
    } finally {
      if (!deferSwitcherAnnouncementToPoll) {
        await tryAnnounceCreatorSwitcherRosterUpdate();
      }
    }
    return;
  }
  await analyzeFiles(options.analyzeOptions || {});
  await refreshAvailableDoctorsAfterRosterChange({ localOnly: afterRosterRemoval });
  mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
  renderFileSurfaces();
}

async function rosterDoctorsFromSelectedFiles(entries = selectedFiles) {
  const hydrated = [];
  for (const entry of entries || []) {
    if (!entry?.id) continue;
    hydrated.push(await ensureRosterEntrySource(entry).catch(() => entry));
  }
  if (!hydrated.some((entry) => entry?.file)) return [];
  const parsed = await parseRosterEntriesLenient(hydrated);
  return rosterDoctorOptions(parsed.sources.mmc, parsed.sources.ddh, parsed.sources.casey, parsed.sources.mch);
}

function availableDoctorsFromRosterDoctorOptions(doctors = []) {
  return sanitizeAvailableRosterDoctors(
    (doctors || []).flatMap((doctor) => {
      const sourceTypes = normalizedDoctorSourceTypes(doctor);
      if (!sourceTypes.length) return [];
      return sourceTypes.map((sourceType) => ({
        key: doctor.key,
        displayName: doctor.displayName,
        sourceType,
        sourceTypes,
        aliases: Array.isArray(doctor.aliases) ? doctor.aliases : [],
      }));
    }),
  );
}

function mergeAvailableRosterDoctors(localDoctors = [], repositoryDoctors = availableRosterDoctors, options = {}) {
  const localOnly = options.localOnly === true;
  if (!localDoctors.length) {
    return localOnly ? [] : sanitizeAvailableRosterDoctors(repositoryDoctors);
  }
  const claimMetadata = new Map(
    (repositoryDoctors || []).map((doctor) => [doctorIdentityKey(doctor), doctor]),
  );
  const merged = localDoctors.map((doctor) => {
    const existing = claimMetadata.get(doctorIdentityKey(doctor));
    return existing ? {
      ...doctor,
      claimedBy: existing.claimedBy || "",
      claimedByName: existing.claimedByName || "",
      accountEmail: existing.accountEmail || "",
    } : doctor;
  });
  if (localOnly) return sanitizeAvailableRosterDoctors(merged);
  const knownKeys = new Set(merged.map((doctor) => doctorIdentityKey(doctor)));
  for (const doctor of repositoryDoctors || []) {
    const identity = doctorIdentityKey(doctor);
    if (knownKeys.has(identity)) continue;
    merged.push(doctor);
    knownKeys.add(identity);
  }
  return sanitizeAvailableRosterDoctors(merged);
}

async function syncCreatorDoctorPickerWithRemainingRosters(options = {}) {
  if (!canUseCreatorDoctorSwitcher()) return;
  const localOnly = options.localOnly === true;
  if (!selectedFiles.length) {
    availableRosterDoctors = [];
    return;
  }
  await ensureSelectedFilesLoaded().catch(() => null);
  let localDoctors = availableDoctorsFromRosterDoctorOptions(await rosterDoctorsFromSelectedFiles());
  if (localDoctors.length) {
    availableRosterDoctors = mergeAvailableRosterDoctors(localDoctors, availableRosterDoctors, { localOnly });
    return;
  }
  if (localOnly) return;
  if (!localDoctors.length) {
    const snapshotDoctors = options.snapshotDoctors || doctorOptions;
    if (snapshotDoctors.length) {
      localDoctors = availableDoctorsFromRosterDoctorOptions(snapshotDoctors);
    }
  }
  if (localDoctors.length) {
    availableRosterDoctors = mergeAvailableRosterDoctors(localDoctors, availableRosterDoctors);
    return;
  }
  availableRosterDoctors = sanitizeAvailableRosterDoctors(availableRosterDoctors);
}

async function pollCalendarAfterRosterChange() {
  if (calendarImportPollPromise) return calendarImportPollPromise;
  const pollRunId = calendarImportPollRunId;
  calendarImportPollPromise = (async () => {
    try {
      for (const delay of [2000, 5000, 10000, 20000]) {
        await new Promise((resolve) => setTimeout(resolve, delay));
        if (pollRunId !== calendarImportPollRunId) return;
        if (!isViewingCreatorAccount() || !cloudAvailable) return;
        try {
          const loaded = await loadCloudCalendarEvents({
            preserveExistingSnapshot: true,
            allowInlineBuild: false,
            cachedRevision: currentSnapshotStale ? "" : (currentSnapshot?.calendarRevision || currentCalendarRevision || ""),
            doctorKey: OWNER_DOCTOR_KEY,
          });
          if (pollRunId !== calendarImportPollRunId) return;
          if (loaded && currentSnapshot && !currentSnapshotStale) {
            renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
            mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
            try {
              await syncCreatorDoctorPickerWithRemainingRosters();
            } catch {
              // Keep the last merged doctor list.
            }
            renderDoctorState();
            setStatus("Calendar loaded.");
            void syncCreatorFileListFromStore({ includeAvailableDoctors: true }).catch(() => null);
            return;
          }
        } catch {
          // Keep polling until the warmed snapshot is ready.
        }
      }
      if (pollRunId !== calendarImportPollRunId) return;
      setStatus("Calendar snapshot is still building. Switch doctor or refresh again shortly.");
    } finally {
      await tryAnnounceCreatorSwitcherRosterUpdate();
    }
  })().finally(() => {
    calendarImportPollPromise = null;
  });
  return calendarImportPollPromise;
}

async function refreshCreatorSnapshotInBackground(options = {}) {
  if (!isViewingCreatorAccount() || !cloudAvailable) return false;
  try {
    const loaded = await loadCloudCalendarEvents({
      preserveExistingSnapshot: true,
      allowInlineBuild: false,
      cachedRevision: currentSnapshot?.calendarRevision || currentCalendarRevision || "",
      doctorKey: OWNER_DOCTOR_KEY,
      transition: options.transition,
    });
    if (!calendarTransitionStillCurrent(options.transition)) return false;
    if (loaded && currentSnapshot?.preview) {
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      currentSnapshotStale = false;
      currentSnapshotBuiltAt = new Date().toISOString();
      void syncCreatorFileListFromStore().catch(() => null);
      setStatus("Calendar refreshed.");
      return true;
    }
    if (latestPreview) {
      setStatus("Calendar snapshot is building...");
    } else {
      setStatus("Could not reload calendar from roster database.", true);
    }
    return false;
  } catch (error) {
    const overload = /503|CPU|memory|overload/i.test(String(error?.message || ""));
    setStatus(overload
      ? "Calendar snapshot is building..."
      : (error.message || "Could not reload calendar from roster database."), overload);
    return false;
  }
}

async function refreshAvailableDoctorsAfterRosterChange(options = {}) {
  if (!isCreatorAuthenticated() || !cloudAvailable) return;
  const localOnly = options.localOnly === true;
  const mergeAvailableDoctors = options.mergeAvailableDoctors === true;
  mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
  try {
    await syncCreatorDoctorPickerWithRemainingRosters({ localOnly });
  } catch {
    // Keep the last repository-backed doctor list.
  }
  for (const delay of [0, 3000, 10000, 20000]) {
    if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
    await refreshCalendarStoreStatus({
      silent: true,
      includeAvailableDoctors: true,
      mergeAvailableDoctors,
    }).catch(() => null);
    try {
      await syncCreatorDoctorPickerWithRemainingRosters({ localOnly });
    } catch {
      // Keep the last merged doctor list.
    }
  }
  renderDoctorState();
  syncAccountsButton();
  renderFileSurfaces();
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
      defaultDoctorKey: "",
      insightsEnabled: false,
      facilityOverviewEnabled: false,
      nonClinical: false,
      directorViewEnabled: false,
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
    defaultDoctorKey: normalizeRosterName(value?.defaultDoctorKey || ""),
    insightsEnabled: role === "owner" || role === "creator" || value?.insightsEnabled === true,
    facilityOverviewEnabled: role === "owner" || role === "creator" || value?.facilityOverviewEnabled === true,
    nonClinical: value?.nonClinical === true,
    directorViewEnabled: role === "owner" || role === "creator" || value?.directorViewEnabled === true,
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
  const stayLoggedInPreference = localStorage.getItem(STAY_LOGGED_IN_PREFERENCE_KEY);
  if (stayLoggedIn) stayLoggedIn.checked = stayLoggedInPreference === null ? true : stayLoggedInPreference === "true";
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
  beginFacilityOverviewAccountSession();
  try {
    await flushCloudStateSave();
  } catch {
    // Keep logout moving even if cloud persistence fails.
  }
  cancelScheduledCloudStateSave();
  localStorage.removeItem(CURRENT_EMAIL_KEY);
  sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
  localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
  clearActiveViewedAccountState();
  currentUserEmail = "";
  currentUserPassword = "";
  authUserEmail = "";
  authUserPassword = "";
  adminViewingEmail = "";
  viewedAccountId = "";
  viewedAccountType = "claimed-user";
  isImpersonating = false;
  impersonatedByCreator = false;
  returnToCreatorAvailable = false;
  currentDefaultDoctorKey = "";
  currentSnapshotOwnerType = "";
  currentSnapshotOwnerId = "";
  currentUserRole = "user";
  cloudAvailable = false;
  setActiveCalendarContext("claimed-account", { email: "" });
  currentRosterClaims = [];
  latestNameMatches = [];
  availableRosterDoctors = [];
  currentSubscription = null;
  currentInsightsEnabled = false;
  currentFacilityOverviewEnabled = false;
  currentNonClinical = false;
  currentDirectorViewEnabled = false;
  closeFacilityOverview();
  currentSuggestedClaims = [];
  selectedFiles = [];
  switchTargetPrefetchRunId += 1;
  switchTargetPrefetchPromise = null;
  resetDerivedState();
  renderLoginState();
  openLoginModal();
  setStatus("Log in to load a roster workspace.");
}

function renderLoginState() {
  const loggedIn = Boolean(currentUserEmail && currentUserPassword);
  // Non-clinical Directors have no personal calendar. Their workspace has its
  // own polished header, so do not leave the technical account-status strip
  // floating over the top of it.
  document.body.classList.toggle("is-non-clinical-director", Boolean(
    loggedIn && currentNonClinical && currentDirectorViewEnabled,
  ));
  loginBar.classList.toggle("hidden", !loggedIn);
  appShell.classList.toggle("hidden", !loggedIn);
  entrancePage.classList.toggle("hidden", loggedIn);
  if (!loggedIn) mobileActionBar.classList.add("hidden");
  const me = currentAccount();
  const displayName = me.realName ? `${me.realName} · ` : "";
  const viewingText = activeDoctorProfile
    ? `Viewing as ${activeDoctorProfile.displayName} · doctor profile`
    : isImpersonating
      ? `Creator God Mode · Viewing as ${displayName}${viewedAccountEmail()}`
      : adminViewingEmail
        ? `Viewing as ${displayName}${currentUserEmail}`
        : `${displayName}${currentUserEmail}`;
  const accountTypeLabel = viewedAccountType === "creator"
    ? "Creator"
    : viewedAccountType === "unclaimed-user"
      ? "Unclaimed account"
      : "Claimed account";
  loginIdentity.textContent = loggedIn
    ? `${viewingText} · ${accountTypeLabel}${cloudAvailable ? " · Cloud sync on" : " · Cloud sync required"}`
    : "";
  backToCreatorButton.classList.toggle("hidden", !canReturnToCreator());
  syncAccountsButton();
  syncActionState();
  syncMobileChrome();
}

function isNonClinicalDirectorWorkspace() {
  return currentNonClinical && currentDirectorViewEnabled && canUseFacilityOverview();
}

function launchNonClinicalDirectorWorkspace(options = {}, loginStartedAt = 0) {
  if (!isNonClinicalDirectorWorkspace() || !calendarTransitionStillCurrent(options.transition)) return false;
  setStatus("Loading Director overview...");
  // Open the Director UI synchronously, then let its selected data view fetch
  // in the background. Non-clinical Directors do not have a personal calendar,
  // so waiting for calendar hydration here only leaves them at a blank screen.
  void openFacilityOverview().then(() => {
    if (calendarTransitionStillCurrent(options.transition)) {
      markLoginPhase("directorOverviewLoaded", loginStartedAt);
    }
  }).catch((error) => {
    if (!calendarTransitionStillCurrent(options.transition)) return;
    setStatus(normalizeAuthMessage(error?.message || "Could not load Director overview."), true);
  });
  markLoginPhase("directorOverviewOpened", loginStartedAt);
  return true;
}

async function loginWithEmail(email, password, options = {}) {
  const previousEmail = currentUserEmail;
  const loginStartedAt = performance.now();
  try {
    beginFacilityOverviewAccountSession();
    await flushCloudStateSave().catch(() => {});
    cancelScheduledCloudStateSave();
    clearActiveViewedAccountState();
    ensureLocalAccountLogin(email, password, options);
    currentUserEmail = normalizeEmail(email);
    currentUserPassword = password;
    authUserEmail = currentUserEmail;
    authUserPassword = currentUserPassword;
    adminViewingEmail = "";
    viewedAccountId = currentUserEmail;
    viewedAccountType = currentUserEmail === OWNER_EMAIL ? "creator" : "claimed-user";
    isImpersonating = false;
    impersonatedByCreator = false;
    returnToCreatorAvailable = false;
    currentDefaultDoctorKey = currentUserEmail === OWNER_EMAIL ? OWNER_DOCTOR_KEY : "";
    currentSnapshotOwnerType = currentUserEmail === OWNER_EMAIL ? "creator-account" : "user-account";
    currentSnapshotOwnerId = currentUserEmail;
    currentUserRole = currentUserEmail === OWNER_EMAIL ? "creator" : "user";
    setActiveCalendarContext(currentUserRole === "creator" ? "creator-account" : "claimed-account", { email: currentUserEmail });
    primeInsightsAccessForCurrentView();
    const transition = beginCalendarTransition();
    localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
    sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
    if (!options.adminTargetEmail) {
      localStorage.setItem(STAY_LOGGED_IN_PREFERENCE_KEY, options.stayLoggedIn ? "true" : "false");
    }
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
    switchTargetPrefetchRunId += 1;
    switchTargetPrefetchPromise = null;
    const loginCacheContext = accountCalendarContextForEmail(currentUserEmail);
    const cachedBeforeAuthentication = loadCachedCalendarSnapshotForContext(loginCacheContext);
    const loginData = await restoreCloudState({
      ...options,
      deferHydration: true,
      deferContext: true,
      deferSnapshotPersistence: true,
      skipSnapshotCacheWriteIfCurrent: true,
      responseMode: "fast",
      cachedRevision: cachedBeforeAuthentication?.calendarRevision || "",
      preserveExistingSnapshot: true,
      loginStartedAt,
      transition,
    });
    if (!currentUserEmail || !calendarTransitionStillCurrent(transition)) return;
    renderLoginState();
    closeLoginModal();
    setEntranceStatus("");
    markLoginPhase("shellRendered", loginStartedAt);
    if (isNonClinicalDirectorWorkspace()) {
      if ((loginData?.responseMode || "full") === "fast") {
        queueDeferredAccountContextLoad({
          loginStartedAt,
          targetEmail: "",
          responseMode: loginData?.responseMode || "fast",
          delayMs: 50,
          transition,
        });
      }
      launchNonClinicalDirectorWorkspace({ transition }, loginStartedAt);
      queueStoredCalendarSnapshotMaintenance();
      return;
    }
    const inlineSnapshotReady = loginSnapshotReadyForRender();
    if (inlineSnapshotReady) {
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      markLoginPhase("cachedCalendarRendered", loginStartedAt);
    }
    const renderedCachedSnapshot = inlineSnapshotReady
      ? true
      : await renderCachedCalendarSnapshotForContextAsync(loginCacheContext, {
          loginStartedAt,
          transition,
          expectedRevision: currentCalendarRevision,
        });
    if ((loginData?.responseMode || "full") === "fast") {
      queueDeferredAccountContextLoad({
        loginStartedAt,
        targetEmail: "",
        responseMode: loginData?.responseMode || "fast",
        delayMs: 50,
        transition,
      });
    }
    markLoginPhase(renderedCachedSnapshot ? "firstCalendarPaint" : "firstShellPaint", loginStartedAt);
    if (renderedCachedSnapshot) markLoginPaintCommitted(loginStartedAt);
    setStatus(renderedCachedSnapshot ? "Checking calendar for updates..." : "Loading calendar...");
    queuePostLoginHydration({
      ...options,
      includeBootstrap: true,
      forceCalendarRefresh: renderedCachedSnapshot,
      allowInlineBuild: !renderedCachedSnapshot,
      transition,
    }, loginStartedAt);
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
        responseMode: options.responseMode || "full",
        realName: options.realName || "",
        cachedRevision: options.cachedRevision || "",
        allowInlineBuild: options.allowInlineBuild !== false,
      }),
    });
    const data = await readJsonResponse(response, "Login failed.");
    if (!calendarTransitionStillCurrent(options.transition)) return null;
    await applyCloudStateData(data, {
      deferContext: options.deferContext === true,
      deferSnapshotPersistence: options.deferSnapshotPersistence === true,
      skipSnapshotCacheWriteIfCurrent: options.skipSnapshotCacheWriteIfCurrent === true,
      preserveExistingSnapshot: options.preserveExistingSnapshot === true,
      transition: options.transition,
    });
    if (!calendarTransitionStillCurrent(options.transition)) return null;
    recordLoginServerTimings(data?.diagnostics?.login, options.loginStartedAt);
    markLoginPhase("authenticated", options.loginStartedAt);
    markAccountSwitchPhase("adminLoadUser", options.accountSwitchStartedAt);
    if (!options.deferHydration) await hydrateAuthenticatedWorkspace({ ...options, includeBootstrap: false }, options.loginStartedAt);
    return data;
  } catch (error) {
    if (!calendarTransitionStillCurrent(options.transition)) return null;
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
    currentFacilityOverviewEnabled = false;
    currentNonClinical = false;
    currentDirectorViewEnabled = false;
    closeFacilityOverview();
    localStorage.removeItem(CURRENT_EMAIL_KEY);
    sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
    localStorage.removeItem(PERSISTENT_PASSWORD_KEY);
    if (options.mode === "create") {
      accountState.users = accountState.users.filter((user) => user.email !== currentUserEmail);
      saveAccountState();
    }
    currentUserEmail = "";
    currentUserPassword = "";
    viewedAccountId = "";
    viewedAccountType = "claimed-user";
    isImpersonating = false;
    impersonatedByCreator = false;
    returnToCreatorAvailable = false;
    renderLoginState();
    openLoginModal(attemptedEmail);
    setStatus(message, true);
    setEntranceStatus(message, true);
  }
}

async function hydrateAuthenticatedWorkspace(options = {}, loginStartedAt = 0) {
  if (!currentUserEmail) return;
  if (!calendarTransitionStillCurrent(options.transition)) return;
  try {
    const adminTargetEmail = normalizeEmail(options.adminTargetEmail);
    if (adminTargetEmail && adminTargetEmail !== OWNER_EMAIL && !currentRosterClaims.length) {
      accountClaimResolutionTransition = options.transition;
      try {
        if (!currentNonClinical) await resolveCurrentAccountClaims(adminTargetEmail);
      } finally {
        accountClaimResolutionTransition = null;
      }
      if (!calendarTransitionStillCurrent(options.transition)) return;
      markLoginPhase("claimsResolved", loginStartedAt);
      markAccountSwitchPhase("claimsResolved", options.accountSwitchStartedAt);
    } else if (!adminTargetEmail && currentUserEmail !== OWNER_EMAIL && !currentRosterClaims.length) {
      accountClaimResolutionTransition = options.transition;
      try {
        if (!currentNonClinical) await resolveCurrentAccountClaims();
      } finally {
        accountClaimResolutionTransition = null;
      }
      if (!calendarTransitionStillCurrent(options.transition)) return;
      markLoginPhase("claimsResolved", loginStartedAt);
    }
    if (!adminTargetEmail && currentUserEmail === OWNER_EMAIL) {
      forceCreatorDoctorSession();
    }
    const inlineSnapshotReady = visibleSnapshotIsCurrent();
    const cachedRevision = inlineSnapshotReady
      ? ""
      : options.cachedRevision || (currentSnapshot?.cacheKey === currentCalendarSnapshotCacheKey()
        ? currentSnapshot.calendarRevision || currentCalendarRevision || ""
        : "");
    const shouldRefreshCalendar = options.forceCalendarRefresh === true || !inlineSnapshotReady;
    const loadedFreshCalendar = shouldRefreshCalendar
      ? await loadCloudCalendarEvents({
          adminTargetEmail,
          doctorKey: options.doctorKey || "",
          cachedRevision,
          allowInlineBuild: options.allowInlineBuild !== false,
          preserveExistingSnapshot: true,
          transition: options.transition,
        })
      : true;
    if (!calendarTransitionStillCurrent(options.transition)) return;
    markLoginPhase("calendarLoaded", loginStartedAt);
    markAccountSwitchPhase("calendarLoaded", options.accountSwitchStartedAt);
    if (options.includeBootstrap !== false) {
      if (inlineSnapshotReady && currentSnapshot?.preview) {
        queueDeferredBootstrapImports({
          expectedKey: activeBootstrapContextKey(),
          loginStartedAt,
          accountSwitchStartedAt: options.accountSwitchStartedAt,
          allowInlineBuild: options.allowInlineBuild !== false,
          transition: options.transition,
        });
        markLoginPhase("workspaceRendered", loginStartedAt);
        markAccountSwitchPhase("workspaceRendered", options.accountSwitchStartedAt);
      } else {
        await bootstrapImports({
          allowInlineBuild: options.allowInlineBuild !== false,
          transition: options.transition,
        });
        if (!calendarTransitionStillCurrent(options.transition)) return;
        if (!lastLoginTimings?.firstCalendarPaint) {
          markLoginPhase("firstCalendarPaint", loginStartedAt);
          markLoginPaintCommitted(loginStartedAt);
        }
        markLoginPhase("workspaceRendered", loginStartedAt);
        markAccountSwitchPhase("workspaceRendered", options.accountSwitchStartedAt);
      }
    }
    if (latestNameMatches.length) {
      const sites = [...new Set(latestNameMatches.map((claim) => claim.sourceType.toUpperCase()))].join(", ");
      setStatus(`Suggested roster name${latestNameMatches.length === 1 ? "" : "s"} for ${sites || "uploaded rosters"}. Please confirm in Account.`);
    }
    if (isCreatorAuthenticated()) {
      window.setTimeout(() => {
        if (!calendarTransitionStillCurrent(options.transition)) return;
        void loadServerUsers();
      }, 0);
      void refreshCalendarStoreStatus({ silent: true, syncSwitcher: false }).then(async () => {
        if (!calendarTransitionStillCurrent(options.transition)) return;
        try {
          await syncCreatorDoctorPickerWithRemainingRosters({ snapshotDoctors: currentSnapshot?.doctorOptions || [] });
        } catch {
          // Keep the last merged doctor list.
        }
        if (latestPreview) renderDoctorState();
        queueCreatorSwitchTargetPrefetch();
      }).catch(() => null);
    }
    if (currentSnapshotStale) {
      queuePostLoginSnapshotRefresh({
        loginStartedAt,
        adminTargetEmail,
        transition: options.transition,
      });
    }
    // A non-clinical Director has no personal calendar to land on. Clinical
    // users, including Directors who still work shifts, retain their calendar
    // until they explicitly choose the Director overview.
    if (currentNonClinical && currentDirectorViewEnabled && canUseFacilityOverview() && !isFacilityOverviewOpen()) {
      await openFacilityOverview();
    }
  } catch (error) {
    if (!calendarTransitionStillCurrent(options.transition)) return;
    const message = normalizeAuthMessage(error.message || "Workspace hydration failed.");
    setStatus(message, true);
    console.warn("Post-login workspace hydration failed", { message, email: currentUserEmail, error, timings: lastLoginTimings });
  }
}

function queuePostLoginHydration(options = {}, loginStartedAt = 0) {
  const expectedEmail = viewedAccountEmail();
  const run = () => {
    if (!calendarTransitionStillCurrent(options.transition)) return;
    void hydrateAuthenticatedWorkspace(options, loginStartedAt).then(() => {
      if (viewedAccountEmail() === expectedEmail && calendarTransitionStillCurrent(options.transition)) markLoginPhase("backgroundHydrationComplete", loginStartedAt);
    });
  };
  const afterFrame = () => window.setTimeout(run, 0);
  if (typeof requestAnimationFrame === "function") {
    requestAnimationFrame(afterFrame);
  } else {
    afterFrame();
  }
}

function queuePostLoginSnapshotRefresh(options = {}) {
  const expectedKey = activeCalendarTransitionKey();
  void (async () => {
    for (const delayMs of [750, 1500, 3000, 6000]) {
      await new Promise((resolve) => window.setTimeout(resolve, delayMs));
      if (!calendarTransitionStillCurrent(options.transition) || activeCalendarTransitionKey() !== expectedKey) return;
      const loaded = await loadCloudCalendarEvents({
        adminTargetEmail: options.adminTargetEmail || "",
        cachedRevision: "",
        allowInlineBuild: false,
        preserveExistingSnapshot: true,
        transition: options.transition,
      }).catch(() => false);
      if (!calendarTransitionStillCurrent(options.transition) || activeCalendarTransitionKey() !== expectedKey) return;
      if (loaded && currentSnapshot?.preview && !currentSnapshotStale) {
        renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {}, { preserveScroll: true });
        markLoginPhase("backgroundCalendarUpdated", options.loginStartedAt);
        setStatus("Calendar refreshed.");
        return;
      }
    }
  })();
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

function recordLoginServerTimings(timings, loginStartedAt = 0) {
  if (!timings || typeof timings !== "object") return;
  if (!lastLoginTimings || (loginStartedAt && lastLoginTimings.startedAt !== loginStartedAt)) {
    lastLoginTimings = { startedAt: loginStartedAt || performance.now() };
  }
  lastLoginTimings.server = { ...timings };
  window.__rosterLoginTimings = { ...lastLoginTimings };
  console.info("Login server timings", timings);
}

function markLoginPaintCommitted(loginStartedAt = 0) {
  if (!loginStartedAt || typeof requestAnimationFrame !== "function") return;
  requestAnimationFrame(() => {
    requestAnimationFrame(() => markLoginPhase("firstCalendarPaintCommitted", loginStartedAt));
  });
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

function activeBootstrapContextKey() {
  if (activeCalendarMode() === "doctor-profile") {
    return `doctor-profile:${activeDoctorProfile?.id || ""}:${normalizeRosterName(activeDoctorProfile?.doctorKey || "")}`;
  }
  return `${activeCalendarMode()}:${viewedAccountEmail()}:${normalizeRosterName(currentDefaultDoctorKey || "")}`;
}

function activeCalendarTransitionKey() {
  if (activeCalendarMode() === "doctor-profile") {
    return `doctor-profile:${activeDoctorProfile?.id || activeDoctorProfile?.ownerId || ""}`;
  }
  return `${activeCalendarMode()}:${viewedAccountEmail()}`;
}

function beginCalendarTransition(expectedKey = activeCalendarTransitionKey()) {
  calendarTransitionRunId += 1;
  cancelDeferredBootstrapImports();
  cancelDeferredAccountContextLoad();
  cancelCalendarImportPoll();
  return {
    runId: calendarTransitionRunId,
    expectedKey: expectedKey || activeCalendarTransitionKey(),
  };
}

function calendarTransitionStillCurrent(transition = null) {
  if (!transition) return true;
  return transition.runId === calendarTransitionRunId
    && activeCalendarTransitionKey() === transition.expectedKey;
}

function cancelDeferredBootstrapImports() {
  clearTimeout(deferredBootstrapTimer);
  deferredBootstrapTimer = 0;
  deferredBootstrapRunId += 1;
}

function cancelDeferredAccountContextLoad() {
  deferredAccountContextRunId += 1;
}

function cancelCalendarImportPoll() {
  calendarImportPollRunId += 1;
}

function queueDeferredBootstrapImports(options = {}) {
  cancelDeferredBootstrapImports();
  const runId = deferredBootstrapRunId;
  const expectedKey = options.expectedKey || activeBootstrapContextKey();
  const delayMs = Math.max(0, Number(options.delayMs || 0));
  deferredBootstrapTimer = window.setTimeout(() => {
    deferredBootstrapTimer = 0;
    void (async () => {
      if (runId !== deferredBootstrapRunId) return;
      if (activeBootstrapContextKey() !== expectedKey) return;
      if (!calendarTransitionStillCurrent(options.transition)) return;
      try {
        await bootstrapImports({
          allowInlineBuild: options.allowInlineBuild !== false,
          transition: options.transition,
        });
        if (runId !== deferredBootstrapRunId || activeBootstrapContextKey() !== expectedKey || !calendarTransitionStillCurrent(options.transition)) return;
        markLoginPhase("workspaceRendered", options.loginStartedAt);
        markAccountSwitchPhase("workspaceRendered", options.accountSwitchStartedAt);
      } catch (error) {
        if (runId !== deferredBootstrapRunId || activeBootstrapContextKey() !== expectedKey || !calendarTransitionStillCurrent(options.transition)) return;
        setStatus(error.message || "Could not finish loading the calendar workspace.", true);
      }
    })();
  }, delayMs);
}

function calendarSnapshotMatchesActiveContext(snapshot = currentSnapshot) {
  if (!snapshot?.preview) return false;
  if (activeCalendarMode() === "doctor-profile") {
    if (!activeDoctorProfile?.ownerId) return false;
    const expectedKey = calendarSnapshotCacheKeyForContext({
      mode: "doctor-profile",
      ownerId: activeDoctorProfile.ownerId,
      doctorKey: activeDoctorProfile.doctorKey,
    });
    if (!expectedKey) return false;
    const snapshotDoctorKey = normalizeRosterName(
      snapshot?.session?.doctorKey
      || snapshot?.doctorOptions?.[0]?.key
      || "",
    );
    return snapshotDoctorKey === normalizeRosterName(activeDoctorProfile.doctorKey);
  }
  if (activeCalendarMode() === "creator-account") {
    return normalizeRosterName(snapshot?.session?.doctorKey || "") === OWNER_DOCTOR_KEY;
  }
  if (activeCalendarMode() === "claimed-account") {
    const snapshotDoctorKey = normalizeRosterName(snapshot?.session?.doctorKey || "");
    return !currentDefaultDoctorKey || snapshotDoctorKey === normalizeRosterName(currentDefaultDoctorKey);
  }
  return true;
}

function loginSnapshotReadyForRender() {
  if (!currentSnapshot?.preview) return false;
  if (String(currentSnapshot.calendarRevision || currentCalendarRevision || "") !== String(currentCalendarRevision || "")) return false;
  return calendarSnapshotMatchesActiveContext(currentSnapshot);
}

function visibleSnapshotIsCurrent(options = {}) {
  if (!currentSnapshot?.preview) return false;
  if (options.requireNotStale && currentSnapshotStale) return false;
  const revision = String(currentSnapshot.calendarRevision || currentCalendarRevision || "");
  if (!revision || revision !== String(currentCalendarRevision || "")) return false;
  return loginSnapshotReadyForRender();
}

function reportBackgroundValidationError(error, options = {}) {
  const message = normalizeAuthMessage(error?.message || options.fallbackMessage || "Could not update the calendar.");
  const overload = /502|503|CPU|memory|overload|Bad Gateway/i.test(String(error?.message || ""));
  if (options.preserveRenderedSnapshot && visibleSnapshotIsCurrent() && overload) {
    console.warn("Background calendar validation failed; keeping cached calendar.", { message, error });
    return;
  }
  if (options.preserveRenderedSnapshot && overload && !calendarSnapshotMatchesActiveContext(currentSnapshot)) {
    console.warn("Background calendar validation failed; calendar did not match active profile.", { message, error });
  }
  setStatus(message, true);
}

async function loadDeferredAccountContext(options = {}) {
  const requestEmail = normalizeEmail(options.targetEmail) ? authUserEmail : currentUserEmail;
  const requestPassword = normalizeEmail(options.targetEmail) ? authUserPassword : currentUserPassword;
  if (!requestEmail || !requestPassword || !cloudAvailable) return null;
  const response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      action: "loadAccountContext",
      email: requestEmail,
      password: requestPassword,
      targetEmail: normalizeEmail(options.targetEmail || ""),
    }),
  });
  return await readJsonResponse(response, "Could not load account context.");
}

function queueDeferredAccountContextLoad(options = {}) {
  const runId = ++deferredAccountContextRunId;
  const delayMs = Math.max(0, Number(options.delayMs || 0));
  window.setTimeout(() => {
    void (async () => {
      try {
        if (!calendarTransitionStillCurrent(options.transition)) return;
        const data = await loadDeferredAccountContext(options);
        if (runId !== deferredAccountContextRunId || !data || viewedAccountEmail() !== normalizeEmail(options.targetEmail || currentUserEmail) || !calendarTransitionStillCurrent(options.transition)) return;
        applyCloudStateContext(data);
        renderLoginState();
      } catch (error) {
        if (runId !== deferredAccountContextRunId || !calendarTransitionStillCurrent(options.transition)) return;
        console.warn("Deferred account context load failed", {
          email: normalizeEmail(options.targetEmail || currentUserEmail),
          message: error?.message || String(error),
        });
      }
    })();
  }, delayMs);
}

async function resolveCurrentAccountClaims(targetEmailOverride = "", options = {}) {
  const transition = options.transition || accountClaimResolutionTransition;
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
  if (!calendarTransitionStillCurrent(transition)) return;
  await applyCloudStateData(data, { transition });
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

function saveLocalAccountIdentity(realName = "") {
  if (!realName) return;
  const localAccount = accountState.users.find((user) => user.email === currentUserEmail);
  if (localAccount) {
    localAccount.realName = realName;
  } else {
    accountState.users.push({
      email: currentUserEmail,
      realName,
      password: "",
      role: currentUserEmail === OWNER_EMAIL ? "owner" : "user",
    });
  }
  saveAccountState();
}

function applyCloudStateIdentity(data) {
  cloudAvailable = data.cloudAvailable === true;
  currentCalendarRevision = String(data.snapshotRevision || data.calendarRevision || currentCalendarRevision || "");
  currentUserRole = data.role || currentUserRole;
  // The fast login envelope already includes these two display permissions.
  // Apply them before the shell is first painted so non-clinical Directors do
  // not briefly see the ordinary calendar controls while their overview loads.
  currentNonClinical = data.nonClinical === true;
  currentDirectorViewEnabled = currentUserRole === "creator" || data.directorViewEnabled === true;
  currentDefaultDoctorKey = normalizeRosterName(data.defaultDoctorKey || currentDefaultDoctorKey || "");
  currentSnapshotOwnerType = String(data.snapshotOwnerType || currentSnapshotOwnerType || "");
  currentSnapshotOwnerId = String(data.snapshotOwnerId || currentSnapshotOwnerId || "").trim();
  viewedAccountId = normalizeEmail(data.viewedAccountId || viewedAccountEmail() || currentUserEmail);
  viewedAccountType = String(data.viewedAccountType || (currentUserRole === "creator" ? "creator" : "claimed-user"));
  isImpersonating = data.isImpersonating === true;
  impersonatedByCreator = data.impersonatedByCreator === true;
  returnToCreatorAvailable = data.returnToCreatorAvailable === true || Boolean(isImpersonating);
  adminViewingEmail = isImpersonating ? viewedAccountId : "";
  currentRosterClaims = sanitizeRosterClaims(data.claims || []);
  if (typeof data.insightsEnabled === "boolean") {
    primeInsightsAccessForCurrentView({ insightsEnabled: data.insightsEnabled });
  }
  if (typeof data.facilityOverviewEnabled === "boolean") {
    currentFacilityOverviewEnabled = currentUserRole === "creator" || data.facilityOverviewEnabled === true;
  }
  syncFacilityOverviewAccess();
  if (data.realName) saveLocalAccountIdentity(data.realName);
}

function applyAvailableRosterDoctorsFromData(data) {
  const incomingAvailableDoctors = sanitizeAvailableRosterDoctors(data.availableDoctors || []);
  if (!incomingAvailableDoctors.length && isCreatorAuthenticated()) return false;
  availableRosterDoctors = incomingAvailableDoctors;
  return incomingAvailableDoctors.length > 0;
}

function applyCloudStateContext(data) {
  const previousInsightsEnabled = currentInsightsEnabled;
  currentInsightsEnabled = currentUserRole === "creator" || data.insightsEnabled === true;
  currentFacilityOverviewEnabled = currentUserRole === "creator" || data.facilityOverviewEnabled === true;
  currentNonClinical = data.nonClinical === true;
  currentDirectorViewEnabled = currentUserRole === "creator" || data.directorViewEnabled === true;
  currentSuggestedClaims = sanitizeRosterClaims(data.suggestedClaims || data.nameMatches || []);
  latestNameMatches = currentSuggestedClaims;
  applyAvailableRosterDoctorsFromData(data);
  currentSubscription = sanitizeSubscription(data.subscription);
  applyIssueConfig(data.issueConfig);
  if (previousInsightsEnabled !== currentInsightsEnabled && latestPreview) rebuildClientPreview();
  syncFacilityOverviewAccess();
}

async function applyCloudStateSnapshot(data, options = {}) {
  if (!cloudAvailable) return;
  if (!calendarTransitionStillCurrent(options.transition)) return;
  if (!data.state) {
    selectedFiles = [];
    currentSnapshot = null;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = "";
    restoredSessionState = null;
    await replaceStoredImports([]);
    if (!calendarTransitionStillCurrent(options.transition)) return;
    clearWorkspaceStoreEntry(currentUserEmail);
    return;
  }
  if (data.snapshotCurrent === true) {
    currentCalendarRevision = String(data.snapshotRevision || data.calendarRevision || currentCalendarRevision || "");
    if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
    currentSnapshotStale = false;
    currentSnapshotBuiltAt = String(data.snapshotBuiltAt || currentSnapshotBuiltAt || "");
    return;
  }
  if (!data.snapshot && options.preserveExistingSnapshot === true && currentSnapshot?.preview) {
    currentSnapshotStale = data.snapshotStale === true || currentSnapshotStale;
    currentSnapshotBuiltAt = String(data.snapshotBuiltAt || currentSnapshotBuiltAt || "");
    return;
  }
  currentSnapshot = sanitizeWorkspaceSnapshot(data.snapshot);
  if (currentSnapshot && currentCalendarRevision && data.snapshotStale !== true) {
    currentSnapshot.calendarRevision = currentCalendarRevision;
    currentSnapshot.cacheKey = currentCalendarSnapshotCacheKey({ ownerEmail: viewedAccountEmail(), doctorKey: currentDefaultDoctorKey || currentSnapshot?.session?.doctorKey || "" });
  }
  currentSnapshotStale = data.snapshotStale === true;
  currentSnapshotBuiltAt = String(data.snapshotBuiltAt || "");
  const stateImports = Array.isArray(data.state.imports) && data.state.imports.length ? data.state.imports : null;
  selectedFiles = importRefsToClientEntries(stateImports || currentSnapshot?.fileRefs || []);
  restoredSessionState = currentSnapshot?.session || (data.state.session && typeof data.state.session === "object" ? data.state.session : null);
  rememberCreatorCalendarSourceRefs();
  saveWorkspaceSnapshotForEmail(activeWorkspaceOwnerKey(), {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: restoredSessionState || {},
    snapshot: currentSnapshot,
  });
  if (currentSnapshot && !currentSnapshotStale) {
    const cacheContext = {
      ownerEmail: viewedAccountEmail(),
      doctorKey: currentDefaultDoctorKey || currentSnapshot?.session?.doctorKey || "",
    };
    if (options.deferSnapshotPersistence === true) {
      queueDeferredSnapshotCachePersist(currentSnapshot, cacheContext, {
        skipIfRevisionMatches: options.skipSnapshotCacheWriteIfCurrent === true,
      });
    } else {
      saveCalendarSnapshotCacheForContext(currentSnapshot, cacheContext, {
        snapshotReady: true,
      });
    }
  }
}

async function applyCloudStateData(data, options = {}) {
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  applyCloudStateIdentity(data);
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  if (options.deferContext) applyAvailableRosterDoctorsFromData(data);
  else applyCloudStateContext(data);
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  await applyCloudStateSnapshot(data, options);
  return calendarTransitionStillCurrent(options.transition);
}

async function loadCloudCalendarEvents(options = {}) {
  if (!cloudAvailable) return false;
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  const expectedKey = options.expectedKey || activeCalendarTransitionKey();
  const adminTargetEmail = normalizeEmail(options.adminTargetEmail);
  const requestEmail = adminTargetEmail ? authUserEmail : currentUserEmail;
  const requestPassword = adminTargetEmail ? authUserPassword : currentUserPassword;
  if (!requestEmail || !requestPassword) return false;
  const preferredDoctorKey = normalizeRosterName(
    options.doctorKey
    || (adminTargetEmail ? preferredDoctorKeyForAccountEmail(adminTargetEmail) : currentDefaultDoctorKey)
    || restoredSessionState?.doctorKey
    || selectedDoctor()?.key
    || (isCreatorAuthenticated() && !adminTargetEmail ? OWNER_DOCTOR_KEY : currentDefaultDoctorKey || currentRosterClaims[0]?.key)
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
      allowInlineBuild: options.allowInlineBuild !== false,
    }),
  });
  const data = await readJsonResponse(response, "Calendar load failed.");
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  if (activeCalendarTransitionKey() !== expectedKey) return false;
  currentCalendarRevision = String(data.snapshotRevision || data.calendarRevision || currentCalendarRevision || "");
  if (data.snapshotCurrent === true) {
    if (currentSnapshot && currentCalendarRevision) currentSnapshot.calendarRevision = currentCalendarRevision;
    if (currentSnapshot) {
      saveCalendarSnapshotCacheForContext(currentSnapshot, {
        ownerEmail: adminTargetEmail || viewedAccountEmail(),
        doctorKey: preferredDoctorKey || currentSnapshot.session?.doctorKey || "",
      });
    }
    return Boolean(currentSnapshot);
  }
  if (!data.snapshot && options.preserveExistingSnapshot === true) {
    currentSnapshotStale = data.snapshotStale === true || currentSnapshotStale;
    currentSnapshotBuiltAt = String(data.snapshotBuiltAt || currentSnapshotBuiltAt || "");
    return Boolean(currentSnapshot);
  }
  currentSnapshot = sanitizeWorkspaceSnapshot(clearCloudLoadedSnapshotFilters(data.snapshot));
  if (currentSnapshot && data.snapshotStale !== true) {
    currentSnapshot.calendarRevision = currentCalendarRevision;
    currentSnapshot.cacheKey = currentCalendarSnapshotCacheKey({
      ownerEmail: adminTargetEmail || viewedAccountEmail(),
      doctorKey: preferredDoctorKey || currentSnapshot.session?.doctorKey || "",
    });
  }
  currentSnapshotStale = data.snapshotStale === true;
  currentSnapshotBuiltAt = String(data.snapshotBuiltAt || "");
  if (!currentSnapshot) return false;
  applyLoadedCalendarFileRefs(currentSnapshot);
  restoredSessionState = currentSnapshot.session || clearCloudLoadedSessionFilters(restoredSessionState || {});
  saveWorkspaceSnapshotForEmail(activeWorkspaceOwnerKey(), {
    fileRefs: selectedFiles.map(importRefForWorkspace),
    session: restoredSessionState || {},
    snapshot: currentSnapshot,
  });
  if (!currentSnapshotStale) {
    saveCalendarSnapshotCacheForContext(currentSnapshot, {
      ownerEmail: adminTargetEmail || viewedAccountEmail(),
      doctorKey: preferredDoctorKey || currentSnapshot.session?.doctorKey || "",
    });
  }
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

function reportCloudSaveFailure(error, payload = null, options = {}) {
  if (options.reportErrors === false) {
    if (typeof options.onError === "function") options.onError(error);
    return;
  }
  const payloadStillMatchesView = payload ? savePayloadMatchesActiveCalendar(payload) : true;
  if (!payloadStillMatchesView) return;
  if (!error?.isRosterPersistenceError) cloudAvailable = false;
  renderLoginState();
  setStatus(error.message || "Cloud save failed.", true);
}

function scheduleCloudStateSave() {
  if (!currentUserEmail) return;
  cancelScheduledCloudStateSave();
  const snapshot = snapshotCloudSavePayload();
  pendingCloudSaveSnapshot = snapshot;
  cloudSaveTimer = setTimeout(() => {
    const queued = pendingCloudSaveSnapshot;
    pendingCloudSaveSnapshot = null;
    saveCloudState(queued || snapshot).catch((error) => reportCloudSaveFailure(error, queued || snapshot));
  }, 700);
}

function cancelScheduledCloudStateSave() {
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = 0;
  clearTimeout(backgroundCloudSaveTimer);
  backgroundCloudSaveTimer = 0;
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
  clearTimeout(backgroundCloudSaveTimer);
  const run = () => {
    backgroundCloudSaveTimer = 0;
    saveCloudState(payload).catch((error) => reportCloudSaveFailure(error, payload, options));
  };
  const delayMs = Math.max(0, Number(options.delayMs || 0));
  if (delayMs) {
    backgroundCloudSaveTimer = window.setTimeout(run, delayMs);
  } else {
    run();
  }
}

function outgoingSnapshotSavePayload(state = captureCalendarViewState()) {
  if (!state) return null;
  if (state.activeDoctorProfile?.id) {
    return {
      accountEmail: state.currentUserEmail,
      requestEmail: authUserEmail || state.currentUserEmail,
      requestPassword: authUserPassword || state.currentUserPassword,
      doctorProfile: { ...state.activeDoctorProfile },
      imports: [],
      session: state.restoredSessionState || state.currentSnapshot?.session || {},
    };
  }
  if (normalizeEmail(state.currentUserEmail) === OWNER_EMAIL && !state.adminViewingEmail) {
    return {
      accountEmail: state.currentUserEmail,
      requestEmail: authUserEmail || state.currentUserEmail,
      requestPassword: authUserPassword || state.currentUserPassword,
      targetEmail: "",
      imports: (state.selectedFiles || []).map((entry) => ({ ...entry })),
      session: state.restoredSessionState || state.currentSnapshot?.session || {},
      removedImportIds: [],
    };
  }
  const viewedEmail = normalizeEmail(state.adminViewingEmail || state.currentUserEmail);
  if (!viewedEmail) return null;
  return {
    accountEmail: viewedEmail,
    requestEmail: authUserEmail || state.currentUserEmail,
    requestPassword: authUserPassword || state.currentUserPassword,
    targetEmail: state.adminViewingEmail ? viewedEmail : "",
    imports: (state.selectedFiles || []).map((entry) => ({ ...entry })),
    session: state.restoredSessionState || state.currentSnapshot?.session || {},
    removedImportIds: [],
  };
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

function snapshotMatchesDoctorProfile(snapshot, profile) {
  if (!snapshot?.preview || !profile?.doctorKey) return false;
  const snapshotDoctorKey = normalizeRosterName(
    snapshot?.session?.doctorKey
    || snapshot?.doctorOptions?.[0]?.key
    || "",
  );
  return snapshotDoctorKey === normalizeRosterName(profile.doctorKey);
}

function creatorSwitchTargetsForPrefetch() {
  if (!isViewingCreatorAccount()) return [];
  const seen = new Set();
  const targets = [];
  for (const doctor of doctorPickerOptions()) {
    const accountEmail = currentClaimedAccountEmail(doctor.accountEmail || doctor.claimedBy || "");
    if (normalizeRosterName(doctor.key) === OWNER_DOCTOR_KEY) continue;
    const profile = accountEmail ? null : doctorProfileForDoctor(doctor);
    if (!accountEmail && !profile?.id) continue;
    const context = accountEmail
      ? accountCalendarContextForEmail(accountEmail)
      : calendarSnapshotContext({
          mode: "doctor-profile",
          ownerId: profile.ownerId,
          doctorKey: profile.doctorKey,
        });
    const key = calendarSnapshotCacheKeyForContext(context);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    targets.push(accountEmail
      ? { kind: "account", email: accountEmail, context }
      : { kind: "doctor-profile", profile, context });
  }
  for (const user of serverUsers.map(normalizeServerUser)) {
    if (!user?.email || user.email === OWNER_EMAIL) continue;
    const context = accountCalendarContextForEmail(user.email);
    const key = calendarSnapshotCacheKeyForContext(context);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    targets.push({
      kind: "account",
      email: user.email,
      context,
    });
  }
  return targets;
}

async function prefetchCreatorSwitchTarget(target) {
  if (!isViewingCreatorAccount() || !cloudAvailable) return;
  if ((await loadCachedCalendarSnapshotForContextAsync(target.context))?.preview) return;
  const requestEmail = authUserEmail || currentUserEmail;
  const requestPassword = authUserPassword || currentUserPassword;
  if (!requestEmail || !requestPassword) return;
  let response;
  if (target.kind === "account") {
    response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "loadCalendarEvents",
        email: requestEmail,
        password: requestPassword,
        targetEmail: target.email,
        doctorKey: target.context?.doctorKey || "",
        startDate: target.context?.range?.startDate || "",
        endDate: target.context?.range?.endDate || "",
        allowInlineBuild: false,
        skipRebuild: true,
      }),
    });
  } else if (target.kind === "doctor-profile") {
    response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        action: "loadDoctorProfile",
        email: requestEmail,
        password: requestPassword,
        profileId: target.profile.id,
        doctorKey: target.profile.doctorKey,
        displayName: target.profile.displayName,
        sourceTypes: target.profile.sourceTypes,
        aliases: target.profile.aliases,
        allowInlineBuild: false,
        skipRebuild: true,
      }),
    });
  } else {
    return;
  }
  const data = await readJsonResponse(response, "Could not prefetch account snapshot.");
  const snapshot = sanitizeWorkspaceSnapshot(data.snapshot);
  if (!snapshot?.preview) return;
  snapshot.calendarRevision = String(data.snapshotRevision || data.calendarRevision || snapshot.calendarRevision || "");
  saveCalendarSnapshotCacheForContext(snapshot, target.context);
}

function queueCreatorSwitchTargetPrefetch() {
  if (!isViewingCreatorAccount() || !cloudAvailable || !isCreatorAuthenticated()) return;
  if (switchTargetPrefetchPromise) return;
  const runId = ++switchTargetPrefetchRunId;
  switchTargetPrefetchPromise = (async () => {
    const targets = creatorSwitchTargetsForPrefetch();
    let nextIndex = 0;
    const workerCount = Math.min(4, targets.length);
    await Promise.all(Array.from({ length: workerCount }, async () => {
      while (nextIndex < targets.length) {
        const target = targets[nextIndex];
        nextIndex += 1;
        if (runId !== switchTargetPrefetchRunId || !isViewingCreatorAccount()) return;
        await prefetchCreatorSwitchTarget(target).catch(() => null);
      }
    }));
  })().finally(() => {
    if (runId === switchTargetPrefetchRunId) switchTargetPrefetchPromise = null;
  });
}

async function saveCloudStateNow(snapshot = null) {
  const payload = snapshot || snapshotCloudSavePayload();
  if (!payload.accountEmail || !payload.requestEmail || !payload.requestPassword || !cloudAvailable) return;
  const isDeleteSave = (payload.removedImportIds || []).length > 0;
  const snapshotPayload = isDeleteSave ? null : await buildWorkspaceSnapshotPayload(payload.session);
  const shouldApplySavedSnapshot = savePayloadMatchesActiveCalendar(payload) && !isDeleteSave
    && (!payload.doctorProfile || snapshotMatchesDoctorProfile(snapshotPayload, payload.doctorProfile));
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
  const state = (payload.removedImportIds || []).length
    ? await buildCloudStateWithoutRosterSync(payload.imports, payload.session)
    : await buildCloudState(payload.imports, payload.session);
  const saveBody = {
    action: "save",
    email: payload.requestEmail,
    password: payload.requestPassword,
    targetEmail: payload.targetEmail,
    state,
    snapshot: snapshotPayload,
    removedImportIds: payload.removedImportIds || [],
    ...(payload.repositorySynced === true ? { repositorySynced: true } : {}),
  };
  let response = await fetch("/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(saveBody),
  });
  if (response.status === 503 && isDeleteSave) {
    await new Promise((resolve) => setTimeout(resolve, 800));
    response = await fetch("/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(saveBody),
    });
  }
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
    await refreshCalendarStoreStatus({ silent: true });
    if (hasPendingRosterAutomation()) {
      throw new Error("Wait for queued roster updates to finish before using advanced recovery.");
    }
    const confirmed = window.confirm(
      "Advanced recovery force-reparses every retained roster and replaces the active roster set. Continue only if the derived roster database is corrupted.",
    );
    if (!confirmed) return;
    const confirmation = window.prompt('Type REBUILD to confirm this recovery operation.');
    if (confirmation !== "REBUILD") {
      setStatus("Roster rebuild cancelled.");
      return;
    }
    setStatus("Rebuilding roster database from roster files...");
    const sourceEntries = retainedRosterEntriesFromStatus();
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
        confirmation,
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

function hasPendingRosterAutomation() {
  return (calendarStoreStatus?.rosterSourceStatuses || [])
    .some((source) => ["queued", "processing"].includes(source?.state));
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
  const fileId = String(id || "");
  if (!fileId || activeManualReparseIds.has(fileId)) return;
  const statusEntry = (calendarStoreStatus?.files || []).find((file) => file.id === id);
  const entry = selectedFiles.find((item) => item.id === id) || (statusEntry ? {
    id: statusEntry.id,
    repoId: statusEntry.id,
    name: statusEntry.name,
    sourceType: statusEntry.sourceType,
    addedAt: "",
  } : null);
  if (!entry) return;
  activeManualReparseIds.add(fileId);
  renderFileSurfaces();
  try {
    const activeState = rosterSyncStates.get(entry.id);
    if (activeState && ["pending", "uploading-source", "parsing", "saving"].includes(activeState.status)) {
      setStatus(`Waiting for ${entry.name} to finish saving...`);
      await waitForRosterFilePersistence(entry, selectedFiles.map((item) => item.id));
      await refreshCalendarStoreStatus({ silent: true });
      await loadCloudCalendarEvents();
      if (currentSnapshot) renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      renderFileSurfaces();
      setStatus(`${entry.name} confirmed in roster database.`);
      return;
    }
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
  } finally {
    activeManualReparseIds.delete(fileId);
    renderFileSurfaces();
  }
}

async function refreshAutomatedRosterSource(sourceId, historicalRange = null) {
  const id = String(sourceId || "");
  if (!id || activeAutomatedSourceRefreshIds.has(id)) return;
  const source = (calendarStoreStatus?.rosterSourceStatuses || []).find((item) => item?.id === sourceId);
  const label = String(source?.label || "Automated roster");
  const canCheckProvider = source?.provider === "findmyshift";
  const range = /^\d{4}-\d{2}-\d{2}$/.test(String(historicalRange?.from || "")) && /^\d{4}-\d{2}-\d{2}$/.test(String(historicalRange?.to || ""))
    ? { from: historicalRange.from, to: historicalRange.to }
    : null;
  const operationLabel = range ? `${label} (${range.from} to ${range.to})` : label;
  const previousSuccessAt = String(source?.lastSuccessAt || "");
  activeAutomatedSourceRefreshIds.add(id);
  renderFileSurfaces();
  try {
    setStatus(range ? `Queueing ${operationLabel} for historical recovery...` : canCheckProvider ? `Checking ${label} for a newer roster...` : `Queueing ${label} to reprocess...`);
    const result = await calendarStoreRequest("refreshAutomatedRosterSource", { sourceId: id, ...(range ? { range } : {}) });
    await refreshCalendarStoreStatus({ silent: true });
    const status = String(result?.status || "queued");
    if (status === "reprocess-queued") {
      setStatus(`${operationLabel} reprocessing has been queued.`);
    } else if (status === "queued") {
      setStatus(`${operationLabel} has been queued for processing.`);
    } else if (status === "processing") {
      setStatus(`${label} is already being processed.`);
    } else {
      setStatus(`${operationLabel} refresh is ${status.replace(/-/g, " ")}.`);
    }
    if (["queued", "processing", "reprocess-queued"].includes(status)) {
      await waitForAutomatedRosterSourceRefresh(id, operationLabel, previousSuccessAt, result?.queue?.runIds);
    }
  } catch (error) {
    setStatus(error.message || `Could not refresh ${label}.`, true);
  } finally {
    activeAutomatedSourceRefreshIds.delete(id);
    renderFileSurfaces();
  }
}

async function waitForAutomatedRosterSourceRefresh(sourceId, label, previousSuccessAt = "", expectedRunIds = []) {
  let sawPendingState = false;
  const trackedRunIds = new Set((Array.isArray(expectedRunIds) ? expectedRunIds : []).map((id) => String(id || "")).filter(Boolean));
  for (;;) {
    const source = (calendarStoreStatus?.rosterSourceStatuses || []).find((item) => item?.id === sourceId);
    if (["queued", "processing"].includes(source?.state)) sawPendingState = true;
    const runsById = new Map((Array.isArray(source?.recentRuns) ? source.recentRuns : []).map((run) => [String(run?.id || ""), run]));
    const trackedRuns = [...trackedRunIds].map((id) => runsById.get(id)).filter(Boolean);
    const failedRun = trackedRuns.find((run) => run.status === "failed");
    if (failedRun) throw new Error(failedRun.message || `${label} update failed.`);
    if (trackedRunIds.size && trackedRuns.length === trackedRunIds.size && trackedRuns.every((run) => run.status === "success")) {
      await loadCloudCalendarEvents();
      if (currentSnapshot) renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      setStatus(`${label} imported.`);
      return;
    }
    if (trackedRunIds.size && trackedRuns.length < trackedRunIds.size) sawPendingState = true;
    if (trackedRunIds.size) {
      await new Promise((resolve) => setTimeout(resolve, 3000));
      await refreshCalendarStoreStatus({ silent: true });
      continue;
    }
    if (source?.state === "received" && (sawPendingState || String(source.lastSuccessAt || "") !== previousSuccessAt)) {
      await loadCloudCalendarEvents();
      if (currentSnapshot) renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      setStatus(`${label} imported.`);
      return;
    }
    if (source?.state === "failed") throw new Error(source.lastError || `${label} update failed.`);
    await new Promise((resolve) => setTimeout(resolve, 3000));
    await refreshCalendarStoreStatus({ silent: true });
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
    if (Array.isArray(data.availableDoctors)) {
      applyAuthoritativeAvailableDoctors(data.availableDoctors);
    }
    applyIssueConfig(data.issueConfig);
    syncAccountsButton();
    queueCreatorSwitchTargetPrefetch();
    if (isViewingCreatorAccount() && latestPreview) renderDoctorState();
  } catch {
    // Keep the last available local list.
  }
}

function adminParserSurfaceReadyForRosterIssueLoad() {
  return isCreatorAuthenticated()
    && isViewingCreatorAccount()
    && cloudAvailable
    && currentAdminTab === "parser"
    && accountsModal
    && !accountsModal.classList.contains("hidden")
    && Boolean(latestPreview || currentSnapshot);
}

function queueGlobalUnresolvedShiftCodeLoad(options = {}) {
  if (!adminParserSurfaceReadyForRosterIssueLoad()) return;
  if (globalUnresolvedShiftCodesLoading) return;
  if (globalUnresolvedShiftCodesLoaded && options.force !== true) return;
  void loadGlobalUnresolvedShiftCodes(options);
}

function markGlobalUnresolvedShiftCodesStale() {
  globalUnresolvedShiftCodesLoaded = false;
}

function refreshGlobalUnresolvedShiftCodesAfterRuleChange() {
  markGlobalUnresolvedShiftCodesStale();
  queueGlobalUnresolvedShiftCodeLoad({ force: true });
}

async function loadGlobalUnresolvedShiftCodes() {
  if (!adminParserSurfaceReadyForRosterIssueLoad()) return;
  const runId = ++globalUnresolvedShiftCodeRunId;
  globalUnresolvedShiftCodesLoading = true;
  globalUnresolvedShiftCodesError = "";
  renderAccountsModal();
  if (shiftCodeReviewModal && !shiftCodeReviewModal.classList.contains("hidden")) renderShiftCodeReviewResults();
  try {
    const data = await calendarStoreRequest("listUnresolvedShiftCodes");
    if (runId !== globalUnresolvedShiftCodeRunId) return;
    globalUnresolvedShiftCodes = sanitizeGlobalUnresolvedShiftCodes(data.unresolvedShiftCodes);
    globalUnresolvedShiftCodesLoaded = true;
  } catch (error) {
    if (runId !== globalUnresolvedShiftCodeRunId) return;
    globalUnresolvedShiftCodesError = error.message || "Could not load unresolved shift codes.";
  } finally {
    if (runId === globalUnresolvedShiftCodeRunId) {
      globalUnresolvedShiftCodesLoading = false;
      if (accountsModal && !accountsModal.classList.contains("hidden") && currentAdminTab === "parser") renderAccountsModal();
      if (shiftCodeReviewModal && !shiftCodeReviewModal.classList.contains("hidden")) renderShiftCodeReviewModal();
    }
  }
}

function sanitizeGlobalUnresolvedShiftCodes(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => {
      const source = sanitizeIssueSource(item?.source);
      const seniority = sanitizeRuleSeniority(item?.seniority);
      const rawValue = String(item?.rawValue || "").trim();
      const code = parserRuleCodeForIssue({ ...item, source, rawValue });
      const message = String(item?.message || "").trim();
      if (!source || !code || !rawValue || !message) return null;
      const sampleDate = String(item?.sampleDate || item?.date || item?.startDay || "").slice(0, 10);
      const fingerprint = issueFingerprint(source, rawValue || code, seniority);
      return {
        id: String(item?.id || `roster::${fingerprint}`).trim(),
        origin: "roster",
        source,
        seniority,
        code,
        rawValue,
        message,
        sampleName: String(item?.sampleName || "Roster").trim(),
        doctorKey: normalizeRosterName(item?.doctorKey || ""),
        displayName: String(item?.displayName || item?.sampleName || "").trim(),
        sampleDate,
        count: Math.max(1, Math.floor(Number(item?.count || 1))),
        firstSeenAt: String(item?.firstSeenAt || sampleDate || ""),
        lastSeenAt: String(item?.lastSeenAt || sampleDate || ""),
        suggestedTitle: String(item?.suggestedTitle || "").trim(),
        timeLabel: String(item?.timeLabel || "").trim(),
        examples: (Array.isArray(item?.examples) ? item.examples : [])
          .map((example) => ({
            id: String(example?.id || "").trim(),
            doctorKey: normalizeRosterName(example?.doctorKey || ""),
            displayName: String(example?.displayName || example?.doctorKey || "").trim(),
            date: String(example?.date || example?.startDay || "").slice(0, 10),
            rawValue: String(example?.rawValue || rawValue).trim(),
            timeLabel: String(example?.timeLabel || "").trim(),
            fileName: String(example?.fileName || "").trim(),
          }))
          .filter((example) => example.displayName || example.doctorKey || example.date),
      };
    })
    .filter(Boolean);
}

async function refreshCalendarStoreStatus(options = {}) {
  if (!isCreatorAuthenticated() || !cloudAvailable) return;
  const statusPayload = {
    selectedDoctorKey: rosterStatusDoctorKey(),
    expectedFileIds: selectedFiles.map((entry) => entry.id),
    ...(options.includeAvailableDoctors ? { includeAvailableDoctors: true } : {}),
    ...(options.lightweight !== false ? { lightweight: true } : {}),
  };
  const requestStatus = options.useRetry === false
    ? () => calendarStoreRequest("calendarStoreStatus", statusPayload)
    : () => calendarStoreRequestWithRetry("calendarStoreStatus", statusPayload, { attempts: options.attempts || 4 });
  try {
    const data = await requestStatus();
    calendarStoreStatus = { ...data, checkedAt: new Date().toISOString() };
    if (options.includeAvailableDoctors === true && Array.isArray(data.availableDoctors)) {
      const incomingDoctors = sanitizeAvailableRosterDoctors(data.availableDoctors);
      const shouldMergeDoctors = options.mergeAvailableDoctors === true;
      availableRosterDoctors = shouldMergeDoctors
        ? mergeAvailableRosterDoctors(availableRosterDoctors, incomingDoctors)
        : incomingDoctors;
    }
    if (isCreatorAuthenticated() && Array.isArray(data.files) && data.files.length) {
      restoreCreatorImportFilesIfNeeded();
      mergeSelectedFilesWithRosterStoreStatus(calendarStoreStatus, { force: true });
    }
    calendarStoreStatusError = "";
    reconcileRosterSyncStates(calendarStoreStatus);
    if (!options.silent) setStatus("Roster database status checked.");
    if (options.syncSwitcher === true && isCreatorAuthenticated()) {
      try {
        await syncCreatorDoctorPickerWithRemainingRosters({ snapshotDoctors: doctorOptions });
      } catch {
        // Keep the last merged doctor list.
      }
      renderDoctorState();
    }
  } catch (error) {
    if (!options.silent) {
      calendarStoreStatusError = error.message || "Could not check roster database status.";
      setStatus(calendarStoreStatusError, true);
    }
  }
  renderFileSurfaces();
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

function appendLiveAdminConsoleMessage(message, isError = false) {
  const text = String(message || "").trim();
  if (!text) return;
  adminConsoleMessages = [{
    message: text,
    isError: isError === true,
    createdAt: new Date().toISOString(),
  }, ...adminConsoleMessages].slice(0, 50);
}

function refreshAdminConsoleMarkupIfVisible() {
  if (!adminConsoleOpen || currentAdminTab !== "system" || adminConsoleLoading) return;
  const consoleContainer = accountsBody?.querySelector(".console-history");
  if (!consoleContainer) return;
  consoleContainer.outerHTML = renderAdminConsoleMarkup();
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
  const providerDoctors = sourceType === "ddh" ? findmyshiftProviderStaffOptions(sourceEntries.ddh) : [];
  const membershipDoctors = mergeMembershipDoctors(doctors, providerDoctors);
  const uniqueDoctors = [];
  const seenDoctors = new Set();
  for (const doctor of membershipDoctors) {
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
    doctors: applyRosterEventSeniorities(uniqueDoctors, eventsByDoctor),
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

async function calendarStoreRequestWithRetry(action, extra = {}, options = {}) {
  const attempts = Math.max(1, Number(options.attempts || 1));
  let lastError = null;
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    try {
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
      if ((response.status === 503 || response.status === 429) && attempt + 1 < attempts) {
        await new Promise((resolve) => setTimeout(resolve, 1200 * (attempt + 1)));
        continue;
      }
      return await readJsonResponse(response, "SQL calendar store request failed.");
    } catch (error) {
      lastError = error;
      const retryable = /503|429|overload|timed out|network/i.test(String(error?.message || ""));
      if (!retryable || attempt + 1 >= attempts) throw error;
      await new Promise((resolve) => setTimeout(resolve, 1200 * (attempt + 1)));
    }
  }
  throw lastError || new Error("SQL calendar store request failed.");
}

async function waitForRosterFilePersistence(entry, expectedFileIds = [], options = {}) {
  const timeoutMs = Number(options.timeoutMs || 180000);
  const pollMs = Number(options.pollMs || 2500);
  const started = Date.now();
  let pollCount = 0;
  while (Date.now() - started < timeoutMs) {
    pollCount += 1;
    let latestStatus = null;
    try {
      latestStatus = await calendarStoreRequestWithRetry("calendarStoreStatus", {
        selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
        expectedFileIds,
      }, { attempts: 4 });
    } catch (error) {
      await new Promise((resolve) => setTimeout(resolve, pollMs));
      continue;
    }
    calendarStoreStatus = { ...latestStatus, checkedAt: new Date().toISOString() };
    reconcileRosterSyncStates(calendarStoreStatus);
    renderFileSurfaces();
    const statusFile = (latestStatus.files || []).find((file) => file.id === entry.id);
    if (isRosterFileStatusHealthy(statusFile)) return latestStatus;
    await new Promise((resolve) => setTimeout(resolve, pollMs));
  }
  throw new Error(`${entry.name || "Roster file"} was not confirmed in D1 before the save timed out.`);
}

async function buildCloudState(imports = selectedFiles, session = buildActiveSessionState()) {
  await saveSelectedRosterFilesToD1(imports);
  return buildCloudStateWithoutRosterSync(imports, session);
}

async function buildCloudStateWithoutRosterSync(imports = selectedFiles, session = buildActiveSessionState()) {
  const subscriptionFeeds = await buildSubscriptionFeeds(session);
  return {
    version: 1,
    imports: (imports || []).map(importRefForWorkspace),
    session,
    subscriptionFeeds,
  };
}

const LARGE_ROSTER_EVENT_THRESHOLD = 1200;
const LARGE_ROSTER_DOCTOR_CHUNK = 18;

function rosterSaveUsesChunkedUpload(payload = {}) {
  return Number(payload?.eventCount || 0) >= LARGE_ROSTER_EVENT_THRESHOLD
    || (payload?.doctors || []).length >= 80;
}

function slimDerivedCalendarRequest(payload = {}, extra = {}) {
  const { eventsByDoctor: _eventsByDoctor, issuesByDoctor: _issuesByDoctor, ...rest } = payload;
  return { ...rest, ...extra };
}

async function saveDerivedCalendarFilePayload(payload, entry, expectedFileIds) {
  if (!rosterSaveUsesChunkedUpload(payload)) {
    return calendarStoreRequestWithRetry("saveDerivedCalendarFile", slimDerivedCalendarRequest(payload, {
      expectedFileIds,
      skipStatus: true,
    }), { attempts: 4 });
  }
  await calendarStoreRequestWithRetry("saveDerivedCalendarFile", slimDerivedCalendarRequest(payload, {
    phase: "start",
    eventsByDoctor: {},
    issuesByDoctor: {},
    expectedFileIds,
    skipStatus: true,
  }), { attempts: 4 });
  const doctorKeys = (payload.doctors || []).map((doctor) => doctor.key).filter(Boolean);
  for (let index = 0; index < doctorKeys.length; index += LARGE_ROSTER_DOCTOR_CHUNK) {
    const keys = doctorKeys.slice(index, index + LARGE_ROSTER_DOCTOR_CHUNK);
    const chunkDoctors = (payload.doctors || []).filter((doctor) => keys.includes(doctor.key));
    const eventsByDoctor = Object.fromEntries(keys.map((key) => [key, payload.eventsByDoctor?.[key] || []]));
    const issuesByDoctor = Object.fromEntries(keys.map((key) => [key, payload.issuesByDoctor?.[key] || []]));
    const chunkEvents = keys.reduce((total, key) => total + (payload.eventsByDoctor?.[key]?.length || 0), 0);
    setRosterSyncState(entry, "saving", `Saving ${Math.min(index + keys.length, doctorKeys.length)}/${doctorKeys.length} doctors…`);
    await calendarStoreRequestWithRetry("saveDerivedCalendarFile", slimDerivedCalendarRequest(payload, {
      phase: "events",
      doctors: chunkDoctors,
      eventsByDoctor,
      issuesByDoctor,
      expectedFileIds,
      skipStatus: true,
    }), { attempts: 4 });
  }
  return calendarStoreRequestWithRetry("saveDerivedCalendarFile", slimDerivedCalendarRequest(payload, {
    phase: "finish",
    eventsByDoctor: {},
    issuesByDoctor: {},
    expectedFileIds,
    skipStatus: true,
  }), { attempts: 4 });
}

async function saveSelectedRosterFilesToD1(imports = selectedFiles, options = {}) {
  if (!cloudAvailable || !currentUserEmail) return emptyRosterPersistenceSummary();
  const entries = (imports || []).filter((entry) => entry?.file);
  if (!entries.length) return emptyRosterPersistenceSummary();
  const allExpectedFileIds = [...new Set(selectedFiles.map((entry) => entry.id).filter(Boolean))];
  const expectedFileIds = allExpectedFileIds.length ? allExpectedFileIds : entries.map((entry) => entry.id);
  const failedIds = new Set([...rosterSyncStates.entries()].filter(([, state]) => state.status === "failed").map(([id]) => id));
  const entriesToSave = options.force === true
    ? entries
    : entries.filter((entry) => !isLocalRosterFileSyncedToD1(entry) || failedIds.has(entry.id));
  const saveResults = [];
  let latestStatus = calendarStoreStatus;
  if (!entriesToSave.length) return summarizeRosterPersistence(entries, latestStatus, saveResults);
  beginRosterSync(entriesToSave, options.force === true ? "rebuild" : "sync");
  for (const entry of entriesToSave) {
    let failStep = "init";
    try {
      setRosterSyncState(entry, "parsing");
      failStep = "parse";
      const payload = await buildDerivedCalendarFilePayload(entry, entry);
      if (options.retainSources !== false) {
        failStep = "retain-source";
        setRosterSyncState(entry, "uploading-source");
        try {
          await retainRosterSource(entry);
        } catch (error) {
        }
      }
      failStep = "d1-save";
      setRosterSyncState(entry, "saving");
      let saveResponse = await saveDerivedCalendarFilePayload(payload, entry, expectedFileIds);
      if (saveResponse.indexing === "scheduled" || saveResponse.indexing === "in-progress" || !isRosterFileStatusHealthy(saveResponse?.fileStatus)) {
        setRosterSyncState(entry, "saving", "Saving to roster database…");
        latestStatus = await waitForRosterFilePersistence(entry, expectedFileIds);
      } else {
        latestStatus = mergeLightweightRosterStatus(saveResponse, payload.file, saveResponse.fileStatus, expectedFileIds);
      }
      saveResults.push({ entry, ok: true });
      entry.sourceType = payload.file?.sourceType || entry.sourceType;
      entry.needsD1Resync = false;
      setRosterSyncState(entry, "synced");
    } catch (error) {
      const errorMessage = error?.message || String(error);
      saveResults.push({ entry, ok: false, error, failStep, message: `${failStep}: ${errorMessage}` });
      setRosterSyncState(entry, "failed", `${failStep}: ${errorMessage}`);
    }
  }
  if (isCreatorAuthenticated()) {
    try {
      latestStatus = await calendarStoreRequest("calendarStoreStatus", {
        selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
        expectedFileIds,
        includeAvailableDoctors: true,
      });
      if (Array.isArray(latestStatus.availableDoctors)) {
        applyAuthoritativeAvailableDoctors(latestStatus.availableDoctors);
      }
      calendarStoreStatusError = "";
    } catch (error) {
      calendarStoreStatusError = error.message || "Could not check roster database status.";
      // Keep the most recent save response if the creator status refresh fails.
    }
  }
  if (latestStatus) calendarStoreStatus = { ...latestStatus, checkedAt: new Date().toISOString() };
  reconcileRosterSyncStates(calendarStoreStatus);
  const summary = summarizeRosterPersistence(entries, calendarStoreStatus, saveResults);
  lastRosterPersistence = summary;
  renderFileSurfaces();
  if (!summary.complete) {
    finishRosterSync();
    const error = new Error(rosterPersistenceFailureMessage(summary));
    error.isRosterPersistenceError = true;
    throw error;
  }
  invalidateCalendarSnapshotCachesForChangedRosterFiles(entriesToSave);
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

function isRosterFileStatusHealthy(file) {
  if (!file?.id) return false;
  if (file.retainedSourceOnly === true) return false;
  return file.status === "populated" || Number(file.eventCount || 0) > 0;
}

function isLocalRosterFileSyncedToD1(entry, status = calendarStoreStatus) {
  if (!entry?.id) return false;
  const statusFile = (status?.files || []).find((file) => file.id === entry.id);
  return isRosterFileStatusHealthy(statusFile);
}

function reconcileRosterSyncStates(status = calendarStoreStatus) {
  if (!status || !Array.isArray(status.files)) return false;
  const filesById = new Map(status.files.map((file) => [file.id, file]));
  const knownIds = new Set([
    ...selectedFiles.map((entry) => entry.id),
    ...status.files.map((file) => file.id),
  ]);
  const activeStatuses = new Set(["pending", "uploading-source", "parsing", "saving"]);
  let changed = false;
  for (const [id, state] of rosterSyncStates.entries()) {
    if (activeStatuses.has(state.status)) continue;
    const file = filesById.get(id);
    if (state.status === "failed" && isRosterFileStatusHealthy(file)) {
      rosterSyncStates.delete(id);
      changed = true;
      continue;
    }
    if ((state.status === "failed" || state.status === "synced") && !knownIds.has(id)) {
      rosterSyncStates.delete(id);
      changed = true;
    }
  }
  if (changed) renderFileSurfaces();
  return changed;
}

function rosterSyncLabel(entry) {
  const state = rosterSyncStates.get(entry.id);
  if (!state || state.status === "synced") return "";
  if (state.status === "failed") {
    const statusFile = (calendarStoreStatus?.files || []).find((file) => file.id === entry.id);
    if (isRosterFileStatusHealthy(statusFile)) return "";
  }
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
    (imports || []).filter((entry) => entry?.id),
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
    : (status?.files || [])
      .filter((file) => expectedFileIdSet.has(file.id) && isRosterFileStatusHealthy(file))
      .map((file) => file.id);
  const persistedSet = new Set(persistedFileIds);
  const activeFileIds = statusMatchesEntries
    ? (statusExpected.activeFileIds || []).filter((id) => expectedFileIdSet.has(id))
    : (status?.files || []).filter((file) => file.active !== false).map((file) => file.id).filter((id) => expectedFileIdSet.has(id));
  const failedEntries = saveResults
    .filter((result) => result && result.ok === false && result.entry?.id)
    .map((result) => ({
      id: result.entry.id,
      name: result.entry.name,
      message: result.message || result.error?.message || "D1 save failed.",
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
    statusMatchesEntries,
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
  const failedMessages = [...new Set(summary.failedEntries.map((entry) => entry.message).filter(Boolean))];
  const missingDetail = missingNames.length ? ` Missing from D1: ${[...new Set(missingNames)].join(", ")}.` : "";
  const failedDetail = failedNames.length
    ? ` Failed to save this upload: ${[...new Set(failedNames)].join(", ")}.${failedMessages.length ? ` Reason: ${failedMessages[0]}` : ""}`
    : "";
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
  const statusFile = (calendarStoreStatus?.files || []).find((file) => file.id === entry.id);
  if (statusFile?.rawSourceAvailable === true) {
    return;
  }
  await calendarStoreRequestWithRetry("uploadRawRosterFile", {
    file: importRefForWorkspace(entry),
    type: entry.file.type || "application/octet-stream",
    dataUrl: await fileToDataUrl(entry.file),
  }, { attempts: 3 });
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

function removeLegacyCalendarSnapshotCaches() {
  for (const key of LEGACY_CALENDAR_SNAPSHOT_CACHE_KEYS) {
    try {
      localStorage.removeItem(key);
    } catch {
      // Legacy cache cleanup must never delay calendar rendering.
    }
  }
}

function compactCalendarSnapshotCacheStore(store, limit = MAX_HOT_SNAPSHOT_CACHE_ENTRIES) {
  const entries = Object.entries(store || {})
    .filter(([, entry]) => (entry?.snapshot || entry)?.preview)
    .sort((left, right) => String(right[1]?.cachedAt || right[1]?.snapshot?.cachedAt || "").localeCompare(String(left[1]?.cachedAt || left[1]?.snapshot?.cachedAt || "")))
    .slice(0, Math.max(1, limit));
  return Object.fromEntries(entries);
}

function loadCalendarSnapshotCacheStore() {
  try {
    const store = JSON.parse(localStorage.getItem(CALENDAR_SNAPSHOT_CACHE_KEY) || "{}");
    return store && typeof store === "object" ? store : {};
  } catch {
    return {};
  }
}

function normalizeCalendarSnapshotCacheEntry(entry, key = "") {
  if (!entry) return null;
  const snapshot = sanitizeWorkspaceSnapshot(entry?.snapshot || entry);
  if (!snapshot?.preview) return null;
  const cachedAt = String(entry?.cachedAt || snapshot.cachedAt || new Date().toISOString());
  const calendarRevision = String(entry?.calendarRevision || snapshot.calendarRevision || "");
  snapshot.cacheKey = key || snapshot.cacheKey || "";
  snapshot.calendarRevision = calendarRevision;
  snapshot.cachedAt = cachedAt;
  return {
    calendarRevision,
    cachedAt,
    snapshot,
  };
}

function rememberCalendarSnapshotCacheEntry(key, entry) {
  if (!key || !entry?.snapshot?.preview) return;
  calendarSnapshotMemoryCache.delete(key);
  calendarSnapshotMemoryCache.set(key, entry);
  if (calendarSnapshotMemoryCache.size <= MAX_MEMORY_SNAPSHOT_CACHE_ENTRIES) return;
  const oldest = [...calendarSnapshotMemoryCache.entries()]
    .sort((left, right) => String(left[1]?.cachedAt || "").localeCompare(String(right[1]?.cachedAt || "")))[0]?.[0];
  if (oldest) calendarSnapshotMemoryCache.delete(oldest);
}

function calendarSnapshotFromCacheEntry(entry, key = "") {
  const normalized = normalizeCalendarSnapshotCacheEntry(entry, key);
  if (!normalized) return null;
  rememberCalendarSnapshotCacheEntry(key || normalized.snapshot.cacheKey || "", normalized);
  return normalized.snapshot;
}

function saveCalendarSnapshotCacheStore(store) {
  const compactStore = compactCalendarSnapshotCacheStore(store);
  try {
    localStorage.setItem(CALENDAR_SNAPSHOT_CACHE_KEY, JSON.stringify(compactStore));
  } catch (error) {
    if (!isStorageQuotaError(error)) throw error;
    removeLegacyCalendarSnapshotCaches();
    const newestOnly = compactCalendarSnapshotCacheStore(compactStore, 1);
    try {
      localStorage.setItem(CALENDAR_SNAPSHOT_CACHE_KEY, JSON.stringify(newestOnly));
    } catch (retryError) {
      if (!isStorageQuotaError(retryError)) throw retryError;
      // IndexedDB remains the durable cache if localStorage is unavailable.
      try {
        localStorage.removeItem(CALENDAR_SNAPSHOT_CACHE_KEY);
      } catch {
        // Ignore browser storage failures.
      }
    }
  }
}

function calendarSnapshotContext(options = {}) {
  const mode = options.mode || activeCalendarMode();
  const explicitTarget = Object.prototype.hasOwnProperty.call(options, "ownerEmail")
    || Object.prototype.hasOwnProperty.call(options, "ownerId")
    || Object.prototype.hasOwnProperty.call(options, "doctorKey");
  const owner = mode === "doctor-profile"
    ? String(options.ownerId || activeDoctorProfile?.ownerId || activeWorkspaceOwnerKey() || "").trim()
    : normalizeEmail(options.ownerEmail || options.ownerId || activeWorkspaceOwnerKey() || viewedAccountEmail() || currentUserEmail);
  const doctorKey = normalizeRosterName(
    options.doctorKey
    || (mode === "doctor-profile"
      ? activeDoctorProfile?.doctorKey
      : explicitTarget
        ? preferredDoctorKeyForAccountEmail(options.ownerEmail || options.ownerId || "")
        : restoredSessionState?.doctorKey || selectedDoctor()?.key || preferredDoctorKeyForCurrentAccount())
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
  const memoryEntry = calendarSnapshotMemoryCache.get(key);
  if (memoryEntry) return calendarSnapshotFromCacheEntry(memoryEntry, key);
  const entry = loadCalendarSnapshotCacheStore()[key];
  return calendarSnapshotFromCacheEntry(entry, key);
}

function loadCachedCalendarSnapshot() {
  return loadCachedCalendarSnapshotForContext();
}

function saveCalendarSnapshotCacheForContext(snapshot = currentSnapshot, context = {}, options = {}) {
  const sanitized = options.snapshotReady === true ? snapshot : sanitizeWorkspaceSnapshot(snapshot);
  if (!sanitized?.preview) return;
  const key = calendarSnapshotCacheKeyForContext({
    ...context,
    doctorKey: sanitized.session?.doctorKey || context.doctorKey || selectedDoctor()?.key || "",
  });
  if (!key) return;
  const store = loadCalendarSnapshotCacheStore();
  if (options.skipIfRevisionMatches === true && String(store[key]?.calendarRevision || "") === String(sanitized.calendarRevision || currentCalendarRevision || "")) {
    return;
  }
  sanitized.cacheKey = key;
  sanitized.calendarRevision = String(sanitized.calendarRevision || currentCalendarRevision || "");
  sanitized.cachedAt = new Date().toISOString();
  const entry = {
    calendarRevision: sanitized.calendarRevision,
    cachedAt: sanitized.cachedAt,
    snapshot: sanitized,
  };
  rememberCalendarSnapshotCacheEntry(key, entry);
  store[key] = entry;
  saveCalendarSnapshotCacheStore(store);
  queueStoredCalendarSnapshotPersist(key, entry);
}

function saveCalendarSnapshotCache(snapshot = currentSnapshot) {
  saveCalendarSnapshotCacheForContext(snapshot);
}

function queueDeferredSnapshotCachePersist(snapshot = currentSnapshot, context = {}, options = {}) {
  const run = () => {
    try {
      saveCalendarSnapshotCacheForContext(snapshot, context, {
        snapshotReady: true,
        skipIfRevisionMatches: options.skipIfRevisionMatches === true,
      });
    } catch {
      // Snapshot cache persistence must not block the foreground workflow.
    }
  };
  const scheduleIdle = () => {
    if (typeof requestIdleCallback === "function") {
      requestIdleCallback(run, { timeout: 1200 });
    } else {
      window.setTimeout(run, 0);
    }
  };
  if (typeof requestAnimationFrame === "function") {
    requestAnimationFrame(scheduleIdle);
  } else {
    scheduleIdle();
  }
}

function invalidateCalendarSnapshotCache() {
  invalidateCalendarSnapshotCachesForSourceTypes(["mmc", "ddh", "casey", "mch"], { includeCreator: true, includeAllProfiles: true });
}

function sourceTypesFromCalendarSnapshotCacheKey(key = "", entry = null) {
  const [mode = "", ownerId = "", doctorKey = ""] = String(key || "").split("|");
  if (mode === "doctor-profile") {
    const profileId = ownerId.startsWith("doctor-profile:") ? ownerId.slice("doctor-profile:".length) : ownerId;
    const sourceText = profileId.includes("::") ? profileId.split("::").slice(1).join("::") : "";
    return sourceText.split("+").map((item) => String(item || "").toLowerCase()).filter((item) => item === "mmc" || item === "ddh" || item === "casey" || item === "mch");
  }
  if (mode === "creator-account") {
    return ["mmc", "ddh", "casey", "mch"];
  }
  const snapshot = entry?.snapshot || loadCalendarSnapshotCacheStore()[key]?.snapshot;
  const doctor = Array.isArray(snapshot?.doctorOptions)
    ? snapshot.doctorOptions.find((entry) => normalizeRosterName(entry?.key || "") === normalizeRosterName(doctorKey))
      || snapshot.doctorOptions[0]
    : null;
  const fromDoctor = normalizedDoctorSourceTypes(doctor);
  if (fromDoctor.length) return fromDoctor;
  const claims = sanitizeRosterClaims(snapshot?.session?.claims || currentRosterClaims || []);
  return [...new Set(claims.map((claim) => String(claim?.sourceType || "").toLowerCase()).filter((item) => item === "mmc" || item === "ddh" || item === "casey" || item === "mch"))];
}

function calendarSnapshotCacheAffectedBySourceTypes(key = "", entry = null, changedSourceTypes = [], options = {}) {
  const changed = new Set(
    (Array.isArray(changedSourceTypes) ? changedSourceTypes : [])
      .map((item) => String(item || "").toLowerCase())
      .filter((item) => item === "mmc" || item === "ddh" || item === "casey" || item === "mch"),
  );
  if (!changed.size || options.includeAllProfiles === true) return true;
  const mode = String(key || "").split("|")[0] || "";
  if (options.includeCreator !== false && mode === "creator-account") return true;
  const profileSources = sourceTypesFromCalendarSnapshotCacheKey(key, entry);
  if (!profileSources.length) return false;
  return profileSources.some((sourceType) => changed.has(sourceType));
}

function invalidateCalendarSnapshotCachesForSourceTypes(changedSourceTypes = [], options = {}) {
  try {
    const store = loadCalendarSnapshotCacheStore();
    let changed = false;
    for (const [key, entry] of Object.entries(store)) {
      if (!calendarSnapshotCacheAffectedBySourceTypes(key, entry, changedSourceTypes, options)) continue;
      delete store[key];
      changed = true;
    }
    for (const [key, entry] of calendarSnapshotMemoryCache.entries()) {
      if (!calendarSnapshotCacheAffectedBySourceTypes(key, entry, changedSourceTypes, options)) continue;
      calendarSnapshotMemoryCache.delete(key);
      changed = true;
    }
    if (changed) saveCalendarSnapshotCacheStore(store);
    void deleteStoredCalendarSnapshotsForSourceTypes(changedSourceTypes, options).catch(() => null);
  } catch {
    // Cache invalidation must not block the foreground workflow.
  }
}

function invalidateCalendarSnapshotCachesForChangedRosterFiles(entries = [], options = {}) {
  const sourceTypes = [...new Set(
    (Array.isArray(entries) ? entries : [])
      .map((entry) => String(entry?.sourceType || "").toLowerCase())
      .filter((item) => item === "mmc" || item === "ddh" || item === "casey" || item === "mch"),
  )];
  invalidateCalendarSnapshotCachesForSourceTypes(sourceTypes, options);
}

function renderCachedCalendarSnapshot(options = {}) {
  return renderCachedCalendarSnapshotForContext(calendarSnapshotContext(options), options);
}

function renderCachedCalendarSnapshotForContext(context = {}, options = {}) {
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  const cached = loadCachedCalendarSnapshotForContext(context);
  if (!cached?.preview) return false;
  return applyCachedCalendarSnapshot(cached, options);
}

async function renderCachedCalendarSnapshotForContextAsync(context = {}, options = {}) {
  if (renderCachedCalendarSnapshotForContext(context, options)) return true;
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  const cached = await loadCachedCalendarSnapshotForContextAsync(context);
  if (!cached?.preview) return false;
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  return applyCachedCalendarSnapshot(cached, options);
}

function applyCachedCalendarSnapshot(cached, options = {}) {
  const expectedRevision = String(options.expectedRevision || "");
  if (expectedRevision && String(cached.calendarRevision || "") !== expectedRevision) return false;
  if (!calendarTransitionStillCurrent(options.transition)) return false;
  currentSnapshot = cached;
  currentSnapshotStale = false;
  currentSnapshotBuiltAt = cached.cachedAt || cached.preview?.lastParsed || "";
  currentCalendarRevision = expectedRevision || cached.calendarRevision || currentCalendarRevision;
  applyLoadedCalendarFileRefs(cached);
  renderWorkspaceFromSnapshot(cached, cached.session || {});
  markLoginPhase("cachedCalendarRendered", options.loginStartedAt);
  markAccountSwitchPhase("cachedCalendarRendered", options.accountSwitchStartedAt);
  setStatus("Calendar loaded from cache.");
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
const DB_VERSION = 3;
const IMPORT_STORE = "imports";
const SNAPSHOT_STORE = "calendarSnapshots";
const CONFLICT_SELECTIONS_KEY = "roster-conflict-selections";

async function openImportsDb() {
  if (!("indexedDB" in window)) {
    throw new Error("Browser storage is unavailable.");
  }
  return await new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (request.oldVersion < 3 && db.objectStoreNames.contains(SNAPSHOT_STORE)) {
        db.deleteObjectStore(SNAPSHOT_STORE);
      }
      if (!db.objectStoreNames.contains(IMPORT_STORE)) {
        db.createObjectStore(IMPORT_STORE, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(SNAPSHOT_STORE)) {
        db.createObjectStore(SNAPSHOT_STORE, { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open import storage."));
  });
}

async function loadStoredCalendarSnapshotEntry(key) {
  if (!key || !("indexedDB" in window)) return null;
  const db = await openImportsDb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readonly");
      const request = tx.objectStore(SNAPSHOT_STORE).get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Could not load snapshot cache."));
    });
  } finally {
    db.close();
  }
}

async function loadCachedCalendarSnapshotForContextAsync(context = {}) {
  const cached = loadCachedCalendarSnapshotForContext(context);
  if (cached?.preview) return cached;
  const key = calendarSnapshotCacheKeyForContext(context);
  if (!key) return null;
  const stored = await loadStoredCalendarSnapshotEntry(key).catch(() => null);
  const entry = normalizeCalendarSnapshotCacheEntry(stored, key);
  if (!entry) return null;
  rememberCalendarSnapshotCacheEntry(key, entry);
  return entry.snapshot;
}

async function saveStoredCalendarSnapshotEntry(key, entry) {
  if (!key || !entry?.snapshot?.preview || !("indexedDB" in window)) return;
  const db = await openImportsDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readwrite");
      tx.objectStore(SNAPSHOT_STORE).put({
        key,
        calendarRevision: String(entry.calendarRevision || ""),
        cachedAt: String(entry.cachedAt || new Date().toISOString()),
        snapshot: entry.snapshot,
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not save snapshot cache."));
    });
  } finally {
    db.close();
  }
}

async function pruneStoredCalendarSnapshots() {
  if (!("indexedDB" in window)) return;
  const db = await openImportsDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readwrite");
      const store = tx.objectStore(SNAPSHOT_STORE);
      const request = store.getAll();
      request.onsuccess = () => {
        const cutoff = Date.now() - MAX_STORED_SNAPSHOT_CACHE_AGE_MS;
        const entries = (request.result || [])
          .map((entry) => ({
            entry,
            cachedAtMs: Date.parse(entry?.cachedAt || entry?.snapshot?.cachedAt || "") || 0,
          }))
          .sort((left, right) => right.cachedAtMs - left.cachedAtMs);
        entries.forEach(({ entry, cachedAtMs }, index) => {
          if (index >= MAX_STORED_SNAPSHOT_CACHE_ENTRIES || (cachedAtMs > 0 && cachedAtMs < cutoff)) {
            store.delete(entry.key);
          }
        });
      };
      request.onerror = () => reject(request.error || new Error("Could not inspect snapshot cache."));
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not prune snapshot cache."));
    });
  } finally {
    db.close();
  }
}

function queueStoredCalendarSnapshotMaintenance(options = {}) {
  storedSnapshotWritesSinceMaintenance += options.afterWrite === true ? 1 : 0;
  if (storedSnapshotMaintenanceQueued) return;
  if (options.afterWrite === true && storedSnapshotWritesSinceMaintenance < 25) return;
  storedSnapshotMaintenanceQueued = true;
  const run = () => {
    void pruneStoredCalendarSnapshots()
      .catch(() => null)
      .finally(() => {
        storedSnapshotMaintenanceQueued = false;
        storedSnapshotWritesSinceMaintenance = 0;
      });
  };
  if (typeof requestIdleCallback === "function") {
    requestIdleCallback(run, { timeout: 4000 });
  } else {
    window.setTimeout(run, 1500);
  }
}

function queueStoredCalendarSnapshotPersist(key, entry) {
  const run = () => {
    void saveStoredCalendarSnapshotEntry(key, entry)
      .then(() => queueStoredCalendarSnapshotMaintenance({ afterWrite: true }))
      .catch(() => null);
  };
  if (typeof requestIdleCallback === "function") {
    requestIdleCallback(run, { timeout: 1500 });
  } else {
    window.setTimeout(run, 0);
  }
}

async function deleteStoredCalendarSnapshots(keys = []) {
  const uniqueKeys = [...new Set(keys.filter(Boolean))];
  if (!uniqueKeys.length || !("indexedDB" in window)) return;
  const db = await openImportsDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readwrite");
      const store = tx.objectStore(SNAPSHOT_STORE);
      for (const key of uniqueKeys) store.delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not clear snapshot cache."));
    });
  } finally {
    db.close();
  }
}

async function deleteStoredCalendarSnapshotsForSourceTypes(changedSourceTypes = [], options = {}) {
  if (!("indexedDB" in window)) return;
  const db = await openImportsDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readwrite");
      const store = tx.objectStore(SNAPSHOT_STORE);
      const request = store.getAll();
      request.onsuccess = () => {
        for (const entry of request.result || []) {
          if (!calendarSnapshotCacheAffectedBySourceTypes(entry.key, entry, changedSourceTypes, options)) continue;
          store.delete(entry.key);
        }
      };
      request.onerror = () => reject(request.error || new Error("Could not inspect snapshot cache."));
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not clear snapshot cache."));
    });
  } finally {
    db.close();
  }
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

function accountImportsSavePayload(imports = selectedFiles) {
  return {
    ...snapshotCloudSavePayload(),
    imports: (imports || []).map((entry) => ({ ...entry })),
    removedImportIds: [],
    repositorySynced: true,
  };
}

function queueAccountImportsSave(options = {}) {
  const payload = accountImportsSavePayload();
  queueBackgroundCloudStateSave(payload, {
    delayMs: options.delayMs ?? 0,
    reportErrors: false,
    onError: () => setStatus("Account sync will retry in the background."),
  });
}

function keepFileIdsAfterRemoval() {
  return selectedFiles.map((entry) => entry.id).filter(Boolean);
}

function syncResponsePurgedRemovedId(syncResult, removedId) {
  if (!removedId) return syncResult?.allPurged === true;
  const verification = syncResult?.verification || [];
  const entry = verification.find((item) => item.fileId === removedId);
  if (entry) return entry.purged === true;
  return !(syncResult?.removedFileIds || []).includes(removedId);
}

async function syncRosterRepositoryToSelection(removedId = null, options = {}) {
  const keepFileIds = keepFileIdsAfterRemoval();
  const syncResult = await calendarStoreRequestWithRetry("syncRosterRepository", {
    keepFileIds,
    selectedDoctorKey: selectedDoctor()?.key || OWNER_DOCTOR_KEY,
  }, { attempts: options.attempts || 4 });
  if (removedId && !syncResponsePurgedRemovedId(syncResult, removedId)) {
    throw new Error("Roster repository sync did not confirm removal.");
  }
  if (syncResult?.files || syncResult?.keptFileIds) {
    calendarStoreStatus = { ...syncResult, checkedAt: new Date().toISOString() };
    calendarStoreStatusError = "";
  }
  if (Array.isArray(syncResult?.availableDoctors)) {
    applyAuthoritativeAvailableDoctors(syncResult.availableDoctors);
  }
  return syncResult;
}

function restoreRemovedImportAfterFailedRemoval(id, removedEntry) {
  pendingRemovedImportIds.delete(id);
  if (!removedEntry || selectedFiles.some((entry) => entry.id === removedEntry.id)) {
    renderFileSurfaces();
    return;
  }
  selectedFiles = [...selectedFiles, {
    id: removedEntry.id,
    repoId: removedEntry.repoId || removedEntry.id,
    name: removedEntry.name,
    sourceType: removedEntry.sourceType,
    size: removedEntry.size || 0,
    lastModified: removedEntry.lastModified || 0,
    addedAt: removedEntry.addedAt || "",
    file: removedEntry.file || null,
  }];
  renderFileSurfaces();
  void syncCreatorDoctorPickerWithRemainingRosters({ localOnly: true })
    .then(() => renderDoctorState())
    .catch(() => null);
}

async function completeRosterRemovalAfterSync(id, removedEntry, removedName) {
  invalidateCalendarSnapshotCachesForChangedRosterFiles([removedEntry]);

  if (!selectedFiles.length) {
    if (isViewingCreatorAccount() && cloudAvailable) {
      applyAuthoritativeAvailableDoctors([]);
      await refreshAvailableDoctorsAfterRosterChange({ localOnly: true, mergeAvailableDoctors: false });
      await waitForCreatorSwitcherRemovalSettled(removedEntry);
      const announced = await tryAnnounceCreatorSwitcherRosterUpdate();
      if (!announced && creatorSwitcherAnnouncementBaseline !== null) {
        setStatus("Roster removed from storage but switcher did not refresh yet; retrying…", true);
        throw new Error("Switcher did not refresh after roster removal.");
      }
    }
    resetDerivedState({ preserveSession: true });
    queueAccountImportsSave();
    if (!canUseCreatorDoctorSwitcher() || creatorSwitcherAnnouncementBaseline === null) {
      setStatus("Add a roster file to begin.");
    }
    return;
  }

  await refreshAvailableDoctorsAfterRosterChange({ localOnly: true, mergeAvailableDoctors: false });

  const loaded = await loadCloudCalendarEvents({
    preserveExistingSnapshot: false,
    cachedRevision: "",
    allowInlineBuild: true,
  });
  if (loaded && currentSnapshot) {
    currentSnapshot = filterSnapshotDoctorsAfterRemoval(currentSnapshot, removedEntry);
    renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
  } else if (!loaded) {
    throw new Error("Could not reload calendar after roster removal.");
  }

  try {
    await syncCreatorDoctorPickerWithRemainingRosters({ localOnly: true });
    renderDoctorState();
  } catch {
    // Keep the last merged doctor list.
  }

  await waitForCreatorSwitcherRemovalSettled(removedEntry);

  const announced = await tryAnnounceCreatorSwitcherRosterUpdate();
  if (!announced && isViewingCreatorAccount() && cloudAvailable && creatorSwitcherAnnouncementBaseline !== null) {
    setStatus("Roster removed from storage but switcher did not refresh yet; retrying…", true);
    throw new Error("Switcher did not refresh after roster removal.");
  }

  queueAccountImportsSave();
}

function scheduleRosterRemovalRetry(id, removedEntry, removedName) {
  const runId = ++rosterRemovalRetryRunId;
  void (async () => {
    const delays = [1500, 4000, 10000, 20000, 45000];
    for (const delay of delays) {
      if (runId !== rosterRemovalRetryRunId) return;
      if (!pendingRemovedImportIds.has(id)) return;
      await new Promise((resolve) => setTimeout(resolve, delay));
      if (runId !== rosterRemovalRetryRunId) return;
      if (!pendingRemovedImportIds.has(id)) return;
      try {
        setStatus(`Retrying removal of ${removedName}...`);
        await syncRosterRepositoryToSelection(id, { attempts: 4 });
        pendingRemovedImportIds.delete(id);
        await completeRosterRemovalAfterSync(id, removedEntry, removedName);
        renderFileSurfaces();
        return;
      } catch (error) {
        const retryable = /503|429|CPU|memory|overload|timed out|network/i.test(String(error?.message || ""));
        if (!retryable) {
          setStatus(error.message || `Could not remove ${removedName} from roster storage.`, true);
          restoreRemovedImportAfterFailedRemoval(id, removedEntry);
          return;
        }
      }
    }
    if (!pendingRemovedImportIds.has(id)) return;
    setStatus(`Could not remove ${removedName} from roster storage after multiple retries.`, true);
  })();
}

async function removeStoredImport(id) {
  if (!id || pendingRemovedImportIds.has(id)) return;
  const removedEntry = selectedFiles.find((entry) => entry.id === id)
    || (calendarStoreStatus?.files || []).find((file) => file.id === id);
  const removedName = removedEntry?.name || "roster file";
  pendingRemovedImportIds.add(id);
  if (isViewingCreatorAccount() && cloudAvailable) {
    creatorSwitcherAnnouncementBaseline = null;
    captureCreatorSwitcherVisibleBaseline();
  }
  cancelScheduledCloudStateSave();
  rosterSyncStates.delete(id);
  selectedFiles = selectedFiles.filter((entry) => entry.id !== id);
  creatorCalendarSourceFileRefs = creatorCalendarSourceFileRefs.filter((entry) => entry.id !== id);
  removeImportRefsFromCurrentSnapshot(id);
  removeImportRefsFromWorkspaceStore(id);
  saveCurrentSessionState();
  renderFileSurfaces();
  setStatus(`Removing ${removedName}...`);
  let rosterDataRemoved = false;
  let removalCompleted = false;
  try {
    try {
      await deleteStoredImportRecords([id]);
      await garbageCollectStoredImports();
    } catch {
      // Keep in-memory removal even if persistent storage is unavailable.
    }
    if (cloudAvailable && isCreatorAuthenticated()) {
      try {
        await syncRosterRepositoryToSelection(id, { attempts: 4 });
        rosterDataRemoved = true;
      } catch (error) {
        const overload = /503|429|CPU|memory|overload/i.test(String(error?.message || ""));
        if (!overload) {
          setStatus(error.message || `Could not remove ${removedName} from roster storage.`, true);
        }
      }
    }
    if (rosterDataRemoved) {
      removalCompleted = true;
      await completeRosterRemovalAfterSync(id, removedEntry, removedName);
      pendingRemovedImportIds.delete(id);
      return;
    }
    if (cloudAvailable && isCreatorAuthenticated()) {
      setStatus(`Could not remove ${removedName} from roster storage yet. Retrying in the background...`, true);
      scheduleRosterRemovalRetry(id, removedEntry, removedName);
    } else {
      queueAccountImportsSave();
      removalCompleted = true;
      pendingRemovedImportIds.delete(id);
    }
  } finally {
    if (removalCompleted) pendingRemovedImportIds.delete(id);
    renderFileSurfaces();
  }
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

async function bootstrapImports(options = {}) {
  try {
    if (!calendarTransitionStillCurrent(options.transition)) return;
    syncAccountsButton();
    if (!selectedFiles.length) {
      if (cloudAvailable) {
        restoreCreatorImportFilesIfNeeded();
        if (!selectedFiles.length) {
          void syncCreatorFileListFromStore().catch(() => null);
        }
      } else {
        const workspace = loadCurrentWorkspace();
        selectedFiles = await loadStoredImportsByRefs(workspace?.fileRefs || []);
        restoredSessionState = workspace?.session || restoredSessionState;
        currentSnapshot = sanitizeWorkspaceSnapshot(workspace?.snapshot);
        currentSnapshotStale = false;
      }
    }
    renderFilesList();
    if (!calendarTransitionStillCurrent(options.transition)) return;
    if (currentSnapshot?.preview && currentSnapshot.doctorOptions?.length) {
      const snapshotInvalid = snapshotHasUnresolvablePreviewEvents(currentSnapshot);
      if (isViewingCreatorAccount() && selectedFiles.length) {
        try {
          await syncCreatorDoctorPickerWithRemainingRosters({ snapshotDoctors: currentSnapshot.doctorOptions || [] });
        } catch {
          // Keep the last merged doctor list.
        }
      }
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      scheduleInsightWarmup();
      if (currentSnapshotStale || snapshotInvalid) {
        setStatus("Refreshing calendar...");
        const canRefreshFromBrowserFiles = selectedFiles.length && (selectedFiles.every((entry) => entry.file) || await ensureSelectedFilesLoaded());
        if (isViewingCreatorAccount() && cloudAvailable) {
          void refreshCreatorSnapshotInBackground({ transition: options.transition });
        } else if (canRefreshFromBrowserFiles) {
          void refreshSnapshotInBackground();
        } else if (cloudAvailable) {
          void loadCloudCalendarEvents({
            adminTargetEmail: adminViewingEmail ? viewedAccountEmail() : "",
            allowInlineBuild: options.allowInlineBuild !== false,
            preserveExistingSnapshot: true,
            cachedRevision: currentSnapshot?.calendarRevision || currentCalendarRevision || "",
            transition: options.transition,
          })
            .then((loadedCalendar) => {
              if (!calendarTransitionStillCurrent(options.transition)) return;
              if (!loadedCalendar || !currentSnapshot?.preview) return;
              if (currentSnapshotStale) {
                setStatus("Calendar loaded. Checking for updates...");
                return;
              }
              renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
              currentSnapshotStale = false;
              currentSnapshotBuiltAt = new Date().toISOString();
              setStatus("Calendar refreshed.");
            })
            .catch((error) => {
              if (!calendarTransitionStillCurrent(options.transition)) return;
              setStatus(error.message || "Could not refresh the calendar.", true);
            });
        } else {
          setStatus("Calendar loaded.");
        }
      } else {
        setStatus("Calendar loaded.");
      }
      return;
    }
    if (cloudAvailable && selectedFiles.length && selectedFiles.some((entry) => !entry.file)) {
      const loadedCalendar = await loadCloudCalendarEvents({
        adminTargetEmail: adminViewingEmail ? viewedAccountEmail() : "",
        allowInlineBuild: options.allowInlineBuild !== false,
        preserveExistingSnapshot: true,
        transition: options.transition,
      });
      if (!calendarTransitionStillCurrent(options.transition)) return;
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
      setStatus(currentNonClinical
        ? "This non-clinical account is ready. No roster name or clinical shifts are linked."
        : availableRosterDoctors.length && !currentRosterClaims.length
          ? "Choose your roster name, or upload a roster if your name is not listed."
          : "Add a roster file to begin.");
    }
  } catch (error) {
    if (currentSnapshot?.preview) {
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
    }
    renderFilesList();
    setStatus(error.message || "Could not restore browser-stored roster files.", true);
  }
}

function snapshotHasUnresolvablePreviewEvents(snapshot) {
  const events = Array.isArray(snapshot?.preview?.events) ? snapshot.preview.events : [];
  if (!events.length) return false;
  const reviewIds = new Set(Array.isArray(snapshot?.preview?.review) ? snapshot.preview.review.map((item) => item?.id).filter(Boolean) : []);
  if (!reviewIds.size) return true;
  return events.some((event) => event?.id && !reviewIds.has(event.id));
}

function renderWorkspaceFromSnapshot(snapshot, session = {}, options = {}) {
  const preservedScroll = options.preserveScroll === true
    ? {
        pageY: window.scrollY || document.documentElement.scrollTop || 0,
        previewY: previewSection?.scrollTop || 0,
      }
    : null;
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
  pendingPreviewSnapToToday = options.preserveScroll !== true;
  renderSettings();
  renderFilesList();
  renderDoctorState();
  indexReviewItems(latestPreview.review || []);
  rebuildClientPreview();
  refreshFacilityOverviewPreferredFacility();
  scheduleInsightWarmup();
  saveCurrentWorkspace();
  if (preservedScroll) {
    requestAnimationFrame(() => {
      if (isMobileLayout()) window.scrollTo({ top: preservedScroll.pageY, behavior: "auto" });
      else if (previewSection) previewSection.scrollTop = preservedScroll.previewY;
    });
  }
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
      if (isViewingCreatorAccount() && cloudAvailable) {
        await refreshCreatorSnapshotInBackground();
        return;
      }
      await analyzeFiles({ resetRange: false, preserveVisiblePreview: true });
      if (latestPreview) {
        currentSnapshotStale = false;
        currentSnapshotBuiltAt = new Date().toISOString();
        setStatus("Calendar refreshed.");
      }
    } catch (error) {
      setStatus(error.message || "Could not refresh the calendar.", true);
    } finally {
      snapshotRefreshPromise = null;
    }
  })();
  return snapshotRefreshPromise;
}

async function bootstrapApp() {
  try {
    removeLegacyCalendarSnapshotCaches();
    if (new URLSearchParams(window.location.search).get("invite")) {
      entrancePage.classList.remove("hidden");
      appShell.classList.add("hidden");
      loginForm.classList.add("hidden");
      createAccountForm.classList.add("hidden");
      inviteAccountForm?.classList.remove("hidden");
      loginTabButton?.classList.add("hidden");
      createTabButton?.classList.add("hidden");
      hideLoadingScreen();
      return;
    }
    renderLoginState();
    if (!currentUserEmail || !currentUserPassword) {
      openLoginModal();
      setStatus("Log in with an email address to load your roster workspace.");
      queueStoredCalendarSnapshotMaintenance();
      return;
    }
    const loginStartedAt = performance.now();
    const transition = beginCalendarTransition();
    forceConsoleSkin();
    const cacheContext = accountCalendarContextForEmail(currentUserEmail);
    const renderedCachedSnapshot = await renderCachedCalendarSnapshotForContextAsync(cacheContext, {
      loginStartedAt,
      transition,
    });
    if (renderedCachedSnapshot) {
      renderLoginState();
      hideLoadingScreen();
      markLoginPhase("firstCalendarPaint", loginStartedAt);
      markLoginPaintCommitted(loginStartedAt);
      setStatus("Checking calendar for updates...");
    } else {
      setStatus("Loading calendar...");
    }
    const loginData = await restoreCloudState({
      mode: "login",
      responseMode: "fast",
      deferHydration: true,
      deferContext: true,
      deferSnapshotPersistence: true,
      skipSnapshotCacheWriteIfCurrent: true,
      cachedRevision: renderedCachedSnapshot ? currentSnapshot?.calendarRevision || "" : "",
      preserveExistingSnapshot: renderedCachedSnapshot,
      loginStartedAt,
      transition,
    });
    if (!currentUserEmail || !calendarTransitionStillCurrent(transition)) return;
    renderLoginState();
    if (isNonClinicalDirectorWorkspace()) {
      if ((loginData?.responseMode || "full") === "fast") {
        queueDeferredAccountContextLoad({
          loginStartedAt,
          targetEmail: "",
          responseMode: loginData?.responseMode || "fast",
          delayMs: 50,
          transition,
        });
      }
      launchNonClinicalDirectorWorkspace({ transition }, loginStartedAt);
      queueStoredCalendarSnapshotMaintenance();
      return;
    }
    const inlineSnapshotReady = loginSnapshotReadyForRender();
    if (!renderedCachedSnapshot && inlineSnapshotReady) {
      renderWorkspaceFromSnapshot(currentSnapshot, restoredSessionState || currentSnapshot.session || {});
      markLoginPhase("cachedCalendarRendered", loginStartedAt);
      markLoginPhase("firstCalendarPaint", loginStartedAt);
      markLoginPaintCommitted(loginStartedAt);
    }
    if ((loginData?.responseMode || "full") === "fast") {
      queueDeferredAccountContextLoad({
        loginStartedAt,
        targetEmail: "",
        responseMode: loginData?.responseMode || "fast",
        delayMs: 50,
        transition,
      });
    }
    queuePostLoginHydration({
      includeBootstrap: true,
      forceCalendarRefresh: renderedCachedSnapshot || inlineSnapshotReady,
      allowInlineBuild: !(renderedCachedSnapshot || inlineSnapshotReady),
      transition,
    }, loginStartedAt);
    queueStoredCalendarSnapshotMaintenance();
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
  if (adminConsoleOpen && currentAdminTab === "system" && !adminConsoleLoading) {
    appendLiveAdminConsoleMessage(text, isError);
    refreshAdminConsoleMarkupIfVisible();
  }
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
