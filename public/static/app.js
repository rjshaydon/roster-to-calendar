import {
  applyEventOverrides,
  buildRosterView,
  customEventsToEvents,
  defaultSettings as rosterDefaultSettings,
  doctorOptions as rosterDoctorOptions,
  exportIcs,
  parseUploadForm,
  previewSummary,
  serializeConflict,
  serializeEvent,
  serializeReviewItem,
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
const dropZone = document.querySelector("#dropZone");
const chooseFilesButton = document.querySelector("#chooseFilesButton");
const fmsUrlInput = document.querySelector("#fmsUrlInput");
const fmsUrlButton = document.querySelector("#fmsUrlButton");
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
const insightsModal = document.querySelector("#insightsModal");
const insightsCloseButton = document.querySelector("#insightsCloseButton");
const insightsModalTitle = document.querySelector("#insightsModalTitle");
const insightsModalSubtitle = document.querySelector("#insightsModalSubtitle");
const insightsModalBody = document.querySelector("#insightsModalBody");
const loginBar = document.querySelector("#loginBar");
const loginIdentity = document.querySelector("#loginIdentity");
const logoutButton = document.querySelector("#logoutButton");
const backToCreatorButton = document.querySelector("#backToCreatorButton");
const skinControl = document.querySelector("#skinControl");
const skinSelect = document.querySelector("#skinSelect");
const loginForm = document.querySelector("#loginForm");
const loginEmail = document.querySelector("#loginEmail");
const loginPassword = document.querySelector("#loginPassword");
const createAccountForm = document.querySelector("#createAccountForm");
const createRealName = document.querySelector("#createRealName");
const createEmail = document.querySelector("#createEmail");
const createPassword = document.querySelector("#createPassword");
const currentDayPreview = document.querySelector("#currentDayPreview");
const exportButton = document.querySelector("#exportButton");
const mobileExportButton = document.querySelector("#mobileExportButton");
const mobileSettingsButton = document.querySelector("#mobileSettingsButton");
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
const customEventWhoButton = document.querySelector("#customEventWhoButton");
const contextMenu = document.querySelector("#contextMenu");

const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const OWNER_EMAIL = "rhaydon@gmail.com";
const OWNER_DOCTOR_KEY = "RICHARD HAYDON";
const DEFAULT_MMC_LOCATION = "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
const DEFAULT_DDH_LOCATION = "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";
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
const CURRENT_EMAIL_KEY = "roster-current-email";
const CURRENT_PASSWORD_KEY = "roster-current-password";
const SKIN_KEY = "roster-active-skin";
const SIX_MONTH_LIMIT_DAYS = 183;
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
  "hospitalFilter",
  "dateFrom",
  "dateTo",
];

let doctorOptions = [];
let detectedSources = {};
let selectedFiles = [];
let parsedRosterSources = null;
let doctorRoleIndex = null;
let parsedImportDoctors = new Map();
let settings = defaultSettings();
let overrides = {};
let latestPreview = null;
let reviewIndex = new Map();
let customEvents = [];
let currentPreviewEvents = new Map();
let availablePreviewHospitals = [];
let dragEventId = null;
let copiedEvent = null;
let previewGesture = null;
let suppressPreviewClickUntil = 0;
let openReviewId = "";
let conflictSelections = {};
let accountState = loadAccountState();
let restoredSessionState = null;
let currentUserEmail = loadCurrentUserEmail();
let currentUserPassword = sessionStorage.getItem(CURRENT_PASSWORD_KEY) || "";
let currentUserRole = currentUserEmail === OWNER_EMAIL ? "creator" : "user";
let authUserEmail = currentUserEmail;
let authUserPassword = currentUserPassword;
let adminViewingEmail = "";
let currentSkin = loadSkin();
let cloudAvailable = false;
let cloudSaveTimer = 0;
let pendingCloudSaveSnapshot = null;
let enforcingRosterLimit = false;
let serverUsers = [];
let currentRosterClaims = [];
let latestNameMatches = [];
let availableRosterDoctors = [];
let insightsState = null;
let doctorAnalysisCacheKey = "";
let doctorAnalysisCache = new Map();

const settingsInputs = Object.fromEntries(
  SETTINGS_FIELDS.map((id) => [id, document.querySelector(`#${id}`)]),
);

applySkin(currentSkin);
applyShiftColours(settings);
applyCurrentDayHighlight(settings);
setEntranceTab("login");

chooseFilesButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  fileInput.click();
});

fmsUrlButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  const value = fmsUrlInput.value.trim();
  if (!value) {
    setStatus("Paste a FindMyShift webcal:// first.", true);
    return;
  }
  setStatus("FindMyShift URL subscriptions are not active yet. Please upload a FindMyShift export for now.", true);
});

dropZone.addEventListener("click", (event) => {
  if (event.target.closest("button, input, label, select, textarea, a")) return;
  fileInput.click();
});
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    if (event.target.closest("button, input, label, select, textarea, a")) return;
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", async () => {
  const accepted = validateIncomingFiles([...fileInput.files]);
  if (!accepted.length) {
    fileInput.value = "";
    return;
  }
  await mergeFiles(accepted);
  fileInput.value = "";
  await analyzeFiles();
});

for (const eventName of ["dragenter", "dragover"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
    document.body.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "dragend", "drop"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (eventName === "dragleave" && dropZone.contains(event.relatedTarget)) return;
    dropZone.classList.remove("is-dragging");
    document.body.classList.remove("is-dragging");
  });
}

dropZone.addEventListener("drop", async (event) => {
  const accepted = validateIncomingFiles([...event.dataTransfer.files]);
  if (!accepted.length) return;
  await mergeFiles(accepted);
  await analyzeFiles();
});

filesButton.addEventListener("click", openFilesModal);
filesCloseButton.addEventListener("click", closeFilesModal);
filesModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-files]")) closeFilesModal();
});
filesList.addEventListener("click", async (event) => {
  const removeButton = event.target.closest("[data-remove-import]");
  if (!removeButton) return;
  if (!canRemoveImports()) return;
  await removeStoredImport(removeButton.dataset.removeImport);
});
accountsButton.addEventListener("click", async () => {
  if (isOwnerAccount()) {
    await loadServerUsers();
  }
  renderAccountsModal();
  accountsModal.classList.remove("hidden");
  accountsModal.setAttribute("aria-hidden", "false");
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
  updateAccountDetails(email, { password, realName });
});
accountsBody.addEventListener("click", (event) => {
  const deleteButton = event.target.closest("[data-delete-account]");
  if (deleteButton) {
    deleteAccount(deleteButton.dataset.deleteAccount);
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
  await loginWithEmail(email, password, { mode: "login" });
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
  returnToCreatorAccount();
});
skinSelect.addEventListener("change", () => {
  if (!isOwnerAccount()) return;
  applySkin(skinSelect.value);
  syncSkinControl();
});

doctorSelect.addEventListener("change", async () => {
  clearPreviewData();
  saveCurrentSessionState();
  syncActionState();
  if (selectedDoctor()) await updatePreview({ resetRange: true });
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
  settingsPanel.classList.toggle("hidden");
});
settingsCloseButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  settingsPanel.classList.add("hidden");
});
mobileSettingsButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  settingsPanel.classList.toggle("hidden");
});

for (const [key, input] of Object.entries(settingsInputs)) {
  input.addEventListener("change", () => {
    settings[key] = input.type === "checkbox" ? input.checked : input.value;
    if (!settings.showNormalizedTitles && !settings.showRawValues) {
      settings.showNormalizedTitles = true;
      settingsInputs.showNormalizedTitles.checked = true;
    }
    saveCurrentSessionState();
    if (latestPreview && (key === "dateFrom" || key === "dateTo")) {
      rebuildClientPreview();
      setStatus("Preview range updated.");
      return;
    }
    if (latestPreview && key === "hospitalFilter") {
      updatePreview();
      return;
    }
    if (latestPreview && (key === "defaultLocationMmc" || key === "defaultLocationDdh")) {
      updatePreview();
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
      if (latestPreview) rebuildClientPreview();
      setStatus("Current day highlight updated.");
      return;
    }
    setStatus("Settings updated.");
  });
}

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
  const whoButton = event.target.closest("[data-open-who-on]");
  if (!whoButton) return;
  closeReviewModal();
  openWhoInsight(whoButton.dataset.openWhoOn, whoButton.dataset.openWhoOn);
});

mobileExportButton.addEventListener("click", () => form.requestSubmit());
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
    returnToCreatorAccount();
    return;
  }
  const rangeTrigger = event.target.closest("[data-range-trigger]");
  if (rangeTrigger) {
    openPreviewRangePicker(rangeTrigger.dataset.rangeTrigger);
    return;
  }
  const whoTrigger = event.target.closest("[data-insight-who]");
  if (whoTrigger) {
    openWhoInsight(whoTrigger.dataset.insightWho, whoTrigger.dataset.insightWhoEnd);
    return;
  }
  const whenTrigger = event.target.closest("[data-insight-when]");
  if (whenTrigger) {
    openWhenInsight(whenTrigger.dataset.insightWhen, whenTrigger.dataset.insightWhenEnd);
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
  startPreviewGesture(event, chip);
});
preview.addEventListener("change", (event) => {
  const doctorPicker = event.target.closest("[data-preview-doctor-select]");
  if (doctorPicker) {
    doctorSelect.value = doctorPicker.value;
    clearPreviewData();
    saveCurrentSessionState();
    syncActionState();
    updatePreview();
    return;
  }

  const rangeInput = event.target.closest("[data-range-input]");
  if (!rangeInput) return;
  applyPreviewRangeChange(rangeInput.dataset.rangeInput, rangeInput.value);
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
    renderInsightsModal();
    return;
  }
  const whenDoctorSelect = event.target.closest("[data-insights-when-doctor]");
  if (whenDoctorSelect && insightsState?.mode === "when") {
    insightsState.comparisonDoctorKey = whenDoctorSelect.value;
    renderInsightsModal();
  }
});
issuesList.addEventListener("click", (event) => {
  const card = event.target.closest("[data-review-id]");
  if (!card) return;
  openReviewModal(card.dataset.reviewId);
});
conflictsList.addEventListener("change", async (event) => {
  const select = event.target.closest("[data-conflict-key]");
  if (!select) return;
  conflictSelections[select.dataset.conflictKey] = select.value;
  saveConflictSelections();
  saveCurrentSessionState();
  await updatePreview();
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
  if (event.key === "Escape") {
    closeReviewModal();
    closeCustomEventModal();
    closeContextMenu();
    closeFilesModal();
    closeAccountsModal();
    closeInsightsModal();
    settingsPanel.classList.add("hidden");
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
    settingsPanel.classList.add("hidden");
  }
}, true);
document.addEventListener("pointermove", (event) => {
  if (!previewGesture || event.pointerId !== previewGesture.pointerId) return;
  updatePreviewGesture(event);
});
document.addEventListener("pointerup", (event) => {
  if (!previewGesture || event.pointerId !== previewGesture.pointerId) return;
  finishPreviewGesture(event);
});
document.addEventListener("pointercancel", (event) => {
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
customEventDeleteButton.addEventListener("click", () => {
  const id = customEventId.value;
  if (!id) return;
  removeCustomEventForActiveCalendar(id);
  closeCustomEventModal();
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Manual event removed.");
});
customEventWhoButton.addEventListener("click", () => {
  const date = customEventWhoButton.dataset.openWhoOn || customEventStartDate.value || formatDateKey(new Date());
  closeCustomEventModal();
  openWhoInsight(date, date);
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

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedFiles.length) {
    setStatus("Add at least one roster file first.", true);
    return;
  }
  const doctor = selectedDoctor();
  if (!doctor) {
    setStatus("Choose a doctor before exporting.", true);
    return;
  }

  setStatus("Building calendar file...");
  try {
    const ics = await buildBrowserIcs(doctor);
    const payload = new Blob([ics], { type: "text/calendar; charset=utf-8" });
    const url = URL.createObjectURL(payload);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${doctor.displayName} roster.ics`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("Calendar file ready.");
  } catch (error) {
    setStatus(error.message, true);
  }
});

function defaultSettings() {
  return {
    showSourcePrefix: true,
    showAmPm: true,
    showTimes: true,
    showRawValues: false,
    showNormalizedTitles: true,
    currentDayBorderColor: "#c44949",
    currentDayBorderOpacity: "42",
    currentDayBackgroundColor: "#c44949",
    currentDayBackgroundOpacity: "8",
    currentDayFillStyle: "gradient",
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
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
  };
}

async function mergeFiles(files) {
  let persistenceFailed = false;
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

async function analyzeFiles() {
  if (!selectedFiles.length) {
    setStatus("Add a roster file to begin.");
    return;
  }
  clearPreviewData();
  doctorOptions = [];
  detectedSources = {};
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
    settings = {
      ...defaultSettings(),
      ...settings,
      ...(data.settings || {}),
      ...(restoredSessionState?.settings || {}),
    };
    overrides = sanitizeOverrideState(restoredSessionState?.overrides);
    customEvents = sanitizeCustomEvents(restoredSessionState?.customEvents, activeCalendarEmail());
    conflictSelections = {
      ...loadConflictSelections(),
      ...(restoredSessionState?.conflictSelections || {}),
    };
    renderSettings();
    renderFilesList();
    renderDoctorState();
    saveCurrentWorkspace();
    scheduleCloudStateSave();
    if (selectedDoctor()) {
      await updatePreview({ resetRange: true });
      return;
    }
  } catch (error) {
    doctorOptions = [];
    detectedSources = {};
    clearPreviewData();
    renderFilesList();
    syncActionState();
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
    doctors: rosterDoctorOptions(parsed.sources.mmc, parsed.sources.ddh),
    settings: rosterDefaultSettings(),
  };
}

async function parseCurrentRosterForm(doctor = null) {
  return await parseUploadForm(new Request(`${window.location.origin}/browser-roster-parse`, {
    method: "POST",
    body: createFormData(doctor),
  }));
}

function sourceImports(sources) {
  return [
    ...sources.mmc.map((entry) => sourceImportMeta(entry, "mmc")),
    ...sources.ddh.map((entry) => sourceImportMeta(entry, "ddh")),
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
  for (const entry of sources.mmc) {
    result.set(entry.id, rosterDoctorOptions([entry], []).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "mmc",
    })));
  }
  for (const entry of sources.ddh) {
    result.set(entry.id, rosterDoctorOptions([], [entry]).map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: "ddh",
    })));
  }
  return result;
}

function renderSettings() {
  for (const [key, input] of Object.entries(settingsInputs)) {
    if (!input) continue;
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else if (input.type === "color") {
      input.value = isHexColour(settings[key]) ? settings[key] : defaultShiftColourForField(key);
    } else {
      input.value = settings[key] || "";
    }
  }
  applyShiftColours(settings);
  applyCurrentDayHighlight(settings);
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

function renderFilesList() {
  if (!filesList) return;
  if (!selectedFiles.length) {
    const emptyMessage = canRemoveImports()
      ? "Add rosters and they will stay here until removed."
      : "Add rosters and they will be retained in the shared roster repository.";
    filesList.innerHTML = `<article class="issue-card"><strong>No files imported yet.</strong><p>${emptyMessage}</p></article>`;
    return;
  }

  const canRemove = canRemoveImports();
  filesList.innerHTML = selectedFiles.map((entry) => `
    <article class="issue-card">
      <div>
        <strong>${escapeHtml(entry.name)}</strong>
        <p>${escapeHtml(String(entry.sourceType || "").toUpperCase())} · Imported ${escapeHtml(formatTimestamp(entry.addedAt))}</p>
      </div>
      ${canRemove
        ? `<button type="button" class="button button-secondary" data-remove-import="${entry.id}">Remove</button>`
        : `<span class="file-readonly">Repository file</span>`}
    </article>
  `).join("");
}

function renderDoctorState() {
  doctorSelect.innerHTML = "";
  doctorName.textContent = "";
  doctorName.classList.add("hidden");
  doctorSelect.classList.add("hidden");
  doctorSection.classList.add("hidden");
  controlBar.classList.toggle("hidden", !selectedFiles.length);
  mobileActionBar.classList.toggle("hidden", !selectedFiles.length);
  settingsPanel.classList.add("hidden");

  if (!doctorOptions.length) {
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

  if (doctorOptions.length === 1) {
    doctorName.textContent = doctorOptions[0].displayName;
    doctorName.classList.remove("hidden");
    setStatus("Loading calendar...");
  } else {
    for (const doctor of doctorOptions) {
      const option = document.createElement("option");
      option.value = doctor.key;
      option.textContent = doctor.displayName;
      doctorSelect.append(option);
    }
    const preferredDoctorKey = preferredDoctorKeyForCurrentAccount();
    if (preferredDoctorKey && doctorOptions.some((doctor) => doctor.key === preferredDoctorKey)) {
      doctorSelect.value = preferredDoctorKey;
    } else if (restoredSessionState?.doctorKey && doctorOptions.some((doctor) => doctor.key === restoredSessionState.doctorKey)) {
      doctorSelect.value = restoredSessionState.doctorKey;
    }
    doctorSelect.classList.remove("hidden");
    setStatus(preferredDoctorKey ? "Loading calendar..." : "Choose a doctor to load the calendar.");
  }

  syncActionState();
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

async function claimSelectedRosterName() {
  const index = Number(claimDoctorSelect.value);
  const candidate = Number.isInteger(index) ? availableRosterDoctors[index] : null;
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
    await bootstrapImports();
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
    if (options.resetRange || !settings.dateFrom || !settings.dateTo) {
      const range = deriveDefaultPreviewRange(data.events || []);
      if (!settings.dateFrom) settings.dateFrom = range.start;
      if (!settings.dateTo) settings.dateTo = range.end;
      renderSettings();
    }
    indexReviewItems(data.review || []);
    rebuildClientPreview();
    saveCurrentSessionState();
    setStatus("Calendar loaded.");
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
  const validDoctors = new Set(rosterDoctorOptions(parsedRosterSources.mmc, parsedRosterSources.ddh).map((item) => item.key));
  const requestedKeys = new Set([doctor.key, ...(doctor.aliases || []).map((alias) => alias.key)].filter(Boolean));
  if (![...requestedKeys].some((key) => validDoctors.has(key))) {
    throw new Error("The selected doctor was not found in the uploaded roster files.");
  }
  const view = buildRosterView(
    parsedRosterSources.mmc,
    parsedRosterSources.ddh,
    doctor.key,
    settings,
    overrides,
    conflictSelections,
    doctor.aliases || [],
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
  const view = buildClientPreviewData(latestPreview);
  if (enforceSixMonthLimit(view)) return;
  renderConflicts(view.conflicts || []);
  renderPreviewGrid(doctor, view);
  renderIssues(view.issues || []);
  saveCurrentSessionState();
}

function buildClientPreviewData(baseData) {
  const baseEvents = new Map((baseData.events || []).map((event) => [event.id, { ...event }]));
  const range = deriveRangeBounds(baseData.events || []);
  const events = [];
  const deletedItems = [];
  const hospitals = availableHospitalsForPreview(baseData.events || []);
  if (settings.hospitalFilter === "all" || hospitals.length > availablePreviewHospitals.length) {
    availablePreviewHospitals = hospitals;
  }

  for (const item of reviewIndex.values()) {
    const event = baseEvents.get(item.id);
    if (!event) continue;
    const override = overrides[item.id] || {};
    const include = typeof override.include === "boolean" ? override.include : item.include;
    if (!include) {
      deletedItems.push({
        ...item,
        status: "deleted",
        message: "Deleted from preview/export. Open to restore or edit.",
      });
      continue;
    }
    events.push(buildEventOverridePatch(event, item, override));
  }

  for (const event of customEventsForActiveCalendar()) {
    events.push(customEventToPreviewEvent(event));
  }

  const previewStart = settings.dateFrom || range.start;
  const previewEnd = settings.dateTo || range.end;
  const visibleEvents = filterEventsByPreviewRange(events, previewStart, previewEnd);
  visibleEvents.sort(comparePreviewEvents);
  return {
    ...baseData,
    events: visibleEvents,
    count: visibleEvents.length,
    date_range: formatPreviewRange(previewStart, previewEnd) || (visibleEvents.length ? summarizeEvents(visibleEvents) : "No events found"),
    previewStart,
    previewEnd,
    hospitals: availablePreviewHospitals,
    lastImport: latestImportTimestamp(),
    issues: [
      ...(baseData.issues || []).filter((issue) => {
      const override = overrides[issue.id] || {};
      const reviewItem = reviewIndex.get(issue.id);
      const include = typeof override.include === "boolean" ? override.include : reviewItem?.include ?? true;
      return include;
      }),
      ...deletedItems,
    ],
  };
}

function availableHospitalsForPreview(events) {
  const codes = new Set();
  for (const event of events || []) {
    if (event.source) codes.add(String(event.source).toUpperCase());
    const titlePrefix = String(event.title || "").match(/^(MMC|DDH):/i)?.[1];
    if (titlePrefix) codes.add(titlePrefix.toUpperCase());
  }
  return [...codes]
    .filter((code) => code === "MMC" || code === "DDH")
    .sort();
}

async function buildBrowserIcs(doctor) {
  if (!parsedRosterSources) {
    const data = await analyzeFilesInBrowser();
    doctorOptions = doctorOptionsForCurrentAccount(data.doctors || []);
  }
  const doctors = rosterDoctorOptions(parsedRosterSources.mmc, parsedRosterSources.ddh);
  const requestedKeys = new Set([doctor.key, ...(doctor.aliases || []).map((alias) => alias.key)].filter(Boolean));
  const selectedDoctor = doctors.find((item) => requestedKeys.has(item.key));
  if (!selectedDoctor) {
    throw new Error("The selected doctor was not found in the uploaded roster files.");
  }
  const rosterEvents = applyEventOverrides(
    buildRosterView(
      parsedRosterSources.mmc,
      parsedRosterSources.ddh,
      doctor.key,
      settings,
      overrides,
      conflictSelections,
      doctor.aliases || [],
    ).events,
    overrides,
  );
  const events = [...rosterEvents, ...customEventsToEvents(customEventsForActiveCalendar(), settings)].sort(comparePreviewEvents);
  if (!events.length) {
    throw new Error("No calendar events were found for the selected doctor.");
  }
  return exportIcs(events, doctor.displayName || selectedDoctor.displayName);
}

function enforceSixMonthLimit(view) {
  if (currentUserRole === "creator" || enforcingRosterLimit || !view.events?.length) return false;
  const range = deriveRangeBounds(view.events);
  if (!range.start || !range.end) return false;
  const latest = parseDateOnly(range.end);
  const cutoff = addDays(latest, -SIX_MONTH_LIMIT_DAYS);
  const cutoffKey = formatDateKey(cutoff);
  if (range.start >= cutoffKey || settings.dateFrom >= cutoffKey) return false;

  const ok = window.confirm(
    `Standard accounts can keep the latest 6 months of roster active. This upload extends before ${formatDate(cutoffKey)}. Delete/hide events before that date and keep the latest roster period?`,
  );
  if (!ok) {
    exportButton.disabled = true;
    mobileExportButton.disabled = true;
    setStatus("Export disabled until the roster is limited to 6 months.", true);
    return false;
  }

  enforcingRosterLimit = true;
  settings.dateFrom = cutoffKey;
  renderSettings();
  saveCurrentSessionState();
  rebuildClientPreview();
  enforcingRosterLimit = false;
  setStatus("Roster limited to the latest 6 months for this account.");
  return true;
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
      </div>
      <span>${escapeHtml(item.rawValue)}</span>
    </article>
  `).join("");
  issuesPanel.classList.remove("hidden");
}

function indexReviewItems(items) {
  reviewIndex = new Map(items.map((item) => [item.id, item]));
}

function renderPreviewGrid(doctor, data) {
  const events = data.events || [];
  currentPreviewEvents = new Map(events.map((event) => [event.id, event]));
  const days = buildPreviewDays(events, data.previewStart, data.previewEnd);
  document.body.classList.add("has-calendar-preview");
  if (!days.length) {
    preview.innerHTML = `
      ${renderPreviewHeader(doctor, data)}
      <div class="preview-empty">No events match the current settings.</div>
    `;
    preview.classList.remove("hidden");
    previewSection.classList.remove("hidden");
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
        ${adminViewingEmail && isCreatorAuthenticated()
          ? `<button type="button" class="button button-secondary preview-back-button" data-preview-back-to-creator>Back to creator</button>`
          : ""}
        <button type="button" class="button button-secondary preview-logout-button" data-preview-logout>Log out</button>
      </div>
    </div>
  `;
}

function renderPreviewDoctorControl(doctor) {
  if (canUseDoctorPicker() && doctorOptions.length > 1) {
    return `
      <label class="preview-doctor-control">
        <span>Doctor</span>
        <select data-preview-doctor-select>
          ${doctorOptions.map((option) => `
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
  const cards = day.events.length
    ? day.events.map((event) => renderPreviewChip(event, formatDateKey(day.date))).join("")
    : `<div class="preview-chip preview-chip-empty"></div>`;
  const currentDayClass = isCurrentDay(day.date) ? " is-current-day" : "";
  return `
    <div class="preview-cell${currentDayClass}" data-add-date="${formatDateKey(day.date)}">
      <div class="preview-date">${day.date.getDate()}</div>
      <div class="preview-stack">${cards}</div>
    </div>
  `;
}

function renderPreviewChip(event, dayKey) {
  const lines = [];
  const startKey = event.start.slice(0, 10);
  if (settings.showNormalizedTitles || !settings.showRawValues) {
    const marker = event.isEditedImport ? '<span class="preview-chip-marker" aria-label="Imported event edited">*</span>' : "";
    lines.push(`<strong>${escapeHtml(event.title)}${marker}</strong>`);
  }
  if (settings.showRawValues) {
    lines.push(`<span class="preview-chip-raw">${escapeHtml(event.rawValue)}</span>`);
  }
  const meta = [];
  if (!event.allDay && settings.showTimes && startKey === dayKey && event.timeLabel) meta.push(event.timeLabel);
  const metaMarkup = meta.length ? `<span class="preview-chip-meta">${escapeHtml(meta.join(" · "))}</span>` : "";
  return `<button type="button" class="preview-chip preview-chip-${eventTone(event)}" data-review-id="${event.id}" data-review-date="${dayKey}">${lines.join("")}${metaMarkup}</button>`;
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
  const headerCells = DAY_NAMES.map((day) => `<div class="preview-day-name">${day}</div>`).join("");
  const bodyRows = [];
  let lastMonthKey = "";
  const firstMonday = section.weeks[0]?.week?.[0]?.date;
  const lastSunday = section.weeks.at(-1)?.week?.at(-1)?.date;

  section.weeks.forEach(({ week, index }) => {
    const monday = week[0]?.date;
    const monthKey = monday ? `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, "0")}` : "";
    if (monthKey && monthKey !== lastMonthKey) {
      bodyRows.push(`<div class="preview-month-row">${formatMonth(monday)}</div>`);
      lastMonthKey = monthKey;
    }
    bodyRows.push(`
      <div class="preview-week-label">
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
        <div class="preview-term-actions">
          <button type="button" class="button button-secondary preview-term-button" data-insight-who="${formatDateKey(firstMonday)}" data-insight-who-end="${formatDateKey(lastSunday)}">Who</button>
          <button type="button" class="button button-secondary preview-term-button" data-insight-when="${formatDateKey(firstMonday)}" data-insight-when-end="${formatDateKey(lastSunday)}">When</button>
        </div>
      </div>
      <div class="preview-grid">
        <div class="preview-week-label preview-week-label-head">Week</div>
        ${headerCells}
        ${bodyRows.join("")}
      </div>
    </section>
  `;
}

function openWhoInsight(termStart, termEnd) {
  const date = defaultInsightDate(termStart, termEnd);
  insightsState = {
    mode: "who",
    termStart,
    termEnd,
    date,
  };
  renderInsightsModal();
}

function openWhenInsight(termStart, termEnd) {
  const options = comparisonDoctorOptions();
  insightsState = {
    mode: "when",
    termStart,
    termEnd,
    comparisonDoctorKey: options[0]?.key || "",
  };
  renderInsightsModal();
}

function closeInsightsModal() {
  insightsState = null;
  insightsModal.classList.add("hidden");
  insightsModal.setAttribute("aria-hidden", "true");
  insightsModalBody.innerHTML = "";
}

function renderInsightsModal() {
  if (!insightsState) return;
  if (insightsState.mode === "who") {
    renderWhoInsight();
  } else if (insightsState.mode === "when") {
    renderWhenInsight();
  }
  insightsModal.classList.remove("hidden");
  insightsModal.setAttribute("aria-hidden", "false");
}

function renderWhoInsight() {
  const date = insightsState.date;
  const mine = selectedDoctorEventsForInsights(date, date).filter(isRosterShiftEvent);
  const activeSources = new Set(mine.map(eventSourceCode).filter(Boolean));
  const coworkers = mine.length
    ? comparisonDoctorOptions()
      .map((doctor) => ({
        doctor,
        events: comparisonDoctorEvents(doctor.key, date, date)
          .filter(isRosterShiftEvent)
          .filter((event) => !activeSources.size || activeSources.has(eventSourceCode(event))),
      }))
      .filter((entry) => entry.events.length)
      .flatMap((entry) => buildWhoAssignments(entry.doctor, entry.events))
    : [];
  const grouped = groupWhoAssignments(coworkers);

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
    ${coworkers.length
      ? renderWhoGroups(grouped)
      : `<article class="issue-card"><p>No other doctors are rostered on this date.</p></article>`}
  `;
}

function renderWhenInsight() {
  const options = comparisonDoctorOptions();
  const selectedKey = options.some((doctor) => doctor.key === insightsState.comparisonDoctorKey)
    ? insightsState.comparisonDoctorKey
    : options[0]?.key || "";
  insightsState.comparisonDoctorKey = selectedKey;
  const selectedComparison = options.find((doctor) => doctor.key === selectedKey) || null;
  const mine = selectedDoctorEventsForInsights(insightsState.termStart, insightsState.termEnd).filter(isRosterShiftEvent);
  const theirs = selectedComparison
    ? comparisonDoctorEvents(selectedComparison.key, insightsState.termStart, insightsState.termEnd).filter(isRosterShiftEvent)
    : [];
  const overlaps = buildOverlapDays(mine, theirs);
  const nextOverlapDate = chooseNextOverlapDate(overlaps);

  insightsModalTitle.textContent = "When";
  insightsModalSubtitle.textContent = "Find the dates where both doctors are working in this term.";
  insightsModalBody.innerHTML = `
    <div class="insights-controls">
      <label class="field">
        <span>Doctor</span>
        <select data-insights-when-doctor>
          ${options.map((doctor) => `
            <option value="${escapeHtml(doctor.key)}" ${doctor.key === selectedKey ? "selected" : ""}>${escapeHtml(doctor.displayName)}</option>
          `).join("")}
        </select>
      </label>
    </div>
    ${selectedComparison
      ? overlaps.length
        ? overlaps.map((entry) => `
          <article class="issue-card${entry.date === nextOverlapDate ? " is-next-overlap" : ""}" ${entry.date === nextOverlapDate ? 'data-insight-next="true"' : ""}>
            <strong>${escapeHtml(formatInsightDate(entry.date))}</strong>
            <p><strong>${escapeHtml(selectedDoctor()?.displayName || "Selected doctor")}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.mine))}</p>
            <p><strong>${escapeHtml(selectedComparison.displayName)}:</strong> ${escapeHtml(renderInsightShiftSummary(entry.theirs))}</p>
          </article>
        `).join("")
        : `<article class="issue-card"><p>No overlapping working days were found in ${escapeHtml(detectAustralianTerm(parseDateOnly(insightsState.termStart)).label)}.</p></article>`
      : `<article class="issue-card"><p>No comparison doctors are available in these roster files.</p></article>`}
  `;
  const nextCard = insightsModalBody.querySelector("[data-insight-next='true']");
  if (nextCard) {
    requestAnimationFrame(() => nextCard.scrollIntoView({ block: "nearest", behavior: "smooth" }));
  }
}

function comparisonDoctorOptions() {
  if (!parsedRosterSources || (!parsedRosterSources.mmc?.length && !parsedRosterSources.ddh?.length)) return [];
  return prioritizeDoctorOptions(rosterDoctorOptions(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || []))
    .filter((doctor) => doctor.key !== selectedDoctor()?.key);
}

function selectedDoctorEventsForInsights(start, end) {
  if (!latestPreview) return [];
  return buildCurrentDoctorPreviewEvents(start, end);
}

function comparisonDoctorEvents(doctorKey, start, end) {
  const cache = getDoctorAnalysisCache();
  const events = cache.get(doctorKey) || [];
  return filterInsightEvents(events, start, end);
}

function buildCurrentDoctorPreviewEvents(start, end) {
  const baseEvents = new Map((latestPreview?.events || []).map((event) => [event.id, { ...event }]));
  const events = [];
  for (const item of reviewIndex.values()) {
    const event = baseEvents.get(item.id);
    if (!event) continue;
    const override = overrides[item.id] || {};
    const include = typeof override.include === "boolean" ? override.include : item.include;
    if (!include) continue;
    events.push(buildEventOverridePatch(event, item, override));
  }
  for (const event of customEventsForActiveCalendar()) {
    if (event.include === false) continue;
    events.push(customEventToPreviewEvent(event));
  }
  return filterInsightEvents(events, start, end);
}

function filterInsightEvents(events, start, end) {
  const startDate = parseDateOnly(start);
  const endDate = parseDateOnly(end);
  return events
    .filter((event) => matchesPreviewHospitalFilter(event, settings.hospitalFilter))
    .filter((event) => eventOverlapsDateRange(event, startDate, endDate))
    .sort(comparePreviewEvents);
}

function matchesPreviewHospitalFilter(event, hospitalFilter) {
  if (!hospitalFilter || hospitalFilter === "all") return true;
  return String(event.source || "").toLowerCase() === String(hospitalFilter).toLowerCase();
}

function buildOverlapDays(mine, theirs) {
  const mineByDay = indexEventsByDay(mine);
  const theirsByDay = indexEventsByDay(theirs);
  return [...mineByDay.keys()]
    .filter((date) => theirsByDay.has(date))
    .sort()
    .map((date) => ({
      date,
      mine: mineByDay.get(date) || [],
      theirs: theirsByDay.get(date) || [],
    }));
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
  const metadata = doctorMetadataForKey(doctor.key);
  return dedupeInsightEvents(events)
    .map((event) => buildWhoAssignment(doctor, metadata, event))
    .filter(Boolean);
}

function buildWhoAssignment(doctor, metadata, event) {
  const source = eventSourceCode(event);
  const role = metadata[source]?.role || metadata.any?.role || "";
  const period = whoPeriodLabel(event);
  const team = whoTeamLabel(event);
  return {
    doctorName: doctor.displayName,
    role,
    roleLabel: role || "",
    roleRank: whoRoleRank(role),
    source,
    period,
    team,
    teamRank: whoTeamRank(team, source),
    specialTime: whoSpecialTimeLabel(event, period),
    event,
  };
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
              <div class="who-team-person">
                <span class="who-team-name">${escapeHtml(item.doctorName)}${item.roleLabel ? ` (${escapeHtml(item.roleLabel)})` : ""}</span>
                ${item.specialTime ? `<span class="who-team-time">${escapeHtml(item.specialTime)}</span>` : ""}
              </div>
            `).join("")}
          </div>
        </article>
      `).join("")}
    </section>
  `).join("");
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
    .replace(/^(MMC|DDH):\s*/i, "")
    .replace(/\s+(AM|PM)\b/i, "")
    .trim() || "Other";
}

function whoTeamRank(team, source) {
  const normalized = String(team || "").toLowerCase();
  const sourceCode = String(source || "").toUpperCase();
  const ranks = sourceCode === "DDH"
    ? ["avao", "orange", "silver", "resus", "float", "clinic", "fast track", "ssu", "hith", "vhh", "paeds", "extra", "other"]
    : ["green", "amber", "resus", "float", "clinic", "fast track", "ssu", "other"];
  const index = ranks.indexOf(normalized);
  return index >= 0 ? index : ranks.length;
}

function whoRoleRank(role) {
  const ranks = {
    SMS: 0,
    SR: 1,
    CMO: 2,
    IR: 3,
    JR: 4,
    HMO: 5,
    I: 6,
    ENP: 7,
    AMP: 8,
  };
  return Object.prototype.hasOwnProperty.call(ranks, role) ? ranks[role] : 99;
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
  };
  const standardStarts = standard[source]?.[period] || new Set();
  return standardStarts.has(start) ? "" : `${start}-${end}`;
}

function eventSourceCode(event) {
  const explicit = String(event?.source || "").trim().toUpperCase();
  if (explicit === "MMC" || explicit === "DDH") return explicit;
  const titlePrefix = String(event?.title || "").match(/^(MMC|DDH):/i)?.[1];
  return titlePrefix ? titlePrefix.toUpperCase() : "";
}

function isRosterShiftEvent(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  return !(
    text.includes("annual leave")
    || text.includes("conference leave")
    || text.includes("sick leave")
    || text.includes("clinical support")
    || /\bcso?\b/.test(text)
    || text.includes("phnw")
    || text.includes("public holiday")
  );
}

function chooseNextOverlapDate(overlaps) {
  if (!overlaps.length) return "";
  const today = formatDateKey(new Date());
  return overlaps.find((entry) => entry.date >= today)?.date || overlaps[0].date;
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
  if (!parsedRosterSources || (!parsedRosterSources.mmc?.length && !parsedRosterSources.ddh?.length)) {
    clearDoctorAnalysisCache();
    return doctorAnalysisCache;
  }
  const cacheKey = JSON.stringify({
    imports: currentImportStateKey(),
    sourcePrefix: settings.showSourcePrefix,
    showAmPm: settings.showAmPm,
    includeAnnualLeave: settings.includeAnnualLeave,
    includeConferenceLeave: settings.includeConferenceLeave,
    includePublicHoliday: settings.includePublicHoliday,
    includeSickLeave: settings.includeSickLeave,
  });
  if (doctorAnalysisCacheKey === cacheKey && doctorAnalysisCache.size) return doctorAnalysisCache;
  const cache = new Map();
  const analysisSettings = {
    ...settings,
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
    includeLocations: false,
  };
  for (const doctor of rosterDoctorOptions(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [])) {
    const view = buildRosterView(parsedRosterSources?.mmc || [], parsedRosterSources?.ddh || [], doctor.key, analysisSettings);
    cache.set(doctor.key, view.events);
  }
  doctorAnalysisCacheKey = cacheKey;
  doctorAnalysisCache = cache;
  return doctorAnalysisCache;
}

function clearDoctorAnalysisCache() {
  doctorAnalysisCacheKey = "";
  doctorAnalysisCache = new Map();
}

function syncActionState() {
  const ready = Boolean(selectedDoctor());
  exportButton.disabled = !ready;
  mobileExportButton.disabled = !ready;
}

function createFormData(doctor = null) {
  const body = new FormData();
  for (const entry of selectedFiles) {
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
  if (event.source === "Custom") {
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
    id: `custom-${Date.now().toString(36)}`,
    ownerEmail: activeCalendarEmail(),
    source: "Custom",
    isEditedImport: false,
  }, targetDate);
  customEvents.push(previewEventToCustomEvent(shifted));
  closeContextMenu();
  rebuildClientPreview();
  saveCurrentSessionState();
  setStatus("Event pasted.");
}

function deletePreviewEvent(id) {
  const event = currentPreviewEvents.get(id);
  if (!event) return;
  if (event.source === "Custom") {
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
  if (!doctorOptions.length) return null;
  if (doctorOptions.length === 1) return doctorOptions[0];
  const preferredDoctorKey = preferredDoctorKeyForCurrentAccount();
  return doctorOptions.find((doctor) => doctor.key === doctorSelect.value)
    || doctorOptions.find((doctor) => doctor.key === preferredDoctorKey)
    || doctorOptions[0];
}

function preferredDoctorKeyForCurrentAccount() {
  if (currentUserEmail === OWNER_EMAIL && !adminViewingEmail) return OWNER_DOCTOR_KEY;
  return "";
}

function canUseDoctorPicker() {
  return isOwnerAccount() && !adminViewingEmail;
}

function doctorOptionsForCurrentAccount(doctors) {
  const options = (doctors || []).map((doctor) => ({
    ...doctor,
    sourceTypes: Array.isArray(doctor.sourceTypes) ? doctor.sourceTypes : [],
  }));
  if (canUseDoctorPicker()) return prioritizeDoctorOptions(options);
  const matches = options.filter((doctor) => doctorMatchesCurrentAccount(doctor));
  if (!matches.length) return [];
  const aliases = matches.flatMap((doctor) => {
    const sourceTypes = doctor.sourceTypes.length ? doctor.sourceTypes : sourceTypesForClaimedDoctor(doctor.key);
    return sourceTypes.map((sourceType) => ({
      sourceType,
      key: doctor.key,
      displayName: doctor.displayName,
    }));
  });
  const displayName = currentAccount().realName || matches[0].displayName;
  return [{
    key: matches[0].key,
    displayName,
    aliases: dedupeDoctorAliases(aliases),
    sourceTypes: [...new Set(aliases.map((alias) => alias.sourceType))],
  }];
}

function prioritizeDoctorOptions(options) {
  const preferredDoctorKey = preferredDoctorKeyForCurrentAccount();
  if (!preferredDoctorKey) return options;
  return [...options].sort((left, right) => {
    const leftPreferred = left.key === preferredDoctorKey ? 1 : 0;
    const rightPreferred = right.key === preferredDoctorKey ? 1 : 0;
    if (leftPreferred !== rightPreferred) return rightPreferred - leftPreferred;
    return left.displayName.localeCompare(right.displayName);
  });
}

function doctorMatchesCurrentAccount(doctor) {
  const claimKeys = new Set(currentRosterClaims.map((claim) => claim.key));
  if (claimKeys.has(doctor.key)) return true;
  return likelySameRosterName(currentAccount().realName, doctor.displayName);
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
  return String(value || "")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .trim()
    .toUpperCase()
    .split(/\s+/)
    .filter(Boolean);
}

function doctorMetadataForKey(doctorKey) {
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
    ["AMP", "AMP"],
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

function resetDerivedState() {
  doctorOptions = [];
  detectedSources = {};
  availableRosterDoctors = [];
  overrides = {};
  customEvents = [];
  restoredSessionState = null;
  doctorRoleIndex = null;
  clearDoctorAnalysisCache();
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
  settingsPanel.classList.add("hidden");
  claimSection.classList.add("hidden");
  clearPreviewData();
}

function clearPreviewData() {
  latestPreview = null;
  reviewIndex = new Map();
  currentPreviewEvents = new Map();
  availablePreviewHospitals = [];
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
  const currentTerm = australianTermForDate(new Date());
  const nextTerm = nextAustralianTerm(currentTerm);
  const endTerm = events.some((event) => eventOverlapsDateRange(event, nextTerm.start, addDays(nextTerm.end, -1)))
    ? nextTerm
    : currentTerm;
  return {
    start: formatDateKey(currentTerm.start),
    end: formatDateKey(addDays(endTerm.end, -1)),
  };
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
  return new Date(value).toLocaleString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatIssueHeading(item) {
  const status = item.status === "unknown"
    ? "Unknown"
    : item.status === "deleted"
      ? "Deleted"
      : "Review";
  return `${status} · ${item.source} · ${formatDate(item.startDay)}`;
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
          <span>Normalized result</span>
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
        <div class="custom-event-grid">
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
        <div class="custom-event-grid ${allDay ? "hidden" : ""}" data-override-time-fields="${item.id}">
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
              ${buildLocationOptionMarkup(preset.mode)}
            </select>
          </label>
          <label class="field ${preset.mode === "custom" ? "" : "hidden"}" data-override-custom-location-field="${item.id}">
            <span>Custom location</span>
            <input type="text" value="${escapeHtml(preset.customValue)}" data-override-custom-location="${item.id}">
          </label>
        </div>
      </div>
      <div class="review-meta">
        <span>Suggested title: ${escapeHtml(item.suggestedTitle || "No normalized result")}</span>
        ${item.timeLabel ? `<span>Times: ${escapeHtml(item.timeLabel)}</span>` : ""}
        ${item.location ? `<span>Location: ${escapeHtml(item.location)}</span>` : ""}
      </div>
      <div class="modal-actions">
        <button type="button" class="button button-secondary" data-open-who-on="${escapeHtml(insightDate)}">Who else am I working with?</button>
      </div>
      ${resetButton}
      ${warnings}
    </article>
  `;
  reviewModal.classList.remove("hidden");
  reviewModal.setAttribute("aria-hidden", "false");
}

function closeReviewModal() {
  openReviewId = "";
  reviewModal.classList.add("hidden");
  reviewModal.setAttribute("aria-hidden", "true");
  reviewModalBody.innerHTML = "";
}

function openCustomEventModal(event = null, presetDate = null) {
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
  customEventDeleteButton.classList.toggle("hidden", !event);
  customEventWhoButton.dataset.openWhoOn = event?.startDate || presetDate || now;
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
}

function populateLocationOptions() {
  customEventLocationMode.innerHTML = buildLocationOptionMarkup();
}

function buildLocationOptionMarkup(selectedMode = "") {
  const options = [];
  if (detectedSources.mmc?.length) options.push({ value: "mmc", label: "MMC Car Park" });
  if (detectedSources.ddh?.length) options.push({ value: "ddh", label: "DDH Car Park" });
  options.push({ value: "offsite", label: "Off-site" });
  options.push({ value: "custom", label: "Custom location" });
  return options.map((option) => `<option value="${option.value}" ${option.value === selectedMode ? "selected" : ""}>${option.label}</option>`).join("");
}

function detectLocationPreset(location) {
  if (!location) return { mode: "offsite", customValue: "" };
  if (location === settings.defaultLocationMmc || location === DEFAULT_MMC_LOCATION) return { mode: "mmc", customValue: "" };
  if (location === settings.defaultLocationDdh || location === DEFAULT_DDH_LOCATION) return { mode: "ddh", customValue: "" };
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
    id: customEventId.value || `custom-${Date.now().toString(36)}`,
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

function resolveCustomEventLocation() {
  if (customEventLocationMode.value === "mmc") return settings.defaultLocationMmc || DEFAULT_MMC_LOCATION;
  if (customEventLocationMode.value === "ddh") return settings.defaultLocationDdh || DEFAULT_DDH_LOCATION;
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
}

function closeFilesModal() {
  filesModal.classList.add("hidden");
  filesModal.setAttribute("aria-hidden", "true");
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
      displayName: String(doctor?.displayName || "").trim(),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
      claimedBy: normalizeEmail(doctor?.claimedBy || ""),
      claimedByName: String(doctor?.claimedByName || "").trim(),
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
  const email = currentUserEmail || accountState.currentEmail;
  return accountState.users.find((user) => user.email === email) || {
    email,
    realName: "",
    password: "",
    role: currentUserRole === "creator" ? "owner" : "user",
    claims: [],
  };
}

function isOwnerAccount() {
  return currentUserRole === "creator" || currentAccount()?.role === "owner";
}

function canRemoveImports() {
  return isOwnerAccount() && !adminViewingEmail;
}

function isCreatorAuthenticated() {
  return normalizeEmail(authUserEmail || currentUserEmail) === OWNER_EMAIL && Boolean(authUserPassword || currentUserPassword);
}

function syncAccountsButton() {
  accountsButton.textContent = isOwnerAccount() ? "Accounts" : "Account";
}

function renderAccountsModal() {
  const me = currentAccount();
  const ownerView = isOwnerAccount();
  accountsModalTitle.textContent = ownerView ? "Accounts" : "Account";
  accountsModalSubtitle.textContent = ownerView
    ? "Manage your account and review other users."
    : "Manage your account details.";

  const serverOtherUsers = serverUsers
    .map(normalizeServerUser)
    .filter((user) => user.email !== me.email);
  const localOtherUsers = accountState.users.filter((user) => user.email !== me.email);
  const otherUsers = serverOtherUsers.length ? serverOtherUsers : localOtherUsers;
  const linkedNames = renderLinkedRosterNames(currentRosterClaims);
  accountsBody.innerHTML = `
    <article class="review-card">
      <div class="review-top">
        <div>
          <strong>${ownerView ? "Owner account" : "Your account"}</strong>
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
      ${linkedNames}
    </article>
    ${ownerView ? `
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
            <span>${otherUsers.length ? `${otherUsers.length} account${otherUsers.length === 1 ? "" : "s"}` : "No other users have logged in yet."}</span>
          </div>
        </div>
        <div class="issues-list">
          ${otherUsers.length ? otherUsers.map((user) => `
            <article class="issue-card">
              <div>
                <strong>${escapeHtml(user.realName || "Name not set")}</strong>
                <p>${escapeHtml(user.email)} · ${user.role === "owner" ? "Creator" : "Standard user"} · ${formatUserSites(user)} · storage limit: latest 6 months active</p>
              </div>
              <div class="account-actions">
                <button type="button" class="button button-secondary" data-enter-account="${escapeHtml(user.email)}">Enter account</button>
                ${user.email !== OWNER_EMAIL ? `<button type="button" class="button button-danger" data-delete-account="${escapeHtml(user.email)}">Delete</button>` : ""}
              </div>
            </article>
          `).join("") : `<article class="issue-card"><p>No additional users yet.</p></article>`}
        </div>
      </article>
    ` : ""}
  `;
}

function renderLinkedRosterNames(claims) {
  const items = sanitizeRosterClaims(claims);
  if (!items.length) {
    return `<p class="status">No roster names are linked to this account yet.</p>`;
  }
  return `
    <div class="issues-list account-claim-list">
      ${items.map((claim) => `
        <article class="issue-card">
          <div>
            <strong>${escapeHtml(claim.sourceType.toUpperCase())}</strong>
            <p>${escapeHtml(claim.displayName)}</p>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function formatUserSites(user) {
  const sites = Array.isArray(user.sites) ? user.sites.filter(Boolean) : [];
  return sites.length ? `sites: ${sites.join(", ")}` : "no linked sites";
}

function updateAccountDetails(email, patch) {
  accountState.users = accountState.users.map((user) => user.email === email ? {
    ...user,
    password: patch.password || user.password || "",
    realName: patch.realName || user.realName || "",
  } : user);
  saveAccountState();
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
    await enterUserAccount(email);
  } catch (error) {
    if (error.message === "Cloud storage is not configured.") {
      setStatus(serverStorageRequiredMessage(), true);
      return;
    }
    setStatus(error.message || "Could not create account.", true);
  }
}

function removeLocalAccount(email) {
  accountState.users = accountState.users.filter((user) => user.email !== email);
  if (!accountState.users.some((user) => user.email === accountState.currentEmail)) {
    accountState.currentEmail = OWNER_EMAIL;
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

  const deletingCurrentAccount = targetEmail === currentUserEmail;
  const confirmed = window.confirm(`Delete account ${targetEmail}? This removes the account login and saved workspace. This cannot be undone.`);
  if (!confirmed) return;

  const creatorCanDelete = isCreatorAuthenticated();
  const requestEmail = creatorCanDelete ? authUserEmail || currentUserEmail : currentUserEmail;
  const requestPassword = creatorCanDelete ? authUserPassword || currentUserPassword : currentUserPassword;

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
    serverUsers = serverUsers.filter((user) => normalizeServerUser(user).email !== targetEmail);
    closeAccountsModal();

    if (deletingCurrentAccount && adminViewingEmail && creatorCanDelete) {
      await returnToCreatorAccount();
      setStatus(`Deleted ${targetEmail}.`);
      return;
    }

    if (deletingCurrentAccount) {
      localStorage.removeItem(CURRENT_EMAIL_KEY);
      sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
      currentUserEmail = "";
      currentUserPassword = "";
      authUserEmail = "";
      authUserPassword = "";
      adminViewingEmail = "";
      currentUserRole = "user";
      cloudAvailable = false;
      await clearLocalWorkspace();
      renderLoginState();
      openLoginModal();
      setEntranceStatus("Account deleted.");
      setStatus("Account deleted.");
      return;
    }

    renderAccountsModal();
    setStatus(`Deleted ${targetEmail}.`);
  } catch (error) {
    setStatus(error.message || "Account deletion failed.", true);
  }
}

function deleteLocalAccountData(email) {
  accountState.users = accountState.users.filter((user) => user.email !== email);
  if (!accountState.users.some((user) => user.email === accountState.currentEmail)) {
    accountState.currentEmail = OWNER_EMAIL;
  }
  saveAccountState();
  const store = loadWorkspaceStore();
  delete store[email];
  saveWorkspaceStore(store);
}

async function enterUserAccount(email) {
  const targetEmail = normalizeEmail(email);
  if (!targetEmail || !isOwnerAccount()) return;
  const creatorEmail = authUserEmail || currentUserEmail;
  const creatorPassword = authUserPassword || currentUserPassword;
  if (normalizeEmail(creatorEmail) !== OWNER_EMAIL || !creatorPassword) {
    setStatus("Creator authentication is required to enter another account.", true);
    return;
  }

  try {
    await flushCloudStateSave().catch(() => {});
    cancelScheduledCloudStateSave();
    await saveCloudState();
  } catch {
    // Cloud save failures are surfaced elsewhere; local state is still saved.
  }

  closeAccountsModal();
  authUserEmail = creatorEmail;
  authUserPassword = creatorPassword;
  adminViewingEmail = targetEmail;
  currentUserEmail = targetEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = targetEmail === OWNER_EMAIL ? "creator" : "user";
  applyTemporarySkin("console");
  setStatus(`Entering ${targetEmail}...`);
  await clearLocalWorkspace();
  await restoreCloudState({ adminTargetEmail: targetEmail });
  await bootstrapImports();
  renderLoginState();
}

async function returnToCreatorAccount() {
  const creatorEmail = authUserEmail || OWNER_EMAIL;
  const creatorPassword = authUserPassword || currentUserPassword;
  await flushCloudStateSave().catch(() => {});
  cancelScheduledCloudStateSave();
  adminViewingEmail = "";
  currentUserEmail = creatorEmail;
  currentUserPassword = creatorPassword;
  currentUserRole = "creator";
  localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
  sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
  applySkin(loadSkin());
  setStatus("Returning to creator account...");
  await clearLocalWorkspace();
  await restoreCloudState();
  await bootstrapImports();
  renderLoginState();
}

async function clearLocalWorkspace() {
  selectedFiles = [];
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

function applyCurrentDayHighlight(sourceSettings = settings) {
  const borderColour = isHexColour(sourceSettings.currentDayBorderColor) ? sourceSettings.currentDayBorderColor : "#c44949";
  const backgroundColour = isHexColour(sourceSettings.currentDayBackgroundColor) ? sourceSettings.currentDayBackgroundColor : borderColour;
  const borderOpacity = normalizeOpacity(sourceSettings.currentDayBorderOpacity, 42);
  const backgroundOpacity = normalizeOpacity(sourceSettings.currentDayBackgroundOpacity, 8);
  const fillStyle = sourceSettings.currentDayFillStyle === "solid" ? "solid" : "gradient";
  const borderRgb = hexToRgb(borderColour);
  const backgroundRgb = hexToRgb(backgroundColour);
  document.documentElement.style.setProperty("--today-border-color", `rgba(${borderRgb.r}, ${borderRgb.g}, ${borderRgb.b}, ${borderOpacity})`);
  document.documentElement.style.setProperty("--today-fill-color", `rgba(${backgroundRgb.r}, ${backgroundRgb.g}, ${backgroundRgb.b}, ${backgroundOpacity})`);
  document.documentElement.style.setProperty(
    "--today-fill-surface",
    fillStyle === "solid"
      ? `rgba(${backgroundRgb.r}, ${backgroundRgb.g}, ${backgroundRgb.b}, ${backgroundOpacity})`
      : `linear-gradient(180deg, rgba(${backgroundRgb.r}, ${backgroundRgb.g}, ${backgroundRgb.b}, ${backgroundOpacity}), rgba(255, 255, 255, 0.82))`,
  );
  if (currentDayPreview) {
    currentDayPreview.style.borderColor = `rgba(${borderRgb.r}, ${borderRgb.g}, ${borderRgb.b}, ${borderOpacity})`;
  }
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

function loadSkin() {
  const stored = localStorage.getItem(SKIN_KEY);
  const email = loadCurrentUserEmail();
  if (!email) return "console";
  if (stored === "original" && email !== OWNER_EMAIL) return "console";
  if (stored === "original" || stored === "console") return stored;
  return "console";
}

function applySkin(skin) {
  currentSkin = skin === "console" ? "console" : "original";
  document.body.dataset.skin = currentSkin;
  localStorage.setItem(SKIN_KEY, currentSkin);
}

function applyTemporarySkin(skin) {
  currentSkin = skin === "original" ? "original" : "console";
  document.body.dataset.skin = currentSkin;
  skinSelect.value = currentSkin;
}

function syncSkinControl() {
  const authenticated = Boolean(currentUserEmail && currentUserPassword);
  const canChooseSkins = isOwnerAccount() && authenticated && !adminViewingEmail;
  skinControl.classList.toggle("hidden", !canChooseSkins);
  skinSelect.value = currentSkin;
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

function normalizeServerUser(value) {
  if (typeof value === "string") {
    return {
      email: value,
      realName: "",
      role: value === OWNER_EMAIL ? "owner" : "user",
      sites: [],
      claims: [],
    };
  }
  const email = normalizeEmail(value?.email);
  const role = value?.role || (email === OWNER_EMAIL ? "owner" : "user");
  return {
    email,
    realName: String(value?.realName || "").trim(),
    role: role === "creator" ? "owner" : role,
    sites: Array.isArray(value?.sites) ? value.sites : [],
    claims: sanitizeRosterClaims(value?.claims || []),
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
  applyTemporarySkin("console");
  loginEmail.value = prefillEmail;
  loginPassword.value = "";
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
  currentUserEmail = "";
  currentUserPassword = "";
  authUserEmail = "";
  authUserPassword = "";
  adminViewingEmail = "";
  currentUserRole = "user";
  cloudAvailable = false;
  currentRosterClaims = [];
  latestNameMatches = [];
  availableRosterDoctors = [];
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
  const viewingText = adminViewingEmail ? `Viewing as ${displayName}${currentUserEmail}` : `${displayName}${currentUserEmail}`;
  loginIdentity.textContent = loggedIn
    ? `${viewingText} · ${currentUserRole === "creator" ? "Creator" : "Standard account"}${cloudAvailable ? " · Cloud sync on" : " · Cloud sync required"}`
    : "";
  backToCreatorButton.classList.toggle("hidden", !adminViewingEmail || !isCreatorAuthenticated());
  syncAccountsButton();
  syncSkinControl();
}

async function loginWithEmail(email, password, options = {}) {
  const previousEmail = currentUserEmail;
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
    localStorage.setItem(CURRENT_EMAIL_KEY, currentUserEmail);
    sessionStorage.setItem(CURRENT_PASSWORD_KEY, currentUserPassword);
    applySkin(loadSkin());
    setStatus("Loading account workspace...");
    setEntranceStatus("Loading account workspace...");
    if (previousEmail !== currentUserEmail) {
      await clearLocalWorkspace();
    }
    await restoreCloudState(options);
    if (!currentUserEmail) return;
    await bootstrapImports();
    if (latestNameMatches.length) {
      const sites = [...new Set(latestNameMatches.map((claim) => claim.sourceType.toUpperCase()))].join(", ");
      setStatus(`Matched roster name${latestNameMatches.length === 1 ? "" : "s"} for ${sites || "uploaded rosters"}.`);
    }
    renderLoginState();
    closeLoginModal();
    setEntranceStatus("");
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
  } catch (error) {
    cancelScheduledCloudStateSave();
    const attemptedEmail = currentUserEmail;
    const message = error.message === "Cloud storage is not configured."
      ? serverStorageRequiredMessage()
      : normalizeAuthMessage(error.message || "Login failed.");
    cloudAvailable = false;
    localStorage.removeItem(CURRENT_EMAIL_KEY);
    sessionStorage.removeItem(CURRENT_PASSWORD_KEY);
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

async function applyCloudStateData(data) {
  cloudAvailable = data.cloudAvailable === true;
  currentUserRole = data.role || currentUserRole;
  currentRosterClaims = sanitizeRosterClaims(data.claims || []);
  latestNameMatches = sanitizeRosterClaims(data.nameMatches || []);
  availableRosterDoctors = sanitizeAvailableRosterDoctors(data.availableDoctors || []);
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
    restoredSessionState = null;
    await replaceStoredImports([]);
    clearWorkspaceStoreEntry(currentUserEmail);
    return;
  }
  const imports = await deserializeCloudImports(data.state.imports || []);
  selectedFiles = imports;
  await replaceStoredImports(imports);
  if (data.state.session) {
    restoredSessionState = data.state.session;
  }
  saveWorkspaceSnapshotForEmail(currentUserEmail, {
    fileRefs: imports.map((entry) => ({
      id: entry.id,
      name: entry.name,
      size: entry.size,
      lastModified: entry.lastModified,
      addedAt: entry.addedAt,
      sourceType: entry.sourceType,
    })),
    session: data.state.session && typeof data.state.session === "object" ? data.state.session : {},
  });
}

function serverStorageRequiredMessage() {
  return "Server storage is not configured. Add a Cloudflare KV namespace binding named ROSTER_STORE to the Pages project, redeploy, then log in again.";
}

function scheduleCloudStateSave() {
  if (!currentUserEmail) return;
  cancelScheduledCloudStateSave();
  const snapshot = snapshotCloudSavePayload();
  pendingCloudSaveSnapshot = snapshot;
  cloudSaveTimer = setTimeout(() => {
    const queued = pendingCloudSaveSnapshot;
    pendingCloudSaveSnapshot = null;
    saveCloudState(queued || snapshot).catch(() => {
      cloudAvailable = false;
      renderLoginState();
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

function snapshotCloudSavePayload() {
  return {
    accountEmail: currentUserEmail,
    requestEmail: adminViewingEmail ? authUserEmail : currentUserEmail,
    requestPassword: adminViewingEmail ? authUserPassword : currentUserPassword,
    targetEmail: adminViewingEmail ? currentUserEmail : "",
    imports: selectedFiles.map((entry) => ({ ...entry })),
    session: buildActiveSessionState(),
  };
}

async function saveCloudState(snapshot = null) {
  const payload = snapshot || snapshotCloudSavePayload();
  if (!payload.accountEmail || !payload.requestEmail || !payload.requestPassword || !cloudAvailable) return;
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
    }),
  });
  const data = await readJsonResponse(response, "Cloud save failed.");
  if (data.claims && payload.accountEmail === currentUserEmail) currentRosterClaims = sanitizeRosterClaims(data.claims);
  renderLoginState();
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
  } catch {
    // Keep the last available local list.
  }
}

async function buildCloudState(imports = selectedFiles, session = buildActiveSessionState()) {
  return {
    version: 1,
    imports: await serializeCloudImports(imports),
    session,
  };
}

function buildActiveSessionState() {
  return {
    doctorKey: doctorOptions.length > 1 ? doctorSelect.value : doctorOptions[0]?.key || "",
    settings: { ...settings },
    overrides: cleanOverrides(),
    customEvents: customEventsForActiveCalendar(),
    conflictSelections: { ...conflictSelections },
    hadPreview: Boolean(latestPreview),
    savedAt: new Date().toISOString(),
  };
}

async function serializeCloudImports(imports) {
  return await Promise.all(imports.map(async (entry) => ({
    id: entry.id,
    name: entry.name,
    size: entry.size,
    lastModified: entry.lastModified,
    addedAt: entry.addedAt,
    sourceType: entry.sourceType,
    doctors: parsedImportDoctors.get(entry.id) || [],
    type: entry.file?.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    dataUrl: await fileToDataUrl(entry.file),
  })));
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
    return JSON.parse(localStorage.getItem(ACCOUNT_WORKSPACES_KEY) || "{}");
  } catch {
    return {};
  }
}

function saveWorkspaceStore(store) {
  localStorage.setItem(ACCOUNT_WORKSPACES_KEY, JSON.stringify(store));
}

function saveWorkspaceSnapshotForEmail(email, snapshot) {
  if (!email) return;
  const store = loadWorkspaceStore();
  store[email] = snapshot;
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
    fileRefs: selectedFiles.map((entry) => ({
      id: entry.id,
      name: entry.name,
      size: entry.size,
      lastModified: entry.lastModified,
      addedAt: entry.addedAt,
      sourceType: entry.sourceType,
    })),
    session: {
      doctorKey: doctorOptions.length > 1 ? doctorSelect.value : doctorOptions[0]?.key || "",
      settings: { ...settings },
      overrides: cleanOverrides(),
      customEvents: customEventsForActiveCalendar(),
      conflictSelections: { ...conflictSelections },
      hadPreview: Boolean(latestPreview),
      savedAt: new Date().toISOString(),
    },
  };
}

function loadCurrentWorkspace() {
  if (!currentUserEmail) return null;
  const store = loadWorkspaceStore();
  return store[currentUserEmail] || null;
}

function saveCurrentWorkspace() {
  if (!currentUserEmail) return;
  saveWorkspaceSnapshotForEmail(currentUserEmail, currentWorkspaceSnapshot());
}

function loadCurrentSessionState() {
  const workspace = loadCurrentWorkspace();
  return workspace?.session || null;
}

function saveCurrentSessionState() {
  try {
    saveCurrentWorkspace();
    scheduleCloudStateSave();
  } catch {
    // Ignore persistence failures for session-only state.
  }
}

function sanitizeOverrideState(value) {
  if (!value || typeof value !== "object") return {};
  return JSON.parse(JSON.stringify(value));
}

function activeCalendarEmail() {
  return normalizeEmail(currentUserEmail);
}

function customEventsForActiveCalendar() {
  const ownerEmail = activeCalendarEmail();
  return sanitizeCustomEvents(customEvents).filter((item) => item.ownerEmail === ownerEmail);
}

function removeCustomEventForActiveCalendar(id) {
  const ownerEmail = activeCalendarEmail();
  customEvents = customEvents.filter((item) => !(item.id === id && normalizeEmail(item.ownerEmail) === ownerEmail));
}

function sanitizeCustomEvents(items, defaultOwnerEmail = "") {
  if (!Array.isArray(items)) return [];
  const fallbackOwnerEmail = normalizeEmail(defaultOwnerEmail);
  return items
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
  const db = await openImportsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readwrite");
    const store = tx.objectStore(IMPORT_STORE);
    for (const id of unreferenced) {
      store.delete(id);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Could not clean stored imports."));
  });
  db.close();
}

async function removeStoredImport(id) {
  selectedFiles = selectedFiles.filter((entry) => entry.id !== id);
  saveCurrentSessionState();
  try {
    await garbageCollectStoredImports();
  } catch {
    // Keep in-memory removal even if persistent storage is unavailable.
  }
  renderFilesList();
  scheduleCloudStateSave();
  if (!selectedFiles.length) {
    resetDerivedState();
    setStatus("Add a roster file to begin.");
    return;
  }
  await analyzeFiles();
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
      }
    }
    renderFilesList();
    if (selectedFiles.length) {
      await analyzeFiles();
    } else {
      renderClaimSection();
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

async function bootstrapApp() {
  renderLoginState();
  if (!currentUserEmail || !currentUserPassword) {
    openLoginModal();
    setStatus("Log in with an email address to load your roster workspace.");
    return;
  }
  await restoreCloudState();
  renderLoginState();
  await bootstrapImports();
}

function setStatus(message, isError = false) {
  status.textContent = message;
  status.dataset.error = isError ? "true" : "false";
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
      return `${fallbackMessage} Cloudflare exceeded its CPU or memory limit while parsing the roster. Try again after the latest deploy; if it persists, the app needs a higher Workers CPU limit or client-side parsing.`;
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
