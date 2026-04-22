const form = document.querySelector("#roster-form");
const fileInput = document.querySelector("#rosterFiles");
const dropZone = document.querySelector("#dropZone");
const chooseFilesButton = document.querySelector("#chooseFilesButton");
const filesButton = document.querySelector("#filesButton");
const accountsButton = document.querySelector("#accountsButton");
const filesModal = document.querySelector("#filesModal");
const filesList = document.querySelector("#filesList");
const filesCloseButton = document.querySelector("#filesCloseButton");
const accountsModal = document.querySelector("#accountsModal");
const accountsCloseButton = document.querySelector("#accountsCloseButton");
const previewButton = document.querySelector("#previewButton");
const exportButton = document.querySelector("#exportButton");
const mobilePreviewButton = document.querySelector("#mobilePreviewButton");
const mobileExportButton = document.querySelector("#mobileExportButton");
const mobileSettingsButton = document.querySelector("#mobileSettingsButton");
const doctorSection = document.querySelector("#doctorSection");
const doctorSelect = document.querySelector("#doctorSelect");
const doctorName = document.querySelector("#doctorName");
const controlBar = document.querySelector("#controlBar");
const settingsToggle = document.querySelector("#settingsToggle");
const settingsPanel = document.querySelector("#settingsPanel");
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
const contextMenu = document.querySelector("#contextMenu");

const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const SETTINGS_FIELDS = [
  "showSourcePrefix",
  "showAmPm",
  "showTimes",
  "showLocations",
  "showRawValues",
  "showNormalizedTitles",
  "includeLocations",
  "includeAnnualLeave",
  "includeConferenceLeave",
  "includePublicHoliday",
  "includeSickLeave",
  "hospitalFilter",
  "dateFrom",
  "dateTo",
];

let doctorOptions = [];
let detectedSources = {};
let selectedFiles = [];
let settings = defaultSettings();
let overrides = {};
let latestPreview = null;
let reviewIndex = new Map();
let customEvents = [];
let currentPreviewEvents = new Map();
let dragEventId = null;
let copiedEvent = null;
let previewGesture = null;
let suppressPreviewClickUntil = 0;
let openReviewId = "";
let conflictSelections = {};

const settingsInputs = Object.fromEntries(
  SETTINGS_FIELDS.map((id) => [id, document.querySelector(`#${id}`)]),
);

chooseFilesButton.addEventListener("click", (event) => {
  event.stopPropagation();
  fileInput.click();
});

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
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
  await removeStoredImport(removeButton.dataset.removeImport);
});
accountsButton.addEventListener("click", () => {
  accountsModal.classList.remove("hidden");
  accountsModal.setAttribute("aria-hidden", "false");
});
accountsCloseButton.addEventListener("click", closeAccountsModal);
accountsModal.addEventListener("click", (event) => {
  if (event.target.matches("[data-close-accounts]")) closeAccountsModal();
});

doctorSelect.addEventListener("change", () => {
  clearPreviewData();
  syncActionState();
});

settingsToggle.addEventListener("click", () => {
  settingsPanel.classList.toggle("hidden");
});
mobileSettingsButton.addEventListener("click", () => {
  settingsPanel.classList.toggle("hidden");
});

for (const [key, input] of Object.entries(settingsInputs)) {
  input.addEventListener("change", () => {
    settings[key] = input.type === "checkbox" ? input.checked : input.value;
    if (!settings.showNormalizedTitles && !settings.showRawValues) {
      settings.showNormalizedTitles = true;
      settingsInputs.showNormalizedTitles.checked = true;
    }
    setStatus("Settings updated. Use Preview to refresh the grid.");
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
  if (!resetButton) return;
  resetImportedEvent(resetButton.dataset.overrideReset);
});

previewButton.addEventListener("click", () => updatePreview());
mobilePreviewButton.addEventListener("click", () => updatePreview());
mobileExportButton.addEventListener("click", () => form.requestSubmit());
preview.addEventListener("click", (event) => {
  if (Date.now() < suppressPreviewClickUntil) return;
  closeContextMenu();
  const chip = event.target.closest("[data-review-id]");
  if (chip) {
    openReviewModal(chip.dataset.reviewId);
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
  }
});
document.addEventListener("click", (event) => {
  if (!event.target.closest("#contextMenu")) {
    closeContextMenu();
  }
});
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
  customEvents = customEvents.filter((item) => item.id !== id);
  closeCustomEventModal();
  rebuildClientPreview();
  setStatus("Manual event removed.");
});
customEventForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const entry = readCustomEventForm();
  if (!entry) return;
  const index = customEvents.findIndex((item) => item.id === entry.id);
  if (index >= 0) {
    customEvents[index] = entry;
    setStatus("Manual event updated.");
  } else {
    customEvents.push(entry);
    setStatus("Manual event added.");
  }
  closeCustomEventModal();
  rebuildClientPreview();
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
    const response = await fetch("/api/export", {
      method: "POST",
      body: createFormData(doctor),
    });
    const payload = await response.blob();
    if (!response.ok) {
      const text = await payload.text();
      throw new Error(parseError(text));
    }
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
    showLocations: false,
    showRawValues: false,
    showNormalizedTitles: true,
    includeLocations: true,
    includeAnnualLeave: true,
    includeConferenceLeave: true,
    includePublicHoliday: true,
    includeSickLeave: true,
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
  };
}

async function mergeFiles(files) {
  const existing = new Set(selectedFiles.map((entry) => fileFingerprint(entry.file)));
  for (const file of files) {
    const fingerprint = fileFingerprint(file);
    if (existing.has(fingerprint)) continue;
    const entry = {
      id: `import-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
      file,
      name: file.name,
      size: file.size,
      lastModified: file.lastModified,
      addedAt: new Date().toISOString(),
      sourceType: "pending",
    };
    selectedFiles.push(entry);
    await saveStoredImport(entry);
  }
  selectedFiles.sort((left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || left.name.localeCompare(right.name));
  renderFilesList();
}

function validateIncomingFiles(files) {
  if (files.some((file) => !file.name.match(/\.(xlsx|xlsm|xltx|xltm)$/i))) {
    setStatus("Only Excel roster files in .xlsx, .xlsm, .xltx, or .xltm format are supported.", true);
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
  setStatus("Detecting roster sources and consultants...");
  try {
    const data = await postForm("/api/analyze");
    doctorOptions = data.doctors || [];
    detectedSources = summarizeDetectedSources(data.imports || []);
    selectedFiles = selectedFiles.map((entry) => {
      const serverEntry = (data.imports || []).find((item) => item.id === entry.id);
      return serverEntry ? { ...entry, sourceType: serverEntry.sourceType } : entry;
    });
    settings = { ...defaultSettings(), ...settings, ...(data.settings || {}) };
    conflictSelections = loadConflictSelections();
    renderSettings();
    renderFilesList();
    renderDoctorState();
  } catch (error) {
    doctorOptions = [];
    detectedSources = {};
    clearPreviewData();
    renderFilesList();
    setStatus(error.message, true);
  }
}

function renderSettings() {
  for (const [key, input] of Object.entries(settingsInputs)) {
    if (!input) continue;
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else {
      input.value = settings[key] || "";
    }
  }
}

function renderFilesList() {
  if (!filesList) return;
  if (!selectedFiles.length) {
    filesList.innerHTML = `<article class="issue-card"><strong>No files imported yet.</strong><p>Add rosters and they will stay here until removed.</p></article>`;
    return;
  }

  filesList.innerHTML = selectedFiles.map((entry) => `
    <article class="issue-card">
      <div>
        <strong>${escapeHtml(entry.name)}</strong>
        <p>${escapeHtml(String(entry.sourceType || "").toUpperCase())} · Imported ${escapeHtml(formatTimestamp(entry.addedAt))}</p>
      </div>
      <button type="button" class="button button-secondary" data-remove-import="${entry.id}">Remove</button>
    </article>
  `).join("");
}

function renderDoctorState() {
  doctorSelect.innerHTML = "";
  doctorName.textContent = "";
  doctorName.classList.add("hidden");
  doctorSelect.classList.add("hidden");
  doctorSection.classList.add("hidden");
  controlBar.classList.add("hidden");
  mobileActionBar.classList.add("hidden");

  if (!doctorOptions.length) {
    setStatus("No consultant names could be matched from the uploaded roster files.", true);
    return;
  }

  doctorSection.classList.remove("hidden");
  controlBar.classList.remove("hidden");
  mobileActionBar.classList.remove("hidden");
  settingsPanel.classList.add("hidden");

  if (doctorOptions.length === 1) {
    doctorName.textContent = doctorOptions[0].displayName;
    doctorName.classList.remove("hidden");
    setStatus("Preview when ready.");
  } else {
    for (const doctor of doctorOptions) {
      const option = document.createElement("option");
      option.value = doctor.key;
      option.textContent = doctor.displayName;
      doctorSelect.append(option);
    }
    doctorSelect.classList.remove("hidden");
    setStatus("Choose a doctor, then preview or export.");
  }

  syncActionState();
}

async function updatePreview() {
  const doctor = selectedDoctor();
  if (!doctor) {
    setStatus("Choose a doctor before previewing.", true);
    return;
  }
  setStatus("Building preview...");
  try {
    const data = await postPreviewForm("/api/preview", doctor);
    latestPreview = data;
    if (!settings.dateFrom || !settings.dateTo) {
      const range = deriveRangeBounds(data.events || []);
      if (!settings.dateFrom) settings.dateFrom = range.start;
      if (!settings.dateTo) settings.dateTo = range.end;
      renderSettings();
    }
    indexReviewItems(data.review || []);
    rebuildClientPreview();
    setStatus("Preview loaded.");
  } catch (error) {
    clearPreviewData();
    setStatus(error.message, true);
  }
}

function rebuildClientPreview() {
  if (!latestPreview) return;
  const doctor = selectedDoctor();
  if (!doctor) return;
  const view = buildClientPreviewData(latestPreview);
  renderConflicts(view.conflicts || []);
  renderPreviewGrid(doctor, view);
  renderIssues(view.issues || []);
}

function buildClientPreviewData(baseData) {
  const baseEvents = new Map((baseData.events || []).map((event) => [event.id, { ...event }]));
  const range = deriveRangeBounds(baseData.events || []);
  const events = [];
  const deletedItems = [];

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

  for (const event of customEvents) {
    events.push(customEventToPreviewEvent(event));
  }

  events.sort(comparePreviewEvents);
  const previewStart = settings.dateFrom || range.start;
  const previewEnd = settings.dateTo || range.end;
  return {
    ...baseData,
    events,
    count: events.length,
    date_range: formatPreviewRange(previewStart, previewEnd) || (events.length ? summarizeEvents(events) : "No events found"),
    previewStart,
    previewEnd,
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
  if (!days.length) {
    preview.innerHTML = `
      <div class="preview-head">
        <strong>${escapeHtml(doctor.displayName)}</strong>
        <span>${data.count} events</span>
        <span>${escapeHtml(data.date_range)}</span>
        <span>Last import: ${escapeHtml(formatTimestamp(data.lastImport))}</span>
      </div>
      <div class="preview-empty">No events match the current settings.</div>
    `;
    preview.classList.remove("hidden");
    previewSection.classList.remove("hidden");
    return;
  }
  const weeks = chunkWeeks(days);
  const termSections = buildTermSections(weeks);

  preview.innerHTML = `
    <div class="preview-head">
      <strong>${escapeHtml(doctor.displayName)}</strong>
      <span>${data.count} events</span>
      <span>${escapeHtml(data.date_range)}</span>
      <span>Last import: ${escapeHtml(formatTimestamp(data.lastImport))}</span>
    </div>
    ${termSections}
  `;
  preview.classList.remove("hidden");
  previewSection.classList.remove("hidden");
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
  return `
    <div class="preview-cell" data-add-date="${formatDateKey(day.date)}">
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
  if (settings.showLocations && event.location) meta.push(event.location);
  const metaMarkup = meta.length ? `<span class="preview-chip-meta">${escapeHtml(meta.join(" · "))}</span>` : "";
  return `<button type="button" class="preview-chip" data-review-id="${event.id}">${lines.join("")}${metaMarkup}</button>`;
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
      <div class="preview-term-title">${escapeHtml(section.label)}</div>
      <div class="preview-grid">
        <div class="preview-week-label preview-week-label-head">Week</div>
        ${headerCells}
        ${bodyRows.join("")}
      </div>
    </section>
  `;
}

function syncActionState() {
  const ready = Boolean(selectedDoctor());
  previewButton.disabled = !ready;
  exportButton.disabled = !ready;
  mobilePreviewButton.disabled = !ready;
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
  }
  body.append("settings", JSON.stringify(settings));
  body.append("overrides", JSON.stringify(cleanOverrides()));
  body.append("customEvents", JSON.stringify(customEvents));
  body.append("conflictSelections", JSON.stringify(conflictSelections));
  return body;
}

function createPreviewFormData(doctor = null) {
  const body = new FormData();
  for (const entry of selectedFiles) {
    body.append("rosterFiles", entry.file);
    body.append("rosterFileId", entry.id);
    body.append("rosterFileAddedAt", entry.addedAt || "");
  }
  if (doctor) {
    body.append("doctorKey", doctor.key);
    body.append("doctorDisplay", doctor.displayName);
  }
  body.append("settings", JSON.stringify(settings));
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
    customEvents = customEvents.map((item) => item.id === id ? previewEventToCustomEvent(updated, item) : item);
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
    source: "Custom",
    isEditedImport: false,
  }, targetDate);
  customEvents.push(previewEventToCustomEvent(shifted));
  closeContextMenu();
  rebuildClientPreview();
  setStatus("Event pasted.");
}

function deletePreviewEvent(id) {
  const event = currentPreviewEvents.get(id);
  if (!event) return;
  if (event.source === "Custom") {
    customEvents = customEvents.filter((item) => item.id !== id);
    if (openReviewId === id) closeReviewModal();
  } else {
    syncImportedOverride(id, { include: false });
    if (openReviewId === id) closeReviewModal();
  }
  closeContextMenu();
  rebuildClientPreview();
  setStatus("Event deleted.");
}

function resetImportedEvent(id) {
  if (!hasImportedOverride(id)) return;
  delete overrides[id];
  closeContextMenu();
  rebuildClientPreview();
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
    moved: false,
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
  const canTimeShift = !gesture.sourceEvent.allDay && hoverDay === gesture.originDay && Math.abs(dy) >= Math.abs(dx);

  if (canTimeShift) {
    const slotOffset = calculateTimeShiftSlots(dy);
    gesture.slotOffset = slotOffset;
    clearDropTargets();
    applyPreviewGestureTime(gesture);
    return;
  }

  if (gesture.slotOffset !== 0) {
    gesture.slotOffset = 0;
    restorePreviewGestureMeta(gesture);
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
  if (gesture.slotOffset !== 0 && hoverDay === gesture.originDay) {
    commitPreviewGestureTime(gesture);
    suppressPreviewClickUntil = Date.now() + 200;
  } else if (hoverDay && hoverDay !== gesture.originDay) {
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
  const slots = Math.floor(distance / 18 + (distance / 90) ** 2 * 4);
  return direction * slots;
}

function applyPreviewGestureTime(gesture) {
  const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.slotOffset * 15);
  setPreviewChipMeta(gesture.chip, shifted, true);
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
  const shifted = shiftTimedEventByMinutes(gesture.sourceEvent, gesture.slotOffset * 15);
  if (shifted.source === "Custom") {
    customEvents = customEvents.map((item) => item.id === shifted.id ? previewEventToCustomEvent(shifted, item) : item);
  } else {
    syncImportedOverride(shifted.id, {
      start: shifted.start,
      end: shifted.end,
      allDay: shifted.allDay,
    });
  }
  rebuildClientPreview();
  setStatus("Event time updated.");
}

function setPreviewChipMeta(chip, event, forceTime = false) {
  const metaParts = [];
  if ((!event.allDay && settings.showTimes) || forceTime) metaParts.push(event.timeLabel);
  if (settings.showLocations && event.location) metaParts.push(event.location);
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

async function postForm(url, doctor = null) {
  const response = await fetch(url, {
    method: "POST",
    body: createFormData(doctor),
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(parseError(text));
  }
  return JSON.parse(text);
}

async function postPreviewForm(url, doctor = null) {
  const response = await fetch(url, {
    method: "POST",
    body: createPreviewFormData(doctor),
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(parseError(text));
  }
  return JSON.parse(text);
}

function selectedDoctor() {
  if (!doctorOptions.length) return null;
  if (doctorOptions.length === 1) return doctorOptions[0];
  return doctorOptions.find((doctor) => doctor.key === doctorSelect.value) || doctorOptions[0];
}

function resetDerivedState() {
  doctorOptions = [];
  detectedSources = {};
  overrides = {};
  customEvents = [];
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
  clearPreviewData();
}

function clearPreviewData() {
  latestPreview = null;
  reviewIndex = new Map();
  currentPreviewEvents = new Map();
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

function formatPreviewRange(start, end) {
  if (!start || !end) return "";
  return `${start} to ${end}`;
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
  const year = date.getFullYear();
  const candidates = [
    buildAustralianTerm(year, 1, 1),
    buildAustralianTerm(year, 2, 4),
    buildAustralianTerm(year, 3, 7),
    buildAustralianTerm(year, 4, 10),
    buildAustralianTerm(year - 1, 4, 10),
  ];
  const match = candidates.find((term) => date >= term.start && date < term.end);
  return match || { label: "Schedule" };
}

function buildAustralianTerm(year, termNumber, startMonthIndex) {
  const start = firstMondayOfMonth(year, startMonthIndex);
  const end = addDays(start, 91);
  return {
    label: `Term ${termNumber}`,
    start,
    end,
  };
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

function openReviewModal(id) {
  const item = reviewIndex.get(id);
  if (!item) {
    const customEvent = customEvents.find((entry) => entry.id === id);
    if (customEvent) {
      openCustomEventModal(customEvent);
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
            <input type="time" value="${startTime || "09:00"}" data-override-start-time="${item.id}">
          </label>
          <label class="field">
            <span>End time</span>
            <input type="time" value="${endTime || "17:00"}" data-override-end-time="${item.id}">
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
  customEventStartTime.value = event?.startTime || "09:00";
  customEventEndTime.value = event?.endTime || "10:00";
  const preset = detectLocationPreset(event?.location || "");
  customEventLocationMode.value = preset.mode;
  customEventCustomLocation.value = preset.customValue;
  customEventCustomLocationField.classList.toggle("hidden", preset.mode !== "custom");
  customEventTimeFields.classList.toggle("hidden", customEventAllDay.checked);
  customEventDeleteButton.classList.toggle("hidden", !event);
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
  if (location === "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia") return { mode: "mmc", customValue: "" };
  if (location === "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia") return { mode: "ddh", customValue: "" };
  return { mode: "custom", customValue: location };
}

function readCustomEventForm() {
  const title = customEventTitle.value.trim();
  const startDate = customEventStartDate.value;
  const endDate = customEventEndDate.value || startDate;
  const allDay = customEventAllDay.checked;
  const startTime = customEventStartTime.value;
  const endTime = customEventEndTime.value;
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
  if (customEventLocationMode.value === "mmc") return "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
  if (customEventLocationMode.value === "ddh") return "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";
  if (customEventLocationMode.value === "custom") return customEventCustomLocation.value.trim();
  return "";
}

function applyImportedEventFormState(id) {
  const baseEvent = getBaseImportedEvent(id);
  const startDate = reviewModalBody.querySelector(`[data-override-start-date="${id}"]`)?.value || baseEvent?.start?.slice(0, 10) || "";
  const rawEndDate = reviewModalBody.querySelector(`[data-override-end-date="${id}"]`)?.value || startDate;
  const endDateInput = rawEndDate < startDate ? startDate : rawEndDate;
  const allDay = reviewModalBody.querySelector(`[data-override-all-day="${id}"]`)?.checked ?? baseEvent?.allDay ?? false;
  const startTime = reviewModalBody.querySelector(`[data-override-start-time="${id}"]`)?.value || extractTimePortion(baseEvent?.start || "") || "09:00";
  const endTime = reviewModalBody.querySelector(`[data-override-end-time="${id}"]`)?.value || extractTimePortion(baseEvent?.end || "") || "17:00";
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
  if (mode === "mmc") return "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
  if (mode === "ddh") return "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";
  if (mode === "custom") {
    return reviewModalBody.querySelector(`[data-override-custom-location="${id}"]`)?.value.trim() || "";
  }
  return "";
}

function extractTimePortion(value) {
  const match = String(value).match(/T(\d{2}:\d{2})/);
  return match ? match[1] : "";
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

const DB_NAME = "roster-converter";
const DB_VERSION = 1;
const IMPORT_STORE = "imports";
const CONFLICT_SELECTIONS_KEY = "roster-conflict-selections";

async function openImportsDb() {
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

async function loadStoredImports() {
  if (!("indexedDB" in window)) return [];
  const db = await openImportsDb();
  const records = await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readonly");
    const request = tx.objectStore(IMPORT_STORE).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error || new Error("Could not load imports."));
  });
  db.close();
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

async function removeStoredImport(id) {
  selectedFiles = selectedFiles.filter((entry) => entry.id !== id);
  const db = await openImportsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMPORT_STORE, "readwrite");
    tx.objectStore(IMPORT_STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Could not remove import."));
  });
  db.close();
  renderFilesList();
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
    selectedFiles = await loadStoredImports();
    renderFilesList();
    if (selectedFiles.length) {
      await analyzeFiles();
    } else {
      setStatus("Add a roster file to begin.");
    }
  } catch (error) {
    setStatus(error.message || "Could not load saved imports.", true);
  }
}

function setStatus(message, isError = false) {
  status.textContent = message;
  status.dataset.error = isError ? "true" : "false";
}

function parseError(text) {
  try {
    return JSON.parse(text).error || "Request failed.";
  } catch {
    return "Request failed.";
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

bootstrapImports();
