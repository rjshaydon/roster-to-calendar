const form = document.querySelector("#roster-form");
const fileInput = document.querySelector("#rosterFiles");
const dropZone = document.querySelector("#dropZone");
const chooseFilesButton = document.querySelector("#chooseFilesButton");
const fileSummary = document.querySelector("#fileSummary");
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
const overview = document.querySelector("#overview");
const overviewSources = document.querySelector("#overviewSources");
const overviewCount = document.querySelector("#overviewCount");
const overviewRange = document.querySelector("#overviewRange");
const overviewParsed = document.querySelector("#overviewParsed");
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
  mergeFiles(accepted);
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
  mergeFiles(accepted);
  await analyzeFiles();
});

fileSummary.addEventListener("click", async (event) => {
  const removeButton = event.target.closest("[data-remove-file]");
  if (!removeButton) return;
  const fingerprint = removeButton.dataset.removeFile;
  selectedFiles = selectedFiles.filter((file) => fileFingerprint(file) !== fingerprint);
  resetDerivedState();
  renderFileSummary();
  if (!selectedFiles.length) {
    setStatus("Add one or two roster files to begin.");
    return;
  }
  await analyzeFiles();
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
  if (!titleInput) return;
  const id = titleInput.dataset.overrideTitle;
  ensureOverride(id).title = titleInput.value;
  rebuildClientPreview();
  setStatus("Mapping override updated.");
});

reviewModalBody.addEventListener("change", (event) => {
  const includeInput = event.target.closest("[data-override-include]");
  if (!includeInput) return;
  const id = includeInput.dataset.overrideInclude;
  ensureOverride(id).include = includeInput.checked;
  rebuildClientPreview();
  setStatus("Inclusion override updated.");
});

previewButton.addEventListener("click", () => updatePreview());
mobilePreviewButton.addEventListener("click", () => updatePreview());
mobileExportButton.addEventListener("click", () => form.requestSubmit());
preview.addEventListener("click", (event) => {
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
issuesList.addEventListener("click", (event) => {
  const card = event.target.closest("[data-review-id]");
  if (!card) return;
  openReviewModal(card.dataset.reviewId);
});
preview.addEventListener("contextmenu", (event) => {
  const chip = event.target.closest("[data-review-id]");
  const cell = event.target.closest("[data-add-date]");
  if (!chip && !cell) return;
  event.preventDefault();
  const items = [];
  if (chip) {
    items.push({ label: "Copy Event", action: () => copyPreviewEvent(chip.dataset.reviewId) });
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
  }
});
document.addEventListener("click", (event) => {
  if (!event.target.closest("#contextMenu")) {
    closeContextMenu();
  }
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

function mergeFiles(files) {
  const map = new Map(selectedFiles.map((file) => [fileFingerprint(file), file]));
  for (const file of files) {
    map.set(fileFingerprint(file), file);
  }
  selectedFiles = [...map.values()];
  renderFileSummary();
}

function validateIncomingFiles(files) {
  if (files.some((file) => !file.name.match(/\.(xlsx|xlsm|xltx|xltm)$/i))) {
    setStatus("Only Excel roster files in .xlsx, .xlsm, .xltx, or .xltm format are supported.", true);
    return [];
  }
  const map = new Map(selectedFiles.map((file) => [fileFingerprint(file), file]));
  for (const file of files) {
    map.set(fileFingerprint(file), file);
  }
  if (map.size > 2) {
    setStatus("You can add up to two roster files. Remove one before adding another.", true);
    return [];
  }
  return files;
}

async function analyzeFiles() {
  if (!selectedFiles.length) {
    setStatus("Add one or two roster files to begin.");
    return;
  }
  resetDerivedState();
  setStatus("Detecting roster sources and consultants...");
  try {
    const data = await postForm("/api/analyze");
    doctorOptions = data.doctors || [];
    detectedSources = data.sources || {};
    settings = { ...defaultSettings(), ...(data.settings || {}) };
    overrides = {};
    renderSettings();
    renderFileSummary();
    renderDoctorState();
  } catch (error) {
    resetDerivedState();
    renderFileSummary();
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

function renderFileSummary() {
  if (!selectedFiles.length) {
    fileSummary.classList.add("hidden");
    fileSummary.innerHTML = "";
    return;
  }

  const sourceCards = [];
  if (detectedSources.mmc) sourceCards.push({ label: "MMC", name: detectedSources.mmc });
  if (detectedSources.ddh) sourceCards.push({ label: "DDH", name: detectedSources.ddh });

  const pendingFiles = selectedFiles.filter((file) => file.name !== detectedSources.mmc && file.name !== detectedSources.ddh);

  fileSummary.innerHTML = [
    ...sourceCards.map((item) => renderBadge(item.label, item.name, selectedFiles.find((file) => file.name === item.name))),
    ...pendingFiles.map((file) => renderBadge("Uploaded", file.name, file)),
  ].join("");
  fileSummary.classList.remove("hidden");
}

function renderBadge(label, value, file) {
  const removeMarkup = file
    ? `
      <button
        type="button"
        class="file-remove"
        data-remove-file="${fileFingerprint(file)}"
        aria-label="Remove ${value}"
        title="Remove ${value}"
      >
        ×
      </button>
    `
    : "";
  return `
    <article class="file-pill">
      ${removeMarkup}
      <span>${label}</span>
      <strong>${value}</strong>
    </article>
  `;
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
    const data = await postForm("/api/preview", doctor);
    latestPreview = data;
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
  renderOverview(view);
  renderPreviewGrid(doctor, view);
  renderIssues(view.issues || []);
}

function buildClientPreviewData(baseData) {
  const baseEvents = new Map((baseData.events || []).map((event) => [event.id, { ...event }]));
  const events = [];

  for (const item of reviewIndex.values()) {
    const event = baseEvents.get(item.id);
    if (!event) continue;
    const override = overrides[item.id] || {};
    const include = typeof override.include === "boolean" ? override.include : item.include;
    if (!include) continue;
    events.push({
      ...event,
      title: (override.title || "").trim() || item.suggestedTitle || event.title,
      start: override.start || event.start,
      end: override.end || event.end,
      allDay: typeof override.allDay === "boolean" ? override.allDay : event.allDay,
      location: override.location || event.location,
      timeLabel: override.start && override.end ? summarizeEventTimes(override.start, override.end, typeof override.allDay === "boolean" ? override.allDay : event.allDay) : event.timeLabel,
      isEditedImport: Boolean(override.start || override.end || override.location),
    });
  }

  for (const event of customEvents) {
    events.push(customEventToPreviewEvent(event));
  }

  events.sort(comparePreviewEvents);
  return {
    ...baseData,
    events,
    count: events.length,
    date_range: events.length ? summarizeEvents(events) : "No events found",
    issues: (baseData.issues || []).filter((issue) => {
      const override = overrides[issue.id] || {};
      const reviewItem = reviewIndex.get(issue.id);
      const include = typeof override.include === "boolean" ? override.include : reviewItem?.include ?? true;
      return include;
    }),
  };
}

function renderOverview(data) {
  const files = [data.sources?.mmc, data.sources?.ddh].filter(Boolean);
  overviewSources.innerHTML = files.length
    ? files.map((file) => `<span>${escapeHtml(file)}</span>`).join("")
    : "<span>Single source</span>";
  overviewCount.textContent = `${data.count}`;
  overviewRange.textContent = data.date_range;
  overviewParsed.textContent = formatTimestamp(data.lastParsed);
  overview.classList.remove("hidden");
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
  if (!events.length) {
    preview.innerHTML = `
      <div class="preview-head">
        <strong>${escapeHtml(doctor.displayName)}</strong>
        <span>0 events</span>
        <span>${escapeHtml(data.date_range)}</span>
      </div>
      <div class="preview-empty">No events match the current settings.</div>
    `;
    preview.classList.remove("hidden");
    previewSection.classList.remove("hidden");
    return;
  }
  const days = buildPreviewDays(events);
  const weeks = chunkWeeks(days);
  const termSections = buildTermSections(weeks);

  preview.innerHTML = `
    <div class="preview-head">
      <strong>${escapeHtml(doctor.displayName)}</strong>
      <span>${data.count} events</span>
      <span>${escapeHtml(data.date_range)}</span>
    </div>
    ${termSections}
  `;
  preview.classList.remove("hidden");
  previewSection.classList.remove("hidden");
}

function buildPreviewDays(events) {
  if (!events.length) return [];
  const eventMap = new Map();
  let firstDay = null;
  let lastDay = null;

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
  return `<button type="button" class="preview-chip" data-review-id="${event.id}" draggable="true">${lines.join("")}${metaMarkup}</button>`;
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
  for (const file of selectedFiles) {
    body.append("rosterFiles", file);
  }
  if (doctor) {
    body.append("doctorKey", doctor.key);
    body.append("doctorDisplay", doctor.displayName);
  }
  body.append("settings", JSON.stringify(settings));
  body.append("overrides", JSON.stringify(cleanOverrides()));
  body.append("customEvents", JSON.stringify(customEvents));
  return body;
}

function cleanOverrides() {
  const next = {};
  for (const [id, value] of Object.entries(overrides)) {
    const title = (value.title || "").trim();
    const include = value.include;
    const start = value.start || "";
    const end = value.end || "";
    const location = value.location || "";
    const allDay = value.allDay;
    if (!title && typeof include !== "boolean" && !start && !end && !location && typeof allDay !== "boolean") continue;
    next[id] = {};
    if (title) next[id].title = title;
    if (typeof include === "boolean") next[id].include = include;
    if (start) next[id].start = start;
    if (end) next[id].end = end;
    if (location) next[id].location = location;
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
    const override = ensureOverride(id);
    override.start = updated.start;
    override.end = updated.end;
    override.allDay = updated.allDay;
    if (updated.location) {
      override.location = updated.location;
    }
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
  overview.classList.add("hidden");
  issuesPanel.classList.add("hidden");
  previewSection.classList.add("hidden");
  preview.innerHTML = "";
  preview.classList.add("hidden");
  issuesList.innerHTML = "";
  closeReviewModal();
  closeCustomEventModal();
  closeContextMenu();
}

function ensureOverride(id) {
  if (!overrides[id]) overrides[id] = {};
  return overrides[id];
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
  const status = item.status === "unknown" ? "Unknown" : "Review";
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
  const overrideValue = escapeHtml((overrides[id]?.title ?? item.overrideTitle ?? ""));
  const includeValue = typeof overrides[id]?.include === "boolean" ? overrides[id].include : item.include;
  const warnings = item.warnings.length
    ? `<ul class="review-warnings">${item.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("")}</ul>`
    : "";
  const badge = item.status === "ok" ? "" : `<span class="review-badge review-badge-${item.status}">${item.status}</span>`;

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
      </div>
      <div class="review-meta">
        <span>Suggested title: ${escapeHtml(item.suggestedTitle || "No normalized result")}</span>
        ${item.timeLabel ? `<span>Times: ${escapeHtml(item.timeLabel)}</span>` : ""}
        ${item.location ? `<span>Location: ${escapeHtml(item.location)}</span>` : ""}
      </div>
      ${warnings}
    </article>
  `;
  reviewModal.classList.remove("hidden");
  reviewModal.setAttribute("aria-hidden", "false");
}

function closeReviewModal() {
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
  const options = [];
  if (detectedSources.mmc) options.push({ value: "mmc", label: "MMC Car Park" });
  if (detectedSources.ddh) options.push({ value: "ddh", label: "DDH Car Park" });
  options.push({ value: "offsite", label: "Off-site" });
  options.push({ value: "custom", label: "Custom location" });
  customEventLocationMode.innerHTML = options.map((option) => `<option value="${option.value}">${option.label}</option>`).join("");
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
