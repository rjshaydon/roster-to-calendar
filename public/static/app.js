const form = document.querySelector("#roster-form");
const fileInput = document.querySelector("#rosterFiles");
const dropZone = document.querySelector("#dropZone");
const chooseFilesButton = document.querySelector("#chooseFilesButton");
const fileSummary = document.querySelector("#fileSummary");
const exportButton = document.querySelector("#exportButton");
const previewButton = document.querySelector("#previewButton");
const doctorSection = document.querySelector("#doctorSection");
const doctorSelect = document.querySelector("#doctorSelect");
const doctorName = document.querySelector("#doctorName");
const preview = document.querySelector("#preview");
const status = document.querySelector("#status");

const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

let doctorOptions = [];
let detectedSources = {};
let selectedFiles = [];

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
  clearPreview();
  previewButton.disabled = !selectedDoctor();
  exportButton.disabled = !selectedDoctor();
});

previewButton.addEventListener("click", async () => {
  const doctor = selectedDoctor();
  if (!doctor) {
    setStatus("Choose a doctor before previewing.", true);
    return;
  }
  await updatePreview();
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
  const workbookFiles = files;
  const map = new Map(selectedFiles.map((file) => [fileFingerprint(file), file]));
  for (const file of workbookFiles) {
    map.set(fileFingerprint(file), file);
  }
  if (map.size > 2) {
    setStatus("You can add up to two roster files. Remove one before adding another.", true);
    return [];
  }
  return workbookFiles;
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
    renderFileSummary();
    renderDoctorState();
  } catch (error) {
    resetDerivedState();
    renderFileSummary();
    setStatus(error.message, true);
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
  doctorSelect.disabled = false;

  if (!doctorOptions.length) {
    doctorSection.classList.add("hidden");
    setStatus("No consultant names could be matched from the uploaded roster files.", true);
    return;
  }

  doctorSection.classList.remove("hidden");
  if (doctorOptions.length === 1) {
    doctorName.textContent = doctorOptions[0].displayName;
    doctorName.classList.remove("hidden");
    previewButton.disabled = false;
    exportButton.disabled = false;
    setStatus("Consultant detected.");
    return;
  }

  for (const doctor of doctorOptions) {
    const option = document.createElement("option");
    option.value = doctor.key;
    option.textContent = doctor.displayName;
    doctorSelect.append(option);
  }
  doctorSelect.classList.remove("hidden");
  previewButton.disabled = false;
  exportButton.disabled = false;
  setStatus("Choose a doctor, then preview or export.");
}

async function updatePreview() {
  const doctor = selectedDoctor();
  if (!doctor) return;
  setStatus("Building preview...");
  try {
    const data = await postForm("/api/preview", doctor);
    renderPreviewGrid(doctor, data);
    setStatus("Preview loaded.");
  } catch (error) {
    clearPreview();
    setStatus(error.message, true);
  }
}

function renderPreviewGrid(doctor, data) {
  const events = data.events || [];
  const days = buildPreviewDays(events);
  const weeks = chunkWeeks(days);
  const headerCells = DAY_NAMES.map((day) => `<div class="preview-day-name">${day}</div>`).join("");
  const weekRows = weeks.map((week, index) => {
    const cells = week.map((day) => renderDayCell(day)).join("");
    const monday = week[0]?.date;
    return `
      <div class="preview-week-label">
        <strong>Week ${index + 1}</strong>
        <span>starting</span>
        <time datetime="${formatDateKey(monday)}">${formatLongDate(monday)}</time>
      </div>
      ${cells}
    `;
  }).join("");

  preview.innerHTML = `
    <div class="preview-head">
      <strong>${doctor.displayName}</strong>
      <span>${data.count} events</span>
      <span>${data.date_range}</span>
    </div>
    <div class="preview-grid">
      <div class="preview-week-label preview-week-label-head">Week</div>
      ${headerCells}
      ${weekRows}
    </div>
  `;
  preview.classList.remove("hidden");
}

function buildPreviewDays(events) {
  if (!events.length) return [];
  const eventMap = new Map();
  let firstDay = null;
  let lastDay = null;

  for (const event of events) {
    const startDate = parseDateOnly(event.start);
    const endDate = parseDateOnly(event.end);
    const inclusiveEnd = event.allDay ? addDays(endDate, -1) : startDate;
    if (!firstDay || startDate < firstDay) firstDay = startDate;
    if (!lastDay || inclusiveEnd > lastDay) lastDay = inclusiveEnd;

    let cursor = new Date(startDate);
    while (cursor <= inclusiveEnd) {
      const key = formatDateKey(cursor);
      if (!eventMap.has(key)) eventMap.set(key, []);
      const label = renderEventLabel(event, cursor, startDate);
      if (label) {
        eventMap.get(key).push(label);
      }
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
      entries: eventMap.get(key) || [],
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
  const items = day.entries.length
    ? day.entries.map((entry) => `<li class="preview-item">${entry}</li>`).join("")
    : `<li class="preview-item preview-item-empty"> </li>`;
  return `
    <div class="preview-cell">
      <div class="preview-date">${day.date.getDate()}</div>
      <ul class="preview-list">${items}</ul>
    </div>
  `;
}

function renderEventLabel(event, cursor, startDate) {
  if (event.allDay) {
    return event.title;
  }
  if (formatDateKey(cursor) !== formatDateKey(startDate)) {
    return "";
  }
  const start = new Date(event.start);
  const end = new Date(event.end);
  return `${formatTime(start)}-${formatTime(end)} ${event.title}`;
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
  return body;
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
  doctorSelect.innerHTML = "";
  doctorName.textContent = "";
  doctorName.classList.add("hidden");
  doctorSelect.classList.add("hidden");
  doctorSection.classList.add("hidden");
  previewButton.disabled = true;
  exportButton.disabled = true;
  clearPreview();
}

function clearPreview() {
  preview.innerHTML = "";
  preview.classList.add("hidden");
}

function fileFingerprint(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
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

function formatTime(date) {
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function formatLongDate(date) {
  return date.toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
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
