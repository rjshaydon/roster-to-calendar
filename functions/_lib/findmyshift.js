import * as XLSX from "xlsx";

const API_BASE = "https://www.findmyshift.com/api/1.4";

export async function findmyshiftLastModified(apiKey, teamId) {
  const payload = await findmyshiftRequest("teams/last-modified", apiKey, { teamId });
  const value = firstDateLikeValue(payload);
  if (!value) throw new Error("FindMyShift did not return a team modification time.");
  return value;
}

export async function findmyshiftRosterWorkbook(apiKey, teamId, range) {
  const report = await findmyshiftRequest("reports/shifts", apiKey, {
    teamId,
    from: range.from,
    to: range.to,
    publishedShifts: "yes",
    times: "yes",
    facilities: "yes",
    comments: "yes",
    groupByStaff: "yes",
  });
  const shifts = extractShiftRows(report);
  if (!shifts.length) throw new Error("FindMyShift returned no usable roster shifts for the configured date range.");
  return reportToDdhWorkbook(shifts);
}

async function findmyshiftRequest(path, apiKey, params) {
  const body = new URLSearchParams({ apiKey: String(apiKey), ...Object.fromEntries(Object.entries(params).map(([key, value]) => [key, String(value)])) });
  const response = await fetch(`${API_BASE}/${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body,
  });
  const text = await response.text();
  if (!response.ok) throw new Error(`FindMyShift ${path} returned HTTP ${response.status}.`);
  try {
    const parsed = JSON.parse(text);
    if (parsed?.error || parsed?.success === false) throw new Error(String(parsed.error?.message || parsed.error || "FindMyShift rejected the request."));
    return parsed;
  } catch (error) {
    if (error instanceof SyntaxError) throw new Error("FindMyShift returned an unexpected response instead of JSON.");
    throw error;
  }
}

function firstDateLikeValue(value) {
  if (typeof value === "string" && !Number.isNaN(Date.parse(value))) return new Date(value).toISOString();
  if (value && typeof value === "object") {
    for (const child of Array.isArray(value) ? value : Object.values(value)) {
      const found = firstDateLikeValue(child);
      if (found) return found;
    }
  }
  return "";
}

function extractShiftRows(report) {
  const rows = [];
  const visit = (value, inherited = {}) => {
    if (Array.isArray(value)) return value.forEach((item) => visit(item, inherited));
    if (!value || typeof value !== "object") return;
    const context = {
      name: nameFrom(value) || inherited.name || "",
      seniority: String(value.jobTitle || value.department || inherited.seniority || "Unknown"),
    };
    const date = dateFrom(value);
    const label = String(value.shiftName || value.name || value.title || value.label || value.role || "").trim();
    const start = timeFrom(value, ["startTime", "start", "timeStart", "from"]);
    const end = timeFrom(value, ["endTime", "end", "timeEnd", "to"]);
    if (context.name && date && (label || (start && end))) {
      rows.push({ name: context.name, seniority: context.seniority, date, label: label || "Shift", start, end });
    }
    for (const child of Object.values(value)) if (child && typeof child === "object") visit(child, context);
  };
  visit(report);
  const unique = new Map();
  for (const row of rows) unique.set(`${row.name}|${row.date}|${row.label}|${row.start}|${row.end}`, row);
  return [...unique.values()];
}

function nameFrom(value) {
  const direct = value.staffName || value.employeeName || value.displayName || value.fullName || "";
  if (direct) return String(direct).trim();
  const first = value.firstName || value.staffFirstName || "";
  const last = value.lastName || value.staffLastName || "";
  return `${first} ${last}`.trim();
}

function dateFrom(value) {
  for (const key of ["date", "shiftDate", "startDate", "startsAt", "startDateTime"]) {
    const raw = value[key];
    const date = raw && new Date(raw);
    if (date && !Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  return "";
}

function timeFrom(value, keys) {
  for (const key of keys) {
    const raw = String(value[key] || "").trim();
    const match = raw.match(/(?:T|^)(\d{1,2}):(\d{2})/);
    if (match) return `${match[1].padStart(2, "0")}:${match[2]}`;
  }
  return "";
}

function reportToDdhWorkbook(rows) {
  const byWeek = new Map();
  for (const row of rows) {
    const monday = mondayFor(row.date);
    if (!byWeek.has(monday)) byWeek.set(monday, []);
    byWeek.get(monday).push(row);
  }
  const output = [];
  for (const [monday, weekRows] of [...byWeek.entries()].sort()) {
    const days = Array.from({ length: 7 }, (_, index) => addDays(monday, index));
    output.push(["", ...days.map(findmyshiftDateLabel)]);
    const bySeniority = new Map();
    for (const row of weekRows) {
      if (!bySeniority.has(row.seniority)) bySeniority.set(row.seniority, new Map());
      const people = bySeniority.get(row.seniority);
      if (!people.has(row.name)) people.set(row.name, []);
      people.get(row.name).push(row);
    }
    for (const [seniority, people] of bySeniority) {
      output.push([seniority]);
      for (const [name, personRows] of people) {
        const labels = days.map((day) => personRows.filter((row) => row.date === day).map((row) => row.label).join(" / "));
        const times = days.map((day) => personRows.filter((row) => row.date === day).map((row) => row.start && row.end ? `${row.start}-${row.end}` : "").filter(Boolean).join(" / "));
        output.push([name, ...labels]);
        if (times.some(Boolean)) output.push(["", ...times]);
      }
    }
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(output), "FindMyShift");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function mondayFor(value) {
  const date = new Date(`${value}T00:00:00Z`);
  const day = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - day);
  return date.toISOString().slice(0, 10);
}

function addDays(value, count) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + count);
  return date.toISOString().slice(0, 10);
}

function findmyshiftDateLabel(value) {
  const date = new Date(`${value}T00:00:00Z`);
  const weekdays = ["Sun.", "Mon.", "Tue.", "Wed.", "Thu.", "Fri.", "Sat."];
  const months = ["Jan.", "Feb.", "Mar.", "Apr.", "May.", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."];
  return `${weekdays[date.getUTCDay()]} ${months[date.getUTCMonth()]} ${String(date.getUTCDate()).padStart(2, "0")}, ${date.getUTCFullYear()}`;
}
