import * as XLSX from "xlsx";

const API_BASE = "https://www.findmyshift.com/api/1.4";

export function findmyshiftConfiguredRosterRange(env = {}, now = new Date()) {
  const year = now.getUTCFullYear();
  // A report returns only shifts actually available in FindMyShift. Ask for a
  // generously complete current window rather than coupling automation to the
  // smaller, read-only diagnostic range. Optional explicit bounds are useful
  // for a provider that has a known archival or future publication limit.
  return {
    from: validDateKey(env.FINDMYSHIFT_FROM) || `${year}-01-01`,
    to: validDateKey(env.FINDMYSHIFT_TO) || `${year + 1}-12-31`,
  };
}

export async function findmyshiftLastModified(apiKey, teamId) {
  const payload = await findmyshiftRequest("teams/last-modified", apiKey, { teamId });
  const value = firstDateLikeValue(payload);
  if (!value) throw new Error("FindMyShift did not return a team modification time.");
  return value;
}

export async function findmyshiftRosterWorkbook(apiKey, teamId, range) {
  // FindMyShift allows only one concurrent request per API key.  Keep these
  // dependent lookups serial even though they are otherwise independent.
  const report = await findmyshiftShiftReport(apiKey, teamId, range);
  const staff = await findmyshiftStaffList(apiKey, teamId);
  const facilities = await findmyshiftFacilityList(apiKey, teamId);
  const shifts = extractShiftRows(report, { staff, facilities });
  if (!shifts.length) throw new Error("FindMyShift returned no usable roster shifts for the configured date range.");
  return findmyshiftRowsWorkbook(shifts);
}

export async function findmyshiftShiftReport(apiKey, teamId, range) {
  // This is deliberately an unfiltered team report. In particular, do not send
  // `filters` or a facility id: doctors can be rostered across all facilities.
  return findmyshiftRequest("reports/shifts", apiKey, {
    teamId,
    from: range.from,
    to: range.to,
    publishedShifts: "yes",
    timesheetData: "no",
    daysToInclude: "0123456",
    comments: "yes",
    times: "yes",
    facilities: "yes",
    groupingInterval: "day",
    groupByStaff: "yes",
  });
}

export async function findmyshiftStaffList(apiKey, teamId) {
  return findmyshiftRequest("staff/list", apiKey, { teamId });
}

export async function findmyshiftFacilityList(apiKey, teamId) {
  return findmyshiftRequest("facilities/list", apiKey, { teamId });
}

export function findmyshiftReportDiagnostics(report, range = {}, options = {}) {
  const fieldCounts = {
    staff: 0,
    seniority: 0,
    date: 0,
    label: 0,
    timed: 0,
    allDay: 0,
    facility: 0,
    comment: 0,
  };
  const sourceDates = new Set();
  let objectCount = 0;
  let arrayCount = 0;
  let candidateCount = 0;
  const visit = (value, depth = 0) => {
    if (depth > 12 || value === null || value === undefined) return;
    if (Array.isArray(value)) {
      arrayCount += 1;
      for (const item of value) visit(item, depth + 1);
      return;
    }
    if (typeof value !== "object") return;
    objectCount += 1;
    const candidate = shiftCandidate(value);
    if (candidate) {
      candidateCount += 1;
      for (const field of candidateFields(value)) fieldCounts[field] += 1;
      const date = dateFrom(value);
      if (date) sourceDates.add(date);
    }
    for (const child of Object.values(value)) visit(child, depth + 1);
  };
  visit(report);
  const shifts = extractShiftRows(report, options);
  const dates = [...sourceDates].sort();
  return {
    responseFormat: Array.isArray(report) ? "array" : report && typeof report === "object" ? "object" : typeof report,
    objectCount,
    arrayCount,
    shiftCandidateCount: candidateCount,
    recognisedShiftCount: shifts.length,
    staffCount: new Set(shifts.map((shift) => shift.name)).size,
    dateRange: dates.length ? { from: dates[0], to: dates.at(-1) } : null,
    requestedDateRange: { from: String(range.from || ""), to: String(range.to || ""), label: String(range.label || "") },
    recognisedFields: fieldCounts,
    flatShiftRows: findmyshiftFlatShiftRowDiagnostics(report),
    responseShape: findmyshiftReportShape(report),
    facilityFilterApplied: false,
  };
}

export function findmyshiftStaffDiagnostics(staff) {
  const entries = Array.isArray(staff) ? staff : [];
  let id = 0;
  let name = 0;
  let seniority = 0;
  for (const entry of entries) {
    if (!entry || typeof entry !== "object") continue;
    if (entry.staffId || entry.id) id += 1;
    if (nameFrom(entry)) name += 1;
    if (entry.jobTitle || entry.department) seniority += 1;
  }
  return {
    responseFormat: Array.isArray(staff) ? "array" : staff && typeof staff === "object" ? "object" : typeof staff,
    itemCount: entries.length,
    recognised: { id, name, seniority },
  };
}

export function findmyshiftFacilityDiagnostics(facilities) {
  const entries = Array.isArray(facilities) ? facilities : [];
  let id = 0;
  let name = 0;
  for (const entry of entries) {
    if (!entry || typeof entry !== "object") continue;
    if (entry.facilityId || entry.id) id += 1;
    if (entry.facilityName || entry.name || entry.title) name += 1;
  }
  return {
    responseFormat: Array.isArray(facilities) ? "array" : facilities && typeof facilities === "object" ? "object" : typeof facilities,
    itemCount: entries.length,
    recognised: { id, name },
  };
}

function findmyshiftFlatShiftRowDiagnostics(report) {
  const rows = Array.isArray(report) ? report.filter((entry) => entry && typeof entry === "object" && !Array.isArray(entry)) : [];
  let staffId = 0;
  let date = 0;
  let shiftText = 0;
  let timeRangeLike = 0;
  let newline = 0;
  for (const row of rows) {
    if (row.staffId !== null && row.staffId !== undefined && String(row.staffId).trim()) staffId += 1;
    if (dateFrom(row)) date += 1;
    const shift = typeof row.shift === "string" ? row.shift.trim() : "";
    if (!shift) continue;
    shiftText += 1;
    if (/\b\d{1,2}(?::\d{2})?\s*(?:am|pm)?\s*(?:-|–|—|to)\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)?\b/i.test(shift)) timeRangeLike += 1;
    if (/\r?\n/.test(shift)) newline += 1;
  }
  return {
    rowCount: rows.length,
    recognised: { staffId, date, shiftText },
    shiftTextCharacteristics: { timeRangeLike, newline },
  };
}

// The report diagnostic must be useful for adapting the parser, while never
// returning names, shift text, comments, facility names, IDs or arbitrary keys
// (the API may use staff names as object keys).  This intentionally reports
// only counts, JSON value types and a small allow-list of ordinary schema keys.
const SAFE_REPORT_SCHEMA_KEYS = new Set([
  "date", "dates", "day", "days", "staff", "staffid", "staffname", "staffmember",
  "employee", "employeename", "shifts", "shift", "shiftid", "shiftname", "shiftlabel",
  "start", "end", "starttime", "endtime", "time", "times", "duration", "hours",
  "facility", "facilities", "facilityname", "location", "locations", "comment", "comments",
  "note", "notes", "role", "title", "jobtitle", "department", "seniority", "leave",
  "timeoff", "absence", "published", "id", "name", "data", "items", "rows", "entries",
]);

function findmyshiftReportShape(report) {
  const levels = Array.from({ length: 4 }, () => ({
    arrays: 0,
    objects: 0,
    scalarValues: { string: 0, number: 0, boolean: 0, null: 0 },
    knownKeys: {},
    unknownKeyCount: 0,
  }));
  const visit = (value, depth = 0) => {
    if (depth >= levels.length) return;
    const level = levels[depth];
    if (value === null) {
      level.scalarValues.null += 1;
      return;
    }
    if (Array.isArray(value)) {
      level.arrays += 1;
      for (const child of value) visit(child, depth + 1);
      return;
    }
    if (typeof value === "object") {
      level.objects += 1;
      for (const [key, child] of Object.entries(value)) {
        const normalised = String(key).replace(/[^a-z0-9]/gi, "").toLowerCase();
        if (SAFE_REPORT_SCHEMA_KEYS.has(normalised)) {
          level.knownKeys[normalised] = Number(level.knownKeys[normalised] || 0) + 1;
        } else {
          level.unknownKeyCount += 1;
        }
        visit(child, depth + 1);
      }
      return;
    }
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      level.scalarValues[typeof value] += 1;
    }
  };
  visit(report);
  return levels.map((level, depth) => ({
    depth,
    ...level,
  })).filter((level) => level.arrays || level.objects || Object.values(level.scalarValues).some(Boolean));
}

async function findmyshiftRequest(path, apiKey, params) {
  const body = new URLSearchParams({ apiKey: String(apiKey), ...Object.fromEntries(Object.entries(params).map(([key, value]) => [key, String(value)])) });
  const response = await fetch(`${API_BASE}/${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body,
  });
  const text = await response.text();
  if (!response.ok) {
    throw findmyshiftRequestError(`FindMyShift ${path} returned HTTP ${response.status}.`, {
      code: "http-error",
      status: response.status,
      contentType: response.headers.get("content-type") || "",
      responseBytes: new TextEncoder().encode(text).byteLength,
    });
  }
  try {
    const parsed = JSON.parse(text);
    if (parsed?.error || parsed?.success === false) {
      throw findmyshiftRequestError("FindMyShift rejected the request.", {
        code: "provider-error",
        status: response.status,
        contentType: response.headers.get("content-type") || "",
        responseBytes: new TextEncoder().encode(text).byteLength,
      });
    }
    return parsed;
  } catch (error) {
    if (error instanceof SyntaxError) {
      throw findmyshiftRequestError("FindMyShift returned an unexpected response instead of JSON.", {
        code: "unexpected-format",
        status: response.status,
        contentType: response.headers.get("content-type") || "",
        responseBytes: new TextEncoder().encode(text).byteLength,
      });
    }
    throw error;
  }
}

function findmyshiftRequestError(message, details = {}) {
  const error = new Error(message);
  error.findmyshiftDiagnostic = {
    code: String(details.code || "request-failed"),
    status: Number(details.status || 0),
    contentType: String(details.contentType || "").split(";")[0].trim().slice(0, 80),
    responseBytes: Number(details.responseBytes || 0),
  };
  return error;
}

function validDateKey(value) {
  const date = String(value || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(date) && !Number.isNaN(Date.parse(`${date}T00:00:00Z`)) ? date : "";
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

export function extractShiftRows(report, options = {}) {
  const flatRows = extractFindmyshiftFlatRows(report, options);
  if (flatRows !== null) return dedupeShiftRows(flatRows);
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
      rows.push({
        name: context.name,
        seniority: context.seniority,
        date,
        label: label || "Shift",
        start,
        end,
        facility: String(value.facilityName || value.facility || value.location || "").trim(),
        comment: String(value.comment || value.comments || value.notes || "").trim(),
      });
    }
    for (const child of Object.values(value)) if (child && typeof child === "object") visit(child, context);
  };
  visit(report);
  return dedupeShiftRows(rows);
}

// `reports/shifts` returns a flat array.  A row has staffId, firstName,
// lastName, date, shift, facilityId, payrollId and occurrences.  The API uses
// the `shift` text for either a named/all-day entry or a plain time range.
// Keep the branch explicit instead of sending it through the generic walker:
// a person's `firstName` would otherwise be mistaken for a shift label.
function extractFindmyshiftFlatRows(report, options) {
  if (!Array.isArray(report)) return null;
  const reportRows = report.filter((entry) => entry && typeof entry === "object" && !Array.isArray(entry));
  if (!reportRows.length || !reportRows.every((entry) => Object.hasOwn(entry, "staffId") && Object.hasOwn(entry, "date") && Object.hasOwn(entry, "shift"))) return null;
  const staffById = indexFindmyshiftStaff(options.staff);
  const facilitiesById = indexFindmyshiftFacilities(options.facilities);
  const rows = [];
  for (const entry of reportRows) {
    const staffId = String(entry.staffId || "").trim();
    const person = staffById.get(staffId) || {};
    const name = nameFrom(entry) || nameFrom(person);
    const date = dateFrom(entry);
    const shiftText = String(entry.shift || "").trim();
    if (!name || !date || !shiftText) continue;
    const time = timeRangeFromShiftText(shiftText);
    const facilityId = String(entry.facilityId || "").trim();
    rows.push({
      sourceStaffId: staffId,
      name,
      seniority: String(person.jobTitle || person.department || entry.jobTitle || entry.department || "Unknown").trim() || "Unknown",
      date,
      label: time ? (shiftText === time.source ? "Shift" : shiftText) : shiftText,
      start: time?.start || "",
      end: time?.end || "",
      facility: facilitiesById.get(facilityId) || String(entry.facilityName || entry.facility || entry.location || facilityId || "").trim(),
      comment: String(entry.comment || entry.comments || entry.notes || "").trim(),
    });
  }
  return rows;
}

function indexFindmyshiftStaff(staff) {
  const index = new Map();
  for (const entry of Array.isArray(staff) ? staff : []) {
    if (!entry || typeof entry !== "object") continue;
    const id = String(entry.staffId || entry.id || "").trim();
    if (id) index.set(id, entry);
  }
  return index;
}

function indexFindmyshiftFacilities(facilities) {
  const index = new Map();
  for (const entry of Array.isArray(facilities) ? facilities : []) {
    if (!entry || typeof entry !== "object") continue;
    const id = String(entry.facilityId || entry.id || "").trim();
    const name = String(entry.facilityName || entry.name || entry.title || "").trim();
    if (id && name) index.set(id, name);
  }
  return index;
}

function timeRangeFromShiftText(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}:\d{2})\s*(?:-|–|—|to)\s*(\d{1,2}:\d{2})$/i);
  if (!match) return null;
  return {
    source: String(value || "").trim(),
    start: normaliseTwentyFourHourTime(match[1]),
    end: normaliseTwentyFourHourTime(match[2]),
  };
}

function normaliseTwentyFourHourTime(value) {
  const [hour, minute] = String(value || "").split(":");
  return `${String(hour || "").padStart(2, "0")}:${minute}`;
}

function dedupeShiftRows(rows) {
  const unique = new Map();
  for (const row of rows) {
    const key = [row.sourceStaffId || row.name, row.date, row.label, row.start, row.end, row.facility, row.comment].join("|");
    unique.set(key, row);
  }
  return [...unique.values()];
}

function shiftCandidate(value) {
  if (!value || typeof value !== "object") return false;
  return Boolean(dateFrom(value) || timeFrom(value, ["startTime", "start", "timeStart", "from"]) || timeFrom(value, ["endTime", "end", "timeEnd", "to"]));
}

function candidateFields(value) {
  const result = [];
  if (nameFrom(value)) result.push("staff");
  if (value.jobTitle || value.department) result.push("seniority");
  if (dateFrom(value)) result.push("date");
  const shiftText = String(value.shift || value.shiftName || value.name || value.title || value.label || value.role || "").trim();
  const directTimed = timeFrom(value, ["startTime", "start", "timeStart", "from"]) && timeFrom(value, ["endTime", "end", "timeEnd", "to"]);
  const shiftTime = timeRangeFromShiftText(shiftText);
  if (shiftText) result.push("label");
  if (directTimed || shiftTime) result.push("timed");
  if (shiftText && !directTimed && !shiftTime) result.push("allDay");
  if (value.facilityId || value.facilityName || value.facility || value.location) result.push("facility");
  if (value.comment || value.comments || value.notes) result.push("comment");
  return result;
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

export function findmyshiftRowsWorkbook(rows) {
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
        const dayRows = (day) => personRows.filter((row) => row.date === day);
        const labels = days.map((day) => dayRows(day).map((row) => row.label === "Shift" && row.start && row.end ? `${row.start}-${row.end}` : row.label).join(" / "));
        const times = days.map((day) => dayRows(day)
          .filter((row) => row.label !== "Shift")
          .map((row) => row.start && row.end ? `${row.start}-${row.end}` : "")
          .filter(Boolean)
          .join(" / "));
        output.push([name, ...labels]);
        if (times.some(Boolean)) output.push(["", ...times]);
      }
    }
  }
  const workbook = XLSX.utils.book_new();
  // This marker makes the existing DDH parser treat single-day leave from the
  // API as a single day, rather than the legacy spreadsheet convention of one
  // leave label standing for a whole week.
  output.unshift(["FindMyShift roster format"]);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(output), "FindMyShift");
  const details = [["Staff ID", "Staff name", "Seniority/job title", "Date", "Shift label", "Start", "End", "Facility", "Comment"]];
  for (const row of rows) {
    details.push([row.sourceStaffId || "", row.name, row.seniority, row.date, row.label, row.start, row.end, row.facility, row.comment]);
  }
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(details), "FindMyShift details");
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
