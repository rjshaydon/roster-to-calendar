import { applyEventOverrides, customEventsToEvents, defaultSettings, exportIcs } from "../_lib/roster.js";
import { hasCalendarDb, loadAccountMirrorBySubscriptionToken, queryAccountCustomEvents, queryDoctorEvents } from "../_lib/d1-calendar.js";
import { normalizeEmail } from "./state.js";

export async function onRequestGet(context) {
  try {
    if (!hasCalendarDb(context.env)) return new Response("D1 database is not configured.", { status: 503 });
    const url = new URL(context.request.url);
    const token = String(url.searchParams.get("token") || "").trim();
    if (!token) return new Response("Import token is required.", { status: 400 });
    const record = await loadAccountMirrorBySubscriptionToken(context.env.ROSTER_DB, token);
    if (!record) return new Response("Import calendar was not found.", { status: 404 });
    const calendar = await buildOneOffImport(context.env.ROSTER_DB, record, url.searchParams);
    if (!calendar?.ics) return new Response("No calendar events are available for this import.", { status: 404 });
    return calendarResponse(calendar.ics, calendar.displayName);
  } catch (error) {
    return new Response(error.message || "Calendar import failed.", { status: 400 });
  }
}

async function buildOneOffImport(db, record, params) {
  const role = record?.role || "";
  const claims = sanitizeClaims(record.claims);
  const session = record?.state?.session && typeof record.state.session === "object" ? record.state.session : {};
  const doctorKeys = (role === "creator" || role === "owner")
    ? [String(session.doctorKey || "").trim()].filter(Boolean)
    : claims.map((claim) => claim.key);
  if (!doctorKeys.length) return null;
  const settings = { ...defaultSettings(), ...(session.settings || {}) };
  const range = normalizeImportRange(params);
  const queryOptions = range.mode === "range" && range.startDate
    ? { startDate: range.startDate, endDate: range.allFuture ? "9999-12-31" : range.endDate || range.startDate }
    : {};
  const hospitals = normalizeHospitals(params.get("hospitals"));
  const rosterEvents = applyEventOverrides(await queryDoctorEvents(db, doctorKeys, queryOptions), session.overrides || {});
  const d1CustomEvents = await queryAccountCustomEvents(db, record.email).catch(() => []);
  const customEvents = customEventsToEvents(latestCustomEventsByIdentity([
    ...sanitizeCustomEvents(session.customEvents, record.email),
    ...sanitizeCustomEvents(d1CustomEvents, record.email),
  ]), settings);
  const events = [...rosterEvents, ...customEvents]
    .filter((event) => eventInRange(event, range))
    .filter((event) => matchesHospitals(event, hospitals))
    .sort(compareEvents);
  if (!events.length) return null;
  return {
    ics: exportIcs(events, record.realName || claims[0]?.displayName || record.email),
    displayName: record.realName || claims[0]?.displayName || record.email,
  };
}

function calendarResponse(ics, displayName) {
  return new Response(ics, {
    status: 200,
    headers: {
      "Content-Type": "text/calendar; charset=utf-8",
      "Cache-Control": "private, no-store",
      "Content-Disposition": `inline; filename="${String(displayName || "roster").replace(/\//g, "-")} roster.ics"`,
    },
  });
}

function sanitizeClaims(value) {
  return (Array.isArray(value) ? value : [])
    .map((claim) => ({ key: String(claim?.key || "").trim(), displayName: String(claim?.displayName || claim?.key || "").trim() }))
    .filter((claim) => claim.key && claim.displayName);
}

function sanitizeCustomEvents(items, defaultOwnerEmail = "") {
  const ownerEmail = normalizeEmail(defaultOwnerEmail);
  return (Array.isArray(items) ? items : [])
    .filter((item) => item && item.id && item.title && item.startDate && item.endDate)
    .map((item) => ({
      id: String(item.id), ownerEmail: normalizeEmail(item.ownerEmail || ownerEmail), title: String(item.title),
      startDate: String(item.startDate), endDate: String(item.endDate), allDay: item.allDay === true,
      startTime: item.allDay ? "" : String(item.startTime || ""), endTime: item.allDay ? "" : String(item.endTime || ""),
      location: String(item.location || ""), include: item.include !== false,
    }))
    .filter((item) => item.ownerEmail === ownerEmail);
}

function latestCustomEventsByIdentity(events) {
  const byId = new Map();
  for (const event of events || []) { byId.delete(event.id); byId.set(event.id, event); }
  const byIdentity = new Map();
  for (const event of byId.values()) {
    const key = [normalizeEmail(event.ownerEmail), event.title, event.startDate, event.endDate, event.allDay ? "all-day" : `${event.startTime}|${event.endTime}`, event.location].join("|");
    byIdentity.delete(key);
    byIdentity.set(key, event);
  }
  return [...byIdentity.values()];
}

function normalizeImportRange(params) {
  const range = String(params.get("view") || "full") === "range";
  return {
    mode: range ? "range" : "full",
    startDate: String(params.get("startDate") || "").slice(0, 10),
    endDate: String(params.get("endDate") || "").slice(0, 10),
    allFuture: String(params.get("allFuture") || "true") !== "false",
  };
}

function normalizeHospitals(value) {
  return String(value || "").split(",").map((item) => item.trim().toUpperCase()).filter(Boolean);
}

function matchesHospitals(event, hospitals) {
  if (!hospitals.length) return true;
  const source = String(event?.source || "").toUpperCase();
  return source === "CUSTOM" || hospitals.includes(source);
}

function eventInRange(event, range) {
  if (range.mode !== "range" || !range.startDate) return true;
  const start = String(event?.start || "").slice(0, 10);
  const end = String(event?.end || event?.start || "").slice(0, 10);
  if (end < range.startDate) return false;
  if (!range.allFuture && range.endDate && start > range.endDate) return false;
  return true;
}

function compareEvents(left, right) {
  const leftDate = String(left.start || "").slice(0, 10);
  const rightDate = String(right.start || "").slice(0, 10);
  if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
  if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
  return String(left.title || "").localeCompare(String(right.title || ""));
}
