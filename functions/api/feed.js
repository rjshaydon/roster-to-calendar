import { applyEventOverrides, customEventsToEvents, defaultSettings, exportIcs } from "../_lib/roster.js";
import { hasCalendarDb, queryDoctorEvents } from "../_lib/d1-calendar.js";
import { accountSnapshotOwner, loadAccountBySubscriptionToken, loadSnapshotRecord, normalizeEmail } from "./state.js";

export async function onRequestGet(context) {
  try {
    if (!context.env.ROSTER_STORE) {
      return new Response("Cloud storage is not configured.", { status: 503 });
    }

    const url = new URL(context.request.url);
    const token = String(url.searchParams.get("token") || "").trim();
    const view = String(url.searchParams.get("view") || "full").trim() === "range" ? "range" : "full";
    if (!token) {
      return new Response("Subscription token is required.", { status: 400 });
    }

    const record = await loadAccountBySubscriptionToken(context.env.ROSTER_STORE, token);
    if (!record) {
      return new Response("Subscription calendar was not found.", { status: 404 });
    }

    const d1Feed = await buildD1SubscriptionFeed(context.env.ROSTER_DB, record, view);
    if (d1Feed?.ics) {
      return calendarResponse(d1Feed.ics, d1Feed.displayName || record.realName || record.email);
    }

    const owner = accountSnapshotOwner(record.email, record.role || "");
    const snapshot = await loadSnapshotRecord(context.env.ROSTER_STORE, owner.ownerType, owner.ownerId);
    const artifact = snapshot?.subscriptionFeeds?.[view] || null;
    if (!artifact?.ics) {
      return new Response("No stored subscription calendar is available for this view.", { status: 404 });
    }

    return calendarResponse(artifact.ics, artifact.doctorDisplay || record.realName || record.email);
  } catch (error) {
    return new Response(error.message || "Subscription feed failed.", { status: 400 });
  }
}

async function buildD1SubscriptionFeed(db, record, view) {
  if (!hasCalendarDb({ ROSTER_DB: db })) return null;
  const role = record?.role || "";
  if (role === "creator" || role === "owner") return null;
  const claims = sanitizeClaims(record.claims);
  if (!claims.length) return null;
  const session = record?.state?.session && typeof record.state.session === "object" ? record.state.session : {};
  const settings = {
    ...defaultSettings(),
    ...(session.settings || {}),
  };
  const range = view === "range" ? normalizeExportRange(session.exportRange) : { mode: "full" };
  const queryOptions = range.mode === "range" && range.startDate
    ? { startDate: range.startDate, endDate: range.allFuture ? "9999-12-31" : range.endDate || range.startDate }
    : {};
  const rosterEvents = applyEventOverrides(await queryDoctorEvents(db, claims.map((claim) => claim.key), queryOptions), session.overrides || {});
  const customEvents = customEventsToEvents(sanitizeCustomEvents(session.customEvents, record.email), settings);
  const events = [...rosterEvents, ...customEvents]
    .filter((event) => range.mode !== "range" || eventInRange(event, range))
    .sort(compareEvents);
  if (!events.length) return null;
  const displayName = record.realName || claims[0]?.displayName || record.email;
  return {
    ics: exportIcs(events, displayName),
    displayName,
  };
}

function calendarResponse(ics, displayName) {
  return new Response(ics, {
    status: 200,
    headers: {
      "Content-Type": "text/calendar; charset=utf-8",
      "Cache-Control": "private, max-age=300",
      "Content-Disposition": `inline; filename="${String(displayName || "roster").replace(/\//g, "-")} subscription.ics"`,
    },
  });
}

function sanitizeClaims(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((claim) => ({
      key: String(claim?.key || "").trim(),
      displayName: String(claim?.displayName || claim?.key || "").trim(),
      sourceType: String(claim?.sourceType || "").trim().toLowerCase(),
    }))
    .filter((claim) => claim.key && claim.displayName);
}

function sanitizeCustomEvents(items, defaultOwnerEmail = "") {
  const ownerEmail = normalizeEmail(defaultOwnerEmail);
  return (Array.isArray(items) ? items : [])
    .filter((item) => item && item.id && item.title && item.startDate && item.endDate)
    .map((item) => ({
      id: String(item.id),
      ownerEmail: normalizeEmail(item.ownerEmail || ownerEmail),
      title: String(item.title),
      startDate: String(item.startDate),
      endDate: String(item.endDate),
      allDay: item.allDay === true,
      startTime: item.allDay ? "" : String(item.startTime || ""),
      endTime: item.allDay ? "" : String(item.endTime || ""),
      location: String(item.location || ""),
      include: item.include !== false,
    }))
    .filter((item) => item.ownerEmail === ownerEmail);
}

function normalizeExportRange(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    mode: "range",
    startDate: String(input.startDate || "").slice(0, 10),
    endDate: String(input.endDate || "").slice(0, 10),
    allFuture: input.allFuture !== false,
  };
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
