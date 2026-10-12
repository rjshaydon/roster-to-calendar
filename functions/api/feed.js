import {subscriptionEvidence,readSubscriptionArtifact,queueSubscriptionArtifact} from '../_lib/subscription-artifacts.js';
import { applyEventOverrides, customEventsToEvents, defaultSettings, exportIcs } from "../_lib/roster.js";
import { applyAccountHospitalLocations, dedupeEventsByIdentity, hasCalendarDb, loadAccountHospitalLocations, loadAccountMirrorBySubscriptionToken, queryAccountCustomEvents, queryDoctorEvents, loadSubscriptionCalendarInputs } from "../_lib/d1-calendar.js";
import { normalizeEmail } from "../_lib/email.js";
import { accountIdentityAliases, expandApprovedIdentityAliases } from '../_lib/doctor-identity.js';

export async function onRequestGet(context) {
  try {
    if (!hasCalendarDb(context.env)) {
      return new Response("D1 database is not configured.", { status: 503 });
    }

    const url = new URL(context.request.url);
    const token = String(url.searchParams.get("token") || "").trim();
    const view = String(url.searchParams.get("view") || "full").trim() === "range" ? "range" : "full";
    if (!token) {
      return new Response("Subscription token is required.", { status: 400 });
    }

    if(context.env.SUBSCRIPTION_ARTIFACT_MODE==='ready'&&context.env.IDENTITY_REVIEW_ENABLED==='true'&&context.env.ROSTER_FILES){
      try {
      const evidence=await subscriptionEvidence(context.env.ROSTER_DB,token,view,context.env.IDENTITY_REVIEW_ENABLED==='true');
      if(!evidence)return new Response('Subscription calendar was not found.',{status:404});
      const artifact=await readSubscriptionArtifact(context.env.ROSTER_FILES,evidence);
      if(artifact)return calendarResponse(artifact.body,artifact.displayName);
      // Preserve complete dynamic calendars during initial build or a deferred
      // rebuild. Never serve a previous identity or an empty stand-in calendar.
      const queued=queueSubscriptionArtifact(context.env.ROSTER_FILES,token,view,evidence);
      if(context.waitUntil)context.waitUntil(queued.catch(()=>undefined));else await queued;
      }catch(error){
        if(error.code!=='SUBSCRIPTION_ARTIFACT_UNSUPPORTED')throw error;
        // Oversized evidence keeps the existing complete dynamic path. A bound
        // is never permission to drop history or disable a valid subscription.
      }
    }

    const record = await loadAccountMirrorBySubscriptionToken(context.env.ROSTER_DB, token);
    if (!record) {
      return new Response("Subscription calendar was not found.", { status: 404 });
    }

    const d1Feed = await buildD1SubscriptionFeed(context.env.ROSTER_DB, record, view, context.env.IDENTITY_REVIEW_ENABLED==='true');
    if (d1Feed?.ics) {
      return calendarResponse(d1Feed.ics, d1Feed.displayName || record.realName || record.email);
    }

    return new Response("No D1 subscription calendar is available for this view.", { status: 404 });
  } catch (error) {
    return new Response(error.message || "Subscription feed failed.", { status: 400 });
  }
}

export async function buildD1SubscriptionFeed(db, record, view, identitiesEnabled=false) {
  if (!hasCalendarDb({ ROSTER_DB: db })) return null;
  const role = record?.role || "";
  let claims = sanitizeClaims(record.claims);
  if(identitiesEnabled && !['creator','owner'].includes(role)) claims=await accountIdentityAliases(db,record.email,claims);
  const session = record?.state?.session && typeof record.state.session === "object" ? record.state.session : {};
  const doctorKeys = [...new Set((role === "creator" || role === "owner")
    ? [String(session.doctorKey || "").trim()].filter(Boolean)
    : claims.map((claim) => claim.key))];
  if(identitiesEnabled && ['creator','owner'].includes(role)) {
    // Creator subscriptions follow the selected roster identity, not the owner's
    // own account link. Preserve the existing cross-site key selection and add
    // only explicitly approved aliases through indexed identity lookups.
    claims=await expandApprovedIdentityAliases(db,doctorKeys.flatMap(key=>['mmc','mch','ddh','vhh','casey'].map(sourceType=>({sourceType,key}))));
  }
  if (!doctorKeys.length && !identitiesEnabled) return null;
  const settings = {
    ...defaultSettings(),
    ...(session.settings || {}),
  };
  const range = view === "range" ? normalizeExportRange(session.exportRange) : { mode: "full" };
  const queryOptions = range.mode === "range" && range.startDate
    ? { startDate: range.startDate, endDate: range.allFuture ? "9999-12-31" : range.endDate || range.startDate }
    : {};
  const inputs=identitiesEnabled?await loadSubscriptionCalendarInputs(db,claims,record.email,session,queryOptions):null;
  const hospitalLocations=inputs?.locations||await loadAccountHospitalLocations(db,record.email,session).catch(()=>null);
  return assembleSubscriptionCalendar(record,view,claims,{
    events:inputs?inputs.events:await queryDoctorEvents(db,doctorKeys,queryOptions),
    locations:hospitalLocations||{},custom:inputs?inputs.custom:await queryAccountCustomEvents(db,record.email).catch(()=>[])
  },identitiesEnabled);
}

// The external runner and request fallback share the exact renderer and UID rules.
export function assembleSubscriptionCalendar(record,view,claims,inputs,identitiesEnabled=true) {
  const session=record?.state?.session||{};
  const settings={...defaultSettings(),...(session.settings||{})};
  const range=view==='range'?normalizeExportRange(session.exportRange):{mode:'full'};
  const hospitalLocations=inputs.locations;
  const rosterEvents=dedupeEventsByIdentity(applyEventOverrides(
    applyAccountHospitalLocations(inputs.events,hospitalLocations||{},{includeLocations:settings.includeLocations!==false}),
    session.overrides||{}));
  const d1CustomEvents=inputs.custom;
  const customEvents = customEventsToEvents(latestCustomEventsByIdentity([
    ...sanitizeCustomEvents(session.customEvents, record.email),
    ...sanitizeCustomEvents(d1CustomEvents, record.email),
  ]), settings, rosterEvents);
  const events = [...rosterEvents, ...customEvents]
    .filter((event) => range.mode !== "range" || eventInRange(event, range))
    .sort(compareEvents);
  if (!events.length && !identitiesEnabled) return null;
  const displayName = claims.find(claim=>claim.preferredName)?.preferredName || record.realName || claims[0]?.displayName || record.email;
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
      "Cache-Control": "private, max-age=30, must-revalidate",
      "Content-Disposition": `inline; filename="${String(displayName || "roster").replace(/[\/\r\n"]/g, "-")} subscription.ics"`,
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

function latestCustomEventsByIdentity(events) {
  const byId = new Map();
  for (const event of events || []) {
    byId.delete(event.id);
    byId.set(event.id, event);
  }
  const byIdentity = new Map();
  for (const event of byId.values()) {
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
