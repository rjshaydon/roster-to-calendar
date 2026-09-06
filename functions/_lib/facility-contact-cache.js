import {
  DDH_CONTACT_LIST_SOURCE_ID,
  MMC_CONTACT_LIST_SOURCE_ID,
  contactAreaForSource,
  contactExtractHasExpired,
  contactsAfterShiftChange,
  normaliseContactListExtract,
  shouldCarryPreviousNightContacts,
  shouldUseCurrentExtractForPreviousNight,
} from "../../public/static/contact-allocations.js";

const SCHEMA_VERSION = 1;

export function facilityContactManifestKey(sourceId) {
  return `facility-overview/v1/contacts/${safePart(sourceId)}/manifest.json`;
}

export function facilityContactObjectKey(sourceId, sourceDate, revision) {
  return `facility-overview/v1/contacts/${safePart(sourceId)}/${sourceDate}/${revision}.json`;
}

export function facilityContactResolutionKey(sourceId, sourceDate) {
  return `facility-overview/v1/contacts/${safePart(sourceId)}/${sourceDate}/resolutions.json`;
}

export async function publishFacilityContactExtract(r2, extractValue, metadata = {}) {
  const extract = normaliseContactListExtract(extractValue);
  if (!r2?.get || !r2?.put || !extract) return { ok: false, unavailable: true };
  const sourceId = extract.sourceId;
  const sourceDate = extract.sourceDate;
  const payload = { schemaVersion: SCHEMA_VERSION, sourceId, sourceDate, extract };
  const revision = await digest(payload);
  const objectKey = facilityContactObjectKey(sourceId, sourceDate, revision);
  const manifestKey = facilityContactManifestKey(sourceId);
  const currentObject = await readJsonObject(r2, manifestKey);
  const current = currentObject.data || { schemaVersion: SCHEMA_VERSION, sourceId, dates: {} };
  if (current.dates?.[sourceDate]?.revision === revision) return { ok: true, unchanged: true, revision };
  await putJson(r2, objectKey, payload);
  const dates = { ...(current.dates || {}), [sourceDate]: {
    key: objectKey,
    revision,
    providerModifiedAt: extract.providerModifiedAt || String(metadata.providerModifiedAt || ""),
    receivedAt: String(metadata.receivedAt || ""),
  } };
  const liveDates = Object.fromEntries(Object.entries(dates).filter(([date]) => !contactExtractHasExpired(date)));
  const manifest = { schemaVersion: SCHEMA_VERSION, sourceId, dates: liveDates, revision: await digest({ sourceId, dates: liveDates }), publishedAt: new Date().toISOString() };
  await putJson(r2, manifestKey, manifest, { onlyIf: currentObject.etag ? { etagMatches: currentObject.etag } : { etagDoesNotMatch: "*" } });
  return { ok: true, changed: true, revision };
}

export async function publishFacilityContactResolutions(r2, sourceId, sourceDate, resolutions = []) {
  if (!r2?.put || !sourceId || !sourceDate) return { ok: false, unavailable: true };
  const stable = { schemaVersion: SCHEMA_VERSION, sourceId, sourceDate, resolutions };
  const revision = await digest(stable);
  await putJson(r2, facilityContactResolutionKey(sourceId, sourceDate), { ...stable, revision, publishedAt: new Date().toISOString() });
  return { ok: true, revision };
}

export async function loadPublishedFacilityContacts(r2, { date, facilityKeys = [] } = {}) {
  if (!r2?.get) return { status: "unavailable", reason: "storage-unavailable" };
  const sourceIds = new Set(facilityKeys.map(contactSourceForFacility).filter(Boolean));
  if (sourceIds.size !== 1) return { status: "unavailable", reason: "multiple-sources" };
  const sourceId = [...sourceIds][0];
  const manifest = (await readJsonObject(r2, facilityContactManifestKey(sourceId))).data;
  if (!manifest) return { status: "unavailable", reason: "no-extract" };
  const entries = Object.entries(manifest.dates || {})
    .filter(([sourceDate]) => !contactExtractHasExpired(sourceDate))
    .sort(([a], [b]) => b.localeCompare(a));
  let selected = entries.find(([sourceDate]) => sourceDate === date);
  let carryMode = "";
  if (!selected) {
    selected = entries.find(([sourceDate]) => shouldUseCurrentExtractForPreviousNight(sourceDate, date));
    carryMode = selected ? "current-for-previous" : "";
  }
  if (!selected) {
    selected = entries.find(([sourceDate]) => shouldCarryPreviousNightContacts(sourceDate, date));
    carryMode = selected ? "previous-night" : "";
  }
  if (!selected) {
    const fallback = entries[0];
    return fallback ? { status: "not-current", revision: fallback[1].revision || "", sourceId, sourceDate: fallback[0], contacts: [], resolutions: [] }
      : { status: "unavailable", reason: "no-extract" };
  }
  const [storedDate, pointer] = selected;
  const payload = (await readJsonObject(r2, pointer.key)).data;
  const extract = normaliseContactListExtract(payload?.extract);
  if (!extract) return { status: "unavailable", reason: "object-missing" };
  const operationalExtract = carryMode ? normaliseContactListExtract({ ...extract, sourceDate: date, contacts: extract.contacts.filter((contact) => contact.shift === "Night") }) : extract;
  const allowedAreas = new Set(facilityKeys.map(contactAreaForSource).filter(Boolean));
  const contacts = operationalExtract.contacts.filter((contact) => allowedAreas.has(contact.area));
  const resolutionPayload = (await readJsonObject(r2, facilityContactResolutionKey(sourceId, operationalExtract.sourceDate))).data;
  return {
    status: "available",
    revision: `${pointer.revision || ""}:${resolutionPayload?.revision || ""}`,
    sourceId,
    sourceDate: operationalExtract.sourceDate,
    providerModifiedAt: extract.providerModifiedAt || pointer.providerModifiedAt || "",
    receivedAt: pointer.receivedAt || "",
    contacts: carryMode ? contacts : contactsAfterShiftChange(contacts, { date }),
    resolutions: resolutionPayload?.resolutions || [],
  };
}

function contactSourceForFacility(value) {
  const code = String(value || "").trim().toUpperCase();
  if (code === "DDH") return DDH_CONTACT_LIST_SOURCE_ID;
  if (code === "MMC" || code === "MCH") return MMC_CONTACT_LIST_SOURCE_ID;
  return "";
}

async function readJsonObject(r2, key) {
  try {
    const object = await r2.get(key);
    if (!object) return { data: null, etag: "" };
    return { data: JSON.parse(await object.text()), etag: String(object.etag || object.httpEtag || "") };
  } catch {
    return { data: null, etag: "" };
  }
}

function putJson(r2, key, value, options = {}) {
  return r2.put(key, JSON.stringify(value), { ...options, httpMetadata: { contentType: "application/json; charset=utf-8" } });
}

async function digest(value) {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  const result = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(result)].map((part) => part.toString(16).padStart(2, "0")).join("");
}

function safePart(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-");
}
