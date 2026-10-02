import {
  DDH_CONTACT_LIST_SOURCE_ID,
  MMC_CONTACT_LIST_SOURCE_ID,
  VHH_CONTACT_LIST_SOURCE_ID,
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
  // Provider timestamps and versions describe the workbook, not the clinical
  // allocation. SharePoint autosave may change them without changing a single
  // contact. Keep the published revision stable for the same allocation so a
  // metadata-only automation run performs no R2 writes.
  const revision = await digest(contactAllocationValue(extract));
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

export function contactAllocationValue(extractValue) {
  const extract = normaliseContactListExtract(extractValue);
  if (!extract) return null;
  const contacts = extract.contacts.map((contact) => ({
    area: String(contact.area || ""),
    shift: String(contact.shift || ""),
    role: String(contact.role || ""),
    name: String(contact.name || ""),
    phone: String(contact.phone || ""),
    isPopulated: contact.isPopulated === true,
  })).sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
  return {
    sourceId: String(extract.sourceId || ""),
    sourceDate: String(extract.sourceDate || ""),
    contacts,
  };
}

export async function publishFacilityContactResolutions(r2, sourceId, sourceDate, resolutions = []) {
  if (!r2?.put || !sourceId || !sourceDate) return { ok: false, unavailable: true };
  const key = facilityContactResolutionKey(sourceId, sourceDate);
  // Publication can finish out of order after concurrent human saves. Merge
  // monotonic per-contact revisions and use R2's conditional write; an older
  // snapshot must never resurrect a rejected allocation. No additional D1 read.
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const current = await readJsonObject(r2, key);
    const merged = new Map((current.data?.resolutions || []).map((resolution) => [resolution.contactKey, resolution]));
    for (const resolution of resolutions) {
      if (Number(resolution.revision || 0) >= Number(merged.get(resolution.contactKey)?.revision || 0)) merged.set(resolution.contactKey, resolution);
    }
    const stable = { schemaVersion: SCHEMA_VERSION, sourceId, sourceDate,
      resolutions: [...merged.values()].sort((left, right) => String(left.contactKey).localeCompare(String(right.contactKey))) };
    const revision = await digest(stable);
    if (revision === current.data?.revision) return { ok: true, revision, unchanged: true };
    try {
      const result = await putJson(r2, key, { ...stable, revision, publishedAt: new Date().toISOString() }, {
        onlyIf: current.etag ? { etagMatches: current.etag } : { etagDoesNotMatch: "*" },
      });
      if (result !== null) return { ok: true, revision };
    } catch (error) {
      if (!/precondition|condition|412/i.test(error?.message || "")) throw error;
    }
  }
  throw new Error("Contact corrections changed during publication. Please retry the save.");
}

export async function loadPublishedFacilityContacts(r2, { date, facilityKeys = [], now = new Date() } = {}) {
  if (!r2?.get) return { status: "unavailable", reason: "storage-unavailable" };
  const sourceIds = new Set(facilityKeys.map(contactSourceForFacility).filter(Boolean));
  if (sourceIds.size !== 1) return { status: "unavailable", reason: "multiple-sources" };
  const sourceId = [...sourceIds][0];
  const manifest = (await readJsonObject(r2, facilityContactManifestKey(sourceId))).data;
  if (!manifest) return { status: "unavailable", reason: "no-extract" };
  const entries = Object.entries(manifest.dates || {})
    .filter(([sourceDate]) => !contactExtractHasExpired(sourceDate, now))
    .sort(([a], [b]) => b.localeCompare(a));
  let selected = entries.find(([sourceDate]) => sourceDate === date);
  let carryMode = "";
  if (sourceId === VHH_CONTACT_LIST_SOURCE_ID) {
    const today = new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit" }).format(now);
    const yesterday = new Date(`${today}T12:00:00Z`); yesterday.setUTCDate(yesterday.getUTCDate() - 1);
    selected = date === today || date === yesterday.toISOString().slice(0, 10) ? entries.find(([sourceDate]) => sourceDate === today) : null;
    if (!selected) return { status: "not-current", sourceId, contacts: [], resolutions: [], revision: "" };
  }
  if (!selected) {
    selected = entries.find(([sourceDate]) => shouldUseCurrentExtractForPreviousNight(sourceDate, date, now));
    carryMode = selected ? "current-for-previous" : "";
  }
  if (!selected) {
    selected = entries.find(([sourceDate]) => shouldCarryPreviousNightContacts(sourceDate, date, now));
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
  const visibleContacts = sourceId === VHH_CONTACT_LIST_SOURCE_ID || carryMode ? contacts : contactsAfterShiftChange(contacts, { date, now });
  return {
    status: "available",
    // Handover changes visibility even if the workbook hasn't changed. A token
    // refresh must deliver the newly visible Night block after 23:00.
    revision: `${pointer.revision || ""}:${resolutionPayload?.revision || ""}:${await digest(visibleContacts.map((contact) => contact.contactKey))}`,
    sourceId,
    sourceDate: operationalExtract.sourceDate,
    providerModifiedAt: extract.providerModifiedAt || pointer.providerModifiedAt || "",
    receivedAt: pointer.receivedAt || "",
    contacts: visibleContacts,
    resolutions: resolutionPayload?.resolutions || [],
  };
}

function contactSourceForFacility(value) {
  const code = String(value || "").trim().toUpperCase();
  if (code === "VHH") return VHH_CONTACT_LIST_SOURCE_ID;
  if (code === "DDH") return DDH_CONTACT_LIST_SOURCE_ID;
  if (code === "MMC" || code === "MCH") return MMC_CONTACT_LIST_SOURCE_ID;
  return "";
}

async function readJsonObject(r2, key) {
  try {
    const object = await r2.get(key);
    if (!object) return { data: null, etag: "" };
    const text = typeof object.text === "function"
      ? await object.text()
      : new TextDecoder().decode(await object.arrayBuffer());
    return { data: JSON.parse(text), etag: String(object.etag || object.httpEtag || "") };
  } catch {
    return { data: null, etag: "" };
  }
}

function putJson(r2, key, value, options = {}) {
  return r2.put(key, new TextEncoder().encode(JSON.stringify(value)), { ...options, httpMetadata: { contentType: "application/json; charset=utf-8" } });
}

async function digest(value) {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  const result = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(result)].map((part) => part.toString(16).padStart(2, "0")).join("");
}

function safePart(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-");
}
