import {
  loadCachedSnapshot,
  queryFacilityStaffDesignations,
  queryFacilityStaffSeniorityOverrides,
  queryFacilityOverviewOnShift,
  queryMaterializedFacilityMetadata,
  queryMaterializedFacilityTermStaff,
  storeCachedSnapshot,
} from "./d1-calendar.js";

const SCHEMA_VERSION = 1;

export function facilityMetadataManifestKey(sourceType) {
  return `facility-overview/v1/${safeSource(sourceType)}/manifest.json`;
}

export function facilityStaffObjectKey(sourceType, termStart, revision) {
  return `facility-overview/v1/${safeSource(sourceType)}/staff/${String(termStart).slice(0, 10)}/${revision}.json.gz`;
}

export function facilityDayObjectKey(sourceType, date, revision) {
  return `facility-overview/v1/${safeSource(sourceType)}/days/${String(date).slice(0, 10)}/${revision}.json.gz`;
}

export function facilityMonthObjectKey(sourceType, month, revision) {
  return `facility-overview/v1/${safeSource(sourceType)}/months/${String(month).slice(0, 7)}/${revision}.json.gz`;
}

export async function publishFacilityStaffMetadata(context, sourceTypes = []) {
  const db = context?.env?.ROSTER_DB;
  const r2 = context?.env?.ROSTER_FILES;
  if (!db?.prepare || !r2?.put || !r2?.get) return { ok: false, unavailable: true };
  const results = [];
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const metadata = await queryMaterializedFacilityMetadata(db, { sourceType });
    const currentManifestObject = await loadJsonObject(r2, facilityMetadataManifestKey(sourceType));
    const currentManifest = currentManifestObject.data;
    const terms = [];
    for (const term of metadata.terms || []) {
      const termEnd = addDays(term.termStart, 90);
      const members = (await queryMaterializedFacilityTermStaff(db, { sourceType, termStart: term.termStart, termEnd }))
        .map((member) => ({
          ...member,
          coverageStart: member.firstApplicableDate || "",
          coverageEnd: member.membershipSource === "sms-continuity" ? "" : member.lastApplicableDate || "",
        }));
      const [designations, seniorityOverrides] = await Promise.all([
        queryFacilityStaffDesignations(db, { sourceType, termStart: term.termStart, termEnd }),
        queryFacilityStaffSeniorityOverrides(db, { sourceType, termStart: term.termStart }),
      ]);
      const staff = { schemaVersion: SCHEMA_VERSION, sourceType, termStart: term.termStart, termEnd, members, events: [], coverage: metadata.coverage, designations, seniorityOverrides };
      const staffRevision = await digest(staff);
      const staffKey = facilityStaffObjectKey(sourceType, term.termStart, staffRevision);
      const currentTerm = (currentManifest?.terms || []).find((entry) => entry.termStart === term.termStart);
      if (currentTerm?.staffRevision !== staffRevision) {
        await storeCachedSnapshot(r2, staffKey, staff, { revision: staffRevision, ownerType: "facility-staff", ownerId: sourceType, rangeKey: term.termStart });
      }
      terms.push({ ...term, termEnd, staffKey, staffRevision });
    }
    const stableManifest = { schemaVersion: SCHEMA_VERSION, sourceType, coverage: metadata.coverage, terms, days: currentManifest?.days || {}, months: currentManifest?.months || {} };
    const revision = await digest(stableManifest);
    if (currentManifest?.revision !== revision) {
      await putJsonGzip(
        r2,
        facilityMetadataManifestKey(sourceType),
        { ...stableManifest, revision, publishedAt: new Date().toISOString() },
        currentManifestObject.etag ? { onlyIf: { etagMatches: currentManifestObject.etag } } : {},
      );
      results.push({ sourceType, changed: true, revision });
    } else {
      results.push({ sourceType, changed: false, revision });
    }
  }
  return { ok: true, results };
}

export async function publishFacilityDays(context, sourceTypeValue, dates = [], options = {}) {
  const db = context?.env?.ROSTER_DB;
  const r2 = context?.env?.ROSTER_FILES;
  const sourceType = safeSource(sourceTypeValue);
  let affectedDates = [...new Set(dates.map((date) => String(date || "").slice(0, 10)).filter((date) => /^\d{4}-\d{2}-\d{2}$/.test(date)))].sort();
  if (!db?.prepare || !r2?.put || !r2?.get || !sourceType) return { ok: false, unavailable: true };
  const now = new Date();
  const nowIso = now.toISOString();
  const existing = await db.prepare("SELECT * FROM facility_day_publications WHERE source_type = ?").bind(sourceType).first();
  const manifestObject = await loadJsonObject(r2, facilityMetadataManifestKey(sourceType));
  if (!affectedDates.length && options.rebuildPublishedDates === true) affectedDates = Object.keys(manifestObject.data?.days || {}).sort();
  if (!affectedDates.length) return { ok: true, unchanged: true };
  if (affectedDates.length > 120) throw new Error(`Day publication exceeds the 120-date safety budget (${affectedDates.length}).`);
  if (existing?.status !== "complete" && existing?.operation_id && manifestObject.data?.publicationOperationId === existing.operation_id) {
    await db.prepare("UPDATE facility_day_publications SET status = 'complete', lease_expires_at = '', last_error = '', updated_at = ? WHERE source_type = ? AND operation_id = ?")
      .bind(nowIso, sourceType, existing.operation_id).run();
  } else if (existing?.status === "publishing" && String(existing.lease_expires_at || "") > nowIso) {
    return { ok: true, busy: true, generation: Number(existing.generation || 0) };
  }
  const generation = Number(existing?.generation || 0) + 1;
  const operationId = crypto.randomUUID();
  const leaseExpiresAt = new Date(now.getTime() + 2 * 60 * 1000).toISOString();
  const baseRevision = String(manifestObject.data?.revision || "");
  await db.prepare(`INSERT INTO facility_day_publications
    (source_type, generation, operation_id, status, base_revision, candidate_revision, lease_expires_at, last_error, updated_at)
    VALUES (?, ?, ?, 'publishing', ?, '', ?, '', ?)
    ON CONFLICT(source_type) DO UPDATE SET generation=excluded.generation, operation_id=excluded.operation_id,
      status='publishing', base_revision=excluded.base_revision, candidate_revision='', lease_expires_at=excluded.lease_expires_at,
      last_error='', updated_at=excluded.updated_at`)
    .bind(sourceType, generation, operationId, baseRevision, leaseExpiresAt, nowIso).run();
  try {
    const currentManifest = manifestObject.data || { schemaVersion: SCHEMA_VERSION, sourceType, coverage: [], terms: [], days: {}, months: {} };
    const days = { ...(currentManifest.days || {}) };
    for (const date of affectedDates) {
      const rows = await queryFacilityOverviewOnShift(db, { facilityKey: sourceType, date, includeOverrides: false });
      const covered = (currentManifest.coverage || []).some((entry) => entry.startDate <= date && entry.endDate >= date);
      if (!rows.length && !covered) {
        delete days[date];
        continue;
      }
      const payload = { schemaVersion: SCHEMA_VERSION, sourceType, date, rows };
      const revision = await digest(payload);
      const key = facilityDayObjectKey(sourceType, date, revision);
      if (days[date]?.revision !== revision) await putJsonGzip(r2, key, payload);
      days[date] = { key, revision };
    }
    const months = { ...(currentManifest.months || {}) };
    for (const month of [...new Set(affectedDates.map((date) => date.slice(0, 7)))]) {
      const monthRows = [];
      const monthDates = Object.entries(days).filter(([date]) => date.startsWith(`${month}-`)).sort(([a], [b]) => a.localeCompare(b));
      for (const [date, pointer] of monthDates) {
        const day = await loadCachedSnapshot(r2, pointer.key);
        if (day?.rows) monthRows.push(...day.rows);
      }
      if (!monthDates.length) {
        delete months[month];
        continue;
      }
      const payload = { schemaVersion: SCHEMA_VERSION, sourceType, month, dates: monthDates.map(([date]) => date), rows: monthRows };
      const revision = await digest(payload);
      const key = facilityMonthObjectKey(sourceType, month, revision);
      if (months[month]?.revision !== revision) await putJsonGzip(r2, key, payload);
      months[month] = { key, revision };
    }
    const stable = { schemaVersion: currentManifest.schemaVersion || SCHEMA_VERSION, sourceType, coverage: currentManifest.coverage || [], terms: currentManifest.terms || [], days, months };
    const candidateRevision = await digest(stable);
    if (candidateRevision === baseRevision) {
      await db.prepare("UPDATE facility_day_publications SET status = 'complete', candidate_revision = ?, lease_expires_at = '', last_error = '', updated_at = ? WHERE source_type = ? AND operation_id = ?")
        .bind(candidateRevision, new Date().toISOString(), sourceType, operationId).run();
      return { ok: true, unchanged: true, generation, revision: candidateRevision, dates: affectedDates };
    }
    const candidate = { ...stable, revision: candidateRevision, generation, publicationOperationId: operationId, publishedAt: nowIso };
    const candidateKey = `facility-overview/v1/${sourceType}/manifests/${generation}-${candidateRevision}.json.gz`;
    await putJsonGzip(r2, candidateKey, candidate);
    await db.prepare("UPDATE facility_day_publications SET candidate_revision = ?, updated_at = ? WHERE source_type = ? AND operation_id = ?")
      .bind(candidateRevision, nowIso, sourceType, operationId).run();
    if (typeof options.beforePointer === "function") await options.beforePointer({ operationId, generation });
    const owner = await db.prepare("SELECT operation_id FROM facility_day_publications WHERE source_type = ? AND status = 'publishing' AND lease_expires_at > ?")
      .bind(sourceType, new Date().toISOString()).first();
    if (owner?.operation_id !== operationId) throw new Error("Day publication lease was superseded.");
    await putJsonGzip(r2, facilityMetadataManifestKey(sourceType), candidate, manifestObject.etag ? { onlyIf: { etagMatches: manifestObject.etag } } : {});
    if (typeof options.afterPointer === "function") await options.afterPointer({ operationId, generation });
    await db.prepare("UPDATE facility_day_publications SET status = 'complete', lease_expires_at = '', last_error = '', updated_at = ? WHERE source_type = ? AND operation_id = ?")
      .bind(new Date().toISOString(), sourceType, operationId).run();
    return { ok: true, changed: candidateRevision !== baseRevision, generation, revision: candidateRevision, dates: affectedDates };
  } catch (error) {
    await db.prepare("UPDATE facility_day_publications SET status = 'error', lease_expires_at = '', last_error = ?, updated_at = ? WHERE source_type = ? AND operation_id = ?")
      .bind(String(error?.message || error).slice(0, 300), new Date().toISOString(), sourceType, operationId).run().catch(() => null);
    throw error;
  }
}

export async function loadPublishedFacilityRange(r2, sourceTypes, startDate, endDate, currentDate, options = {}) {
  if (!r2?.get) return { preparing: true, events: [], coverage: [], revision: "" };
  const start = new Date(`${startDate}T12:00:00Z`);
  const end = new Date(`${endDate}T12:00:00Z`);
  const rangeDays = Math.round((end.getTime() - start.getTime()) / 86400000);
  if (!Number.isFinite(rangeDays) || rangeDays < 0 || rangeDays > 370) return { preparing: true, events: [], coverage: [], revision: "" };
  const months = monthsInRange(startDate, endDate);
  const selected = [];
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const manifest = await loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType));
    if (!manifest) continue;
    const visibleTerms = (manifest.terms || []).filter((term) => term.visibleFrom <= currentDate && term.termEnd >= startDate && term.termStart <= endDate);
    if (!visibleTerms.length) continue;
    const monthPointers = months.map((month) => [month, manifest.months?.[month]]).filter(([, pointer]) => pointer?.key);
    if (monthPointers.length) selected.push({ sourceType, manifest, visibleTerms, monthPointers });
  }
  if (!selected.length) return { preparing: true, events: [], coverage: [], revision: "" };
  const revision = await digest(selected.flatMap(({ visibleTerms, monthPointers }) => [
    ...visibleTerms.map((term) => term.staffRevision || ""),
    ...monthPointers.map(([, pointer]) => pointer.revision || ""),
  ]).sort());
  if (options.cachedRevision && String(options.cachedRevision) === revision) return { preparing: false, unchanged: true, events: [], coverage: [], revision };
  const events = [];
  const coverage = [];
  for (const { manifest, visibleTerms, monthPointers } of selected) {
    const overrides = new Map();
    for (const term of visibleTerms) {
      const staff = term.staffKey ? await loadCachedSnapshot(r2, term.staffKey) : null;
      for (const entry of staff?.seniorityOverrides || []) overrides.set(`${entry.sourceType}|${entry.doctorKey}`, entry);
    }
    for (const [, pointer] of monthPointers) {
      const snapshot = await loadCachedSnapshot(r2, pointer.key);
      if (!snapshot) continue;
      for (const row of snapshot.rows || []) {
        const date = String(row.event?.start || "").slice(0, 10);
        if (date < startDate || date > endDate || !visibleTerms.some((term) => term.termStart <= date && term.termEnd >= date)) continue;
        const override = overrides.get(`${row.sourceType}|${row.doctorKey}`);
        events.push(override && !override.useRosterSeniority
          ? { ...row, seniority: override.seniority, seniorityOverride: override, event: { ...row.event, seniority: override.seniority, facilitySeniorityOverride: true } }
          : row);
      }
    }
    coverage.push(...(manifest.coverage || []));
  }
  return { preparing: false, events, coverage, revision };
}

export async function loadPublishedFacilityDays(r2, sourceTypes, date, currentDate = date) {
  if (!r2?.get) return { preparing: true, rows: [] };
  const rows = [];
  const revisions = [];
  let found = false;
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const manifest = await loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType));
    const pointer = manifest?.days?.[date];
    if (!pointer?.key) continue;
    const day = await loadCachedSnapshot(r2, pointer.key);
    if (!day) continue;
    const termStart = termStartForDate(date);
    const term = (manifest.terms || []).find((entry) => entry.termStart === termStart && entry.visibleFrom <= currentDate);
    if (!term) continue;
    found = true;
    revisions.push(pointer.revision || "", term.staffRevision || "");
    const staff = term?.staffKey ? await loadCachedSnapshot(r2, term.staffKey) : null;
    const overrides = new Map((staff?.seniorityOverrides || []).map((entry) => [`${entry.sourceType}|${entry.doctorKey}`, entry]));
    rows.push(...(day.rows || []).map((row) => {
      const override = overrides.get(`${row.sourceType}|${row.doctorKey}`);
      return override && !override.useRosterSeniority
        ? { ...row, seniority: override.seniority, seniorityOverride: override, event: { ...row.event, seniority: override.seniority, facilitySeniorityOverride: true } }
        : row;
    }));
  }
  return { preparing: !found, rows, revision: await digest(revisions.sort()) };
}

export async function loadPublishedFacilityMetadata(r2, sourceTypes, today) {
  if (!r2?.get) return { preparing: true, facilities: [], catalogEvents: [] };
  const manifests = (await Promise.all([...new Set(sourceTypes.map(safeSource).filter(Boolean))]
    .map((sourceType) => loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType))))).filter(Boolean);
  if (!manifests.length) return { preparing: true, facilities: [], catalogEvents: [] };
  const facilities = manifests.flatMap((manifest) => manifest.coverage || []);
  const catalogEvents = [];
  for (const manifest of manifests) {
    for (const term of manifest.terms || []) {
      if (term.visibleFrom > today || term.termEnd < today) continue;
      for (const [index, item] of (term.catalog || []).entries()) {
        for (const date of [...new Set([item.firstDate, item.lastDate].filter(Boolean))]) {
          const start = item.startTime ? `${date}T${item.startTime}` : date;
          const end = item.endTime ? `${date}T${item.endTime}` : date;
          catalogEvents.push({
            doctorKey: `FACILITY_OVERVIEW_CATALOG_${item.catalogKey || index}`,
            displayName: "Roster catalogue",
            sourceType: manifest.sourceType,
            seniority: item.seniority || "",
            date,
            event: { id: `facility-overview-catalog:${item.catalogKey || index}:${date}`, title: item.title || "", rawValue: item.rawValue || "", location: item.location || "", source: manifest.sourceType, sources: [manifest.sourceType], seniority: item.seniority || "", start, end, allDay: item.allDay === true, timeLabel: item.timeLabel || "" },
          });
        }
      }
    }
  }
  return { preparing: false, facilities, catalogEvents, revision: await digest({ facilities, catalogEvents }) };
}

export async function loadPublishedFacilityStaff(r2, sourceTypes, termStart, today) {
  if (!r2?.get) return { preparing: true };
  const payloads = [];
  const revisions = [];
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const manifest = await loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType));
    if (!manifest) continue;
    const term = (manifest.terms || []).find((entry) => entry.termStart === termStart);
    if (!term || term.visibleFrom > today || !term.staffKey) continue;
    const staff = await loadCachedSnapshot(r2, term.staffKey);
    if (staff) {
      payloads.push(staff);
      revisions.push(term.staffRevision || "");
    }
  }
  if (!payloads.length) return { preparing: true };
  return {
    preparing: false,
    revision: await digest(revisions.sort()),
    members: payloads.flatMap((item) => item.members || []),
    events: payloads.flatMap((item) => item.events || []),
    coverage: payloads.flatMap((item) => item.coverage || []),
    designations: payloads.flatMap((item) => item.designations || []),
    seniorityOverrides: payloads.flatMap((item) => item.seniorityOverrides || []),
  };
}

function safeSource(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9-]+/g, "");
}

function addDays(value, days) {
  const date = new Date(`${String(value).slice(0, 10)}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

async function digest(value) {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  const result = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(result)].map((part) => part.toString(16).padStart(2, "0")).join("");
}

async function putJsonGzip(r2, key, value, options = {}) {
  const json = JSON.stringify(value);
  const bytes = await new Response(new Blob([json]).stream().pipeThrough(new CompressionStream("gzip"))).arrayBuffer();
  return r2.put(key, bytes, { ...options, httpMetadata: { contentType: "application/json", contentEncoding: "gzip" } });
}

async function loadJsonObject(r2, key) {
  try {
    const object = await r2.get(key);
    if (!object) return { data: null, etag: "" };
    const bytes = await object.arrayBuffer();
    const header = new Uint8Array(bytes, 0, Math.min(2, bytes.byteLength));
    const text = header[0] === 0x1f && header[1] === 0x8b
      ? await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).text()
      : new TextDecoder().decode(bytes);
    return { data: JSON.parse(text), etag: String(object.etag || object.httpEtag || "") };
  } catch {
    return { data: null, etag: "" };
  }
}

function termStartForDate(value) {
  const date = new Date(`${String(value).slice(0, 10)}T12:00:00Z`);
  const year = date.getUTCFullYear();
  const candidates = [];
  for (const candidateYear of [year - 1, year]) {
    for (const month of [1, 4, 7, 10]) {
      const first = new Date(Date.UTC(candidateYear, month, 1, 12));
      first.setUTCDate(1 + ((8 - first.getUTCDay()) % 7));
      candidates.push(first.toISOString().slice(0, 10));
    }
  }
  return candidates.filter((candidate) => candidate <= value).sort().at(-1) || "";
}

function monthsInRange(startDate, endDate) {
  const cursor = new Date(`${startDate.slice(0, 7)}-01T12:00:00Z`);
  const endMonth = endDate.slice(0, 7);
  const result = [];
  while (cursor.toISOString().slice(0, 7) <= endMonth && result.length < 13) {
    result.push(cursor.toISOString().slice(0, 7));
    cursor.setUTCMonth(cursor.getUTCMonth() + 1);
  }
  return result;
}
