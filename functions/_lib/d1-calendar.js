import {
  buildRosterViewFromStoredImports,
  defaultSettings,
  previewSummary,
  serializeEvent,
} from "./roster.js";

const SOURCE_TYPES = ["mmc", "ddh", "casey", "mch"];

export function hasCalendarDb(env) {
  return Boolean(env?.ROSTER_DB?.prepare);
}

export async function ensureCalendarSchema(db) {
  if (!db?.prepare) return false;
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_files (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      source_type TEXT NOT NULL,
      active INTEGER NOT NULL DEFAULT 1,
      size INTEGER NOT NULL DEFAULT 0,
      last_modified INTEGER NOT NULL DEFAULT 0,
      added_at TEXT NOT NULL DEFAULT '',
      uploaded_at TEXT NOT NULL DEFAULT '',
      uploaded_by TEXT NOT NULL DEFAULT '',
      parsed_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_doctors (
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      PRIMARY KEY (source_type, doctor_key)
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_file_doctors (
      file_id TEXT NOT NULL,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      PRIMARY KEY (file_id, source_type, doctor_key)
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_events (
      id TEXT PRIMARY KEY,
      file_id TEXT NOT NULL,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      start_date TEXT NOT NULL,
      end_date TEXT NOT NULL,
      start_ts TEXT NOT NULL,
      end_ts TEXT NOT NULL,
      title TEXT NOT NULL,
      raw_value TEXT NOT NULL DEFAULT '',
      seniority TEXT NOT NULL DEFAULT '',
      location TEXT NOT NULL DEFAULT '',
      all_day INTEGER NOT NULL DEFAULT 0,
      time_label TEXT NOT NULL DEFAULT '',
      event_json TEXT NOT NULL
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_doctor_range ON roster_events (doctor_key, start_date, end_date)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_date_source ON roster_events (start_date, source_type)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_file ON roster_events (file_id)").run();
  return true;
}

export async function upsertDerivedRosterFile(db, file, storedImport) {
  if (!db?.prepare || !file?.id || !storedImport?.dataUrl) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType || storedImport.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const doctors = sanitizeFileDoctors(file.doctors, sourceType);
  const parsedAt = new Date().toISOString();
  await db.prepare(`
    INSERT INTO roster_files (id, name, source_type, active, size, last_modified, added_at, uploaded_at, uploaded_by, parsed_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      source_type = excluded.source_type,
      active = excluded.active,
      size = excluded.size,
      last_modified = excluded.last_modified,
      added_at = excluded.added_at,
      uploaded_at = excluded.uploaded_at,
      uploaded_by = excluded.uploaded_by,
      parsed_at = excluded.parsed_at
  `).bind(
    file.id,
    file.name || storedImport.name || "roster.xlsx",
    sourceType,
    file.active === false ? 0 : 1,
    Number(file.size || storedImport.size || 0),
    Number(file.lastModified || storedImport.lastModified || 0),
    String(file.addedAt || ""),
    String(file.uploadedAt || ""),
    String(file.uploadedBy || ""),
    parsedAt,
  ).run();
  await db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(file.id).run();
  await db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(file.id).run();
  for (const doctor of doctors) {
    await db.prepare(`
      INSERT INTO roster_doctors (source_type, doctor_key, display_name, updated_at)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        updated_at = excluded.updated_at
    `).bind(sourceType, doctor.key, doctor.displayName, parsedAt).run();
    await db.prepare(`
      INSERT INTO roster_file_doctors (file_id, source_type, doctor_key, display_name)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET display_name = excluded.display_name
    `).bind(file.id, sourceType, doctor.key, doctor.displayName).run();
    const view = await buildRosterViewFromStoredImports([{ ...storedImport, sourceType, id: file.id, repoId: file.id }], doctor.key, defaultSettings(), {}, {}, []);
    for (const event of view.events.map(serializeEvent)) {
      await db.prepare(`
        INSERT INTO roster_events (
          id, file_id, source_type, doctor_key, display_name, start_date, end_date, start_ts, end_ts,
          title, raw_value, seniority, location, all_day, time_label, event_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(
        `${file.id}:${doctor.key}:${event.id}`,
        file.id,
        sourceType,
        doctor.key,
        doctor.displayName,
        datePart(event.start),
        datePart(event.end || event.start),
        String(event.start || ""),
        String(event.end || event.start || ""),
        String(event.title || ""),
        String(event.rawValue || ""),
        String(event.seniority || ""),
        String(event.location || ""),
        event.allDay === true ? 1 : 0,
        String(event.timeLabel || ""),
        JSON.stringify(event),
      ).run();
    }
  }
  return { ok: true, doctors: doctors.length };
}

export async function replaceDerivedRosterFile(db, file, doctors, eventsByDoctor) {
  if (!db?.prepare || !file?.id) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const safeDoctors = sanitizeFileDoctors(doctors, sourceType);
  const parsedAt = new Date().toISOString();
  await db.prepare(`
    INSERT INTO roster_files (id, name, source_type, active, size, last_modified, added_at, uploaded_at, uploaded_by, parsed_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      source_type = excluded.source_type,
      active = excluded.active,
      size = excluded.size,
      last_modified = excluded.last_modified,
      added_at = excluded.added_at,
      uploaded_at = excluded.uploaded_at,
      uploaded_by = excluded.uploaded_by,
      parsed_at = excluded.parsed_at
  `).bind(
    file.id,
    file.name || "roster.xlsx",
    sourceType,
    file.active === false ? 0 : 1,
    Number(file.size || 0),
    Number(file.lastModified || 0),
    String(file.addedAt || ""),
    String(file.uploadedAt || ""),
    String(file.uploadedBy || ""),
    parsedAt,
  ).run();
  await db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(file.id).run();
  await db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(file.id).run();
  await bulkUpsertDoctors(db, sourceType, safeDoctors, parsedAt);
  await bulkInsertFileDoctors(db, file.id, sourceType, safeDoctors);
  const eventRows = [];
  for (const doctor of safeDoctors) {
    const events = Array.isArray(eventsByDoctor?.[doctor.key]) ? eventsByDoctor[doctor.key] : [];
    for (const event of events.map(sanitizeEvent).filter(Boolean)) {
      eventRows.push([
        `${file.id}:${doctor.key}:${event.id}`,
        file.id,
        sourceType,
        doctor.key,
        doctor.displayName,
        datePart(event.start),
        datePart(event.end || event.start),
        String(event.start || ""),
        String(event.end || event.start || ""),
        String(event.title || ""),
        String(event.rawValue || ""),
        String(event.seniority || ""),
        String(event.location || ""),
        event.allDay === true ? 1 : 0,
        String(event.timeLabel || ""),
        JSON.stringify(event),
      ]);
    }
  }
  await bulkInsertEvents(db, eventRows);
  return { ok: true, doctors: safeDoctors.length, events: eventRows.length };
}

export async function setDerivedRosterFileActive(db, fileId, active) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("UPDATE roster_files SET active = ? WHERE id = ?").bind(active ? 1 : 0, fileId).run();
}

export async function deleteDerivedRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_files WHERE id = ?").bind(fileId).run();
}

export async function queryDoctorEvents(db, doctorKeys, options = {}) {
  if (!db?.prepare || !doctorKeys?.length) return [];
  await ensureCalendarSchema(db);
  const keys = [...new Set(doctorKeys.filter(Boolean))];
  if (!keys.length) return [];
  const placeholders = keys.map(() => "?").join(", ");
  const start = String(options.startDate || "0000-01-01");
  const end = String(options.endDate || "9999-12-31");
  const rows = await db.prepare(`
    SELECT event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.doctor_key IN (${placeholders})
      AND roster_events.start_date <= ?
      AND roster_events.end_date >= ?
    ORDER BY roster_events.start_ts, roster_events.source_type, roster_events.title
  `).bind(...keys, end, start).all();
  return (rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean);
}

export async function countDerivedEventsByFile(db, fileIds = []) {
  if (!db?.prepare || !fileIds?.length) return new Map();
  await ensureCalendarSchema(db);
  const counts = new Map();
  for (const fileId of [...new Set(fileIds.filter(Boolean))]) {
    const row = await db.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ?").bind(fileId).first();
    counts.set(fileId, Number(row?.count || 0));
  }
  return counts;
}

export async function countDerivedDoctorsByFile(db, fileIds = []) {
  if (!db?.prepare || !fileIds?.length) return new Map();
  await ensureCalendarSchema(db);
  const counts = new Map();
  for (const fileId of [...new Set(fileIds.filter(Boolean))]) {
    const row = await db.prepare("SELECT COUNT(*) AS count FROM roster_file_doctors WHERE file_id = ?").bind(fileId).first();
    counts.set(fileId, Number(row?.count || 0));
  }
  return counts;
}

export async function queryCoworkerEvents(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const start = String(options.startDate || options.date || "0000-01-01");
  const end = String(options.endDate || options.date || "9999-12-31");
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes);
  const excludeKeys = new Set((options.excludeDoctorKeys || []).filter(Boolean));
  const sourceSql = sourceTypes.length ? `AND roster_events.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
  const rows = await db.prepare(`
    SELECT doctor_key, display_name, source_type, event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.start_date <= ?
      AND roster_events.end_date >= ?
      ${sourceSql}
    ORDER BY roster_events.display_name, roster_events.start_ts
  `).bind(end, start, ...sourceTypes).all();
  return (rows.results || [])
    .filter((row) => !excludeKeys.has(row.doctor_key))
    .map((row) => ({
      doctorKey: row.doctor_key,
      displayName: row.display_name,
      sourceType: row.source_type,
      event: parseEvent(row.event_json),
    }))
    .filter((row) => row.event);
}

export function buildPreviewFromDerivedEvents(events) {
  const safeEvents = (events || []).map((event) => ({ ...event }));
  return {
    ...previewSummary(safeEvents),
    events: safeEvents,
    review: safeEvents.map((event) => reviewItemForEvent(event)),
    issues: [],
    conflicts: [],
    imports: [],
    sources: {},
    lastParsed: new Date().toISOString(),
    derivedFromD1: true,
  };
}

function reviewItemForEvent(event) {
  const day = datePart(event.start);
  return {
    id: event.id,
    source: event.source,
    seniority: event.seniority || "",
    startDay: day,
    endDay: datePart(event.end || event.start) || day,
    rawValue: event.rawValue || event.title || "",
    normalizedTitle: event.title,
    suggestedTitle: event.title,
    overrideTitle: "",
    status: "derived",
    warnings: [],
    include: true,
    exportable: true,
    location: event.location || "",
    allDay: event.allDay === true,
    timeLabel: event.timeLabel || "",
  };
}

function sanitizeFileDoctors(doctors, fallbackSourceType) {
  return (Array.isArray(doctors) ? doctors : [])
    .map((doctor) => ({
      key: String(doctor?.key || "").trim(),
      displayName: String(doctor?.displayName || doctor?.key || "").trim(),
      sourceType: normalizeSourceType(doctor?.sourceType || fallbackSourceType),
    }))
    .filter((doctor) => doctor.key && doctor.displayName && doctor.sourceType);
}

function sanitizeSourceTypes(values) {
  return [...new Set((Array.isArray(values) ? values : [])
    .map(normalizeSourceType)
    .filter(Boolean))];
}

function normalizeSourceType(value) {
  const source = String(value || "").trim().toLowerCase();
  return SOURCE_TYPES.includes(source) ? source : "";
}

function datePart(value) {
  return String(value || "").slice(0, 10);
}

function parseEvent(value) {
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

async function bulkUpsertDoctors(db, sourceType, doctors, updatedAt) {
  for (const chunk of chunkRows(doctors.map((doctor) => [sourceType, doctor.key, doctor.displayName, updatedAt]), 20)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT INTO roster_doctors (source_type, doctor_key, display_name, updated_at)
      VALUES ${chunk.map(() => "(?, ?, ?, ?)").join(", ")}
      ON CONFLICT(source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        updated_at = excluded.updated_at
    `).bind(...chunk.flat()).run();
  }
}

async function bulkInsertFileDoctors(db, fileId, sourceType, doctors) {
  for (const chunk of chunkRows(doctors.map((doctor) => [fileId, sourceType, doctor.key, doctor.displayName]), 20)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT INTO roster_file_doctors (file_id, source_type, doctor_key, display_name)
      VALUES ${chunk.map(() => "(?, ?, ?, ?)").join(", ")}
      ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET display_name = excluded.display_name
    `).bind(...chunk.flat()).run();
  }
}

async function bulkInsertEvents(db, rows) {
  for (const chunk of chunkRows(rows, 5)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT INTO roster_events (
        id, file_id, source_type, doctor_key, display_name, start_date, end_date, start_ts, end_ts,
        title, raw_value, seniority, location, all_day, time_label, event_json
      ) VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()).run();
  }
}

function chunkRows(rows, size) {
  const chunks = [];
  for (let index = 0; index < rows.length; index += size) {
    chunks.push(rows.slice(index, index + size));
  }
  return chunks;
}

function sanitizeEvent(value) {
  if (!value || typeof value !== "object" || !value.id || !value.start) return null;
  return {
    id: String(value.id),
    source: String(value.source || ""),
    sources: Array.isArray(value.sources) ? value.sources.map((item) => String(item || "")).filter(Boolean) : undefined,
    seniority: String(value.seniority || ""),
    title: String(value.title || ""),
    allDay: value.allDay === true,
    start: String(value.start || ""),
    end: String(value.end || value.start || ""),
    location: String(value.location || ""),
    rawValue: String(value.rawValue || ""),
    timeLabel: String(value.timeLabel || ""),
    monthKey: String(value.monthKey || ""),
  };
}
