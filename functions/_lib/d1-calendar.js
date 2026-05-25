import {
  buildRosterViewFromStoredImports,
  defaultSettings,
  previewSummary,
  serializeEvent,
} from "./roster.js";

const SOURCE_TYPES = ["mmc", "ddh", "casey", "mch"];
const ensuredCalendarDbs = new WeakSet();
const pendingCalendarSchemaEnsures = new WeakMap();

export function hasCalendarDb(env) {
  return Boolean(env?.ROSTER_DB?.prepare);
}

export async function ensureCalendarSchema(db) {
  if (!db?.prepare) return false;
  if (ensuredCalendarDbs.has(db)) return true;
  const pending = pendingCalendarSchemaEnsures.get(db);
  if (pending) return pending;
  const ensurePromise = ensureCalendarSchemaUncached(db)
    .then(() => {
      ensuredCalendarDbs.add(db);
      pendingCalendarSchemaEnsures.delete(db);
      return true;
    })
    .catch((error) => {
      pendingCalendarSchemaEnsures.delete(db);
      throw error;
    });
  pendingCalendarSchemaEnsures.set(db, ensurePromise);
  return ensurePromise;
}

async function ensureCalendarSchemaUncached(db) {
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
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events (source_type, start_date, end_date)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_file ON roster_events (file_id)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_issues (
      id TEXT PRIMARY KEY,
      file_id TEXT NOT NULL,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      start_date TEXT NOT NULL DEFAULT '',
      raw_value TEXT NOT NULL DEFAULT '',
      seniority TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT '',
      message TEXT NOT NULL DEFAULT '',
      resolution_type TEXT NOT NULL DEFAULT '',
      suggested_title TEXT NOT NULL DEFAULT '',
      time_label TEXT NOT NULL DEFAULT '',
      issue_json TEXT NOT NULL
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_issues_doctor_date ON roster_issues (doctor_key, start_date)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_issues_file ON roster_issues (file_id)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_profiles (
      email TEXT PRIMARY KEY,
      real_name TEXT NOT NULL DEFAULT '',
      role TEXT NOT NULL DEFAULT 'user',
      insights_enabled INTEGER NOT NULL DEFAULT 0,
      subscription_token TEXT NOT NULL DEFAULT '',
      password_salt TEXT NOT NULL DEFAULT '',
      password_hash TEXT NOT NULL DEFAULT '',
      admin_issues_json TEXT NOT NULL DEFAULT '[]',
      local_parser_extensions_json TEXT NOT NULL DEFAULT '[]',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await ensureColumn(db, "account_profiles", "password_salt", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "account_profiles", "password_hash", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "account_profiles", "admin_issues_json", "TEXT NOT NULL DEFAULT '[]'");
  await ensureColumn(db, "account_profiles", "local_parser_extensions_json", "TEXT NOT NULL DEFAULT '[]'");
  await ensureColumn(db, "account_profiles", "created_at", "TEXT NOT NULL DEFAULT ''");
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_claims (
      email TEXT NOT NULL,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      matched_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (email, source_type, doctor_key)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_account_claims_doctor ON account_claims (source_type, doctor_key)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_states (
      email TEXT PRIMARY KEY,
      session_json TEXT NOT NULL DEFAULT '{}',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_hospital_locations (
      email TEXT NOT NULL,
      source_type TEXT NOT NULL,
      location TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (email, source_type)
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS canonical_doctors (
      canonical_key TEXT PRIMARY KEY,
      display_name TEXT NOT NULL,
      source_type TEXT NOT NULL DEFAULT '',
      source_types_json TEXT NOT NULL DEFAULT '[]',
      aliases_json TEXT NOT NULL DEFAULT '[]',
      has_events INTEGER NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS custom_events (
      owner_email TEXT NOT NULL,
      id TEXT NOT NULL,
      title TEXT NOT NULL,
      start_date TEXT NOT NULL,
      end_date TEXT NOT NULL,
      all_day INTEGER NOT NULL DEFAULT 0,
      start_time TEXT NOT NULL DEFAULT '',
      end_time TEXT NOT NULL DEFAULT '',
      location TEXT NOT NULL DEFAULT '',
      include INTEGER NOT NULL DEFAULT 1,
      updated_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (owner_email, id)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_custom_events_owner_range ON custom_events (owner_email, start_date, end_date)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS doctor_profiles (
      profile_id TEXT PRIMARY KEY,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      source_types_json TEXT NOT NULL DEFAULT '[]',
      state_json TEXT NOT NULL DEFAULT '{}',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_doctor_profiles_doctor ON doctor_profiles (doctor_key)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS subscription_tokens (
      token TEXT PRIMARY KEY,
      email TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_subscription_tokens_email ON subscription_tokens (email)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS parser_rules (
      id TEXT PRIMARY KEY,
      scope TEXT NOT NULL DEFAULT 'global',
      email TEXT NOT NULL DEFAULT '',
      source_type TEXT NOT NULL,
      seniority TEXT NOT NULL DEFAULT '',
      code TEXT NOT NULL,
      title TEXT NOT NULL DEFAULT '',
      rule_json TEXT NOT NULL DEFAULT '{}',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_parser_rules_scope ON parser_rules (scope, email, source_type)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS parser_rule_suggestions (
      id TEXT PRIMARY KEY,
      email TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'pending',
      suggestion_json TEXT NOT NULL DEFAULT '{}',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS issue_dismissals (
      email TEXT NOT NULL,
      fingerprint TEXT NOT NULL,
      dismissed_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (email, fingerprint)
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS console_messages (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      actor_email TEXT NOT NULL DEFAULT '',
      message TEXT NOT NULL DEFAULT '',
      is_error INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_console_messages_created_at ON console_messages (created_at DESC, id DESC)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS issue_ignores (
      fingerprint TEXT PRIMARY KEY,
      ignored_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS raw_roster_files (
      file_id TEXT PRIMARY KEY,
      name TEXT NOT NULL DEFAULT '',
      source_type TEXT NOT NULL DEFAULT '',
      size INTEGER NOT NULL DEFAULT 0,
      last_modified INTEGER NOT NULL DEFAULT 0,
      object_key TEXT NOT NULL DEFAULT '',
      type TEXT NOT NULL DEFAULT '',
      data_url TEXT NOT NULL DEFAULT '',
      uploaded_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_daily_presence (
      date TEXT NOT NULL,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL,
      event_id TEXT NOT NULL,
      PRIMARY KEY (date, source_type, doctor_key, event_id)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_daily_presence_doctor_date ON roster_daily_presence (doctor_key, date)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_daily_presence_date_source ON roster_daily_presence (date, source_type)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS snapshot_registry (
      owner_type TEXT NOT NULL,
      owner_id TEXT NOT NULL,
      doctor_key TEXT NOT NULL DEFAULT '',
      range_key TEXT NOT NULL,
      requested_revision TEXT NOT NULL DEFAULT '',
      built_revision TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'missing',
      artifact_key TEXT NOT NULL DEFAULT '',
      built_at TEXT NOT NULL DEFAULT '',
      size_bytes INTEGER NOT NULL DEFAULT 0,
      build_ms INTEGER NOT NULL DEFAULT 0,
      last_error TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (owner_type, owner_id, doctor_key, range_key)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_snapshot_registry_owner ON snapshot_registry (owner_type, owner_id, updated_at DESC)").run();
  await ensureColumn(db, "raw_roster_files", "name", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "raw_roster_files", "source_type", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "raw_roster_files", "size", "INTEGER NOT NULL DEFAULT 0");
  await ensureColumn(db, "raw_roster_files", "last_modified", "INTEGER NOT NULL DEFAULT 0");
  await ensureColumn(db, "raw_roster_files", "type", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "raw_roster_files", "data_url", "TEXT NOT NULL DEFAULT ''");
  return true;
}

async function ensureColumn(db, table, column, definition) {
  const info = await db.prepare(`PRAGMA table_info(${table})`).all();
  if ((info.results || []).some((row) => row.name === column)) return;
  await db.prepare(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`).run();
}

export async function upsertDerivedRosterFile(db, file, storedImport) {
  if (!db?.prepare || !file?.id || !storedImport?.dataUrl) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType || storedImport.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const doctors = sanitizeFileDoctors(file.doctors, sourceType);
  const derivedByDoctor = [];
  let totalEvents = 0;
  for (const doctor of doctors) {
    const view = await buildRosterViewFromStoredImports([{ ...storedImport, sourceType, id: file.id, repoId: file.id }], doctor.key, defaultSettings(), {}, {}, []);
    const events = view.events.map(serializeEvent);
    const issues = (view.issues || []).map(sanitizeIssue).filter(Boolean);
    derivedByDoctor.push({ doctor, events, issues });
    totalEvents += events.length;
  }
  if (!doctors.length || totalEvents <= 0) return { ok: false, reason: "zero-events", doctors: doctors.length, events: totalEvents };
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
  await db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(file.id).run();
  for (const { doctor, events, issues } of derivedByDoctor) {
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
    for (const event of events) {
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
    for (const issue of issues) {
      await db.prepare(`
        INSERT INTO roster_issues (
          id, file_id, source_type, doctor_key, display_name, start_date, raw_value, seniority,
          status, message, resolution_type, suggested_title, time_label, issue_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(
        `${file.id}:${doctor.key}:${issue.id}`,
        file.id,
        sourceType,
        doctor.key,
        doctor.displayName,
        String(issue.startDay || issue.date || ""),
        String(issue.rawValue || ""),
        String(issue.seniority || ""),
        String(issue.status || ""),
        String(issue.message || ""),
        String(issue.resolutionType || ""),
        String(issue.suggestedTitle || ""),
        String(issue.timeLabel || ""),
        JSON.stringify(issue),
      ).run();
    }
  }
  const eventsByDoctorForPresence = Object.fromEntries(
    derivedByDoctor.map(({ doctor, events }) => [doctor.key, events])
  );
  await deleteDailyPresenceForFile(db, file.id);
  await populateDailyPresenceForFile(db, file.id, eventsByDoctorForPresence, {
    sourceType,
    doctors,
  });
  return { ok: true, doctors: doctors.length, events: totalEvents };
}

export async function replaceDerivedRosterFile(db, file, doctors, eventsByDoctor, issuesByDoctor = {}) {
  if (!db?.prepare || !file?.id) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const safeDoctors = sanitizeFileDoctors(doctors, sourceType);
  const parsedAt = new Date().toISOString();
  const statements = [db.prepare(`
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
  )];
  statements.push(
    db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(file.id),
    ...bulkUpsertDoctorStatements(db, sourceType, safeDoctors, parsedAt),
    ...bulkInsertFileDoctorStatements(db, file.id, sourceType, safeDoctors),
  );
  const eventRows = [];
  const issueRows = [];
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
    const issues = Array.isArray(issuesByDoctor?.[doctor.key]) ? issuesByDoctor[doctor.key] : [];
    for (const issue of issues.map(sanitizeIssue).filter(Boolean)) {
      issueRows.push([
        `${file.id}:${doctor.key}:${issue.id}`,
        file.id,
        sourceType,
        doctor.key,
        doctor.displayName,
        String(issue.startDay || issue.date || ""),
        String(issue.rawValue || ""),
        String(issue.seniority || ""),
        String(issue.status || ""),
        String(issue.message || ""),
        String(issue.resolutionType || ""),
        String(issue.suggestedTitle || ""),
        String(issue.timeLabel || ""),
        JSON.stringify(issue),
      ]);
    }
  }
  statements.push(...bulkInsertEventStatements(db, eventRows));
  statements.push(...bulkInsertIssueStatements(db, issueRows));
  await deleteDailyPresenceForFile(db, file.id);
  await runTransactionalBatch(db, statements);
  await populateDailyPresenceForFile(db, file.id, eventsByDoctor, {
    sourceType,
    doctors: safeDoctors,
  });
  return { ok: true, doctors: safeDoctors.length, events: eventRows.length, issues: issueRows.length };
}

export async function setDerivedRosterFileActive(db, fileId, active) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("UPDATE roster_files SET active = ? WHERE id = ?").bind(active ? 1 : 0, fileId).run();
  if (!active) {
    await deleteDailyPresenceForFile(db, fileId);
    return;
  }
  await rebuildDailyPresenceForFile(db, fileId);
}

export async function deleteDerivedRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(fileId).run();
  await deleteDailyPresenceForFile(db, fileId);
  await db.prepare("DELETE FROM roster_files WHERE id = ?").bind(fileId).run();
}

export async function upsertRawRosterFile(db, file, raw = {}) {
  if (!db?.prepare || !file?.id || (!raw?.dataUrl && !raw?.objectKey)) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  await db.prepare(`
    INSERT INTO raw_roster_files (file_id, name, source_type, size, last_modified, object_key, type, data_url, uploaded_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(file_id) DO UPDATE SET
      name = excluded.name,
      source_type = excluded.source_type,
      size = excluded.size,
      last_modified = excluded.last_modified,
      object_key = excluded.object_key,
      type = excluded.type,
      data_url = excluded.data_url,
      uploaded_at = excluded.uploaded_at
  `).bind(
    file.id,
    String(file.name || fileNameFromRawRosterFileId(file.id) || "roster.xlsx"),
    normalizeSourceType(file.sourceType || raw.sourceType || "") || inferSourceTypeFromRosterFileName(file.name || file.id),
    Number(file.size || 0),
    Number(file.lastModified || file.last_modified || 0),
    String(raw.objectKey || ""),
    String(raw.type || ""),
    String(raw.dataUrl || ""),
    String(raw.uploadedAt || new Date().toISOString()),
  ).run();
  return { ok: true };
}

export async function queryRawRosterFiles(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT file_id, name, source_type, size, last_modified, object_key, type, data_url, uploaded_at
    FROM raw_roster_files
    ORDER BY uploaded_at, file_id
  `).all();
  return (rows.results || []).map(rawRosterFileFromRow).filter(Boolean);
}

export async function loadRawRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare(`
    SELECT file_id, name, source_type, size, last_modified, object_key, type, data_url, uploaded_at
    FROM raw_roster_files
    WHERE file_id = ?
  `).bind(fileId).first();
  return rawRosterFileFromRow(row);
}

export async function queryRosterFileDoctorsForKeys(db, doctorKeys = []) {
  if (!db?.prepare || !doctorKeys?.length) return [];
  await ensureCalendarSchema(db);
  const keys = [...new Set(doctorKeys.map((key) => String(key || "").trim()).filter(Boolean))];
  if (!keys.length) return [];
  const placeholders = keys.map(() => "?").join(", ");
  const rows = await db.prepare(`
    SELECT
      roster_file_doctors.file_id AS file_id,
      roster_files.name AS file_name,
      roster_files.source_type AS file_source_type,
      roster_files.active AS active,
      roster_file_doctors.source_type AS source_type,
      roster_file_doctors.doctor_key AS doctor_key,
      roster_file_doctors.display_name AS display_name,
      (SELECT COUNT(*) FROM roster_events
        WHERE roster_events.file_id = roster_file_doctors.file_id
          AND roster_events.doctor_key = roster_file_doctors.doctor_key) AS event_count
    FROM roster_file_doctors
    INNER JOIN roster_files ON roster_files.id = roster_file_doctors.file_id
    WHERE roster_files.active = 1
      AND roster_file_doctors.doctor_key IN (${placeholders})
    ORDER BY roster_files.added_at, roster_files.name, roster_file_doctors.display_name
  `).bind(...keys).all();
  return (rows.results || [])
    .map((row) => ({
      fileId: String(row.file_id || "").trim(),
      fileName: String(row.file_name || "").trim(),
      fileSourceType: String(row.file_source_type || "").trim().toLowerCase(),
      active: row.active !== 0,
      sourceType: String(row.source_type || "").trim().toLowerCase(),
      doctorKey: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      eventCount: Number(row.event_count || 0),
    }))
    .filter((row) => row.fileId && row.doctorKey && row.displayName && SOURCE_TYPES.includes(row.sourceType));
}

export async function deleteRawRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("DELETE FROM raw_roster_files WHERE file_id = ?").bind(fileId).run();
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
  return mergeDuplicateLeaveEvents((rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean));
}

export async function queryDoctorSeniorities(db, doctorKeys = []) {
  if (!db?.prepare || !doctorKeys?.length) return [];
  await ensureCalendarSchema(db);
  const keys = [...new Set(doctorKeys.map((key) => String(key || "").trim()).filter(Boolean))];
  if (!keys.length) return [];
  const placeholders = keys.map(() => "?").join(", ");
  const rows = await db.prepare(`
    SELECT DISTINCT roster_events.seniority AS seniority
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.doctor_key IN (${placeholders})
      AND roster_events.seniority <> ''
    ORDER BY roster_events.seniority
  `).bind(...keys).all();
  return (rows.results || []).map((row) => String(row.seniority || "").trim()).filter(Boolean);
}

export async function queryDoctorEventsForFileDoctorPairs(db, pairs = [], options = {}) {
  if (!db?.prepare || !pairs?.length) return [];
  await ensureCalendarSchema(db);
  const safePairs = uniqueFileDoctorPairs(pairs);
  if (!safePairs.length) return [];
  const pairSql = safePairs.map(() => "(roster_events.file_id = ? AND roster_events.doctor_key = ?)").join(" OR ");
  const pairArgs = safePairs.flatMap((pair) => [pair.fileId, pair.doctorKey]);
  const start = String(options.startDate || "0000-01-01");
  const end = String(options.endDate || "9999-12-31");
  const rows = await db.prepare(`
    SELECT event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND (${pairSql})
      AND roster_events.start_date <= ?
      AND roster_events.end_date >= ?
    ORDER BY roster_events.start_ts, roster_events.source_type, roster_events.title
  `).bind(...pairArgs, end, start).all();
  return mergeDuplicateLeaveEvents((rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean));
}

export async function queryDoctorIssues(db, doctorKeys, options = {}) {
  if (!db?.prepare || !doctorKeys?.length) return [];
  await ensureCalendarSchema(db);
  const keys = [...new Set(doctorKeys.filter(Boolean))];
  if (!keys.length) return [];
  const placeholders = keys.map(() => "?").join(", ");
  const start = String(options.startDate || "0000-01-01");
  const end = String(options.endDate || "9999-12-31");
  const rows = await db.prepare(`
    SELECT issue_json
    FROM roster_issues
    INNER JOIN roster_files ON roster_files.id = roster_issues.file_id
    WHERE roster_files.active = 1
      AND roster_issues.doctor_key IN (${placeholders})
      AND roster_issues.start_date <= ?
      AND roster_issues.start_date >= ?
    ORDER BY roster_issues.start_date, roster_issues.source_type, roster_issues.raw_value
  `).bind(...keys, end, start).all();
  return (rows.results || []).map((row) => parseIssue(row.issue_json)).filter(Boolean);
}

export async function queryDoctorIssuesForFileDoctorPairs(db, pairs = [], options = {}) {
  if (!db?.prepare || !pairs?.length) return [];
  await ensureCalendarSchema(db);
  const safePairs = uniqueFileDoctorPairs(pairs);
  if (!safePairs.length) return [];
  const pairSql = safePairs.map(() => "(roster_issues.file_id = ? AND roster_issues.doctor_key = ?)").join(" OR ");
  const pairArgs = safePairs.flatMap((pair) => [pair.fileId, pair.doctorKey]);
  const start = String(options.startDate || "0000-01-01");
  const end = String(options.endDate || "9999-12-31");
  const rows = await db.prepare(`
    SELECT issue_json
    FROM roster_issues
    INNER JOIN roster_files ON roster_files.id = roster_issues.file_id
    WHERE roster_files.active = 1
      AND (${pairSql})
      AND roster_issues.start_date <= ?
      AND roster_issues.start_date >= ?
    ORDER BY roster_issues.start_date, roster_issues.source_type, roster_issues.raw_value
  `).bind(...pairArgs, end, start).all();
  return (rows.results || []).map((row) => parseIssue(row.issue_json)).filter(Boolean);
}

export async function queryRosterFileDoctors(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      roster_file_doctors.file_id AS file_id,
      roster_files.name AS file_name,
      roster_files.source_type AS file_source_type,
      roster_files.active AS active,
      roster_file_doctors.source_type AS source_type,
      roster_file_doctors.doctor_key AS doctor_key,
      roster_file_doctors.display_name AS display_name,
      (SELECT COUNT(*) FROM roster_events
        WHERE roster_events.file_id = roster_file_doctors.file_id
          AND roster_events.doctor_key = roster_file_doctors.doctor_key) AS event_count
    FROM roster_file_doctors
    INNER JOIN roster_files ON roster_files.id = roster_file_doctors.file_id
    WHERE roster_files.active = 1
    ORDER BY roster_files.added_at, roster_files.name, roster_file_doctors.display_name
  `).all();
  return (rows.results || [])
    .map((row) => ({
      fileId: String(row.file_id || "").trim(),
      fileName: String(row.file_name || "").trim(),
      fileSourceType: String(row.file_source_type || "").trim().toLowerCase(),
      active: row.active !== 0,
      sourceType: String(row.source_type || "").trim().toLowerCase(),
      doctorKey: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      eventCount: Number(row.event_count || 0),
    }))
    .filter((row) => row.fileId && row.doctorKey && row.displayName && SOURCE_TYPES.includes(row.sourceType));
}

export async function queryRosterDoctors(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT DISTINCT
      roster_file_doctors.source_type AS source_type,
      roster_file_doctors.doctor_key AS doctor_key,
      roster_file_doctors.display_name AS display_name
    FROM roster_file_doctors
    INNER JOIN roster_files ON roster_files.id = roster_file_doctors.file_id
    WHERE roster_files.active = 1
    ORDER BY roster_file_doctors.display_name, roster_file_doctors.source_type
  `).all();
  return (rows.results || [])
    .map((row) => ({
      key: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      sourceType: String(row.source_type || "").trim().toLowerCase(),
    }))
    .filter((doctor) => doctor.key && doctor.displayName && SOURCE_TYPES.includes(doctor.sourceType));
}

export async function queryRosterFiles(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const includeInactive = options.includeInactive === true;
  const rows = await db.prepare(`
    SELECT
      roster_files.id AS id,
      roster_files.name AS name,
      roster_files.source_type AS source_type,
      roster_files.active AS active,
      roster_files.size AS size,
      roster_files.last_modified AS last_modified,
      roster_files.added_at AS added_at,
      roster_files.uploaded_at AS uploaded_at,
      roster_files.uploaded_by AS uploaded_by,
      (SELECT COUNT(*) FROM roster_file_doctors WHERE roster_file_doctors.file_id = roster_files.id) AS doctor_count,
      (SELECT COUNT(*) FROM roster_events WHERE roster_events.file_id = roster_files.id) AS event_count
    FROM roster_files
    ${includeInactive ? "" : "WHERE roster_files.active = 1"}
    ORDER BY roster_files.added_at, roster_files.name
  `).all();
  const doctorRows = await db.prepare(`
    SELECT file_id, source_type, doctor_key, display_name
    FROM roster_file_doctors
    ORDER BY display_name, source_type
  `).all();
  const doctorsByFile = new Map();
  for (const row of doctorRows.results || []) {
    const fileId = String(row.file_id || "").trim();
    if (!doctorsByFile.has(fileId)) doctorsByFile.set(fileId, []);
    doctorsByFile.get(fileId).push({
      key: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      sourceType: String(row.source_type || "").trim().toLowerCase(),
    });
  }
  return (rows.results || []).map((row) => ({
    id: String(row.id || "").trim(),
    name: String(row.name || "roster.xlsx"),
    sourceType: String(row.source_type || "").trim().toLowerCase(),
    active: row.active !== 0,
    size: Number(row.size || 0),
    lastModified: Number(row.last_modified || 0),
    addedAt: String(row.added_at || ""),
    uploadedAt: String(row.uploaded_at || ""),
    uploadedBy: String(row.uploaded_by || ""),
    doctors: doctorsByFile.get(String(row.id || "").trim()) || [],
    expectedDoctors: Number(row.doctor_count || 0),
    indexedDoctors: Number(row.doctor_count || 0),
    eventCount: Number(row.event_count || 0),
    derivedFromD1: true,
  })).filter((file) => file.id && SOURCE_TYPES.includes(file.sourceType));
}

export async function queryRosterFileRanges(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const includeInactive = options.includeInactive === true;
  const rows = await db.prepare(`
    SELECT
      roster_files.id AS id,
      roster_files.name AS name,
      roster_files.source_type AS source_type,
      roster_files.active AS active,
      roster_files.last_modified AS last_modified,
      roster_files.added_at AS added_at,
      roster_files.uploaded_at AS uploaded_at,
      MIN(roster_events.start_date) AS start_date,
      MAX(roster_events.start_date) AS coverage_end_date,
      MAX(roster_events.end_date) AS end_date,
      COUNT(roster_events.id) AS event_count
    FROM roster_files
    LEFT JOIN roster_events ON roster_events.file_id = roster_files.id
    ${includeInactive ? "" : "WHERE roster_files.active = 1"}
    GROUP BY roster_files.id
    ORDER BY roster_files.source_type, start_date, roster_files.name
  `).all();
  return (rows.results || []).map((row) => ({
    id: String(row.id || "").trim(),
    name: String(row.name || "roster.xlsx"),
    sourceType: String(row.source_type || "").trim().toLowerCase(),
    active: row.active !== 0,
    lastModified: Number(row.last_modified || 0),
    addedAt: String(row.added_at || ""),
    uploadedAt: String(row.uploaded_at || ""),
    startDate: String(row.start_date || ""),
    coverageEndDate: String(row.coverage_end_date || ""),
    endDate: String(row.end_date || ""),
    eventCount: Number(row.event_count || 0),
  })).filter((file) => file.id && SOURCE_TYPES.includes(file.sourceType));
}

export async function queryRosterFileRefsForDoctors(db, doctorKeys = []) {
  if (!db?.prepare || !doctorKeys?.length) return [];
  await ensureCalendarSchema(db);
  const keys = [...new Set(doctorKeys.filter(Boolean))];
  if (!keys.length) return [];
  const rows = await db.prepare(`
    SELECT DISTINCT
      roster_files.id AS id,
      roster_files.name AS name,
      roster_files.source_type AS source_type,
      roster_files.active AS active,
      roster_files.size AS size,
      roster_files.last_modified AS last_modified,
      roster_files.added_at AS added_at,
      roster_files.uploaded_at AS uploaded_at,
      roster_files.uploaded_by AS uploaded_by
    FROM roster_files
    INNER JOIN roster_file_doctors ON roster_file_doctors.file_id = roster_files.id
    WHERE roster_files.active = 1
      AND roster_file_doctors.doctor_key IN (${keys.map(() => "?").join(", ")})
    ORDER BY roster_files.added_at, roster_files.name
  `).bind(...keys).all();
  return (rows.results || []).map((row) => ({
    id: String(row.id || "").trim(),
    repoId: String(row.id || "").trim(),
    name: String(row.name || "roster.xlsx"),
    sourceType: String(row.source_type || "").trim().toLowerCase(),
    active: row.active !== 0,
    size: Number(row.size || 0),
    lastModified: Number(row.last_modified || 0),
    addedAt: String(row.added_at || ""),
    uploadedAt: String(row.uploaded_at || ""),
    uploadedBy: String(row.uploaded_by || ""),
  })).filter((file) => file.id && SOURCE_TYPES.includes(file.sourceType));
}

export async function queryActiveRosterFileRefs(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      id,
      name,
      source_type,
      active,
      size,
      last_modified,
      added_at,
      uploaded_at,
      uploaded_by
    FROM roster_files
    WHERE active = 1
    ORDER BY added_at, name
  `).all();
  return (rows.results || []).map((row) => ({
    id: String(row.id || "").trim(),
    repoId: String(row.id || "").trim(),
    name: String(row.name || "roster.xlsx"),
    sourceType: String(row.source_type || "").trim().toLowerCase(),
    active: row.active !== 0,
    size: Number(row.size || 0),
    lastModified: Number(row.last_modified || 0),
    addedAt: String(row.added_at || ""),
    uploadedAt: String(row.uploaded_at || ""),
    uploadedBy: String(row.uploaded_by || ""),
  })).filter((file) => file.id && SOURCE_TYPES.includes(file.sourceType));
}

export async function queryCalendarRevision(db, ownerEmail = "") {
  if (!db?.prepare) return "";
  await ensureCalendarSchema(db);
  const email = normalizeEmail(ownerEmail);
  const roster = await db.prepare(`
    SELECT
      COUNT(*) AS active_file_count,
      COALESCE(MAX(parsed_at), '') AS max_parsed_at,
      COALESCE(MAX(uploaded_at), '') AS max_uploaded_at,
      COALESCE(MAX(last_modified), 0) AS max_last_modified
    FROM roster_files
    WHERE active = 1
  `).first();
  const accountState = email
    ? await db.prepare("SELECT session_json FROM account_states WHERE email = ?").bind(email).first()
    : null;
  const customEvents = email
    ? await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM custom_events WHERE owner_email = ?").bind(email).first()
    : null;
  const claims = email
    ? await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM account_claims WHERE email = ?").bind(email).first()
    : null;
  const locations = email
    ? await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM account_hospital_locations WHERE email = ?").bind(email).first()
    : null;
  const profiles = email
    ? await db.prepare(`
      SELECT COUNT(DISTINCT doctor_profiles.profile_id) AS count, COALESCE(MAX(doctor_profiles.updated_at), '') AS max_updated_at
      FROM doctor_profiles
      INNER JOIN account_claims ON account_claims.doctor_key = doctor_profiles.doctor_key
      WHERE account_claims.email = ?
    `).bind(email).first()
    : null;
  const materializedSession = materializedSessionStateForRevision(parseJsonObject(accountState?.session_json, {}));
  return [
    Number(roster?.active_file_count || 0),
    String(roster?.max_parsed_at || ""),
    String(roster?.max_uploaded_at || ""),
    Number(roster?.max_last_modified || 0),
    stableJsonStringify(materializedSession),
    Number(claims?.count || 0),
    String(claims?.max_updated_at || ""),
    Number(locations?.count || 0),
    String(locations?.max_updated_at || ""),
    Number(customEvents?.count || 0),
    String(customEvents?.max_updated_at || ""),
    Number(profiles?.count || 0),
    String(profiles?.max_updated_at || ""),
  ].join("|");
}

export async function upsertAccountMirror(db, record, options = {}) {
  if (!db?.prepare || !record?.email) return false;
  await ensureCalendarSchema(db);
  const email = normalizeEmail(record.email);
  const role = String(record.role || (email ? "user" : "") || "user");
  const updatedAt = new Date().toISOString();
  await db.prepare(`
    INSERT INTO account_profiles (
      email, real_name, role, insights_enabled, subscription_token, password_salt, password_hash,
      admin_issues_json, local_parser_extensions_json, created_at, updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(email) DO UPDATE SET
      real_name = excluded.real_name,
      role = excluded.role,
      insights_enabled = excluded.insights_enabled,
      subscription_token = excluded.subscription_token,
      password_salt = CASE WHEN excluded.password_salt <> '' THEN excluded.password_salt ELSE account_profiles.password_salt END,
      password_hash = CASE WHEN excluded.password_hash <> '' THEN excluded.password_hash ELSE account_profiles.password_hash END,
      admin_issues_json = excluded.admin_issues_json,
      local_parser_extensions_json = excluded.local_parser_extensions_json,
      updated_at = excluded.updated_at
  `).bind(
    email,
    String(record.realName || ""),
    role,
    record.insightsEnabled === true ? 1 : 0,
    String(record.subscriptionToken || ""),
    String(record.passwordSalt || ""),
    String(record.passwordHash || ""),
    JSON.stringify(Array.isArray(record.adminIssues) ? record.adminIssues : []),
    JSON.stringify(Array.isArray(record.localParserExtensions) ? record.localParserExtensions : []),
    String(record.createdAt || updatedAt),
    updatedAt,
  ).run();
  await db.prepare("DELETE FROM account_claims WHERE email = ?").bind(email).run();
  await db.prepare("DELETE FROM subscription_tokens WHERE email = ?").bind(email).run();
  if (record.subscriptionToken) {
    await db.prepare(`
      INSERT INTO subscription_tokens (token, email, created_at, updated_at)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(token) DO UPDATE SET
        email = excluded.email,
        updated_at = excluded.updated_at
    `).bind(String(record.subscriptionToken || ""), email, String(record.createdAt || updatedAt), updatedAt).run();
  }
  const claims = sanitizeAccountClaims(record.claims);
  for (const chunk of chunkRows(claims.map((claim) => [
    email,
    claim.sourceType,
    claim.key,
    claim.displayName,
    claim.matchedAt,
    updatedAt,
  ]), 20)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT INTO account_claims (email, source_type, doctor_key, display_name, matched_at, updated_at)
      VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?)").join(", ")}
      ON CONFLICT(email, source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        matched_at = excluded.matched_at,
        updated_at = excluded.updated_at
    `).bind(...chunk.flat()).run();
  }
  const preserveExistingState = options.preserveExistingState === true;
  const existingState = preserveExistingState
    ? await db.prepare("SELECT session_json FROM account_states WHERE email = ?").bind(email).first()
    : null;
  if (!existingState?.session_json) {
    const session = record?.state?.session && typeof record.state.session === "object" ? record.state.session : {};
    await db.prepare(`
      INSERT INTO account_states (email, session_json, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(email) DO UPDATE SET
        session_json = excluded.session_json,
        updated_at = excluded.updated_at
    `).bind(email, JSON.stringify(session), updatedAt).run();
  }
  await upsertAccountHospitalLocations(db, email, hospitalLocationsFromSession(record?.state?.session), { preserveExisting: preserveExistingState });
  return true;
}

export async function deleteAccountMirror(db, email) {
  if (!db?.prepare || !email) return;
  await ensureCalendarSchema(db);
  const normalizedEmail = normalizeEmail(email);
  await db.prepare("DELETE FROM account_claims WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM account_states WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM account_hospital_locations WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM custom_events WHERE owner_email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM subscription_tokens WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM account_profiles WHERE email = ?").bind(normalizedEmail).run();
}

export async function replaceCanonicalDoctors(db, doctors = []) {
  if (!db?.prepare) return false;
  await ensureCalendarSchema(db);
  const now = new Date().toISOString();
  const statements = [db.prepare("DELETE FROM canonical_doctors")];
  for (const doctor of doctors || []) {
    if (!doctor?.key || !doctor?.displayName) continue;
    statements.push(db.prepare(`
      INSERT INTO canonical_doctors (
        canonical_key, display_name, source_type, source_types_json, aliases_json, has_events, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `).bind(
      String(doctor.key || "").trim(),
      String(doctor.displayName || "").trim(),
      String(doctor.sourceType || doctor.sourceTypes?.[0] || "").trim().toLowerCase(),
      JSON.stringify(Array.isArray(doctor.sourceTypes) ? doctor.sourceTypes : []),
      JSON.stringify(Array.isArray(doctor.aliases) ? doctor.aliases : []),
      doctor.hasEvents === false ? 0 : 1,
      now,
    ));
  }
  await runTransactionalBatch(db, statements);
  return true;
}

export async function queryCanonicalDoctors(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT canonical_key, display_name, source_type, source_types_json, aliases_json, has_events
    FROM canonical_doctors
    ${options.includeZeroEventStandalone === false ? "WHERE has_events = 1" : ""}
    ORDER BY display_name, source_type
  `).all();
  return (rows.results || []).map((row) => {
    let sourceTypes = [];
    let aliases = [];
    try { sourceTypes = JSON.parse(row.source_types_json || "[]"); } catch {}
    try { aliases = JSON.parse(row.aliases_json || "[]"); } catch {}
    return {
      key: String(row.canonical_key || "").trim(),
      displayName: String(row.display_name || "").trim(),
      sourceType: String(row.source_type || "").trim().toLowerCase(),
      sourceTypes: Array.isArray(sourceTypes) ? sourceTypes : [],
      aliases: Array.isArray(aliases) ? aliases : [],
      hasEvents: Number(row.has_events || 0) === 1,
    };
  }).filter((doctor) => doctor.key && doctor.displayName);
}

export async function replaceAccountCustomEvents(db, ownerEmail, events = []) {
  if (!db?.prepare || !ownerEmail) return false;
  await ensureCalendarSchema(db);
  const email = normalizeEmail(ownerEmail);
  const now = new Date().toISOString();
  const statements = [db.prepare("DELETE FROM custom_events WHERE owner_email = ?").bind(email)];
  for (const event of events || []) {
    if (!event?.id || !event?.title || !event?.startDate || !event?.endDate) continue;
    statements.push(db.prepare(`
      INSERT INTO custom_events (
        owner_email, id, title, start_date, end_date, all_day, start_time, end_time, location, include, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      email,
      String(event.id),
      String(event.title),
      String(event.startDate).slice(0, 10),
      String(event.endDate).slice(0, 10),
      event.allDay === true ? 1 : 0,
      event.allDay === true ? "" : String(event.startTime || ""),
      event.allDay === true ? "" : String(event.endTime || ""),
      String(event.location || ""),
      event.include === false ? 0 : 1,
      now,
    ));
  }
  await runTransactionalBatch(db, statements);
  return true;
}

export async function queryAccountCustomEvents(db, ownerEmail) {
  if (!db?.prepare || !ownerEmail) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT id, title, start_date, end_date, all_day, start_time, end_time, location, include
    FROM custom_events
    WHERE owner_email = ?
    ORDER BY start_date, start_time, title, id
  `).bind(normalizeEmail(ownerEmail)).all();
  return (rows.results || []).map((row) => ({
    id: String(row.id || ""),
    ownerEmail: normalizeEmail(ownerEmail),
    title: String(row.title || ""),
    startDate: String(row.start_date || ""),
    endDate: String(row.end_date || ""),
    allDay: Number(row.all_day || 0) === 1,
    startTime: String(row.start_time || ""),
    endTime: String(row.end_time || ""),
    location: String(row.location || ""),
    include: Number(row.include ?? 1) !== 0,
  })).filter((event) => event.id && event.title && event.startDate && event.endDate);
}

export async function queryClaimedAccounts(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      account_profiles.email AS email,
      account_profiles.real_name AS real_name,
      account_profiles.role AS role,
      account_claims.source_type AS source_type,
      account_claims.doctor_key AS doctor_key,
      account_claims.display_name AS display_name,
      account_claims.matched_at AS matched_at
    FROM account_profiles
    LEFT JOIN account_claims ON account_claims.email = account_profiles.email
    ORDER BY account_profiles.email, account_claims.source_type, account_claims.display_name
  `).all();
  const accounts = new Map();
  for (const row of rows.results || []) {
    const email = normalizeEmail(row.email);
    if (!email || row.role === "creator" || row.role === "owner") continue;
    if (!accounts.has(email)) {
      accounts.set(email, {
        email,
        realName: String(row.real_name || "").trim(),
        claims: [],
      });
    }
    if (row.doctor_key && row.source_type) {
      accounts.get(email).claims.push({
        key: String(row.doctor_key || "").trim(),
        displayName: String(row.display_name || row.doctor_key || "").trim(),
        sourceType: String(row.source_type || "").trim().toLowerCase(),
        matchedAt: String(row.matched_at || ""),
      });
    }
  }
  return [...accounts.values()];
}

export async function loadAccountMirrorBySubscriptionToken(db, token) {
  if (!db?.prepare || !token) return null;
  await ensureCalendarSchema(db);
  const tokenRow = await db.prepare("SELECT email FROM subscription_tokens WHERE token = ?").bind(String(token || "").trim()).first();
  if (tokenRow?.email) {
    return await loadAccountMirror(db, tokenRow.email);
  }
  const rows = await db.prepare(`
    SELECT
      account_profiles.email AS email,
      account_profiles.real_name AS real_name,
      account_profiles.role AS role,
      account_profiles.insights_enabled AS insights_enabled,
      account_profiles.subscription_token AS subscription_token,
      account_claims.source_type AS source_type,
      account_claims.doctor_key AS doctor_key,
      account_claims.display_name AS display_name,
      account_claims.matched_at AS matched_at,
      account_states.session_json AS session_json
    FROM account_profiles
    LEFT JOIN account_claims ON account_claims.email = account_profiles.email
    LEFT JOIN account_states ON account_states.email = account_profiles.email
    WHERE account_profiles.subscription_token = ?
    ORDER BY account_claims.source_type, account_claims.display_name
  `).bind(String(token || "").trim()).all();
  const first = rows.results?.[0];
  if (!first?.email) return null;
  const claims = [];
  for (const row of rows.results || []) {
    if (!row.doctor_key || !row.source_type) continue;
    claims.push({
      key: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      sourceType: String(row.source_type || "").trim().toLowerCase(),
      matchedAt: String(row.matched_at || ""),
    });
  }
  let session = {};
  try {
    const parsed = first.session_json ? JSON.parse(first.session_json) : {};
    session = parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    session = {};
  }
  return {
    email: normalizeEmail(first.email),
    realName: String(first.real_name || "").trim(),
    role: String(first.role || "user"),
    insightsEnabled: first.insights_enabled === 1,
    subscriptionToken: String(first.subscription_token || ""),
    claims,
    state: {
      version: 1,
      imports: [],
      session,
      subscriptionFeeds: {},
    },
  };
}

export async function loadAccountMirror(db, email) {
  if (!db?.prepare || !email) return null;
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      account_profiles.email AS email,
      account_profiles.real_name AS real_name,
      account_profiles.role AS role,
      account_profiles.insights_enabled AS insights_enabled,
      account_profiles.subscription_token AS subscription_token,
      account_profiles.password_salt AS password_salt,
      account_profiles.password_hash AS password_hash,
      account_profiles.admin_issues_json AS admin_issues_json,
      account_profiles.local_parser_extensions_json AS local_parser_extensions_json,
      account_profiles.created_at AS created_at,
      account_profiles.updated_at AS updated_at,
      account_claims.source_type AS source_type,
      account_claims.doctor_key AS doctor_key,
      account_claims.display_name AS display_name,
      account_claims.matched_at AS matched_at,
      account_states.session_json AS session_json
    FROM account_profiles
    LEFT JOIN account_claims ON account_claims.email = account_profiles.email
    LEFT JOIN account_states ON account_states.email = account_profiles.email
    WHERE account_profiles.email = ?
    ORDER BY account_claims.source_type, account_claims.display_name
  `).bind(normalizeEmail(email)).all();
  return accountMirrorFromRows(rows.results || []);
}

export async function listAccountMirrors(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      account_profiles.email AS email,
      account_profiles.real_name AS real_name,
      account_profiles.role AS role,
      account_profiles.insights_enabled AS insights_enabled,
      account_profiles.subscription_token AS subscription_token,
      account_profiles.password_salt AS password_salt,
      account_profiles.password_hash AS password_hash,
      account_profiles.admin_issues_json AS admin_issues_json,
      account_profiles.local_parser_extensions_json AS local_parser_extensions_json,
      account_profiles.created_at AS created_at,
      account_profiles.updated_at AS updated_at,
      account_claims.source_type AS source_type,
      account_claims.doctor_key AS doctor_key,
      account_claims.display_name AS display_name,
      account_claims.matched_at AS matched_at,
      account_states.session_json AS session_json
    FROM account_profiles
    LEFT JOIN account_claims ON account_claims.email = account_profiles.email
    LEFT JOIN account_states ON account_states.email = account_profiles.email
    ORDER BY account_profiles.email, account_claims.source_type, account_claims.display_name
  `).all();
  const grouped = new Map();
  for (const row of rows.results || []) {
    const email = normalizeEmail(row.email);
    if (!email) continue;
    if (!grouped.has(email)) grouped.set(email, []);
    grouped.get(email).push(row);
  }
  return [...grouped.values()].map(accountMirrorFromRows).filter(Boolean);
}

function accountMirrorFromRows(rows) {
  const first = rows?.[0];
  if (!first?.email) return null;
  const claims = [];
  for (const row of rows || []) {
    if (!row.doctor_key || !row.source_type) continue;
    claims.push({
      key: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || row.doctor_key || "").trim(),
      sourceType: String(row.source_type || "").trim().toLowerCase(),
      matchedAt: String(row.matched_at || ""),
    });
  }
  let session = {};
  let adminIssues = [];
  let localParserExtensions = [];
  try {
    const parsed = first.session_json ? JSON.parse(first.session_json) : {};
    session = parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    session = {};
  }
  try {
    const parsed = first.admin_issues_json ? JSON.parse(first.admin_issues_json) : [];
    adminIssues = Array.isArray(parsed) ? parsed : [];
  } catch {
    adminIssues = [];
  }
  try {
    const parsed = first.local_parser_extensions_json ? JSON.parse(first.local_parser_extensions_json) : [];
    localParserExtensions = Array.isArray(parsed) ? parsed : [];
  } catch {
    localParserExtensions = [];
  }
  return {
    email: normalizeEmail(first.email),
    realName: String(first.real_name || "").trim(),
    role: String(first.role || "user"),
    insightsEnabled: first.insights_enabled === 1,
    subscriptionToken: String(first.subscription_token || ""),
    passwordSalt: String(first.password_salt || ""),
    passwordHash: String(first.password_hash || ""),
    adminIssues,
    localParserExtensions,
    claims: sanitizeAccountClaims(claims),
    createdAt: String(first.created_at || ""),
    updatedAt: String(first.updated_at || ""),
    state: {
      version: 1,
      imports: [],
      session,
      subscriptionFeeds: {},
    },
  };
}

export async function loadAccountStateMirror(db, email) {
  if (!db?.prepare || !email) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare("SELECT session_json FROM account_states WHERE email = ?").bind(normalizeEmail(email)).first();
  if (!row?.session_json) return null;
  try {
    const session = JSON.parse(row.session_json);
    return session && typeof session === "object" ? { session } : null;
  } catch {
    return null;
  }
}

export async function loadAccountHospitalLocations(db, email, fallbackSession = null) {
  if (!db?.prepare || !email) return hospitalLocationsFromSession(fallbackSession);
  await ensureCalendarSchema(db);
  const normalizedEmail = normalizeEmail(email);
  let rows = [];
  try {
    const result = await db.prepare(`
      SELECT source_type, location
      FROM account_hospital_locations
      WHERE email = ?
      ORDER BY source_type
    `).bind(normalizedEmail).all();
    rows = result.results || [];
  } catch {
    rows = [];
  }
  if (!rows.length) {
    const locations = hospitalLocationsFromSession(fallbackSession);
    await upsertAccountHospitalLocations(db, normalizedEmail, locations, { preserveExisting: true }).catch(() => null);
    return locations;
  }
  return normalizeHospitalLocationMap(Object.fromEntries(rows.map((row) => [String(row.source_type || "").toLowerCase(), row.location])));
}

export async function upsertAccountHospitalLocations(db, email, locations = {}, options = {}) {
  if (!db?.prepare || !email) return normalizeHospitalLocationMap(locations);
  await ensureCalendarSchema(db);
  const normalizedEmail = normalizeEmail(email);
  const next = normalizeHospitalLocationMap(locations);
  if (options.preserveExisting === true) {
    const rows = await db.prepare("SELECT source_type, location FROM account_hospital_locations WHERE email = ?").bind(normalizedEmail).all().catch(() => ({ results: [] }));
    const existing = Object.fromEntries((rows.results || []).map((row) => [String(row.source_type || "").toLowerCase(), String(row.location || "")]));
    for (const sourceType of SOURCE_TYPES) {
      if (existing[sourceType]) next[sourceType] = existing[sourceType];
    }
  }
  const updatedAt = new Date().toISOString();
  for (const sourceType of SOURCE_TYPES) {
    await db.prepare(`
      INSERT INTO account_hospital_locations (email, source_type, location, updated_at)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(email, source_type) DO UPDATE SET
        location = excluded.location,
        updated_at = excluded.updated_at
    `).bind(normalizedEmail, sourceType, next[sourceType], updatedAt).run();
  }
  return next;
}

export function hospitalLocationsFromSession(session = null) {
  const settings = session?.settings && typeof session.settings === "object" ? session.settings : {};
  return normalizeHospitalLocationMap({
    mmc: settings.defaultLocationMmc,
    ddh: settings.defaultLocationDdh,
    casey: settings.defaultLocationCasey,
    mch: settings.defaultLocationMch,
  });
}

export function mergeHospitalLocationsIntoSettings(settings = {}, locations = {}) {
  const normalized = normalizeHospitalLocationMap(locations);
  return {
    ...(settings || {}),
    defaultLocationMmc: normalized.mmc,
    defaultLocationDdh: normalized.ddh,
    defaultLocationCasey: normalized.casey,
    defaultLocationMch: normalized.mch,
  };
}

export function applyAccountHospitalLocations(events = [], locations = {}, options = {}) {
  if (options.includeLocations === false) return (events || []).map((event) => ({ ...event, location: "" }));
  const normalized = normalizeHospitalLocationMap(locations);
  return (events || []).map((event) => {
    const sourceType = eventSourceType(event);
    if (!SOURCE_TYPES.includes(sourceType) || isOffsiteClinicalSupportEvent(event) || isLocationlessRosterEvent(event)) return event;
    const location = normalized[sourceType] || "";
    return location ? { ...event, location } : event;
  });
}

function normalizeHospitalLocationMap(value = {}) {
  const defaults = defaultSettings();
  return {
    mmc: String(value.mmc || value.defaultLocationMmc || defaults.defaultLocationMmc).trim() || defaults.defaultLocationMmc,
    ddh: String(value.ddh || value.defaultLocationDdh || defaults.defaultLocationDdh).trim() || defaults.defaultLocationDdh,
    casey: String(value.casey || value.defaultLocationCasey || defaults.defaultLocationCasey).trim() || defaults.defaultLocationCasey,
    mch: String(value.mch || value.defaultLocationMch || defaults.defaultLocationMch).trim() || defaults.defaultLocationMch,
  };
}

function eventSourceType(event) {
  const source = String(event?.source || event?.sourceType || "").trim().toLowerCase();
  if (source === "casey") return "casey";
  return SOURCE_TYPES.includes(source) ? source : "";
}

function isOffsiteClinicalSupportEvent(event) {
  const sourceType = eventSourceType(event);
  if (sourceType !== "mmc" && sourceType !== "ddh") return false;
  const title = String(event?.title || "").replace(/^[A-Z]+:\s*/, "").trim().toUpperCase();
  const raw = String(event?.rawValue || "").trim().toUpperCase();
  if (sourceType === "mmc") return title === "CS" || raw === "CS";
  if (sourceType === "ddh") {
    if (title.includes("ONSITE") || raw.includes("ONSITE")) return false;
    return title === "CS" || raw === "CS" || raw.includes("OFFSITE") || raw.includes("NOT ONSITE") || raw.includes("CS/OFF");
  }
  return false;
}

function isLocationlessRosterEvent(event) {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  return /\b(leave|conference|cme|sick|study|exam|parental|phnw|public holiday)\b/.test(text);
}

export async function accountMirrorStatus(db) {
  if (!db?.prepare) return { unavailable: true, profiles: 0, claims: 0, states: 0, doctorProfiles: 0 };
  await ensureCalendarSchema(db);
  const profiles = await db.prepare("SELECT COUNT(*) AS count FROM account_profiles").first();
  const claims = await db.prepare("SELECT COUNT(*) AS count FROM account_claims").first();
  const states = await db.prepare("SELECT COUNT(*) AS count FROM account_states").first();
  const subscriptionTokens = await db.prepare("SELECT COUNT(*) AS count FROM account_profiles WHERE subscription_token <> ''").first();
  const doctorProfiles = await db.prepare("SELECT COUNT(*) AS count FROM doctor_profiles").first();
  return {
    unavailable: false,
    profiles: Number(profiles?.count || 0),
    claims: Number(claims?.count || 0),
    states: Number(states?.count || 0),
    subscriptionTokens: Number(subscriptionTokens?.count || 0),
    doctorProfiles: Number(doctorProfiles?.count || 0),
  };
}

export async function appendConsoleMessage(db, entry = {}) {
  if (!db?.prepare || !entry?.message) return false;
  await ensureCalendarSchema(db);
  await db.prepare(`
    INSERT INTO console_messages (actor_email, message, is_error, created_at)
    VALUES (?, ?, ?, ?)
  `).bind(
    normalizeEmail(entry.actorEmail || ""),
    String(entry.message || "").trim(),
    entry.isError === true ? 1 : 0,
    String(entry.createdAt || new Date().toISOString()),
  ).run();
  await db.prepare(`
    DELETE FROM console_messages
    WHERE id NOT IN (
      SELECT id FROM console_messages
      ORDER BY created_at DESC, id DESC
      LIMIT 50
    )
  `).run();
  return true;
}

export async function listConsoleMessages(db, limit = 50) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const safeLimit = Math.max(1, Math.min(Number(limit || 50) || 50, 50));
  const result = await db.prepare(`
    SELECT id, actor_email, message, is_error, created_at
    FROM console_messages
    ORDER BY created_at DESC, id DESC
    LIMIT ?
  `).bind(safeLimit).all();
  return (result?.results || []).map((row) => ({
    id: Number(row.id || 0),
    actorEmail: String(row.actor_email || ""),
    message: String(row.message || ""),
    isError: Number(row.is_error || 0) === 1,
    createdAt: String(row.created_at || ""),
  }));
}

export async function upsertDoctorProfileMirror(db, profile) {
  if (!db?.prepare || !profile?.profileId) return false;
  await ensureCalendarSchema(db);
  const now = new Date().toISOString();
  await db.prepare(`
    INSERT INTO doctor_profiles (profile_id, doctor_key, display_name, source_types_json, state_json, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(profile_id) DO UPDATE SET
      doctor_key = excluded.doctor_key,
      display_name = excluded.display_name,
      source_types_json = excluded.source_types_json,
      state_json = excluded.state_json,
      updated_at = excluded.updated_at
  `).bind(
    String(profile.profileId || "").trim(),
    String(profile.doctorKey || "").trim(),
    String(profile.displayName || "").trim(),
    JSON.stringify(sanitizeSourceTypes(profile.sourceTypes)),
    JSON.stringify(profile.state && typeof profile.state === "object" ? profile.state : {}),
    String(profile.createdAt || now),
    String(profile.updatedAt || now),
  ).run();
  return true;
}

export async function deleteDoctorProfileMirror(db, profileId) {
  if (!db?.prepare || !profileId) return;
  await ensureCalendarSchema(db);
  await db.prepare("DELETE FROM doctor_profiles WHERE profile_id = ?").bind(String(profileId || "").trim()).run();
}

export async function loadDoctorProfileMirror(db, profileId) {
  if (!db?.prepare || !profileId) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare("SELECT * FROM doctor_profiles WHERE profile_id = ?").bind(String(profileId || "").trim()).first();
  return doctorProfileFromRow(row);
}

export async function queryDoctorProfileMirrors(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare("SELECT * FROM doctor_profiles ORDER BY display_name, profile_id").all();
  return (rows.results || []).map(doctorProfileFromRow).filter(Boolean);
}

export async function countDerivedEventsByFile(db, fileIds = []) {
  if (!db?.prepare || !fileIds?.length) return new Map();
  await ensureCalendarSchema(db);
  const ids = [...new Set(fileIds.filter(Boolean))];
  const rows = await db.prepare(`
    SELECT file_id, COUNT(*) AS count
    FROM roster_events
    WHERE file_id IN (${ids.map(() => "?").join(", ")})
    GROUP BY file_id
  `).bind(...ids).all();
  const counts = new Map(ids.map((id) => [id, 0]));
  for (const row of rows.results || []) counts.set(String(row.file_id || ""), Number(row.count || 0));
  return counts;
}

export async function countDerivedEventsByFileDoctorPairs(db, pairs = []) {
  if (!db?.prepare || !pairs?.length) return new Map();
  await ensureCalendarSchema(db);
  const safePairs = uniqueFileDoctorPairs(pairs);
  const pairSql = safePairs.map(() => "(file_id = ? AND doctor_key = ?)").join(" OR ");
  const rows = await db.prepare(`
    SELECT file_id, doctor_key, COUNT(*) AS count
    FROM roster_events
    WHERE ${pairSql}
    GROUP BY file_id, doctor_key
  `).bind(...safePairs.flatMap((pair) => [pair.fileId, pair.doctorKey])).all();
  const counts = new Map(safePairs.map((pair) => [`${pair.fileId}:${pair.doctorKey}`, 0]));
  for (const row of rows.results || []) counts.set(`${row.file_id}:${row.doctor_key}`, Number(row.count || 0));
  return counts;
}

export async function countDerivedDoctorsByFile(db, fileIds = []) {
  if (!db?.prepare || !fileIds?.length) return new Map();
  await ensureCalendarSchema(db);
  const ids = [...new Set(fileIds.filter(Boolean))];
  const rows = await db.prepare(`
    SELECT file_id, COUNT(*) AS count
    FROM roster_file_doctors
    WHERE file_id IN (${ids.map(() => "?").join(", ")})
    GROUP BY file_id
  `).bind(...ids).all();
  const counts = new Map(ids.map((id) => [id, 0]));
  for (const row of rows.results || []) counts.set(String(row.file_id || ""), Number(row.count || 0));
  return counts;
}

export async function queryCoworkerEvents(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const start = String(options.startDate || options.date || "0000-01-01");
  const end = String(options.endDate || options.date || "9999-12-31");
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes);
  const excludeKeys = new Set((options.excludeDoctorKeys || []).filter(Boolean));
  const includeKeys = [...new Set((options.doctorKeys || []).filter(Boolean))];
  const overlapKeys = [...new Set((options.overlapDoctorKeys || []).filter(Boolean))];
  const sourceSql = sourceTypes.length ? `AND p.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
  const doctorSql = includeKeys.length ? `AND p.doctor_key IN (${includeKeys.map(() => "?").join(", ")})` : "";
  const excludeDoctorSql = excludeKeys.size ? `AND p.doctor_key NOT IN (${[...excludeKeys].map(() => "?").join(", ")})` : "";
  if (overlapKeys.length) {
    const rows = await db.prepare(`
      SELECT DISTINCT
        ev.event_json,
        p.doctor_key,
        p.display_name,
        p.source_type
      FROM roster_daily_presence AS mine
      INNER JOIN roster_daily_presence AS p
        ON p.date = mine.date
       AND p.source_type = mine.source_type
      INNER JOIN roster_events AS ev ON ev.id = p.event_id
      WHERE mine.doctor_key IN (${overlapKeys.map(() => "?").join(", ")})
        AND mine.date >= ?
        AND mine.date <= ?
        AND p.date >= ?
        AND p.date <= ?
        ${sourceSql}
        ${doctorSql}
        ${excludeDoctorSql}
      ORDER BY p.display_name, ev.start_ts
    `).bind(...overlapKeys, start, end, start, end, ...sourceTypes, ...includeKeys, ...excludeKeys).all();
    return rowsToCoworkerEvents(rows);
  }
  const rows = await db.prepare(`
    SELECT DISTINCT
      ev.event_json,
      p.doctor_key,
      p.display_name,
      p.source_type
    FROM roster_daily_presence AS p
    INNER JOIN roster_events AS ev ON ev.id = p.event_id
    WHERE p.date >= ?
      AND p.date <= ?
      ${sourceSql}
      ${doctorSql}
      ${excludeDoctorSql}
    ORDER BY p.display_name, ev.start_ts
  `).bind(start, end, ...sourceTypes, ...includeKeys, ...excludeKeys).all();
  return rowsToCoworkerEvents(rows);
}

export async function queryCoworkerEventsFromEvents(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const start = String(options.startDate || options.date || "0000-01-01");
  const end = String(options.endDate || options.date || "9999-12-31");
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes);
  const excludeKeys = new Set((options.excludeDoctorKeys || []).filter(Boolean));
  const includeKeys = [...new Set((options.doctorKeys || []).filter(Boolean))];
  const overlapKeys = [...new Set((options.overlapDoctorKeys || []).filter(Boolean))];
  const sourceSql = sourceTypes.length ? `AND roster_events.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
  const doctorSql = includeKeys.length ? `AND roster_events.doctor_key IN (${includeKeys.map(() => "?").join(", ")})` : "";
  const excludeDoctorSql = excludeKeys.size ? `AND roster_events.doctor_key NOT IN (${[...excludeKeys].map(() => "?").join(", ")})` : "";
  if (overlapKeys.length) {
    const overlapSourceSql = sourceTypes.length ? `AND other_events.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
    const overlapDoctorSql = includeKeys.length ? `AND other_events.doctor_key IN (${includeKeys.map(() => "?").join(", ")})` : "";
    const overlapExcludeDoctorSql = excludeKeys.size ? `AND other_events.doctor_key NOT IN (${[...excludeKeys].map(() => "?").join(", ")})` : "";
    const rows = await db.prepare(`
      SELECT DISTINCT
        other_events.doctor_key,
        other_events.display_name,
        other_events.source_type,
        other_events.event_json,
        other_events.start_ts
      FROM roster_events AS mine
      INNER JOIN roster_files AS mine_files ON mine_files.id = mine.file_id
      INNER JOIN roster_events AS other_events
        ON other_events.source_type = mine.source_type
       AND other_events.start_date <= mine.end_date
       AND other_events.end_date >= mine.start_date
      INNER JOIN roster_files AS other_files ON other_files.id = other_events.file_id
      WHERE mine_files.active = 1
        AND other_files.active = 1
        AND mine.doctor_key IN (${overlapKeys.map(() => "?").join(", ")})
        AND mine.start_date <= ?
        AND mine.end_date >= ?
        AND other_events.start_date <= ?
        AND other_events.end_date >= ?
        ${overlapSourceSql}
        ${overlapDoctorSql}
        ${overlapExcludeDoctorSql}
      ORDER BY other_events.display_name, other_events.start_ts
    `).bind(...overlapKeys, end, start, end, start, ...sourceTypes, ...includeKeys, ...excludeKeys).all();
    return rowsToCoworkerEvents(rows);
  }
  const rows = await db.prepare(`
    SELECT
      roster_events.doctor_key,
      roster_events.display_name,
      roster_events.source_type,
      roster_events.event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.start_date <= ?
      AND roster_events.end_date >= ?
      ${sourceSql}
      ${doctorSql}
      ${excludeDoctorSql}
    ORDER BY roster_events.display_name, roster_events.start_ts
  `).bind(end, start, ...sourceTypes, ...includeKeys, ...excludeKeys).all();
  return rowsToCoworkerEvents(rows);
}

function rowsToCoworkerEvents(rows) {
  return (rows.results || [])
    .map((row) => ({
      doctorKey: row.doctor_key,
      displayName: row.display_name,
      sourceType: row.source_type,
      event: parseEvent(row.event_json),
    }))
    .filter((row) => row.event);
}

export async function queryOverlapDoctors(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const start = String(options.startDate || options.date || "0000-01-01");
  const end = String(options.endDate || options.date || "9999-12-31");
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes);
  const overlapKeys = [...new Set((options.overlapDoctorKeys || []).filter(Boolean))];
  const excludeKeys = [...new Set((options.excludeDoctorKeys || []).filter(Boolean))];
  if (!overlapKeys.length) return [];
  const sourceSql = sourceTypes.length ? `AND p.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
  const excludeDoctorSql = excludeKeys.length ? `AND p.doctor_key NOT IN (${excludeKeys.map(() => "?").join(", ")})` : "";
  const rows = await db.prepare(`
    SELECT DISTINCT
      p.doctor_key,
      p.display_name,
      p.source_type
    FROM roster_daily_presence AS mine
    INNER JOIN roster_daily_presence AS p
      ON p.date = mine.date
     AND p.source_type = mine.source_type
    WHERE mine.doctor_key IN (${overlapKeys.map(() => "?").join(", ")})
      AND mine.date >= ?
      AND mine.date <= ?
      AND p.date >= ?
      AND p.date <= ?
      ${sourceSql}
      ${excludeDoctorSql}
    ORDER BY p.display_name, p.source_type
  `).bind(...overlapKeys, start, end, start, end, ...sourceTypes, ...excludeKeys).all();
  return (rows.results || []).map((row) => ({
    doctorKey: row.doctor_key,
    displayName: row.display_name,
    sourceType: row.source_type,
  }));
}

export async function queryOverlapDoctorsFromEvents(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const start = String(options.startDate || options.date || "0000-01-01");
  const end = String(options.endDate || options.date || "9999-12-31");
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes);
  const overlapKeys = [...new Set((options.overlapDoctorKeys || []).filter(Boolean))];
  const excludeKeys = [...new Set((options.excludeDoctorKeys || []).filter(Boolean))];
  if (!overlapKeys.length) return [];
  const sourceSql = sourceTypes.length ? `AND other_events.source_type IN (${sourceTypes.map(() => "?").join(", ")})` : "";
  const excludeDoctorSql = excludeKeys.length ? `AND other_events.doctor_key NOT IN (${excludeKeys.map(() => "?").join(", ")})` : "";
  const rows = await db.prepare(`
    SELECT DISTINCT
      other_events.doctor_key,
      other_events.display_name,
      other_events.source_type
    FROM roster_events AS mine
    INNER JOIN roster_files AS mine_files ON mine_files.id = mine.file_id
    INNER JOIN roster_events AS other_events
      ON other_events.source_type = mine.source_type
     AND other_events.start_date <= mine.end_date
     AND other_events.end_date >= mine.start_date
    INNER JOIN roster_files AS other_files ON other_files.id = other_events.file_id
    WHERE mine_files.active = 1
      AND other_files.active = 1
      AND mine.doctor_key IN (${overlapKeys.map(() => "?").join(", ")})
      AND mine.start_date <= ?
      AND mine.end_date >= ?
      AND other_events.start_date <= ?
      AND other_events.end_date >= ?
      ${sourceSql}
      ${excludeDoctorSql}
    ORDER BY other_events.display_name, other_events.source_type
  `).bind(...overlapKeys, end, start, end, start, ...sourceTypes, ...excludeKeys).all();
  return (rows.results || []).map((row) => ({
    doctorKey: row.doctor_key,
    displayName: row.display_name,
    sourceType: row.source_type,
  }));
}

export async function deleteDailyPresenceForFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  await db.prepare(`
    DELETE FROM roster_daily_presence
    WHERE event_id LIKE ? ESCAPE '\\'
  `).bind(`${escapeLike(fileId)}:%`).run();
}

export async function populateDailyPresenceForFile(db, fileId, eventsByDoctor = {}, options = {}) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  const rows = [];
  const doctorMetadata = new Map((options.doctors || []).map((doctor) => [String(doctor?.key || doctor?.doctorKey || doctor?.doctor_key || "").trim(), doctor]));
  const fallbackSourceType = normalizeSourceType(options.sourceType || "");
  for (const [doctorKey, events] of Object.entries(eventsByDoctor || {})) {
    const metadata = doctorMetadata.get(String(doctorKey || "").trim()) || {};
    for (const event of Array.isArray(events) ? events : []) {
      for (const row of expandEventToDailyPresenceRows(event, {
        sourceType: fallbackSourceType || normalizeSourceType(metadata.sourceType || metadata.source_type || ""),
        doctorKey,
        displayName: metadata.displayName || metadata.display_name || doctorKey,
        storedEventId: options.storedEventIds?.get?.(event) || "",
      })) {
        rows.push([
          row.date,
          row.sourceType,
          row.doctorKey,
          row.displayName,
          row.storedEventId || `${fileId}:${row.doctorKey}:${row.eventId}`,
        ]);
      }
    }
  }
  for (const chunk of chunkRows(rows, 100)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT OR IGNORE INTO roster_daily_presence (date, source_type, doctor_key, display_name, event_id)
      VALUES ${chunk.map(() => "(?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()).run();
  }
}

export async function rebuildDailyPresenceForFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      roster_events.id AS id,
      roster_events.source_type AS source_type,
      roster_events.doctor_key AS doctor_key,
      roster_events.display_name AS display_name,
      roster_events.start_date AS start_date,
      roster_events.end_date AS end_date,
      roster_events.start_ts AS start_ts,
      roster_events.end_ts AS end_ts
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_events.file_id = ?
      AND roster_files.active = 1
  `).bind(fileId).all();
  await deleteDailyPresenceForFile(db, fileId);
  const presenceRows = [];
  for (const row of rows.results || []) {
    const doctorKey = String(row.doctor_key || "").trim();
    const event = {
      id: String(row.id || "").split(":").pop(),
      source: String(row.source_type || ""),
      start: String(row.start_ts || row.start_date || ""),
      end: String(row.end_ts || row.end_date || row.start_ts || row.start_date || ""),
    };
    for (const presence of expandEventToDailyPresenceRows(event, {
      sourceType: row.source_type,
      doctorKey,
      displayName: row.display_name,
      storedEventId: row.id,
    })) {
      presenceRows.push([
        presence.date,
        presence.sourceType,
        presence.doctorKey,
        presence.displayName,
        presence.storedEventId,
      ]);
    }
  }
  for (const chunk of chunkRows(presenceRows, 100)) {
    if (!chunk.length) continue;
    await db.prepare(`
      INSERT OR IGNORE INTO roster_daily_presence (date, source_type, doctor_key, display_name, event_id)
      VALUES ${chunk.map(() => "(?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()).run();
  }
}

export async function rebuildDailyPresenceForActiveFiles(db) {
  if (!db?.prepare) return { files: 0 };
  await ensureCalendarSchema(db);
  const rows = await db.prepare("SELECT id FROM roster_files WHERE active = 1 ORDER BY added_at, name").all();
  let count = 0;
  for (const row of rows.results || []) {
    const fileId = String(row.id || "").trim();
    if (!fileId) continue;
    await rebuildDailyPresenceForFile(db, fileId);
    count += 1;
  }
  return { files: count };
}

function expandEventToDailyPresenceRows(event, context = {}) {
  const start = String(event?.start || "").slice(0, 10);
  const end = String(event?.end || event?.start || "").slice(0, 10);
  const sourceType = eventSourceType(event) || normalizeSourceType(context.sourceType || "");
  const doctorKey = String(event?.doctorKey || event?.doctor_key || context.doctorKey || "").trim();
  const displayName = String(event?.displayName || event?.doctorDisplay || context.displayName || doctorKey || "").trim();
  const eventId = String(event?.id || "").trim();
  if (!start || !end || start > end || !sourceType || !doctorKey || !eventId) return [];
  const rows = [];
  let cursor = start;
  while (cursor <= end) {
    rows.push({ date: cursor, sourceType, doctorKey, displayName, eventId, storedEventId: String(context.storedEventId || "") });
    cursor = nextDateKey(cursor);
  }
  return rows;
}

function nextDateKey(dateKey) {
  const [year, month, day] = String(dateKey || "").split("-").map((part) => Number(part));
  const date = new Date(Date.UTC(year || 1970, (month || 1) - 1, day || 1));
  date.setUTCDate(date.getUTCDate() + 1);
  return date.toISOString().slice(0, 10);
}

function escapeLike(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/%/g, "\\%").replace(/_/g, "\\_");
}

export function buildPreviewFromDerivedEvents(events, options = {}) {
  const safeEvents = mergeDuplicateLeaveEvents(events || []).map((event) => ({ ...event }));
  const issues = (Array.isArray(options.issues) ? options.issues : []).map(sanitizeIssue).filter(Boolean);
  return {
    ...previewSummary(safeEvents),
    events: safeEvents,
    review: safeEvents.map((event) => reviewItemForEvent(event)),
    issues,
    conflicts: [],
    imports: [],
    sources: {},
    lastParsed: new Date().toISOString(),
    derivedFromD1: true,
    customEventsMaterialized: options.customEventsMaterialized === true,
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

function sanitizeIssue(value) {
  if (!value || typeof value !== "object") return null;
  const id = String(value.id || value.fingerprint || "").trim();
  const source = String(value.source || "").trim();
  const rawValue = String(value.rawValue || "").trim();
  const message = String(value.message || "").trim();
  if (!id || !source || !rawValue || !message) return null;
  return {
    id,
    source,
    seniority: String(value.seniority || "").trim(),
    startDay: String(value.startDay || value.date || "").slice(0, 10),
    rawValue,
    status: String(value.status || "unknown").trim() || "unknown",
    message,
    resolutionType: String(value.resolutionType || "").trim(),
    suggestedTitle: String(value.suggestedTitle || "").trim(),
    timeLabel: String(value.timeLabel || "").trim(),
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

function sanitizeAccountClaims(claims) {
  const normalized = (Array.isArray(claims) ? claims : [])
    .map((claim) => ({
      key: String(claim?.key || "").trim(),
      displayName: String(claim?.displayName || claim?.key || "").trim(),
      sourceType: normalizeSourceType(claim?.sourceType || ""),
      matchedAt: String(claim?.matchedAt || ""),
    }))
    .filter((claim) => claim.key && claim.displayName && claim.sourceType);
  const deduped = [];
  const seen = new Set();
  for (const claim of normalized) {
    const marker = `${claim.sourceType}:${claim.key}`;
    if (seen.has(marker)) continue;
    seen.add(marker);
    deduped.push(claim);
  }
  return deduped.sort((left, right) => left.sourceType.localeCompare(right.sourceType) || left.displayName.localeCompare(right.displayName) || left.key.localeCompare(right.key));
}

function doctorProfileFromRow(row) {
  if (!row?.profile_id) return null;
  let sourceTypes = [];
  let state = {};
  try {
    sourceTypes = JSON.parse(row.source_types_json || "[]");
  } catch {
    sourceTypes = [];
  }
  try {
    state = JSON.parse(row.state_json || "{}");
  } catch {
    state = {};
  }
  return {
    profileId: String(row.profile_id || "").trim(),
    doctorKey: String(row.doctor_key || "").trim(),
    displayName: String(row.display_name || "").trim(),
    sourceTypes: sanitizeSourceTypes(sourceTypes),
    state: state && typeof state === "object" ? state : {},
    createdAt: String(row.created_at || ""),
    updatedAt: String(row.updated_at || ""),
  };
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

function rawRosterFileFromRow(row) {
  if (!row?.file_id || (!row?.data_url && !row?.object_key)) return null;
  const fileId = String(row.file_id || "").trim();
  const name = String(row.name || fileNameFromRawRosterFileId(fileId) || "roster.xlsx");
  return {
    fileId,
    id: fileId,
    repoId: fileId,
    name,
    sourceType: normalizeSourceType(row.source_type || "") || inferSourceTypeFromRosterFileName(name),
    size: Number(row.size || 0),
    lastModified: Number(row.last_modified || 0),
    objectKey: String(row.object_key || ""),
    type: String(row.type || ""),
    dataUrl: String(row.data_url || ""),
    uploadedAt: String(row.uploaded_at || ""),
    rawSourceAvailable: true,
  };
}

function fileNameFromRawRosterFileId(fileId) {
  const parts = String(fileId || "").split(":");
  return parts.length > 2 ? parts.slice(0, -2).join(":") : String(fileId || "");
}

function inferSourceTypeFromRosterFileName(name) {
  const text = String(name || "").toLowerCase();
  if (text.includes("dandenong") || text.includes("ddh")) return "ddh";
  if (text.includes("casey")) return "casey";
  if (text.includes("paeds") || text.includes("mch") || text.includes("children")) return "mch";
  if (text.includes("adult") || text.includes("mmc") || text.includes("monash")) return "mmc";
  return "";
}

function uniqueFileDoctorPairs(pairs = []) {
  const seen = new Set();
  const result = [];
  for (const pair of pairs || []) {
    const fileId = String(pair?.fileId || pair?.file_id || "").trim();
    const doctorKey = String(pair?.doctorKey || pair?.doctor_key || pair?.key || "").trim();
    if (!fileId || !doctorKey) continue;
    const marker = `${fileId}:${doctorKey}`;
    if (seen.has(marker)) continue;
    seen.add(marker);
    result.push({ fileId, doctorKey });
  }
  return result;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
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

function parseIssue(value) {
  try {
    return sanitizeIssue(JSON.parse(value));
  } catch {
    return null;
  }
}

function mergeDuplicateLeaveEvents(events) {
  const passthrough = [];
  const leaveEvents = [];
  for (const event of events || []) {
    if (!isMergeableLeaveEvent(event)) {
      passthrough.push(event);
      continue;
    }
    leaveEvents.push(event);
  }
  const merged = [];
  const ordered = leaveEvents.sort((left, right) => String(left.start || "").localeCompare(String(right.start || "")) || String(left.end || "").localeCompare(String(right.end || "")));
  for (const event of ordered) {
    const previous = merged.length ? merged[merged.length - 1] : null;
    const overlaps = previous && String(event.start || "") < String(previous.end || previous.start || "");
    const adjacentSameType = previous && String(event.start || "") === String(previous.end || previous.start || "")
      && preferredLeaveTitle(previous.title, "", previous.rawValue) === preferredLeaveTitle(event.title, "", event.rawValue);
    if (previous && (overlaps || adjacentSameType)) {
      previous.end = String(event.end || "") > String(previous.end || "") ? event.end : previous.end;
      previous.rawValue = mergeRawValues(previous.rawValue, event.rawValue);
      previous.sources = mergeSources(previous.sources, event.sources, previous.source, event.source);
      previous.title = preferredLeaveTitle(previous.title, event.title, previous.rawValue);
      continue;
    }
    merged.push({
      ...event,
      title: preferredLeaveTitle(event.title, "", event.rawValue),
      sources: mergeSources(event.sources, null, event.source),
    });
  }
  return [...passthrough, ...merged]
    .sort((left, right) => String(left.start || "").localeCompare(String(right.start || "")) || String(left.title || "").localeCompare(String(right.title || "")));
}

function isMergeableLeaveEvent(event) {
  if (event?.allDay !== true) return false;
  return leaveTextMatches(`${event.title || ""} ${event.rawValue || ""}`);
}

function leavesOverlap(left, right) {
  return String(right.start || "") <= String(left.end || left.start || "");
}

function leaveTextMatches(value) {
  return /\b(leave|conference|cme|study|annual|sick|personal|exam|sabbatical|parental|long service)\b/i.test(String(value || ""));
}

function preferredLeaveTitle(leftTitle, rightTitle, rawValue = "") {
  const combined = `${leftTitle || ""} ${rightTitle || ""} ${rawValue || ""}`;
  if (/\b(conference|cme)\b/i.test(combined)) return "Conference Leave";
  if (/\bannual\b/i.test(combined)) return "Annual Leave";
  if (/\bsick\b/i.test(combined)) return "Sick Leave";
  if (/\bpersonal\b/i.test(combined)) return "Personal Leave";
  if (/\bstudy\b/i.test(combined)) return "Study Leave";
  if (/\bexam\b/i.test(combined)) return "Exam Leave";
  if (/\bsabbatical\b/i.test(combined)) return "Sabbatical Leave";
  if (/\bparental\b/i.test(combined)) return "Parental Leave";
  if (/\blong service\b/i.test(combined)) return "Long Service Leave";
  return String(leftTitle || rightTitle || "Leave").trim();
}

function mergeRawValues(left, right) {
  const values = [left, right]
    .flatMap((item) => String(item || "").split(" / "))
    .map((item) => item.trim())
    .filter(Boolean);
  return [...new Set(values)].join(" / ");
}

function mergeSources(leftSources, rightSources, leftSource, rightSource) {
  const values = [leftSources, rightSources, leftSource, rightSource]
    .flatMap((item) => Array.isArray(item) ? item : [item])
    .map((item) => String(item || "").trim())
    .filter((item) => /^(MMC|DDH|Casey|MCH)$/i.test(item));
  return [...new Set(values.map((item) => item.toUpperCase() === "CASEY" ? "Casey" : item.toUpperCase()))];
}

async function bulkUpsertDoctors(db, sourceType, doctors, updatedAt) {
  for (const statement of bulkUpsertDoctorStatements(db, sourceType, doctors, updatedAt)) {
    await statement.run();
  }
}

function bulkUpsertDoctorStatements(db, sourceType, doctors, updatedAt) {
  return chunkRows(doctors.map((doctor) => [sourceType, doctor.key, doctor.displayName, updatedAt]), 20)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_doctors (source_type, doctor_key, display_name, updated_at)
      VALUES ${chunk.map(() => "(?, ?, ?, ?)").join(", ")}
      ON CONFLICT(source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        updated_at = excluded.updated_at
    `).bind(...chunk.flat()));
}

async function bulkInsertFileDoctors(db, fileId, sourceType, doctors) {
  for (const statement of bulkInsertFileDoctorStatements(db, fileId, sourceType, doctors)) {
    await statement.run();
  }
}

function bulkInsertFileDoctorStatements(db, fileId, sourceType, doctors) {
  return chunkRows(doctors.map((doctor) => [fileId, sourceType, doctor.key, doctor.displayName]), 20)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_file_doctors (file_id, source_type, doctor_key, display_name)
      VALUES ${chunk.map(() => "(?, ?, ?, ?)").join(", ")}
      ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET display_name = excluded.display_name
    `).bind(...chunk.flat()));
}

async function bulkInsertEvents(db, rows) {
  for (const statement of bulkInsertEventStatements(db, rows)) {
    await statement.run();
  }
}

function bulkInsertEventStatements(db, rows) {
  return chunkRows(rows, 1)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_events (
        id, file_id, source_type, doctor_key, display_name, start_date, end_date, start_ts, end_ts,
        title, raw_value, seniority, location, all_day, time_label, event_json
      ) VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()));
}

function bulkInsertIssueStatements(db, rows) {
  return chunkRows(rows, 1)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_issues (
        id, file_id, source_type, doctor_key, display_name, start_date, raw_value, seniority,
        status, message, resolution_type, suggested_title, time_label, issue_json
      ) VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()));
}

async function runTransactionalBatch(db, statements) {
  if (!statements.length) return [];
  if (typeof db.batch === "function") return await db.batch(statements);
  const results = [];
  for (const statement of statements) results.push(await statement.run());
  return results;
}

function chunkRows(rows, size) {
  const chunks = [];
  for (let index = 0; index < rows.length; index += size) {
    chunks.push(rows.slice(index, index + size));
  }
  return chunks;
}

export function snapshotRegistryRangeKey({ startDate = "", endDate = "" } = {}) {
  return `${String(startDate || "").slice(0, 10)}..${String(endDate || "").slice(0, 10)}`;
}

export function snapshotArtifactKey({ ownerType = "", ownerId = "", doctorKey = "", rangeKey = "" } = {}) {
  const parts = [ownerType, ownerId, doctorKey, rangeKey]
    .map((part) => String(part || "").trim().replace(/[^A-Za-z0-9_.@-]+/g, "_"));
  return `snapshots/${parts.join("/")}.json.gz`;
}

export async function loadSnapshotRegistryEntry(db, key = {}) {
  if (!db?.prepare) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare(`
    SELECT owner_type, owner_id, doctor_key, range_key, requested_revision, built_revision, status, artifact_key,
           built_at, size_bytes, build_ms, last_error, updated_at
    FROM snapshot_registry
    WHERE owner_type = ? AND owner_id = ? AND doctor_key = ? AND range_key = ?
  `).bind(
    String(key.ownerType || ""),
    normalizeOwnerIdValue(key.ownerId || key.owner_id || ""),
    String(key.doctorKey || key.doctor_key || ""),
    String(key.rangeKey || key.range_key || ""),
  ).first();
  return snapshotRegistryEntryFromRow(row);
}

export async function upsertSnapshotRegistryEntry(db, entry = {}) {
  if (!db?.prepare || !entry?.ownerType || !entry?.ownerId || !entry?.rangeKey) return false;
  await ensureCalendarSchema(db);
  const updatedAt = String(entry.updatedAt || new Date().toISOString());
  await db.prepare(`
    INSERT INTO snapshot_registry (
      owner_type, owner_id, doctor_key, range_key, requested_revision, built_revision, status, artifact_key,
      built_at, size_bytes, build_ms, last_error, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(owner_type, owner_id, doctor_key, range_key) DO UPDATE SET
      requested_revision = excluded.requested_revision,
      built_revision = excluded.built_revision,
      status = excluded.status,
      artifact_key = excluded.artifact_key,
      built_at = excluded.built_at,
      size_bytes = excluded.size_bytes,
      build_ms = excluded.build_ms,
      last_error = excluded.last_error,
      updated_at = excluded.updated_at
  `).bind(
    String(entry.ownerType || ""),
    normalizeOwnerIdValue(entry.ownerId || ""),
    String(entry.doctorKey || ""),
    String(entry.rangeKey || ""),
    String(entry.requestedRevision || ""),
    String(entry.builtRevision || ""),
    String(entry.status || "missing"),
    String(entry.artifactKey || ""),
    String(entry.builtAt || ""),
    Number(entry.sizeBytes || 0),
    Number(entry.buildMs || 0),
    String(entry.lastError || ""),
    updatedAt,
  ).run();
  return true;
}

export async function listSnapshotRegistryEntriesForOwner(db, ownerType, ownerId) {
  if (!db?.prepare || !ownerType || !ownerId) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT owner_type, owner_id, doctor_key, range_key, requested_revision, built_revision, status, artifact_key,
           built_at, size_bytes, build_ms, last_error, updated_at
    FROM snapshot_registry
    WHERE owner_type = ? AND owner_id = ?
    ORDER BY updated_at DESC
  `).bind(String(ownerType || ""), normalizeOwnerIdValue(ownerId || "")).all();
  return (rows.results || []).map(snapshotRegistryEntryFromRow).filter(Boolean);
}

export async function deleteSnapshotRegistryEntriesForOwner(db, ownerType, ownerId) {
  if (!db?.prepare || !ownerType || !ownerId) return;
  await ensureCalendarSchema(db);
  await db.prepare("DELETE FROM snapshot_registry WHERE owner_type = ? AND owner_id = ?").bind(
    String(ownerType || ""),
    normalizeOwnerIdValue(ownerId || ""),
  ).run();
}

function snapshotRegistryEntryFromRow(row) {
  if (!row?.owner_type || !row?.owner_id || !row?.range_key) return null;
  return {
    ownerType: String(row.owner_type || ""),
    ownerId: normalizeOwnerIdValue(row.owner_id || ""),
    doctorKey: String(row.doctor_key || ""),
    rangeKey: String(row.range_key || ""),
    requestedRevision: String(row.requested_revision || ""),
    builtRevision: String(row.built_revision || ""),
    status: String(row.status || "missing"),
    artifactKey: String(row.artifact_key || ""),
    builtAt: String(row.built_at || ""),
    sizeBytes: Number(row.size_bytes || 0),
    buildMs: Number(row.build_ms || 0),
    lastError: String(row.last_error || ""),
    updatedAt: String(row.updated_at || ""),
  };
}

export async function loadCachedSnapshot(r2, artifactKey) {
  if (!r2?.get || !artifactKey) return null;
  try {
    const object = await r2.get(artifactKey);
    if (!object) return null;
    const bytes = await object.arrayBuffer();
    const header = new Uint8Array(bytes, 0, Math.min(2, bytes.byteLength));
    const isGzip = header.length === 2 && header[0] === 0x1f && header[1] === 0x8b;
    const text = isGzip
      ? await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).text()
      : new TextDecoder().decode(bytes);
    return JSON.parse(text);
  } catch {
    return null;
  }
}

export async function storeCachedSnapshot(r2, artifactKey, snapshot, metadata = {}) {
  if (!r2?.put || !artifactKey || !snapshot) return false;
  const json = JSON.stringify(snapshot);
  const compressed = await new Response(new Blob([json]).stream().pipeThrough(new CompressionStream("gzip"))).arrayBuffer();
  await r2.put(artifactKey, compressed, {
    customMetadata: {
      revision: String(metadata.revision || ""),
      ownerType: String(metadata.ownerType || ""),
      ownerId: normalizeOwnerIdValue(metadata.ownerId || ""),
      doctorKey: String(metadata.doctorKey || ""),
      rangeKey: String(metadata.rangeKey || ""),
    },
    httpMetadata: { contentType: "application/json", contentEncoding: "gzip" },
  });
  return true;
}

export async function deleteCachedSnapshotsForOwner(r2, ownerType, ownerId) {
  if (!r2?.list || !ownerType || !ownerId) return;
  const prefix = `snapshots/${String(ownerType || "").trim().replace(/[^A-Za-z0-9_.@-]+/g, "_")}/${normalizeOwnerIdValue(ownerId || "")}/`;
  let cursor = undefined;
  do {
    const listed = await r2.list({ prefix, cursor });
    if (listed.objects.length) await r2.delete(listed.objects.map((object) => object.key));
    cursor = listed.truncated ? listed.cursor : undefined;
  } while (cursor);
}

function parseJsonObject(value, fallback = {}) {
  if (typeof value !== "string" || !value.trim()) return fallback;
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function normalizeOwnerIdValue(value) {
  const raw = String(value || "").trim();
  return raw.includes("@") ? normalizeEmail(raw) : raw;
}

function materializedSessionStateForRevision(session = {}) {
  const safe = session && typeof session === "object" ? session : {};
  return {
    doctorKey: String(safe.doctorKey || "").trim().toUpperCase(),
    overrides: sortObjectKeys(safe.overrides && typeof safe.overrides === "object" ? safe.overrides : {}),
    conflictSelections: sortObjectKeys(safe.conflictSelections && typeof safe.conflictSelections === "object" ? safe.conflictSelections : {}),
  };
}

function stableJsonStringify(value) {
  return JSON.stringify(sortObjectKeys(value));
}

function sortObjectKeys(value) {
  if (Array.isArray(value)) return value.map(sortObjectKeys);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, sortObjectKeys(value[key])])
  );
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
