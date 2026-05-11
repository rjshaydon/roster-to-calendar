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
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_profiles (
      email TEXT PRIMARY KEY,
      real_name TEXT NOT NULL DEFAULT '',
      role TEXT NOT NULL DEFAULT 'user',
      insights_enabled INTEGER NOT NULL DEFAULT 0,
      subscription_token TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
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
  return mergeDuplicateLeaveEvents((rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean));
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

export async function upsertAccountMirror(db, record, options = {}) {
  if (!db?.prepare || !record?.email) return false;
  await ensureCalendarSchema(db);
  const email = normalizeEmail(record.email);
  const role = String(record.role || (email ? "user" : "") || "user");
  const updatedAt = new Date().toISOString();
  await db.prepare(`
    INSERT INTO account_profiles (email, real_name, role, insights_enabled, subscription_token, updated_at)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(email) DO UPDATE SET
      real_name = excluded.real_name,
      role = excluded.role,
      insights_enabled = excluded.insights_enabled,
      subscription_token = excluded.subscription_token,
      updated_at = excluded.updated_at
  `).bind(
    email,
    String(record.realName || ""),
    role,
    record.insightsEnabled === true ? 1 : 0,
    String(record.subscriptionToken || ""),
    updatedAt,
  ).run();
  await db.prepare("DELETE FROM account_claims WHERE email = ?").bind(email).run();
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
  return true;
}

export async function deleteAccountMirror(db, email) {
  if (!db?.prepare || !email) return;
  await ensureCalendarSchema(db);
  const normalizedEmail = normalizeEmail(email);
  await db.prepare("DELETE FROM account_claims WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM account_states WHERE email = ?").bind(normalizedEmail).run();
  await db.prepare("DELETE FROM account_profiles WHERE email = ?").bind(normalizedEmail).run();
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
  const safeEvents = mergeDuplicateLeaveEvents(events || []).map((event) => ({ ...event }));
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

function sanitizeAccountClaims(claims) {
  return (Array.isArray(claims) ? claims : [])
    .map((claim) => ({
      key: String(claim?.key || "").trim(),
      displayName: String(claim?.displayName || claim?.key || "").trim(),
      sourceType: normalizeSourceType(claim?.sourceType || ""),
      matchedAt: String(claim?.matchedAt || ""),
    }))
    .filter((claim) => claim.key && claim.displayName && claim.sourceType);
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
    if (previous && leavesOverlap(previous, event)) {
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
  return /\b(leave|conference|cme|study|annual|sick|personal)\b/i.test(String(value || ""));
}

function preferredLeaveTitle(leftTitle, rightTitle, rawValue = "") {
  const combined = `${leftTitle || ""} ${rightTitle || ""} ${rawValue || ""}`;
  if (/\b(conference|cme)\b/i.test(combined)) return "Conference Leave";
  if (/\bannual\b/i.test(combined)) return "Annual Leave";
  if (/\bsick\b/i.test(combined)) return "Sick Leave";
  if (/\bpersonal\b/i.test(combined)) return "Personal Leave";
  if (/\bstudy\b/i.test(combined)) return "Study Leave";
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
