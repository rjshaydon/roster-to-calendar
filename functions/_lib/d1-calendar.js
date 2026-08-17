import {
  buildRosterViewFromStoredImports,
  defaultSettings,
  previewSummary,
  serializeEvent,
} from "./roster.js";

const SOURCE_TYPES = ["mmc", "ddh", "casey", "mch"];
// Only a successfully activated parse may carry this version. Older rows are
// deliberately marked legacy-unverified by migration 0019.
export const ROSTER_PARSER_VERSION = "shared-core-v2";
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
      source_id TEXT NOT NULL DEFAULT '',
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
      seniority TEXT NOT NULL DEFAULT '',
      membership_source TEXT NOT NULL DEFAULT 'roster',
      PRIMARY KEY (file_id, source_type, doctor_key)
    )
  `).run();
  await ensureColumn(db, "roster_file_doctors", "seniority", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "roster_file_doctors", "membership_source", "TEXT NOT NULL DEFAULT 'roster'");
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_file_doctors_source_file ON roster_file_doctors (source_type, file_id)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS facility_staff_designations (
      id TEXT PRIMARY KEY,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL DEFAULT '',
      seniority TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL,
      term_start TEXT NOT NULL,
      term_end TEXT NOT NULL DEFAULT '',
      source_revision TEXT NOT NULL DEFAULT '',
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      cleared_at TEXT NOT NULL DEFAULT '',
      cleared_reason TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_facility_staff_designations_active ON facility_staff_designations (source_type, doctor_key, active, term_start)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS facility_staff_seniority_overrides (
      id TEXT PRIMARY KEY,
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL DEFAULT '',
      seniority TEXT NOT NULL DEFAULT '',
      use_roster_seniority INTEGER NOT NULL DEFAULT 0,
      term_start TEXT NOT NULL,
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      cleared_at TEXT NOT NULL DEFAULT '',
      cleared_reason TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_facility_staff_seniority_overrides_active ON facility_staff_seniority_overrides (source_type, doctor_key, active, term_start)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS facility_sms_memberships (
      source_type TEXT NOT NULL,
      doctor_key TEXT NOT NULL,
      display_name TEXT NOT NULL DEFAULT '',
      first_seen_date TEXT NOT NULL DEFAULT '',
      last_seen_date TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (source_type, doctor_key)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_facility_sms_memberships_first_seen ON facility_sms_memberships (source_type, first_seen_date)").run();
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
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_file_doctor ON roster_events (file_id, doctor_key)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_date_source ON roster_events (start_date, source_type)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events (source_type, start_date, end_date)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_events_file ON roster_events (file_id)").run();
  await backfillFacilitySmsMemberships(db);
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
      facility_overview_enabled INTEGER NOT NULL DEFAULT 0,
      subscription_token TEXT NOT NULL DEFAULT '',
      password_salt TEXT NOT NULL DEFAULT '',
      password_hash TEXT NOT NULL DEFAULT '',
      admin_issues_json TEXT NOT NULL DEFAULT '[]',
      local_parser_extensions_json TEXT NOT NULL DEFAULT '[]',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await ensureColumn(db, "account_profiles", "facility_overview_enabled", "INTEGER NOT NULL DEFAULT 0");
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
    CREATE TABLE IF NOT EXISTS roster_people (
      person_id TEXT PRIMARY KEY, preferred_display_name TEXT NOT NULL DEFAULT '',
      provenance TEXT NOT NULL DEFAULT 'automatic', review_state TEXT NOT NULL DEFAULT 'approved',
      created_at TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL DEFAULT '', approved_by TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_person_aliases (
      source_type TEXT NOT NULL, doctor_key TEXT NOT NULL, display_name TEXT NOT NULL DEFAULT '', person_id TEXT NOT NULL,
      provenance TEXT NOT NULL DEFAULT 'automatic', confidence TEXT NOT NULL DEFAULT 'high', review_state TEXT NOT NULL DEFAULT 'approved',
      created_at TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL DEFAULT '', approved_by TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (source_type, doctor_key)
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_person_aliases_person ON roster_person_aliases (person_id)").run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS account_people (
      email TEXT PRIMARY KEY, person_id TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
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
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_daily_presence_event_id ON roster_daily_presence (event_id)").run();
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
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_sources (
      id TEXT PRIMARY KEY,
      provider TEXT NOT NULL DEFAULT '',
      source_type TEXT NOT NULL DEFAULT '',
      label TEXT NOT NULL DEFAULT '',
      enabled INTEGER NOT NULL DEFAULT 0,
      config_json TEXT NOT NULL DEFAULT '{}',
      cursor_json TEXT NOT NULL DEFAULT '{}',
      provider_version TEXT NOT NULL DEFAULT '',
      provider_modified_at TEXT NOT NULL DEFAULT '',
      last_checked_at TEXT NOT NULL DEFAULT '',
      last_success_at TEXT NOT NULL DEFAULT '',
      last_error TEXT NOT NULL DEFAULT '',
      active_file_id TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_sync_runs (
      id TEXT PRIMARY KEY,
      source_id TEXT NOT NULL,
      trigger_type TEXT NOT NULL DEFAULT '',
      provider_version TEXT NOT NULL DEFAULT '',
      content_hash TEXT NOT NULL DEFAULT '',
      file_id TEXT NOT NULL DEFAULT '',
      source_file_id TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'started',
      message TEXT NOT NULL DEFAULT '',
      doctor_count INTEGER NOT NULL DEFAULT 0,
      event_count INTEGER NOT NULL DEFAULT 0,
      started_at TEXT NOT NULL DEFAULT '',
      completed_at TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS roster_dispatches (
      id TEXT PRIMARY KEY,
      status TEXT NOT NULL DEFAULT 'requested',
      reason TEXT NOT NULL DEFAULT '',
      github_run_id TEXT NOT NULL DEFAULT '',
      requested_at TEXT NOT NULL DEFAULT '',
      accepted_at TEXT NOT NULL DEFAULT '',
      started_at TEXT NOT NULL DEFAULT '',
      completed_at TEXT NOT NULL DEFAULT '',
      retry_after TEXT NOT NULL DEFAULT '',
      attempt_count INTEGER NOT NULL DEFAULT 0,
      last_error TEXT NOT NULL DEFAULT ''
    )
  `).run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_started ON roster_sync_runs (source_id, started_at DESC)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_hash ON roster_sync_runs (source_id, content_hash, status)").run();
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_roster_dispatches_status_retry ON roster_dispatches (status, retry_after DESC)").run();
  await ensureColumn(db, "roster_files", "source_id", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "roster_files", "parser_version", "TEXT NOT NULL DEFAULT 'legacy-unverified'");
  await ensureColumn(db, "roster_sync_runs", "source_file_id", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "roster_sources", "provider_version", "TEXT NOT NULL DEFAULT ''");
  await ensureColumn(db, "roster_sources", "provider_modified_at", "TEXT NOT NULL DEFAULT ''");
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
    INSERT INTO roster_files (id, name, source_type, source_id, active, size, last_modified, added_at, uploaded_at, uploaded_by, parsed_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      source_type = excluded.source_type,
      source_id = excluded.source_id,
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
    String(file.sourceId || ""),
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
      INSERT INTO roster_file_doctors (file_id, source_type, doctor_key, display_name, seniority, membership_source)
      VALUES (?, ?, ?, ?, ?, ?)
      ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        seniority = excluded.seniority,
        membership_source = excluded.membership_source
    `).bind(file.id, sourceType, doctor.key, doctor.displayName, doctor.seniority || "", doctor.membershipSource || "roster").run();
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
  await recordFacilitySmsMembershipsForRosterFile(db, file.id);
  return { ok: true, doctors: doctors.length, events: totalEvents };
}

const D1_MAX_BIND_PARAMS = 100;
const D1_MAX_BATCH_STATEMENTS = 800;
const D1_PRESENCE_BATCH_STATEMENTS = 80;

function derivedRosterFileUpsertStatement(db, file, sourceType, parsedAt) {
  return db.prepare(`
    INSERT INTO roster_files (id, name, source_type, source_id, active, size, last_modified, added_at, uploaded_at, uploaded_by, parsed_at, parser_version)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      source_type = excluded.source_type,
      source_id = excluded.source_id,
      active = excluded.active,
      size = excluded.size,
      last_modified = excluded.last_modified,
      added_at = excluded.added_at,
      uploaded_at = excluded.uploaded_at,
      uploaded_by = excluded.uploaded_by,
      parsed_at = excluded.parsed_at,
      parser_version = excluded.parser_version
  `).bind(
    file.id,
    file.name || "roster.xlsx",
    sourceType,
    String(file.sourceId || ""),
    file.active === false ? 0 : 1,
    Number(file.size || 0),
    Number(file.lastModified || 0),
    String(file.addedAt || ""),
    String(file.uploadedAt || ""),
    String(file.uploadedBy || ""),
    parsedAt,
    String(file.parserVersion || ROSTER_PARSER_VERSION),
  );
}

function collectDerivedEventAndIssueRows(file, sourceType, safeDoctors, eventsByDoctor = {}, issuesByDoctor = {}) {
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
  return { eventRows, issueRows };
}

export async function startDerivedRosterFileSave(db, file, doctors, options = {}) {
  if (!db?.prepare || !file?.id) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const safeDoctors = sanitizeFileDoctors(doctors, sourceType);
  const parsedAt = new Date().toISOString();
  const statements = [
    derivedRosterFileUpsertStatement(db, file, sourceType, parsedAt),
    db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(file.id),
    ...bulkUpsertDoctorStatements(db, sourceType, safeDoctors, parsedAt),
    ...bulkInsertFileDoctorStatements(db, file.id, sourceType, safeDoctors),
  ];
  if (options.clearDailyPresence !== false) {
    await deleteDailyPresenceForFile(db, file.id);
  }
  await runTransactionalBatch(db, statements);
  return { ok: true, doctors: safeDoctors.length, events: 0, issues: 0 };
}

export async function appendDerivedRosterFileEvents(db, file, doctors, eventsByDoctor = {}, issuesByDoctor = {}) {
  if (!db?.prepare || !file?.id) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const safeDoctors = sanitizeFileDoctors(doctors, sourceType);
  const { eventRows, issueRows } = collectDerivedEventAndIssueRows(file, sourceType, safeDoctors, eventsByDoctor, issuesByDoctor);
  const statements = [
    ...bulkInsertEventStatements(db, eventRows),
    ...bulkInsertIssueStatements(db, issueRows),
  ];
  if (statements.length) await runTransactionalBatch(db, statements);
  await recordFacilitySmsMembershipsForRosterFile(db, file.id);
  return { ok: true, doctors: safeDoctors.length, events: eventRows.length, issues: issueRows.length };
}

export async function replaceDerivedRosterFile(db, file, doctors, eventsByDoctor, issuesByDoctor = {}, options = {}) {
  if (!db?.prepare || !file?.id) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(file.sourceType);
  if (!sourceType) return { ok: false, reason: "unsupported-source" };
  const safeDoctors = sanitizeFileDoctors(doctors, sourceType);
  const parsedAt = new Date().toISOString();
  const { eventRows, issueRows } = collectDerivedEventAndIssueRows(file, sourceType, safeDoctors, eventsByDoctor, issuesByDoctor);
  const statements = [
    derivedRosterFileUpsertStatement(db, file, sourceType, parsedAt),
    db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(file.id),
    db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(file.id),
    ...bulkUpsertDoctorStatements(db, sourceType, safeDoctors, parsedAt),
    ...bulkInsertFileDoctorStatements(db, file.id, sourceType, safeDoctors),
    ...bulkInsertEventStatements(db, eventRows),
    ...bulkInsertIssueStatements(db, issueRows),
  ];
  await deleteDailyPresenceForFile(db, file.id);
  await runTransactionalBatch(db, statements);
  await recordFacilitySmsMembershipsForRosterFile(db, file.id);
  if (options.deferDailyPresence !== true) {
    await populateDailyPresenceForFile(db, file.id, eventsByDoctor, {
      sourceType,
      doctors: safeDoctors,
    });
  }
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

export async function activateDerivedRosterFile(db, fileId, parserVersion = ROSTER_PARSER_VERSION) {
  if (!db?.prepare || !fileId) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  await db.prepare("UPDATE roster_files SET active = 1, parser_version = ? WHERE id = ?")
    .bind(String(parserVersion || ROSTER_PARSER_VERSION), String(fileId)).run();
  await rebuildDailyPresenceForFile(db, fileId);
  return { ok: true, fileId: String(fileId) };
}

// Compare the persisted source occurrence, not title/time heuristics. The
// parser's occurrence id is unique only within one doctor, so doctor_key is
// part of the safety identity.
export async function compareDerivedRosterFiles(db, baselineFileId, candidateFileId, options = {}) {
  if (!db?.prepare || !baselineFileId || !candidateFileId) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT file_id, doctor_key, event_json
    FROM roster_events
    WHERE file_id IN (?, ?)
  `).bind(String(baselineFileId), String(candidateFileId)).all();
  const limit = Math.max(1, Math.min(Number(options.limit || 50), 500));
  const eventsByFile = new Map([[String(baselineFileId), new Map()], [String(candidateFileId), new Map()]]);
  for (const row of rows.results || []) {
    const event = parseEvent(row.event_json);
    const fileEvents = eventsByFile.get(String(row.file_id || ""));
    if (!event || !fileEvents) continue;
    const identity = `${String(row.doctor_key || "")}|${String(event.id || "")}`;
    if (identity.endsWith("|")) continue;
    fileEvents.set(identity, { doctorKey: String(row.doctor_key || ""), ...event });
  }
  const baseline = eventsByFile.get(String(baselineFileId)) || new Map();
  const candidate = eventsByFile.get(String(candidateFileId)) || new Map();
  const strictRemoved = [...baseline].filter(([identity]) => !candidate.has(identity)).map(([, event]) => event);
  const strictAdded = [...candidate].filter(([identity]) => !baseline.has(identity)).map(([, event]) => event);
  const unmatchedCandidate = new Set(candidate.keys());
  const omitted = [];
  for (const event of baseline.values()) {
    const matchedIdentity = [...unmatchedCandidate].find((identity) => sameRosterOccurrence(event, candidate.get(identity)));
    if (matchedIdentity) {
      unmatchedCandidate.delete(matchedIdentity);
    } else if ([...candidate.values()].some((candidateEvent) => reusableMergedLeaveOccurrence(event, candidateEvent))) {
      // A current parser may consolidate several formerly separate leave days
      // into a single all-day span. The span can correctly account for more
      // than one original source occurrence, so do not consume it after the
      // first matching day.
    } else {
      omitted.push(event);
    }
  }
  const approvedOmissions = omitted.filter((event) => isApprovedReparseOmission(event, baselineFileId));
  const removed = omitted.filter((event) => !isApprovedReparseOmission(event, baselineFileId));
  const nearestCandidates = removed.map((event) => ({
    event,
    candidate: [...candidate.values()].find((candidateEvent) => sameDoctorSourceDay(event, candidateEvent)) || null,
  }));
  const added = [...unmatchedCandidate].map((identity) => candidate.get(identity));
  return {
    ok: true,
    baselineEvents: baseline.size,
    candidateEvents: candidate.size,
    // Parser event ids change when a valid correction changes title or time.
    // Preserve them as diagnostics, but gate activation on source occurrences.
    strictRemovedCount: strictRemoved.length,
    strictAddedCount: strictAdded.length,
    omittedCount: omitted.length,
    approvedOmissionCount: approvedOmissions.length,
    removedCount: removed.length,
    addedCount: added.length,
    removed: removed.slice(0, limit),
    nearestCandidates: nearestCandidates.slice(0, limit),
    added: added.slice(0, limit),
    approvedOmissions: approvedOmissions.slice(0, limit),
  };
}

export function sameRosterOccurrence(baseline, candidate) {
  if (!baseline || !candidate) return false;
  if (String(baseline.doctorKey || "") !== String(candidate.doctorKey || "")) return false;
  if (String(baseline.source || "") !== String(candidate.source || "")) return false;
  const day = String(baseline.start || "").slice(0, 10);
  if (!day || !eventCoversDay(candidate, day)) return false;
  const baselineRaw = normalizeRosterRawValue(baseline.rawValue);
  const candidateValues = String(candidate.rawValue || "")
    .split(" / ")
    .map(normalizeRosterRawValue)
    .filter(Boolean);
  if (baselineRaw && candidateValues.includes(baselineRaw)) return true;
  // Contiguous leave is intentionally consolidated by the current parser.
  // Its combined raw annotation may therefore differ from a previous
  // single-day row, even though the same leave category still covers that
  // source day. Treat only that narrow case as the same occurrence; a leave
  // changing into a shift, a different leave category, or an absent date is
  // still a hard promotion blocker.
  const baselineLeave = leaveOccurrenceCategory(baseline);
  return Boolean(baselineLeave && baselineLeave === leaveOccurrenceCategory(candidate));
}

function leaveOccurrenceCategory(event) {
  if (event?.allDay !== true) return "";
  const text = normalizeRosterRawValue(`${event.title || ""} ${event.rawValue || ""}`);
  if (/\bANNUAL\b/.test(text) && /\bPARENTAL\b/.test(text)) return "annual-parental";
  if (/\b(?:CONFERENCE|CME)\b/.test(text)) return "conference";
  if (/\bANNUAL\b|\bA\/L\b/.test(text)) return "annual";
  if (/\b(?:SICK|S\/L)\b/.test(text)) return "sick";
  if (/\b(?:CARER'?S|C\/L)\b/.test(text)) return "carers";
  if (/\b(?:FAMILY|F\/L)\b/.test(text)) return "family";
  if (/\bPERSONAL\b/.test(text)) return "personal";
  if (/\bSTUDY\b/.test(text)) return "study";
  if (/\bEXAM\b/.test(text)) return "exam";
  if (/\b(?:SABBATICAL|SAB\/L)\b/.test(text)) return "sabbatical";
  if (/\bPARENTAL\b/.test(text)) return "parental";
  if (/\bLONG SERVICE\b|\bLSL\b/.test(text)) return "long-service";
  if (/\b(?:LEAVE WITHOUT PAY|LWOP|LWP)\b/.test(text)) return "without-pay";
  if (/\bSPECIAL LEAVE\b/.test(text)) return "special";
  return "";
}

function reusableMergedLeaveOccurrence(baseline, candidate) {
  const category = leaveOccurrenceCategory(baseline);
  if (!category || category !== leaveOccurrenceCategory(candidate)) return false;
  if (String(baseline?.doctorKey || "") !== String(candidate?.doctorKey || "")) return false;
  if (String(baseline?.source || "") !== String(candidate?.source || "")) return false;
  const day = String(baseline?.start || "").slice(0, 10);
  return Boolean(day && eventCoversDay(candidate, day));
}

function sameDoctorSourceDay(baseline, candidate) {
  if (!baseline || !candidate) return false;
  if (String(baseline.doctorKey || "") !== String(candidate.doctorKey || "")) return false;
  if (String(baseline.source || "") !== String(candidate.source || "")) return false;
  const day = String(baseline.start || "").slice(0, 10);
  return Boolean(day && eventCoversDay(candidate, day));
}

function eventCoversDay(event, day) {
  const startDay = String(event?.start || "").slice(0, 10);
  const endDay = String(event?.end || "").slice(0, 10);
  if (!startDay || !endDay) return false;
  // A timed event commonly begins and ends on the same calendar date, so its
  // end date cannot be used as an exclusive day boundary. A roster occurrence
  // always belongs to its start date. Only all-day merged leave entries may
  // validly cover later source dates.
  if (startDay === day) return true;
  return event?.allDay === true && startDay < day && day < endDay;
}

function normalizeRosterRawValue(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
}

// Product-approved on 2026-08-13: these historic DDH calendar rows were
// created by the former "leave anywhere in the week" inference. They are
// replaced by the source-faithful entries parsed under the Monday-only rule.
// Keep this as an exact source-file/event allow-list: it is not a general
// permission to remove leave, and is self-retiring once each file is promoted.
const APPROVED_DDH_WEEKLY_LEAVE_REPLACEMENTS = new Set([
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|LEE ROBBINS|Annual Leave|2026-03-30|2026-04-13|JMS AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|MARIAN ISAAC|Annual Leave|2026-04-06|2026-04-13|JMS AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|PETER VAN KOOY|Annual Leave|2026-02-23|2026-03-02|JMS AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|PETER VAN KOOY|Annual Leave|2026-03-30|2026-04-13|JMS AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|SARA HUSSAIN|Conference Leave|2026-03-23|2026-03-30|JMS Conf",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|STEVE GUASTALEGNAME|Annual Leave|2026-02-02|2026-04-06|AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|STEVE GUASTALEGNAME|Annual Leave|2026-04-20|2026-05-04|AL",
  "Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005|YEE ANN SOO|Annual Leave|2026-02-16|2026-03-02|JMS AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|HWEE MIN LEE|Annual Leave|2026-05-04|2026-05-11|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|HWEE MIN LEE|Annual Leave|2026-06-01|2026-06-08|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|HWEE MIN LEE|Annual Leave|2026-06-15|2026-06-22|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|HWEE MIN LEE|Annual Leave|2026-06-29|2026-07-06|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|HWEE MIN LEE|Annual Leave|2026-07-13|2026-07-20|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-05-04|2026-05-11|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-05-18|2026-05-25|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-06-01|2026-06-08|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-06-15|2026-06-22|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-06-29|2026-07-06|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-07-13|2026-07-20|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|KEN HII|Annual Leave|2026-07-27|2026-08-03|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|STEVE GUASTALEGNAME|Annual Leave|2026-05-11|2026-05-18|AL",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007|STEVE GUASTALEGNAME|Annual Leave|2026-07-27|2026-08-03|AL",
]);

// Product-approved on 2026-08-13 after direct roster review. These rows exist
// only in the legacy derived calendar and have no corresponding retained MMC
// roster entry. The source-file scope makes this a one-off correction, not a
// general permission to remove leave.
const APPROVED_MMC_LEGACY_UNSUPPORTED_LEAVE = new Set([
  "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df|MICKEY FERGUSON|A/L",
  "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df|HELEN PSIHOGIOS|A/L / ANNUAL LEAVE",
  "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df|MICHELLE BERTOLUCCI|AL 9.5HRS",
  "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df|SAM HAIFI|A/L",
  "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df|STEVE TROUPAKIS|A/L / ANNUAL LEAVE",
  "automation:monash-adults:75a99896fe1c2e6d1bc33db8|HEATHER LACEY|AL 10HRS",
  "AdultTerm1.2026.xlsx:641068:1776812908257|JOSEPH VU|C/L",
]);

export function isApprovedReparseOmission(event, baselineFileId = "") {
  const source = String(event?.source || "").toUpperCase();
  const raw = normalizeRosterRawValue(event.rawValue);
  // Product-approved on 2026-08-13: this is a DDH roster-writer request
  // ("CS not onsite please"), not Clinical Support work. It is deliberately
  // exact so that real CS, CS onsite, and every other removal still block a
  // staged promotion for review.
  if (source === "DDH" && isApprovedDdhClinicalSupportRequestOmission(raw)) return true;
  // Product-approved on 2026-08-14: these exact DDH free-text requests are
  // roster-writer messages, not allocations. Their removal is intentional
  // when the shared parser stops emitting them as calendar events.
  if (source === "DDH" && isApprovedDdhRosterWriterMessageOmission(raw)) return true;
  // Product-approved: these are DDH clinical-support references entered into
  // MMC rosters to avoid unsafe late/early allocations, not MMC work.
  if (source === "MMC" && /(?:^|\s)CS\s*DH$/.test(raw)) return true;
  // Non-SMS clinicians do not have Conference Leave allocations. If an old
  // calendar row conflicts with an actual shift, the source-faithful shift
  // wins on reparse.
  const seniority = String(event?.seniority || "").trim();
  if (source === "MMC" && seniority && seniority.toUpperCase() !== "SMS"
    && /^CONFERENCE LEAVE$/i.test(String(event?.title || "").trim())) return true;
  const legacyMmcLeaveKey = [
    String(baselineFileId || ""),
    String(event?.doctorKey || "").toUpperCase(),
    raw,
  ].join("|");
  if (APPROVED_MMC_LEGACY_UNSUPPORTED_LEAVE.has(legacyMmcLeaveKey)) return true;
  if (source !== "DDH") return false;
  const exactLeaveReplacement = [
    String(baselineFileId || ""),
    String(event.doctorKey || ""),
    String(event.title || ""),
    String(event.start || "").slice(0, 10),
    String(event.end || "").slice(0, 10),
    String(event.rawValue || "").trim(),
  ].join("|");
  if (APPROVED_DDH_WEEKLY_LEAVE_REPLACEMENTS.has(exactLeaveReplacement)) return true;
  // These are references made by a DDH roster writer to another service, not
  // DDH work. They are explicitly excluded from calendars by product policy.
  if (/\b(?:TOX|HITH|HTIH|VHH|ARV|WARRAGUL|MMC|CASEY|AED|PED)\b/.test(raw)
    || /\b(?:HITH|HTIH|VHH)(?:AM|PM)\b/.test(raw)
    || /^(?:08|10|14)H?(?:00|30)$/.test(raw)) return true;
  if (!raw.startsWith("EXTRA")) return false;
  // A generic Extra entry is a payroll annotation. Explicitly timed and
  // period-labelled Extras remain rostered calendar shifts.
  const hasPeriodOrSwing = /\b(?:AM|PM|SWING)\b/.test(raw);
  const hasExplicitTime = /\b\d{1,2}(?::?\d{2})\s*(?:-|–|TO)\s*\d{1,2}(?::?\d{2})\b/.test(raw);
  return !hasPeriodOrSwing && !hasExplicitTime;
}

function isApprovedDdhRosterWriterMessageOmission(raw) {
  if (/^(?:AM|PM)\s*\(\s*AVOID IF POSSIBLE\s*\)$/.test(raw)) return true;
  if (/^AM\s+OK$/.test(raw)) return true;
  if (/^C\/S\s+FOR\s+\d{1,2}\/\d{1,2}(?:\/\d{2,4})?$/.test(raw)) return true;
  if (/^CAN\s+WORK(?:\s+\d+\s+EXTRA\s+THIS\s+WEEK)?$/.test(raw)) return true;
  if (/^CAN['’]?T\s+DO\s+THIS\s+WEEKEND(?:,?\s+SORRY!?)?$/.test(raw)) return true;
  if (/\b\d+\s+SHIFTS?\s+THIS\s+WEEK\b.*\b(?:MAKE\s+UP|NEXT\s+WEEK)\b/.test(raw)) return true;
  if (/^(?:\(\s*PREFERRED\s*\)|SEE\s+\S+|WORK(?:ING|DED)\s+\d{1,2}\/\d{1,2}|MOVED\s+TO\s+(?:SAT|SUN)\b|CLINICAL\s+SHIFT\s+MOVED\b)/.test(raw)) return true;
  if (/^(?:SUN|SA|SAT|=|L|WWEEEEWE)$/.test(raw)) return true;
  return /^(?:DAY\s+OFF|SHIFT)\s+IN\s+LIEU\b|^IN\s+LL?IEU\b|^PH\s+NOT\s+WORKED$|^PM\s+AUSTIN\s+INSTEAD\s+OF\s+AM$|^MON\s+UNI$|^PM\s+ONLU$|^Y\s+BUT\s+NOT\s+0730$/.test(raw);
}

function isApprovedDdhClinicalSupportRequestOmission(raw) {
  if (!/^CS\b/.test(raw)) return false;
  return /\bNOT\s+ON[- ]?SITE\b|\bPLS?\b|\bWORKED\s+\d{1,2}\/\d{1,2}\b|\/OFF\b|\bONLY\b|\(MONDAY\)|PVUS\s+WORKSHOP|WBA\s+COORDINATOR/.test(raw);
}

// A retained-file reparse is written under a staging id. This method performs
// the final replacement only after the caller has reviewed the comparison.
// The calendar continues to read the old active rows until this batch commits.
export async function promoteVerifiedStagedRosterFile(db, stagingFileId, targetFileId, options = {}) {
  if (!db?.prepare || !stagingFileId || !targetFileId) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const parserVersion = String(options.parserVersion || ROSTER_PARSER_VERSION);
  const comparison = await compareDerivedRosterFiles(db, targetFileId, stagingFileId, { limit: 50 });
  if (!comparison.ok) return comparison;
  const approvedRemovedEventIdentities = new Set(Array.isArray(options.approvedRemovedEventIdentities)
    ? options.approvedRemovedEventIdentities.map((identity) => String(identity || "")).filter(Boolean)
    : []);
  const hasOnlyApprovedRemovals = comparison.removedCount > 0
    && comparison.removed.length === comparison.removedCount
    && comparison.removedCount === approvedRemovedEventIdentities.size
    && comparison.removed.every((event) => approvedRemovedEventIdentities.has(`${String(event.doctorKey || "")}|${String(event.id || "")}`));
  if (comparison.removedCount > 0 && options.allowRemoved !== true && !hasOnlyApprovedRemovals) {
    return { ok: false, reason: "unreviewed-removals", comparison };
  }
  const staged = await db.prepare("SELECT name, source_type, source_id, size, last_modified, added_at, uploaded_at, uploaded_by, parsed_at FROM roster_files WHERE id = ? AND active = 0")
    .bind(String(stagingFileId)).first();
  const target = await db.prepare("SELECT id FROM roster_files WHERE id = ? AND active = 1").bind(String(targetFileId)).first();
  if (!staged || !target) return { ok: false, reason: "staging-or-target-not-active", comparison };
  const now = new Date().toISOString();
  const statements = [
    db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(String(targetFileId)),
    db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(String(targetFileId)),
    db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(String(targetFileId)),
    db.prepare("UPDATE roster_file_doctors SET file_id = ? WHERE file_id = ?").bind(String(targetFileId), String(stagingFileId)),
    db.prepare("UPDATE roster_events SET id = REPLACE(id, ?, ?), file_id = ? WHERE file_id = ?").bind(`${stagingFileId}:`, `${targetFileId}:`, String(targetFileId), String(stagingFileId)),
    db.prepare("UPDATE roster_issues SET id = REPLACE(id, ?, ?), file_id = ? WHERE file_id = ?").bind(`${stagingFileId}:`, `${targetFileId}:`, String(targetFileId), String(stagingFileId)),
    db.prepare(`
      UPDATE roster_files
      SET name = ?, source_type = ?, source_id = ?, active = 1, size = ?, last_modified = ?,
          added_at = ?, uploaded_at = ?, uploaded_by = ?, parsed_at = ?, parser_version = ?
      WHERE id = ?
    `).bind(staged.name, staged.source_type, staged.source_id, staged.size, staged.last_modified, staged.added_at, staged.uploaded_at, staged.uploaded_by, now, parserVersion, String(targetFileId)),
    db.prepare("DELETE FROM roster_files WHERE id = ?").bind(String(stagingFileId)),
  ];
  await runTransactionalBatch(db, statements);
  await deleteDailyPresenceForFile(db, targetFileId);
  await rebuildDailyPresenceForFile(db, targetFileId);
  return { ok: true, fileId: String(targetFileId), comparison };
}

export async function deleteDerivedRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return;
  await ensureCalendarSchema(db);
  const file = await db.prepare("SELECT source_type FROM roster_files WHERE id = ?").bind(fileId).first();
  const sourceType = normalizeSourceType(file?.source_type || "");
  await db.prepare("DELETE FROM roster_events WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_issues WHERE file_id = ?").bind(fileId).run();
  await db.prepare("DELETE FROM roster_file_doctors WHERE file_id = ?").bind(fileId).run();
  if (sourceType) await deleteOrphanRosterDoctors(db, [sourceType]);
  await deleteDailyPresenceForFile(db, fileId);
  await db.prepare("DELETE FROM roster_files WHERE id = ?").bind(fileId).run();
}

export async function trimDerivedRosterFileOverlap(db, fileId, startDate, endDate) {
  if (!db?.prepare || !fileId || !startDate || !endDate) return { removedEvents: 0, remainingEvents: 0, deleted: false };
  await ensureCalendarSchema(db);
  const before = await db.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ?").bind(fileId).first();
  await runTransactionalBatch(db, [
    db.prepare(`
      DELETE FROM roster_events
      WHERE file_id = ?
        AND start_date <= ?
        AND end_date >= ?
    `).bind(fileId, endDate, startDate),
    db.prepare(`
      DELETE FROM roster_issues
      WHERE file_id = ?
        AND start_date >= ?
        AND start_date <= ?
    `).bind(fileId, startDate, endDate),
    db.prepare(`
      DELETE FROM roster_file_doctors
      WHERE file_id = ?
        AND NOT EXISTS (
          SELECT 1
          FROM roster_events
          WHERE roster_events.file_id = roster_file_doctors.file_id
            AND roster_events.doctor_key = roster_file_doctors.doctor_key
        )
    `).bind(fileId),
  ]);
  const after = await db.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ?").bind(fileId).first();
  const beforeCount = Number(before?.count || 0);
  const remainingEvents = Number(after?.count || 0);
  if (!remainingEvents) {
    await deleteDerivedRosterFile(db, fileId);
    return { removedEvents: beforeCount, remainingEvents: 0, deleted: true };
  }
  await deleteDailyPresenceForFile(db, fileId);
  await rebuildDailyPresenceForFile(db, fileId);
  return {
    removedEvents: Math.max(0, beforeCount - remainingEvents),
    remainingEvents,
    deleted: false,
  };
}

export async function deleteOrphanRosterDoctors(db, sourceTypes = []) {
  if (!db?.prepare) return;
  await ensureCalendarSchema(db);
  const safeSourceTypes = [...new Set((sourceTypes || []).map(normalizeSourceType).filter(Boolean))];
  for (const sourceType of safeSourceTypes) {
    await db.prepare(`
      DELETE FROM roster_doctors
      WHERE source_type = ?
        AND NOT EXISTS (
          SELECT 1
          FROM roster_file_doctors
          WHERE roster_file_doctors.source_type = roster_doctors.source_type
            AND roster_file_doctors.doctor_key = roster_doctors.doctor_key
        )
    `).bind(sourceType).run();
  }
}

export async function verifyRosterFilesPurged(db, fileIds = []) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const ids = [...new Set((fileIds || []).map((id) => String(id || "").trim()).filter(Boolean))];
  if (!ids.length) return [];
  const results = [];
  for (const fileId of ids) {
    const d1Row = await db.prepare("SELECT id FROM roster_files WHERE id = ?").bind(fileId).first();
    const rawRow = await db.prepare("SELECT file_id FROM raw_roster_files WHERE file_id = ?").bind(fileId).first();
    const eventRow = await db.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ?").bind(fileId).first();
    const eventCount = Number(eventRow?.count || 0);
    const d1File = Boolean(d1Row?.id);
    const r2Raw = Boolean(rawRow?.file_id);
    results.push({
      fileId,
      d1File,
      r2Raw,
      eventCount,
      purged: !d1File && !r2Raw && eventCount <= 0,
    });
  }
  return results;
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

export async function deleteRetainedRosterSource(db, r2, fileId) {
  const normalizedId = String(fileId || "").trim();
  if (!normalizedId) return;
  if (!db?.prepare) return;
  await ensureCalendarSchema(db);
  const raw = await loadRawRosterFile(db, normalizedId);
  const objectKey = String(raw?.objectKey || "").trim() || `rosters/${normalizedId}`;
  if (r2?.delete) {
    await r2.delete(objectKey);
  }
  await deleteRawRosterFile(db, normalizedId);
}

const FACILITY_STAFF_DESIGNATIONS = new Set([
  "long_service_leave",
  "sabbatical_leave",
  "sick_leave",
  "personal_leave",
  "previous_staff",
]);

const FACILITY_STAFF_DESIGNATION_LABELS = {
  long_service_leave: "Long Service Leave",
  sabbatical_leave: "Sabbatical Leave",
  sick_leave: "Sick Leave",
  personal_leave: "Personal Leave",
  previous_staff: "No longer works for this ED",
};

const FACILITY_STAFF_SENIORITIES = new Set([
  "SMS",
  "CMO",
  "Senior Registrar",
  "Transitional/Intermediate Registrar",
  "Junior Registrar",
  "HMO",
  "Intern",
  "NP",
  "Physio",
  "Unknown",
]);

async function backfillFacilitySmsMemberships(db) {
  const now = new Date().toISOString();
  await db.prepare(`
    INSERT INTO facility_sms_memberships (
      source_type, doctor_key, display_name, first_seen_date, last_seen_date, created_at, updated_at
    )
    SELECT roster_file_doctors.source_type, roster_file_doctors.doctor_key,
      MAX(roster_file_doctors.display_name), MIN(roster_events.start_date), MAX(roster_events.start_date), ?, ?
    FROM roster_file_doctors
    INNER JOIN roster_events ON roster_events.file_id = roster_file_doctors.file_id
      AND roster_events.doctor_key = roster_file_doctors.doctor_key
    WHERE UPPER(roster_file_doctors.seniority) = 'SMS'
    GROUP BY roster_file_doctors.source_type, roster_file_doctors.doctor_key
    ON CONFLICT(source_type, doctor_key) DO UPDATE SET
      display_name = excluded.display_name,
      first_seen_date = MIN(facility_sms_memberships.first_seen_date, excluded.first_seen_date),
      last_seen_date = MAX(facility_sms_memberships.last_seen_date, excluded.last_seen_date),
      updated_at = excluded.updated_at
    WHERE facility_sms_memberships.display_name <> excluded.display_name
      OR facility_sms_memberships.first_seen_date > excluded.first_seen_date
      OR facility_sms_memberships.last_seen_date < excluded.last_seen_date
  `).bind(now, now).run();
}

async function recordFacilitySmsMembershipsForRosterFile(db, fileId) {
  const now = new Date().toISOString();
  await db.prepare(`
    WITH file_coverage AS (
      SELECT MIN(start_date) AS start_date, MAX(start_date) AS end_date
      FROM roster_events WHERE file_id = ?
    )
    INSERT INTO facility_sms_memberships (
      source_type, doctor_key, display_name, first_seen_date, last_seen_date, created_at, updated_at
    )
    SELECT roster_file_doctors.source_type, roster_file_doctors.doctor_key, roster_file_doctors.display_name,
      file_coverage.start_date, file_coverage.end_date, ?, ?
    FROM roster_file_doctors
    CROSS JOIN file_coverage
    WHERE roster_file_doctors.file_id = ?
      AND UPPER(roster_file_doctors.seniority) = 'SMS'
      AND file_coverage.start_date IS NOT NULL
    ON CONFLICT(source_type, doctor_key) DO UPDATE SET
      display_name = excluded.display_name,
      first_seen_date = MIN(facility_sms_memberships.first_seen_date, excluded.first_seen_date),
      last_seen_date = MAX(facility_sms_memberships.last_seen_date, excluded.last_seen_date),
      updated_at = excluded.updated_at
    WHERE facility_sms_memberships.display_name <> excluded.display_name
      OR facility_sms_memberships.first_seen_date > excluded.first_seen_date
      OR facility_sms_memberships.last_seen_date < excluded.last_seen_date
  `).bind(fileId, now, now, fileId).run();
}

export async function setFacilityStaffSeniorityOverride(db, input = {}) {
  if (!db?.prepare) throw new Error("Roster database is unavailable.");
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(input.sourceType || input.facilityKey);
  const doctorKey = String(input.doctorKey || "").trim();
  const termStart = datePart(input.termStart);
  const useRosterSeniority = input.useRosterSeniority === true;
  const seniority = useRosterSeniority ? "" : normalizeFacilityStaffSeniority(input.seniority);
  if (!sourceType || !doctorKey || !termStart || (!useRosterSeniority && !FACILITY_STAFF_SENIORITIES.has(seniority))) {
    throw new Error("A valid ED staff seniority is required.");
  }
  const id = facilityStaffSeniorityOverrideId(sourceType, doctorKey, termStart);
  const now = new Date().toISOString();
  await db.prepare(`
    INSERT INTO facility_staff_seniority_overrides (
      id, source_type, doctor_key, display_name, seniority, use_roster_seniority,
      term_start, active, created_by, created_at, updated_at, cleared_at, cleared_reason
    ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, '', '')
    ON CONFLICT(id) DO UPDATE SET
      display_name = excluded.display_name,
      seniority = excluded.seniority,
      use_roster_seniority = excluded.use_roster_seniority,
      active = 1,
      created_by = excluded.created_by,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at,
      cleared_at = '',
      cleared_reason = ''
  `).bind(
    id,
    sourceType,
    doctorKey,
    String(input.displayName || "").trim(),
    seniority,
    useRosterSeniority ? 1 : 0,
    termStart,
    normalizeEmail(input.createdBy),
    now,
    now,
  ).run();
  return loadFacilityStaffSeniorityOverride(db, id);
}

export async function setFacilityStaffSeniorityOverrides(db, input = {}) {
  if (!db?.prepare) throw new Error("Roster database is unavailable.");
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(input.sourceType || input.facilityKey);
  const termStart = datePart(input.termStart);
  const useRosterSeniority = input.useRosterSeniority === true;
  const seniority = useRosterSeniority ? "" : normalizeFacilityStaffSeniority(input.seniority);
  const suppliedStaff = Array.isArray(input.staff) ? input.staff : [];
  if (!sourceType || !termStart || (!useRosterSeniority && !FACILITY_STAFF_SENIORITIES.has(seniority))) {
    throw new Error("A valid ED staff seniority is required.");
  }
  const staff = [];
  const seen = new Set();
  for (const person of suppliedStaff) {
    const doctorKey = String(person?.doctorKey || "").trim();
    if (!doctorKey) throw new Error("Each selected staff member must be valid.");
    if (seen.has(doctorKey)) continue;
    seen.add(doctorKey);
    staff.push({ doctorKey, displayName: String(person?.displayName || "").trim() });
  }
  if (!staff.length) throw new Error("Choose at least one staff member.");

  const now = new Date().toISOString();
  const createdBy = normalizeEmail(input.createdBy);
  const statements = staff.map((person) => db.prepare(`
    INSERT INTO facility_staff_seniority_overrides (
      id, source_type, doctor_key, display_name, seniority, use_roster_seniority,
      term_start, active, created_by, created_at, updated_at, cleared_at, cleared_reason
    ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, '', '')
    ON CONFLICT(id) DO UPDATE SET
      display_name = excluded.display_name,
      seniority = excluded.seniority,
      use_roster_seniority = excluded.use_roster_seniority,
      active = 1,
      created_by = excluded.created_by,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at,
      cleared_at = '',
      cleared_reason = ''
  `).bind(
    facilityStaffSeniorityOverrideId(sourceType, person.doctorKey, termStart),
    sourceType,
    person.doctorKey,
    person.displayName,
    seniority,
    useRosterSeniority ? 1 : 0,
    termStart,
    createdBy,
    now,
    now,
  ));
  if (typeof db.batch === "function") await db.batch(statements);
  else await Promise.all(statements.map((statement) => statement.run()));
  return Promise.all(staff.map((person) => loadFacilityStaffSeniorityOverride(db, facilityStaffSeniorityOverrideId(sourceType, person.doctorKey, termStart))));
}

export async function queryFacilityStaffSeniorityOverrides(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(options.sourceType || options.facilityKey || "");
  const termStart = datePart(options.termStart);
  if (!termStart) return [];
  const sourceSql = sourceType ? "AND source_type = ?" : "";
  const bindings = sourceType ? [termStart, sourceType] : [termStart];
  const rows = await db.prepare(`
    SELECT *
    FROM facility_staff_seniority_overrides
    WHERE active = 1 AND term_start <= ? ${sourceSql}
    ORDER BY source_type, doctor_key, term_start DESC
  `).bind(...bindings).all();
  const latest = new Map();
  for (const row of rows.results || []) {
    const override = facilityStaffSeniorityOverrideFromRow(row);
    const key = `${override?.sourceType || ""}|${override?.doctorKey || ""}`;
    if (override && !latest.has(key)) latest.set(key, override);
  }
  return [...latest.values()];
}

async function loadFacilityStaffSeniorityOverride(db, id) {
  const row = await db.prepare("SELECT * FROM facility_staff_seniority_overrides WHERE id = ?").bind(String(id || "")).first();
  return facilityStaffSeniorityOverrideFromRow(row);
}

function facilityStaffSeniorityOverrideId(sourceType, doctorKey, termStart) {
  return `facility-seniority:${sourceType}:${String(doctorKey).trim()}:${termStart}`;
}

function facilityStaffSeniorityOverrideFromRow(row) {
  if (!row?.id) return null;
  const useRosterSeniority = Number(row.use_roster_seniority || 0) === 1;
  const seniority = normalizeFacilityStaffSeniority(row.seniority);
  if (!useRosterSeniority && !FACILITY_STAFF_SENIORITIES.has(seniority)) return null;
  return {
    id: String(row.id), sourceType: normalizeSourceType(row.source_type), doctorKey: String(row.doctor_key || "").trim(),
    displayName: String(row.display_name || "").trim(), seniority, useRosterSeniority,
    termStart: datePart(row.term_start), active: Number(row.active || 0) === 1,
    createdAt: String(row.created_at || ""), updatedAt: String(row.updated_at || ""),
  };
}

function normalizeFacilityStaffSeniority(value) {
  const seniority = String(value || "").trim();
  const normalized = seniority.toLowerCase();
  if (normalized === "sr" || normalized.includes("senior registrar")) return "Senior Registrar";
  if (normalized === "tr" || normalized === "ir" || normalized.includes("transitional") || normalized.includes("intermediate")) return "Transitional/Intermediate Registrar";
  if (normalized === "jr" || normalized.includes("junior registrar")) return "Junior Registrar";
  if (normalized === "enp" || normalized === "np" || normalized.includes("nurse practitioner")) return "NP";
  if (normalized === "amp" || normalized === "physio" || normalized.includes("physiotherapist")) return "Physio";
  if (normalized === "sms" || normalized === "cmo" || normalized === "hmo") return normalized.toUpperCase();
  if (normalized === "intern" || normalized === "i") return "Intern";
  if (normalized === "unknown") return "Unknown";
  return seniority;
}

export async function setFacilityStaffDesignation(db, input = {}) {
  if (!db?.prepare) throw new Error("Roster database is unavailable.");
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(input.sourceType || input.facilityKey);
  const doctorKey = String(input.doctorKey || "").trim();
  const designation = String(input.designation || "").trim();
  const termStart = datePart(input.termStart);
  const termEnd = datePart(input.termEnd);
  if (!sourceType || !doctorKey || !FACILITY_STAFF_DESIGNATIONS.has(designation) || !termStart || !termEnd || termEnd < termStart) {
    throw new Error("A valid ED staff designation is required.");
  }
  const seniority = String(input.seniority || "").trim() || "Unknown";
  if (designation === "previous_staff") {
    const smsRecord = await db.prepare(`
      SELECT 1 AS found
      FROM facility_sms_memberships
      WHERE source_type = ? AND doctor_key = ?
      LIMIT 1
    `).bind(sourceType, doctorKey).first() || await db.prepare(`
      SELECT 1 AS found
      FROM roster_file_doctors
      INNER JOIN roster_files ON roster_files.id = roster_file_doctors.file_id
      WHERE roster_files.active = 1
        AND roster_file_doctors.source_type = ?
        AND roster_file_doctors.doctor_key = ?
        AND roster_file_doctors.seniority = 'SMS'
      LIMIT 1
    `).bind(sourceType, doctorKey).first();
    if (!smsRecord) throw new Error("Only SMS staff can be moved to Previous staff.");
  }
  const id = facilityStaffDesignationId(sourceType, doctorKey, termStart);
  const now = new Date().toISOString();
  const sourceRevision = await latestFacilityRosterRevision(db, sourceType);
  await db.prepare(`
    INSERT INTO facility_staff_designations (
      id, source_type, doctor_key, display_name, seniority, designation,
      term_start, term_end, source_revision, active, created_by, created_at,
      updated_at, cleared_at, cleared_reason
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, '', '')
    ON CONFLICT(id) DO UPDATE SET
      display_name = excluded.display_name,
      seniority = excluded.seniority,
      designation = excluded.designation,
      term_end = excluded.term_end,
      source_revision = excluded.source_revision,
      active = 1,
      created_by = excluded.created_by,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at,
      cleared_at = '',
      cleared_reason = ''
  `).bind(
    id,
    sourceType,
    doctorKey,
    String(input.displayName || "").trim(),
    seniority,
    designation,
    termStart,
    termEnd,
    sourceRevision,
    normalizeEmail(input.createdBy),
    now,
    now,
  ).run();
  return loadFacilityStaffDesignation(db, id);
}

export async function clearFacilityStaffDesignation(db, designationId, options = {}) {
  if (!db?.prepare || !designationId) return null;
  await ensureCalendarSchema(db);
  const now = new Date().toISOString();
  await db.prepare(`
    UPDATE facility_staff_designations
    SET active = 0, updated_at = ?, cleared_at = ?, cleared_reason = ?
    WHERE id = ? AND active = 1
  `).bind(now, now, String(options.reason || "creator-undo").slice(0, 80), String(designationId)).run();
  return loadFacilityStaffDesignation(db, designationId);
}

export async function queryFacilityStaffDesignations(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const sourceType = normalizeSourceType(options.sourceType || options.facilityKey || "");
  const termStart = datePart(options.termStart);
  const termEnd = datePart(options.termEnd);
  if (!termStart || !termEnd) return [];
  const sourceSql = sourceType ? "AND source_type = ?" : "";
  const bindings = sourceType ? [termStart, termStart, sourceType] : [termStart, termStart];
  const rows = await db.prepare(`
    SELECT *
    FROM facility_staff_designations
    WHERE active = 1 ${sourceSql}
      AND (
        designation = 'previous_staff' AND term_start <= ?
        OR designation <> 'previous_staff' AND term_start = ?
      )
    ORDER BY source_type, designation, display_name, term_start
  `).bind(...bindings).all();
  return (rows.results || []).map(facilityStaffDesignationFromRow).filter(Boolean);
}

export async function reconcileFacilityStaffDesignationsForRosterFile(db, fileId) {
  if (!db?.prepare || !fileId) return { cleared: 0 };
  await ensureCalendarSchema(db);
  const file = await db.prepare("SELECT id, source_type, parsed_at FROM roster_files WHERE id = ? AND active = 1").bind(String(fileId)).first();
  const sourceType = normalizeSourceType(file?.source_type || "");
  const parsedAt = String(file?.parsed_at || "");
  if (!sourceType || !parsedAt) return { cleared: 0 };
  const coverage = await db.prepare("SELECT MIN(start_date) AS start_date, MAX(start_date) AS end_date FROM roster_events WHERE file_id = ?").bind(String(fileId)).first();
  const coverageStart = datePart(coverage?.start_date);
  const coverageEnd = datePart(coverage?.end_date);
  if (!coverageStart || !coverageEnd) return { cleared: 0 };
  const now = new Date().toISOString();
  const result = await db.prepare(`
    UPDATE facility_staff_designations
    SET active = 0, updated_at = ?, cleared_at = ?, cleared_reason = 'new-roster-membership'
    WHERE active = 1
      AND source_type = ?
      AND source_revision <> ?
      AND created_at < ?
      AND EXISTS (
        SELECT 1 FROM roster_file_doctors
        WHERE roster_file_doctors.file_id = ?
          AND roster_file_doctors.source_type = facility_staff_designations.source_type
          AND roster_file_doctors.doctor_key = facility_staff_designations.doctor_key
      )
      AND (
        designation = 'previous_staff' AND term_start <= ?
        OR designation <> 'previous_staff' AND term_start <= ? AND term_end >= ?
      )
  `).bind(now, now, sourceType, String(fileId), parsedAt, String(fileId), coverageEnd, coverageEnd, coverageStart).run();
  return { cleared: Number(result?.meta?.changes || result?.changes || 0) };
}

async function queryFacilityDesignationLeaveEvents(db, doctorKeys = [], options = {}) {
  const keys = [...new Set((doctorKeys || []).map((key) => String(key || "").trim()).filter(Boolean))];
  if (!keys.length) return [];
  const startDate = datePart(options.startDate || "0000-01-01");
  const endDate = datePart(options.endDate || "9999-12-31");
  const sources = sanitizeSourceTypes(options.sourceTypes || []);
  const sourceSql = sources.length ? `AND source_type IN (${sources.map(() => "?").join(", ")})` : "";
  const rows = await db.prepare(`
    SELECT *
    FROM facility_staff_designations
    WHERE active = 1
      AND designation <> 'previous_staff'
      AND doctor_key IN (${keys.map(() => "?").join(", ")})
      AND term_start <= ? AND term_end >= ?
      ${sourceSql}
    ORDER BY term_start, source_type, display_name
  `).bind(...keys, endDate, startDate, ...sources).all();
  return (rows.results || []).map((row) => {
    const designation = facilityStaffDesignationFromRow(row);
    if (!designation) return null;
    return {
      id: `facility-designation:${designation.id}`,
      source: displaySourceType(designation.sourceType),
      title: FACILITY_STAFF_DESIGNATION_LABELS[designation.designation] || "Leave",
      allDay: true,
      start: designation.termStart,
      end: nextDateKey(designation.termEnd),
      rawValue: "Creator staff designation",
      seniority: designation.seniority,
      location: "",
      timeLabel: "All day",
      designationId: designation.id,
      doctorKey: designation.doctorKey,
      displayName: designation.displayName,
    };
  }).filter(Boolean);
}

async function latestFacilityRosterRevision(db, sourceType) {
  const row = await db.prepare(`
    SELECT id FROM roster_files
    WHERE active = 1 AND source_type = ?
    ORDER BY parsed_at DESC, id DESC LIMIT 1
  `).bind(sourceType).first();
  return String(row?.id || "");
}

async function loadFacilityStaffDesignation(db, id) {
  const row = await db.prepare("SELECT * FROM facility_staff_designations WHERE id = ?").bind(String(id || "")).first();
  return facilityStaffDesignationFromRow(row);
}

function facilityStaffDesignationId(sourceType, doctorKey, termStart) {
  return `facility-designation:${sourceType}:${String(doctorKey).trim()}:${termStart}`;
}

function facilityStaffDesignationFromRow(row) {
  if (!row?.id || !FACILITY_STAFF_DESIGNATIONS.has(String(row.designation || ""))) return null;
  return {
    id: String(row.id), sourceType: normalizeSourceType(row.source_type), doctorKey: String(row.doctor_key || "").trim(),
    displayName: String(row.display_name || "").trim(), seniority: String(row.seniority || "Unknown").trim() || "Unknown",
    designation: String(row.designation), label: FACILITY_STAFF_DESIGNATION_LABELS[String(row.designation)] || "",
    termStart: datePart(row.term_start), termEnd: datePart(row.term_end), active: Number(row.active || 0) === 1,
    createdAt: String(row.created_at || ""), updatedAt: String(row.updated_at || ""), clearedAt: String(row.cleared_at || ""),
    clearedReason: String(row.cleared_reason || ""),
  };
}

function displaySourceType(sourceType) {
  const source = normalizeSourceType(sourceType);
  if (source === "casey") return "Casey";
  if (source === "mch") return "MCH";
  return source.toUpperCase();
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
  const designationEvents = await queryFacilityDesignationLeaveEvents(db, keys, { startDate: start, endDate: end });
  return mergeDuplicateLeaveEvents([
    ...(rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean),
    ...designationEvents,
  ]);
}

// Event ids are generated by the parser from its source occurrence.  Claims
// can legitimately overlap, so only this established identity is safe to
// collapse; two shifts with matching times but different ids remain visible.
export function dedupeEventsByIdentity(events = []) {
  const seen = new Set();
  return (events || []).filter((event) => {
    const identity = String(event?.id || "").trim();
    if (!identity) return true;
    if (seen.has(identity)) return false;
    seen.add(identity);
    return true;
  });
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
  const designationEvents = await queryFacilityDesignationLeaveEvents(
    db,
    [...new Set(safePairs.map((pair) => pair.doctorKey))],
    {
      startDate: start,
      endDate: end,
    },
  );
  return mergeDuplicateLeaveEvents([
    ...(rows.results || []).map((row) => parseEvent(row.event_json)).filter(Boolean),
    ...designationEvents,
  ]);
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

export async function queryUnresolvedRosterShiftIssueRows(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const limit = Math.min(Math.max(Number(options.limit || 5000), 1), 10000);
  const rows = await db.prepare(`
    SELECT
      roster_issues.id AS id,
      roster_issues.file_id AS file_id,
      roster_issues.source_type AS source_type,
      roster_issues.doctor_key AS doctor_key,
      roster_issues.display_name AS display_name,
      roster_issues.start_date AS start_date,
      roster_issues.raw_value AS raw_value,
      roster_issues.seniority AS seniority,
      roster_issues.status AS status,
      roster_issues.message AS message,
      roster_issues.resolution_type AS resolution_type,
      roster_issues.suggested_title AS suggested_title,
      roster_issues.time_label AS time_label,
      roster_issues.issue_json AS issue_json,
      roster_files.name AS file_name
    FROM roster_issues
    INNER JOIN roster_files ON roster_files.id = roster_issues.file_id
    WHERE roster_files.active = 1
      AND roster_issues.raw_value <> ''
      AND roster_issues.message <> ''
      AND (
        roster_issues.resolution_type = 'shift_code'
        OR roster_issues.status = 'unknown'
        OR LOWER(roster_issues.message) LIKE '%shift code not recognised%'
        OR LOWER(roster_issues.message) LIKE '%shift label not recognised%'
        OR LOWER(roster_issues.message) LIKE '%shift code not recognized%'
        OR LOWER(roster_issues.message) LIKE '%shift label not recognized%'
      )
    ORDER BY roster_issues.start_date DESC, roster_issues.source_type, roster_issues.raw_value
    LIMIT ?
  `).bind(limit).all();
  return (rows.results || []).map((row) => {
    const issue = parseIssue(row.issue_json) || sanitizeIssue({
      id: row.id,
      source: row.source_type,
      seniority: row.seniority,
      startDay: row.start_date,
      rawValue: row.raw_value,
      status: row.status,
      message: row.message,
      resolutionType: row.resolution_type,
      suggestedTitle: row.suggested_title,
      timeLabel: row.time_label,
    });
    if (!issue) return null;
    return {
      ...issue,
      fileId: String(row.file_id || ""),
      fileName: String(row.file_name || ""),
      doctorKey: String(row.doctor_key || ""),
      displayName: String(row.display_name || ""),
      startDay: issue.startDay || String(row.start_date || "").slice(0, 10),
      rawValue: issue.rawValue || String(row.raw_value || ""),
      seniority: issue.seniority || String(row.seniority || ""),
      message: issue.message || String(row.message || ""),
      resolutionType: issue.resolutionType || String(row.resolution_type || ""),
      suggestedTitle: issue.suggestedTitle || String(row.suggested_title || ""),
      timeLabel: issue.timeLabel || String(row.time_label || ""),
    };
  }).filter(Boolean);
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
      roster_files.source_id AS source_id,
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
    ${includeInactive ? "" : "WHERE file_id IN (SELECT id FROM roster_files WHERE active = 1)"}
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
    sourceId: String(row.source_id || "").trim(),
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

export async function listActiveRetainedRosterFiles(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT
      roster_files.id AS id, roster_files.name AS name, roster_files.source_type AS source_type,
      roster_files.source_id AS source_id, roster_files.size AS size,
      roster_files.last_modified AS last_modified, raw_roster_files.object_key AS object_key,
      raw_roster_files.data_url AS data_url, raw_roster_files.type AS type
    FROM roster_files
    INNER JOIN raw_roster_files ON raw_roster_files.file_id = roster_files.id
    WHERE roster_files.active = 1
      AND (raw_roster_files.object_key <> '' OR raw_roster_files.data_url <> '')
    ORDER BY roster_files.added_at, roster_files.id
  `).all();
  return (rows.results || []).map((row) => ({
    id: String(row.id || ""), name: String(row.name || "roster.xlsx"),
    sourceType: String(row.source_type || "").toLowerCase(), sourceId: String(row.source_id || ""),
    size: Number(row.size || 0), lastModified: Number(row.last_modified || 0),
    objectKey: String(row.object_key || ""), dataUrl: String(row.data_url || ""), type: String(row.type || ""),
  })).filter((file) => file.id && SOURCE_TYPES.includes(file.sourceType));
}

export async function querySourceTypesForFileIds(db, fileIds = []) {
  if (!db?.prepare) return [];
  const ids = [...new Set((Array.isArray(fileIds) ? fileIds : []).map((id) => String(id || "").trim()).filter(Boolean))];
  if (!ids.length) return [];
  await ensureCalendarSchema(db);
  const placeholders = ids.map(() => "?").join(", ");
  const rows = await db.prepare(`
    SELECT DISTINCT source_type AS source_type
    FROM roster_files
    WHERE id IN (${placeholders})
  `).bind(...ids).all();
  return [...new Set((rows?.results || rows || [])
    .map((row) => String(row?.source_type || "").toLowerCase())
    .filter((sourceType) => SOURCE_TYPES.includes(sourceType)))];
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
      roster_files.source_id AS source_id,
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
    sourceId: String(row.source_id || "").trim(),
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

export async function upsertRosterSource(db, source = {}) {
  if (!db?.prepare || !String(source.id || "").trim()) return { ok: false, reason: "missing-source-id" };
  await ensureCalendarSchema(db);
  const now = new Date().toISOString();
  const id = String(source.id).trim();
  await db.prepare(`
    INSERT INTO roster_sources (
      id, provider, source_type, label, enabled, config_json, cursor_json, provider_version, provider_modified_at,
      last_checked_at, last_success_at, last_error, active_file_id, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      provider = excluded.provider,
      source_type = excluded.source_type,
      label = excluded.label,
      enabled = excluded.enabled,
      config_json = excluded.config_json,
      cursor_json = excluded.cursor_json,
      provider_version = excluded.provider_version,
      provider_modified_at = excluded.provider_modified_at,
      last_checked_at = excluded.last_checked_at,
      last_success_at = excluded.last_success_at,
      last_error = excluded.last_error,
      active_file_id = excluded.active_file_id,
      updated_at = excluded.updated_at
  `).bind(
    id,
    String(source.provider || ""),
    normalizeSourceType(source.sourceType) || "",
    String(source.label || id),
    source.enabled === true ? 1 : 0,
    JSON.stringify(source.config && typeof source.config === "object" ? source.config : {}),
    JSON.stringify(source.cursor && typeof source.cursor === "object" ? source.cursor : {}),
    String(source.providerVersion || ""),
    String(source.providerModifiedAt || ""),
    String(source.lastCheckedAt || ""),
    String(source.lastSuccessAt || ""),
    String(source.lastError || ""),
    String(source.activeFileId || ""),
    String(source.createdAt || now),
    String(source.updatedAt || now),
  ).run();
  return { ok: true, id };
}

export async function loadRosterSource(db, sourceId) {
  if (!db?.prepare || !String(sourceId || "").trim()) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare("SELECT * FROM roster_sources WHERE id = ?").bind(String(sourceId).trim()).first();
  return rosterSourceFromRow(row);
}

export async function listRosterSources(db) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare("SELECT * FROM roster_sources ORDER BY label, id").all();
  return (rows.results || []).map(rosterSourceFromRow).filter(Boolean);
}

export async function listRosterSyncRuns(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const limit = Math.min(Math.max(Number(options.limit || 50), 1), 250);
  const rows = await db.prepare(`
    SELECT * FROM roster_sync_runs
    ORDER BY started_at DESC, id DESC
    LIMIT ?
  `).bind(limit).all();
  return (rows.results || []).map(rosterSyncRunFromRow).filter(Boolean);
}

export async function findSuccessfulRosterSyncByHash(db, sourceId, contentHash, fileName = "") {
  if (!db?.prepare || !sourceId || !contentHash) return null;
  await ensureCalendarSchema(db);
  const hasFileName = Boolean(String(fileName || "").trim());
  const row = await db.prepare(hasFileName ? `
    SELECT roster_sync_runs.*
    FROM roster_sync_runs
    INNER JOIN raw_roster_files ON raw_roster_files.file_id = COALESCE(NULLIF(roster_sync_runs.source_file_id, ''), roster_sync_runs.file_id)
    WHERE roster_sync_runs.source_id = ?
      AND roster_sync_runs.content_hash = ?
      AND roster_sync_runs.status = 'success'
      AND LOWER(raw_roster_files.name) = LOWER(?)
    ORDER BY roster_sync_runs.completed_at DESC LIMIT 1
  ` : `
    SELECT * FROM roster_sync_runs
    WHERE source_id = ? AND content_hash = ? AND status = 'success'
    ORDER BY completed_at DESC LIMIT 1
  `).bind(...(hasFileName
    ? [String(sourceId), String(contentHash), String(fileName)]
    : [String(sourceId), String(contentHash)])).first();
  return rosterSyncRunFromRow(row);
}

export async function findQueuedRosterSyncByHash(db, sourceId, contentHash, fileName = "") {
  if (!db?.prepare || !sourceId || !contentHash) return null;
  await ensureCalendarSchema(db);
  const hasFileName = Boolean(String(fileName || "").trim());
  const row = await db.prepare(hasFileName ? `
    SELECT roster_sync_runs.*
    FROM roster_sync_runs
    INNER JOIN raw_roster_files ON raw_roster_files.file_id = COALESCE(NULLIF(roster_sync_runs.source_file_id, ''), roster_sync_runs.file_id)
    WHERE roster_sync_runs.source_id = ?
      AND roster_sync_runs.content_hash = ?
      AND roster_sync_runs.status IN ('queued', 'processing')
      AND LOWER(raw_roster_files.name) = LOWER(?)
    ORDER BY roster_sync_runs.started_at DESC LIMIT 1
  ` : `
    SELECT * FROM roster_sync_runs
    WHERE source_id = ? AND content_hash = ? AND status IN ('queued', 'processing')
    ORDER BY started_at DESC LIMIT 1
  `).bind(...(hasFileName
    ? [String(sourceId), String(contentHash), String(fileName)]
    : [String(sourceId), String(contentHash)])).first();
  return rosterSyncRunFromRow(row);
}

export async function findRosterSyncByProviderVersion(db, sourceId, providerVersion, fileName) {
  if (!db?.prepare || !sourceId || !providerVersion || !fileName) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare(`
    SELECT roster_sync_runs.*
    FROM roster_sync_runs
    INNER JOIN raw_roster_files ON raw_roster_files.file_id = COALESCE(NULLIF(roster_sync_runs.source_file_id, ''), roster_sync_runs.file_id)
    WHERE roster_sync_runs.source_id = ?
      AND roster_sync_runs.provider_version = ?
      AND LOWER(raw_roster_files.name) = LOWER(?)
      AND roster_sync_runs.status IN ('success', 'queued', 'processing', 'failed')
    ORDER BY
      CASE roster_sync_runs.status
        WHEN 'success' THEN 0
        WHEN 'processing' THEN 1
        WHEN 'queued' THEN 2
        ELSE 3
      END,
      roster_sync_runs.started_at DESC
    LIMIT 1
  `).bind(String(sourceId), String(providerVersion), String(fileName)).first();
  return rosterSyncRunFromRow(row);
}

export async function loadRosterSyncRun(db, runId) {
  if (!db?.prepare || !runId) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare("SELECT * FROM roster_sync_runs WHERE id = ?").bind(String(runId)).first();
  return rosterSyncRunFromRow(row);
}

export async function listQueuedRosterSyncRuns(db, limit = 4) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  await supersedeObsoleteQueuedRosterSyncRuns(db);
  const safeLimit = Math.min(Math.max(Number(limit || 4), 1), 20);
  const rows = await db.prepare(`
    SELECT
      roster_sync_runs.*,
      raw_roster_files.name AS file_name,
      raw_roster_files.source_type AS file_source_type,
      raw_roster_files.size AS file_size,
      raw_roster_files.last_modified AS file_last_modified,
      raw_roster_files.type AS file_type,
      raw_roster_files.object_key AS object_key,
      roster_sources.provider_modified_at AS provider_modified_at
    FROM roster_sync_runs
    INNER JOIN raw_roster_files ON raw_roster_files.file_id = COALESCE(NULLIF(roster_sync_runs.source_file_id, ''), roster_sync_runs.file_id)
    LEFT JOIN roster_sources ON roster_sources.id = roster_sync_runs.source_id
    WHERE roster_sync_runs.status IN ('queued', 'processing')
      AND (
        roster_sync_runs.provider_version = ''
        OR NOT EXISTS (
          SELECT 1
          FROM roster_sync_runs AS earlier_run
          INNER JOIN raw_roster_files AS earlier_file ON earlier_file.file_id = COALESCE(NULLIF(earlier_run.source_file_id, ''), earlier_run.file_id)
          WHERE earlier_run.source_id = roster_sync_runs.source_id
            AND earlier_run.provider_version = roster_sync_runs.provider_version
            AND LOWER(earlier_file.name) = LOWER(raw_roster_files.name)
            AND earlier_run.status IN ('queued', 'processing')
            AND (
              earlier_run.started_at < roster_sync_runs.started_at
              OR (earlier_run.started_at = roster_sync_runs.started_at AND earlier_run.id < roster_sync_runs.id)
            )
        )
      )
      AND NOT EXISTS (
        SELECT 1
        FROM roster_sync_runs AS newer_run
        INNER JOIN raw_roster_files AS newer_file ON newer_file.file_id = COALESCE(NULLIF(newer_run.source_file_id, ''), newer_run.file_id)
        WHERE newer_run.source_id = roster_sync_runs.source_id
          AND LOWER(newer_file.name) = LOWER(raw_roster_files.name)
          AND newer_run.status IN ('queued', 'processing')
          AND (
            newer_run.started_at > roster_sync_runs.started_at
            OR (newer_run.started_at = roster_sync_runs.started_at AND newer_run.id > roster_sync_runs.id)
          )
      )
    ORDER BY roster_sync_runs.started_at ASC
    LIMIT ?
  `).bind(safeLimit).all();
  return (rows.results || []).map((row) => ({
    ...rosterSyncRunFromRow(row),
    fileName: String(row.file_name || "roster.xlsx"),
    sourceType: String(row.file_source_type || "").toLowerCase(),
    size: Number(row.file_size || 0),
    lastModified: Number(row.file_last_modified || 0),
    contentType: String(row.file_type || "application/octet-stream"),
    objectKey: String(row.object_key || ""),
    providerModifiedAt: String(row.provider_modified_at || ""),
  })).filter((run) => run.id && run.fileId && run.objectKey);
}

export async function supersedeObsoleteQueuedRosterSyncRuns(db) {
  if (!db?.prepare) return { ok: false, reason: "missing-db" };
  await ensureCalendarSchema(db);
  const completedAt = new Date().toISOString();
  const result = await db.prepare(`
    UPDATE roster_sync_runs AS stale_run
    SET status = 'superseded', message = ?, completed_at = ?
    WHERE stale_run.status = 'queued'
      AND EXISTS (
        SELECT 1
        FROM roster_sync_runs AS newer_run
        INNER JOIN raw_roster_files AS stale_file ON stale_file.file_id = COALESCE(NULLIF(stale_run.source_file_id, ''), stale_run.file_id)
        INNER JOIN raw_roster_files AS newer_file ON newer_file.file_id = COALESCE(NULLIF(newer_run.source_file_id, ''), newer_run.file_id)
        WHERE newer_run.source_id = stale_run.source_id
          AND LOWER(newer_file.name) = LOWER(stale_file.name)
          AND newer_run.status IN ('queued', 'processing')
          AND (
            newer_run.started_at > stale_run.started_at
            OR (newer_run.started_at = stale_run.started_at AND newer_run.id > stale_run.id)
          )
      )
  `).bind(
    "A newer version of this roster file is queued for processing.",
    completedAt,
  ).run();
  return { ok: true, changes: Number(result?.meta?.changes || result?.changes || 0) };
}

export async function claimRosterDispatch(db, { reason = "", retryAfter = "", now = new Date().toISOString() } = {}) {
  if (!db?.prepare) return { claimed: false, reason: "missing-db" };
  await ensureCalendarSchema(db);
  const pending = await db.prepare("SELECT id FROM roster_sync_runs WHERE status IN ('queued', 'processing') LIMIT 1").first();
  if (!pending?.id) return { claimed: false, reason: "queue-empty" };
  const active = await db.prepare(`
    SELECT * FROM roster_dispatches
    WHERE status IN ('requested', 'accepted', 'running', 'failed')
      AND retry_after > ?
    ORDER BY requested_at DESC
    LIMIT 1
  `).bind(String(now)).first();
  if (active?.id) return { claimed: false, reason: "already-dispatched", dispatch: rosterDispatchFromRow(active) };
  const id = `dispatch:${crypto.randomUUID()}`;
  await db.prepare(`
    INSERT INTO roster_dispatches (
      id, status, reason, requested_at, retry_after, attempt_count
    ) VALUES (?, 'requested', ?, ?, ?, 1)
  `).bind(id, String(reason || ""), String(now), String(retryAfter || now)).run();
  return { claimed: true, dispatch: { id, status: "requested", reason: String(reason || ""), requestedAt: String(now), retryAfter: String(retryAfter || now), attemptCount: 1, lastError: "" } };
}

export async function updateRosterDispatch(db, dispatchId, update = {}) {
  if (!db?.prepare || !dispatchId) return { ok: false, reason: "missing-input" };
  await ensureCalendarSchema(db);
  const existing = await db.prepare("SELECT * FROM roster_dispatches WHERE id = ?").bind(String(dispatchId)).first();
  if (!existing) return { ok: false, reason: "not-found" };
  const next = {
    status: String(update.status || existing.status || "requested"),
    githubRunId: String(update.githubRunId ?? existing.github_run_id ?? ""),
    acceptedAt: String(update.acceptedAt ?? existing.accepted_at ?? ""),
    startedAt: String(update.startedAt ?? existing.started_at ?? ""),
    completedAt: String(update.completedAt ?? existing.completed_at ?? ""),
    retryAfter: String(update.retryAfter ?? existing.retry_after ?? ""),
    attemptCount: Number(update.attemptCount ?? existing.attempt_count ?? 0),
    lastError: String(update.lastError ?? existing.last_error ?? ""),
  };
  await db.prepare(`
    UPDATE roster_dispatches
    SET status = ?, github_run_id = ?, accepted_at = ?, started_at = ?, completed_at = ?,
      retry_after = ?, attempt_count = ?, last_error = ?
    WHERE id = ?
  `).bind(
    next.status, next.githubRunId, next.acceptedAt, next.startedAt, next.completedAt,
    next.retryAfter, next.attemptCount, next.lastError, String(dispatchId),
  ).run();
  return { ok: true, dispatch: { id: String(dispatchId), ...next } };
}

export async function loadLatestRosterDispatch(db) {
  if (!db?.prepare) return null;
  await ensureCalendarSchema(db);
  const row = await db.prepare("SELECT * FROM roster_dispatches ORDER BY requested_at DESC LIMIT 1").first();
  return row ? rosterDispatchFromRow(row) : null;
}

function rosterDispatchFromRow(row = {}) {
  return {
    id: String(row.id || ""),
    status: String(row.status || ""),
    reason: String(row.reason || ""),
    githubRunId: String(row.github_run_id || ""),
    requestedAt: String(row.requested_at || ""),
    acceptedAt: String(row.accepted_at || ""),
    startedAt: String(row.started_at || ""),
    completedAt: String(row.completed_at || ""),
    retryAfter: String(row.retry_after || ""),
    attemptCount: Number(row.attempt_count || 0),
    lastError: String(row.last_error || ""),
  };
}

export async function markRosterSyncRunProcessing(db, runId) {
  if (!db?.prepare || !runId) return { ok: false, reason: "missing-run-id" };
  await ensureCalendarSchema(db);
  await db.prepare(`
    UPDATE roster_sync_runs
    SET status = 'processing', message = ?, completed_at = ''
    WHERE id = ?
  `).bind("Background processing started.", String(runId)).run();
  return { ok: true };
}

export async function createRosterSyncRun(db, run = {}) {
  if (!db?.prepare || !run?.id || !run?.sourceId) return { ok: false, reason: "missing-run-input" };
  await ensureCalendarSchema(db);
  await db.prepare(`
    INSERT INTO roster_sync_runs (
      id, source_id, trigger_type, provider_version, content_hash, file_id, source_file_id, status,
      message, doctor_count, event_count, started_at, completed_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    String(run.id), String(run.sourceId), String(run.triggerType || "manual"),
    String(run.providerVersion || ""), String(run.contentHash || ""), String(run.fileId || ""), String(run.sourceFileId || run.fileId || ""),
    String(run.status || "started"), String(run.message || ""), Number(run.doctorCount || 0),
    Number(run.eventCount || 0), String(run.startedAt || new Date().toISOString()), String(run.completedAt || ""),
  ).run();
  return { ok: true, id: String(run.id) };
}

export async function finishRosterSyncRun(db, runId, update = {}) {
  if (!db?.prepare || !runId) return { ok: false, reason: "missing-run-id" };
  await ensureCalendarSchema(db);
  await db.prepare(`
    UPDATE roster_sync_runs
    SET status = ?, message = ?, file_id = ?, doctor_count = ?, event_count = ?, completed_at = ?
    WHERE id = ?
  `).bind(
    String(update.status || "failed"), String(update.message || ""), String(update.fileId || ""),
    Number(update.doctorCount || 0), Number(update.eventCount || 0),
    String(update.completedAt || new Date().toISOString()), String(runId),
  ).run();
  return { ok: true };
}

export async function supersedeDuplicateRosterSyncRuns(db, run = {}, fileName = "") {
  if (!db?.prepare || !run?.id || !run?.sourceId || !run?.providerVersion || !fileName) {
    return { ok: true, changes: 0 };
  }
  await ensureCalendarSchema(db);
  const completedAt = new Date().toISOString();
  const result = await db.prepare(`
    UPDATE roster_sync_runs
    SET status = 'superseded', message = ?, completed_at = ?
    WHERE source_id = ?
      AND provider_version = ?
      AND id <> ?
      AND status IN ('queued', 'processing')
      AND file_id IN (
        SELECT file_id FROM raw_roster_files WHERE LOWER(name) = LOWER(?)
      )
  `).bind(
    "Duplicate provider version skipped after an identical source version was imported.",
    completedAt,
    String(run.sourceId),
    String(run.providerVersion),
    String(run.id),
    String(fileName),
  ).run();
  return { ok: true, changes: Number(result?.meta?.changes || result?.changes || 0) };
}

function parseStoredJson(value, fallback = {}) {
  try {
    const parsed = JSON.parse(String(value || ""));
    return parsed && typeof parsed === "object" ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function rosterSourceFromRow(row) {
  if (!row?.id) return null;
  return {
    id: String(row.id), provider: String(row.provider || ""),
    sourceType: String(row.source_type || "").toLowerCase(), label: String(row.label || row.id),
    enabled: Number(row.enabled || 0) === 1, config: parseStoredJson(row.config_json),
    cursor: parseStoredJson(row.cursor_json), providerVersion: String(row.provider_version || ""),
    providerModifiedAt: String(row.provider_modified_at || ""), lastCheckedAt: String(row.last_checked_at || ""),
    lastSuccessAt: String(row.last_success_at || ""), lastError: String(row.last_error || ""),
    activeFileId: String(row.active_file_id || ""), createdAt: String(row.created_at || ""),
    updatedAt: String(row.updated_at || ""),
  };
}

function rosterSyncRunFromRow(row) {
  if (!row?.id) return null;
  return {
    id: String(row.id), sourceId: String(row.source_id || ""), triggerType: String(row.trigger_type || ""), status: String(row.status || ""),
    providerVersion: String(row.provider_version || ""), contentHash: String(row.content_hash || ""), fileId: String(row.file_id || ""), sourceFileId: String(row.source_file_id || row.file_id || ""),
    message: String(row.message || ""), doctorCount: Number(row.doctor_count || 0),
    eventCount: Number(row.event_count || 0), startedAt: String(row.started_at || ""),
    completedAt: String(row.completed_at || ""),
  };
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
      COALESCE(MAX(last_modified), 0) AS max_last_modified,
      COALESCE(GROUP_CONCAT(id || ':' || source_id || ':' || parsed_at, '|'), '') AS active_file_fingerprint
    FROM (
      SELECT id, source_id, parsed_at, uploaded_at, last_modified
      FROM roster_files
      WHERE active = 1
      ORDER BY id
    )
  `).first();
  const accountContext = email
    ? await db.prepare(`
      SELECT
        account_states.session_json AS session_json,
        account_profiles.local_parser_extensions_json AS local_parser_extensions_json
      FROM account_profiles
      LEFT JOIN account_states ON account_states.email = account_profiles.email
      WHERE account_profiles.email = ?
    `).bind(email).first()
    : null;
  const parserRules = await db.prepare(`
    SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at
    FROM parser_rules
    WHERE scope = 'global'
  `).first();
  const customEvents = email
    ? await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM custom_events WHERE owner_email = ?").bind(email).first()
    : null;
  const facilityDesignations = await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM facility_staff_designations").first();
  const facilitySeniorityOverrides = await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM facility_staff_seniority_overrides").first();
  const facilitySmsMemberships = await db.prepare("SELECT COUNT(*) AS count, COALESCE(MAX(updated_at), '') AS max_updated_at FROM facility_sms_memberships").first();
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
  const materializedSession = materializedSessionStateForRevision(parseJsonObject(accountContext?.session_json, {}));
  const localParserExtensions = parseJsonObject(accountContext?.local_parser_extensions_json, {});
  return [
    Number(roster?.active_file_count || 0),
    String(roster?.max_parsed_at || ""),
    String(roster?.max_uploaded_at || ""),
    Number(roster?.max_last_modified || 0),
    String(roster?.active_file_fingerprint || ""),
    stableJsonStringify(materializedSession),
    Number(parserRules?.count || 0),
    String(parserRules?.max_updated_at || ""),
    stableJsonStringify(localParserExtensions),
    Number(claims?.count || 0),
    String(claims?.max_updated_at || ""),
    Number(locations?.count || 0),
    String(locations?.max_updated_at || ""),
    Number(customEvents?.count || 0),
    String(customEvents?.max_updated_at || ""),
    Number(facilityDesignations?.count || 0),
    String(facilityDesignations?.max_updated_at || ""),
    Number(facilitySeniorityOverrides?.count || 0),
    String(facilitySeniorityOverrides?.max_updated_at || ""),
    Number(facilitySmsMemberships?.count || 0),
    String(facilitySmsMemberships?.max_updated_at || ""),
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
      email, real_name, role, insights_enabled, facility_overview_enabled, subscription_token, password_salt, password_hash,
      admin_issues_json, local_parser_extensions_json, created_at, updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(email) DO UPDATE SET
      real_name = excluded.real_name,
      role = excluded.role,
      insights_enabled = excluded.insights_enabled,
      facility_overview_enabled = excluded.facility_overview_enabled,
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
    record.facilityOverviewEnabled === true ? 1 : 0,
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
  // A claimed account is a reviewed identity boundary. Seed only its explicit
  // source/key claims; ambiguous names are never inferred here.
  await ensureAccountPersonAliases(db, email, record.realName, claims);
  for (const chunk of chunkRowsForBindLimit(claims.map((claim) => [
    email,
    claim.sourceType,
    claim.key,
    claim.displayName,
    claim.matchedAt,
    updatedAt,
  ]), 6, D1_MAX_BIND_PARAMS)) {
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

export async function ensureAccountPersonAliases(db, email, preferredDisplayName, claims = []) {
  if (!db?.prepare || !email) return null;
  let personId = `account:${normalizeEmail(email)}`;
  const now = new Date().toISOString();
  // Legacy records may already contain the same approved claim on more than
  // one account. Keep those records readable by joining their existing person;
  // new conflicting claims continue to be rejected by the claim endpoint.
  for (const claim of sanitizeAccountClaims(claims)) {
    const existing = await db.prepare("SELECT person_id FROM roster_person_aliases WHERE source_type = ? AND doctor_key = ?").bind(claim.sourceType, claim.key).first();
    if (existing?.person_id) {
      personId = String(existing.person_id);
      break;
    }
  }
  await db.prepare(`
    INSERT INTO roster_people (person_id, preferred_display_name, provenance, review_state, created_at, updated_at)
    VALUES (?, ?, 'automatic', 'approved', ?, ?)
    ON CONFLICT(person_id) DO UPDATE SET preferred_display_name = CASE WHEN excluded.preferred_display_name <> '' THEN excluded.preferred_display_name ELSE roster_people.preferred_display_name END, updated_at = excluded.updated_at
  `).bind(personId, String(preferredDisplayName || ""), now, now).run();
  await db.prepare(`
    INSERT INTO account_people (email, person_id, created_at, updated_at)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(email) DO UPDATE SET person_id = excluded.person_id, updated_at = excluded.updated_at
  `).bind(normalizeEmail(email), personId, now, now).run();
  for (const claim of sanitizeAccountClaims(claims)) {
    const existing = await db.prepare("SELECT person_id FROM roster_person_aliases WHERE source_type = ? AND doctor_key = ?").bind(claim.sourceType, claim.key).first();
    if (existing?.person_id && existing.person_id !== personId) continue;
    await db.prepare(`
      INSERT INTO roster_person_aliases (source_type, doctor_key, display_name, person_id, provenance, confidence, review_state, created_at, updated_at)
      VALUES (?, ?, ?, ?, 'automatic', 'high', 'approved', ?, ?)
      ON CONFLICT(source_type, doctor_key) DO UPDATE SET display_name = excluded.display_name, updated_at = excluded.updated_at
    `).bind(claim.sourceType, claim.key, claim.displayName, personId, now, now).run();
  }
  return personId;
}

export async function queryPersonAliasesForAccount(db, email) {
  if (!db?.prepare || !email) return [];
  await ensureCalendarSchema(db);
  const rows = await db.prepare(`
    SELECT aliases.source_type, aliases.doctor_key, aliases.display_name, aliases.provenance, aliases.confidence, aliases.review_state, aliases.updated_at
    FROM account_people accounts INNER JOIN roster_person_aliases aliases ON aliases.person_id = accounts.person_id
    WHERE accounts.email = ? ORDER BY aliases.source_type, aliases.display_name
  `).bind(normalizeEmail(email)).all();
  return (rows.results || []).map((row) => ({
    sourceType: String(row.source_type || ""), key: String(row.doctor_key || ""), displayName: String(row.display_name || ""),
    provenance: String(row.provenance || ""), confidence: String(row.confidence || ""), reviewState: String(row.review_state || ""), updatedAt: String(row.updated_at || ""),
  }));
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
      account_profiles.facility_overview_enabled AS facility_overview_enabled,
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
    facilityOverviewEnabled: first.facility_overview_enabled === 1,
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
      account_profiles.facility_overview_enabled AS facility_overview_enabled,
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
      account_profiles.facility_overview_enabled AS facility_overview_enabled,
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
    facilityOverviewEnabled: first.facility_overview_enabled === 1,
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
    return applyFacilityStaffSeniorityOverridesToCoworkerEvents(db, rowsToCoworkerEvents(rows));
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
        other_events.seniority,
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
      roster_events.seniority,
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
  return applyFacilityStaffSeniorityOverridesToCoworkerEvents(db, rowsToCoworkerEvents(rows));
}

function rowsToCoworkerEvents(rows) {
  return (rows.results || [])
    .map((row) => ({
      doctorKey: row.doctor_key,
      displayName: row.display_name,
      sourceType: row.source_type,
      seniority: row.seniority,
      event: parseEvent(row.event_json),
    }))
    .filter((row) => row.event);
}

async function applyFacilityStaffSeniorityOverridesToCoworkerEvents(db, rows) {
  const overridesByTerm = new Map();
  return Promise.all((rows || []).map(async (row) => {
    const sourceType = normalizeSourceType(row.sourceType);
    const date = String(row.event?.start || "").slice(0, 10);
    const termStart = australianTermStartForDate(date);
    const cacheKey = `${sourceType}|${termStart}`;
    if (!overridesByTerm.has(cacheKey)) {
      const overrides = await queryFacilityStaffSeniorityOverrides(db, { sourceType, termStart });
      overridesByTerm.set(cacheKey, new Map(overrides.map((override) => [`${override.sourceType}|${override.doctorKey}`, override])));
    }
    const override = overridesByTerm.get(cacheKey).get(`${sourceType}|${row.doctorKey}`);
    if (!override || override.useRosterSeniority) return row;
    return { ...row, seniority: override.seniority, event: { ...row.event, seniority: override.seniority, facilitySeniorityOverride: true } };
  }));
}

export async function queryFacilityOverviewOnShift(db, options = {}) {
  if (!db?.prepare) return [];
  await ensureCalendarSchema(db);
  const date = String(options.date || "").slice(0, 10);
  const facilityKey = normalizeSourceType(options.facilityKey || options.sourceType || "");
  if (!date || !facilityKey) return [];
  const rows = await db.prepare(`
    SELECT
      roster_events.doctor_key,
      roster_events.display_name,
      roster_events.source_type,
      roster_events.seniority,
      roster_events.event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND lower(roster_events.source_type) = ?
      -- At a glance is the roster for this day, rather than everyone whose
      -- shift happens to cross into it.  This excludes the preceding day's
      -- PM shifts that finish at midnight.
      AND roster_events.start_date = ?
    ORDER BY roster_events.start_ts, roster_events.display_name, roster_events.title
  `).bind(facilityKey, date).all();
  const seniorityOverrides = new Map((await queryFacilityStaffSeniorityOverrides(db, {
    sourceType: facilityKey,
    termStart: australianTermStartForDate(date),
  })).map((override) => [`${override.sourceType}|${override.doctorKey}`, override]));
  return (rows.results || [])
    .map((row) => ({
      doctorKey: String(row.doctor_key || "").trim(),
      displayName: String(row.display_name || "").trim(),
      sourceType: normalizeSourceType(row.source_type),
      seniority: String(row.seniority || "").trim(),
      event: parseEvent(row.event_json),
    }))
    .map((row) => {
      const override = seniorityOverrides.get(`${row.sourceType}|${row.doctorKey}`);
      if (!override || override.useRosterSeniority) return row;
      return { ...row, seniority: override.seniority, seniorityOverride: override, event: { ...row.event, seniority: override.seniority, facilitySeniorityOverride: true } };
    })
    .filter((row) => row.doctorKey && row.displayName && row.event);
}

// This intentionally returns the roster events rather than a second server-side
// stream classification. The browser already owns the parser-rule-aware Who/On
// shift classifier, so using the same records keeps the two views in agreement.
export async function queryFacilityOverviewRange(db, options = {}) {
  if (!db?.prepare) return { events: [], coverage: [] };
  await ensureCalendarSchema(db);
  const startDate = String(options.startDate || "").slice(0, 10);
  const endDate = String(options.endDate || "").slice(0, 10);
  const sourceTypes = sanitizeSourceTypes(options.sourceTypes || []);
  if (!startDate || !endDate || endDate < startDate || !sourceTypes.length) return { events: [], coverage: [] };
  const placeholders = sourceTypes.map(() => "?").join(", ");
  const rows = await db.prepare(`
    SELECT roster_events.doctor_key, roster_events.display_name, roster_events.source_type,
      roster_events.seniority, roster_events.start_date, roster_events.event_json
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.source_type IN (${placeholders})
      AND roster_events.start_date >= ? AND roster_events.start_date <= ?
    ORDER BY roster_events.start_date, roster_events.source_type, roster_events.start_ts, roster_events.display_name
  `).bind(...sourceTypes, startDate, endDate).all();
  const coverageRows = await db.prepare(`
    SELECT roster_events.source_type, roster_events.file_id,
      MIN(roster_events.start_date) AS start_date, MAX(roster_events.start_date) AS end_date
    FROM roster_events
    INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1
      AND roster_events.source_type IN (${placeholders})
    GROUP BY roster_events.source_type, roster_events.file_id
  `).bind(...sourceTypes).all();
  const overrideCache = new Map();
  const events = [];
  for (const row of rows.results || []) {
    const sourceType = normalizeSourceType(row.source_type);
    const date = String(row.start_date || "").slice(0, 10);
    const doctorKey = String(row.doctor_key || "").trim();
    const termStart = australianTermStartForDate(date);
    const overrideKey = `${sourceType}|${termStart}`;
    if (!overrideCache.has(overrideKey)) {
      const overrides = await queryFacilityStaffSeniorityOverrides(db, { sourceType, termStart });
      overrideCache.set(overrideKey, new Map(overrides.map((override) => [`${override.sourceType}|${override.doctorKey}`, override])));
    }
    const override = overrideCache.get(overrideKey).get(`${sourceType}|${doctorKey}`);
    const seniority = override && !override.useRosterSeniority ? override.seniority : String(row.seniority || "").trim();
    const event = parseEvent(row.event_json);
    if (!doctorKey || !event) continue;
    events.push({
      doctorKey,
      displayName: String(row.display_name || "").trim() || doctorKey,
      sourceType,
      seniority,
      date,
      event: override && !override.useRosterSeniority ? { ...event, seniority } : event,
    });
  }
  return {
    events,
    coverage: (coverageRows.results || []).map((row) => ({
      sourceType: normalizeSourceType(row.source_type),
      startDate: String(row.start_date || "").slice(0, 10),
      endDate: String(row.end_date || "").slice(0, 10),
    })).filter((row) => row.sourceType && row.startDate && row.endDate),
  };
}

export async function queryFacilityOverviewStaff(db, options = {}) {
  if (!db?.prepare) return { members: [], events: [], coverage: [], designations: [] };
  await ensureCalendarSchema(db);
  const termStart = String(options.termStart || "").slice(0, 10);
  const termEnd = String(options.termEnd || "").slice(0, 10);
  const facilityKey = normalizeSourceType(options.facilityKey || options.sourceType || "");
  if (!termStart || !termEnd) return { members: [], events: [], coverage: [], designations: [] };
  const facilitySql = facilityKey ? "AND roster_files.source_type = ?" : "";
  const bindings = facilityKey ? [facilityKey, termEnd, termStart] : [termEnd, termStart];
  const members = await db.prepare(`
    WITH file_coverage AS (
      SELECT file_id, MIN(start_date) AS coverage_start, MAX(start_date) AS coverage_end
      FROM roster_events GROUP BY file_id
    )
    SELECT roster_file_doctors.doctor_key, roster_file_doctors.display_name,
      roster_file_doctors.source_type, roster_file_doctors.seniority AS membership_seniority,
      roster_file_doctors.membership_source, file_coverage.coverage_start, file_coverage.coverage_end
    FROM roster_file_doctors
    INNER JOIN roster_files ON roster_files.id = roster_file_doctors.file_id
    INNER JOIN file_coverage ON file_coverage.file_id = roster_file_doctors.file_id
    WHERE roster_files.active = 1 ${facilitySql}
      AND file_coverage.coverage_start <= ? AND file_coverage.coverage_end >= ?
    ORDER BY roster_file_doctors.source_type, roster_file_doctors.display_name
  `).bind(...bindings).all();
  const continuingSms = await db.prepare(`
    SELECT doctor_key, display_name, source_type, last_seen_date AS coverage_end
    FROM facility_sms_memberships
    WHERE first_seen_date <= ? ${facilityKey ? "AND source_type = ?" : ""}
    ORDER BY source_type, display_name
  `).bind(...(facilityKey ? [termEnd, facilityKey] : [termEnd])).all();
  const eventBindings = facilityKey ? [facilityKey, termEnd, termStart] : [termEnd, termStart];
  const events = await db.prepare(`
    SELECT roster_events.doctor_key, roster_events.display_name, roster_events.source_type,
      roster_events.seniority, roster_events.event_json
    FROM roster_events INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1 ${facilityKey ? "AND roster_events.source_type = ?" : ""}
      AND roster_events.start_date <= ? AND roster_events.end_date >= ?
    ORDER BY roster_events.source_type, roster_events.display_name, roster_events.start_ts
  `).bind(...eventBindings).all();
  const coverage = await db.prepare(`
    SELECT roster_events.source_type, MIN(roster_events.start_date) AS start_date,
      MAX(roster_events.start_date) AS end_date
    FROM roster_events INNER JOIN roster_files ON roster_files.id = roster_events.file_id
    WHERE roster_files.active = 1 ${facilityKey ? "AND roster_events.source_type = ?" : ""}
    GROUP BY roster_events.source_type
  `).bind(...(facilityKey ? [facilityKey] : [])).all();
  const memberRows = (members.results || []).map((row) => ({
      doctorKey: String(row.doctor_key || "").trim(), displayName: String(row.display_name || "").trim(),
      sourceType: normalizeSourceType(row.source_type), seniority: String(row.membership_seniority || "").trim(),
      membershipSource: String(row.membership_source || "roster"), coverageStart: String(row.coverage_start || ""), coverageEnd: String(row.coverage_end || ""),
    })).filter((row) => row.doctorKey && row.displayName);
  const currentMembership = new Set(memberRows.map((row) => `${row.sourceType}|${row.doctorKey}`));
  for (const row of continuingSms.results || []) {
    const entry = {
      doctorKey: String(row.doctor_key || "").trim(), displayName: String(row.display_name || "").trim(),
      sourceType: normalizeSourceType(row.source_type), seniority: "SMS",
      membershipSource: "sms-continuity", coverageStart: "", coverageEnd: String(row.coverage_end || ""),
    };
    if (entry.doctorKey && entry.displayName && !currentMembership.has(`${entry.sourceType}|${entry.doctorKey}`)) memberRows.push(entry);
  }
  const designations = await queryFacilityStaffDesignations(db, { sourceType: facilityKey, termStart, termEnd });
  const seniorityOverrides = await queryFacilityStaffSeniorityOverrides(db, { sourceType: facilityKey, termStart });
  return {
    members: memberRows,
    events: (events.results || []).map((row) => ({
      doctorKey: String(row.doctor_key || "").trim(), displayName: String(row.display_name || "").trim(),
      sourceType: normalizeSourceType(row.source_type), seniority: String(row.seniority || "").trim(), event: parseEvent(row.event_json),
    })).filter((row) => row.doctorKey && row.event),
    coverage: (coverage.results || []).map((row) => ({ sourceType: normalizeSourceType(row.source_type), startDate: String(row.start_date || ""), endDate: String(row.end_date || "") })),
    designations,
    seniorityOverrides,
  };
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
  const lower = `${String(fileId)}:`;
  const upper = `${String(fileId)};`;
  await db.prepare(`
    DELETE FROM roster_daily_presence
    WHERE event_id >= ? AND event_id < ?
  `).bind(lower, upper).run();
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
  await insertDailyPresenceRows(db, rows);
}

function dailyPresenceInsertStatements(db, rows) {
  return chunkRowsForBindLimit(rows, 5, D1_MAX_BIND_PARAMS)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT OR IGNORE INTO roster_daily_presence (date, source_type, doctor_key, display_name, event_id)
      VALUES ${chunk.map(() => "(?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()));
}

async function insertDailyPresenceRows(db, rows) {
  await runStatementBatches(db, dailyPresenceInsertStatements(db, rows), D1_PRESENCE_BATCH_STATEMENTS);
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
  await insertDailyPresenceRows(db, presenceRows);
}

export async function rebuildDailyPresenceForActiveFiles(db, options = {}) {
  if (!db?.prepare) return { files: 0, done: true, nextOffset: null };
  await ensureCalendarSchema(db);
  const offset = Math.max(0, Number.parseInt(options.offset ?? 0, 10) || 0);
  const limit = Math.max(1, Math.min(Number.parseInt(options.limit ?? 10, 10) || 10, 25));
  const rows = await db.prepare(`
    SELECT id
    FROM roster_files
    WHERE active = 1
    ORDER BY added_at, name, id
    LIMIT ? OFFSET ?
  `).bind(limit + 1, offset).all();
  const fileIds = (rows.results || [])
    .map((row) => String(row.id || "").trim())
    .filter(Boolean);
  const batchFileIds = fileIds.slice(0, limit);
  for (const activeFileId of batchFileIds) {
    await rebuildDailyPresenceForFile(db, activeFileId);
  }
  const done = fileIds.length <= limit;
  return {
    files: batchFileIds.length,
    offset,
    limit,
    nextOffset: done ? null : offset + batchFileIds.length,
    done,
  };
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
      seniority: String(doctor?.seniority || "").trim(),
      membershipSource: String(doctor?.membershipSource || doctor?.membership_source || "roster").trim() || "roster",
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

// Keep this calendar in step with the browser and state API: Victorian medical
// terms start on the first Monday of February, May, August and November.
// These are JavaScript month indexes (zero based), not month numbers.
export function australianTermStartForDate(value) {
  const date = new Date(`${datePart(value)}T12:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  const year = date.getUTCFullYear();
  const candidates = [[year, 1], [year, 4], [year, 7], [year, 10], [year - 1, 10]];
  for (const [candidateYear, month] of candidates) {
    const start = new Date(Date.UTC(candidateYear, month, 1, 12, 0, 0));
    const day = start.getUTCDay();
    start.setUTCDate(start.getUTCDate() + (day === 0 ? 1 : day === 1 ? 0 : 8 - day));
    const end = new Date(start);
    end.setUTCDate(end.getUTCDate() + 91);
    if (date >= start && date < end) return start.toISOString().slice(0, 10);
  }
  return "";
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
  if (/\b(?:sick|s\/l)\b/i.test(combined)) return "Sick leave";
  if (/\bpersonal\b/i.test(combined)) return "Personal Leave";
  if (/\bstudy\b/i.test(combined)) return "Study Leave";
  if (/\bexam\b/i.test(combined)) return "Exam Leave";
  if (/\b(?:sabbatical|sab\/l)\b/i.test(combined)) return "Sabbatical";
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
  return chunkRowsForBindLimit(doctors.map((doctor) => [sourceType, doctor.key, doctor.displayName, updatedAt]), 4, D1_MAX_BIND_PARAMS)
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
  return chunkRowsForBindLimit(doctors.map((doctor) => [fileId, sourceType, doctor.key, doctor.displayName, doctor.seniority || "", doctor.membershipSource || "roster"]), 6, D1_MAX_BIND_PARAMS)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_file_doctors (file_id, source_type, doctor_key, display_name, seniority, membership_source)
      VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?)").join(", ")}
      ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET
        display_name = excluded.display_name,
        seniority = excluded.seniority,
        membership_source = excluded.membership_source
    `).bind(...chunk.flat()));
}

async function bulkInsertEvents(db, rows) {
  for (const statement of bulkInsertEventStatements(db, rows)) {
    await statement.run();
  }
}

function bulkInsertEventStatements(db, rows) {
  return chunkRowsForBindLimit(rows, 16, D1_MAX_BIND_PARAMS)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_events (
        id, file_id, source_type, doctor_key, display_name, start_date, end_date, start_ts, end_ts,
        title, raw_value, seniority, location, all_day, time_label, event_json
      ) VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()));
}

function bulkInsertIssueStatements(db, rows) {
  return chunkRowsForBindLimit(rows, 14, D1_MAX_BIND_PARAMS)
    .filter((chunk) => chunk.length)
    .map((chunk) => db.prepare(`
      INSERT INTO roster_issues (
        id, file_id, source_type, doctor_key, display_name, start_date, raw_value, seniority,
        status, message, resolution_type, suggested_title, time_label, issue_json
      ) VALUES ${chunk.map(() => "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}
    `).bind(...chunk.flat()));
}

async function runStatementBatches(db, statements, batchSize = D1_MAX_BATCH_STATEMENTS) {
  if (!statements.length) return [];
  const results = [];
  for (const chunk of chunkRows(statements, batchSize)) {
    if (!chunk.length) continue;
    if (typeof db.batch === "function") results.push(...(await db.batch(chunk)));
    else for (const statement of chunk) results.push(await statement.run());
  }
  return results;
}

async function runTransactionalBatch(db, statements) {
  return runStatementBatches(db, statements, D1_MAX_BATCH_STATEMENTS);
}

function chunkRowsForBindLimit(rows, columnsPerRow, maxParams = D1_MAX_BIND_PARAMS) {
  const rowChunkSize = Math.max(1, Math.floor(maxParams / Math.max(1, columnsPerRow)));
  return chunkRows(rows, rowChunkSize);
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

export async function listSnapshotRegistryWarmupCandidates(db, options = {}) {
  if (!db?.prepare) return [];
  const ownerTypes = [...new Set((options.ownerTypes || []).map((item) => String(item || "").trim()).filter(Boolean))];
  const statuses = [...new Set((options.statuses || ["ready"]).map((item) => String(item || "").trim()).filter(Boolean))];
  const rangeKey = String(options.rangeKey || "").trim();
  const limit = Math.max(1, Math.min(Number.parseInt(options.limit ?? 25, 10) || 25, 100));
  if (!ownerTypes.length || !statuses.length) return [];
  await ensureCalendarSchema(db);
  const ownerSql = ownerTypes.map(() => "?").join(", ");
  const statusSql = statuses.map(() => "?").join(", ");
  const rangeSql = rangeKey ? "AND range_key = ?" : "";
  const binds = [...ownerTypes, ...statuses];
  if (rangeKey) binds.push(rangeKey);
  binds.push(limit);
  const rows = await db.prepare(`
    SELECT owner_type, owner_id, doctor_key, range_key, requested_revision, built_revision, status, artifact_key,
           built_at, size_bytes, build_ms, last_error, updated_at
    FROM snapshot_registry
    WHERE owner_type IN (${ownerSql})
      AND status IN (${statusSql})
      ${rangeSql}
    ORDER BY updated_at DESC
    LIMIT ?
  `).bind(...binds).all();
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
