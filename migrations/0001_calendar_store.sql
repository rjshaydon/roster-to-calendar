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
);

CREATE TABLE IF NOT EXISTS roster_doctors (
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (source_type, doctor_key)
);

CREATE TABLE IF NOT EXISTS roster_file_doctors (
  file_id TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  PRIMARY KEY (file_id, source_type, doctor_key)
);

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
);

CREATE INDEX IF NOT EXISTS idx_roster_events_doctor_range ON roster_events (doctor_key, start_date, end_date);
CREATE INDEX IF NOT EXISTS idx_roster_events_date_source ON roster_events (start_date, source_type);
CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events (source_type, start_date, end_date);
CREATE INDEX IF NOT EXISTS idx_roster_events_file ON roster_events (file_id);

CREATE TABLE IF NOT EXISTS raw_roster_files (
  file_id TEXT PRIMARY KEY,
  object_key TEXT NOT NULL DEFAULT '',
  type TEXT NOT NULL DEFAULT '',
  data_url TEXT NOT NULL DEFAULT '',
  uploaded_at TEXT NOT NULL DEFAULT ''
);
