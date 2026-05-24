CREATE TABLE IF NOT EXISTS roster_daily_presence (
  date TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  event_id TEXT NOT NULL,
  PRIMARY KEY (date, source_type, doctor_key, event_id)
);

CREATE INDEX IF NOT EXISTS idx_roster_daily_presence_doctor_date
ON roster_daily_presence (doctor_key, date);

CREATE INDEX IF NOT EXISTS idx_roster_daily_presence_date_source
ON roster_daily_presence (date, source_type);

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
);

CREATE INDEX IF NOT EXISTS idx_snapshot_registry_owner
ON snapshot_registry (owner_type, owner_id, updated_at DESC);
