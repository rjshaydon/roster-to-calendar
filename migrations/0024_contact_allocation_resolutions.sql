CREATE TABLE IF NOT EXISTS contact_allocation_resolutions (
  id TEXT PRIMARY KEY,
  source_id TEXT NOT NULL,
  source_date TEXT NOT NULL,
  contact_key TEXT NOT NULL,
  source_type TEXT NOT NULL DEFAULT '',
  doctor_key TEXT NOT NULL DEFAULT '',
  display_name TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  revision INTEGER NOT NULL DEFAULT 0,
  created_by TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_by TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  cleared_by TEXT NOT NULL DEFAULT '',
  cleared_at TEXT NOT NULL DEFAULT '',
  UNIQUE(source_id, source_date, contact_key)
);

CREATE INDEX IF NOT EXISTS idx_contact_allocation_resolutions_lookup
  ON contact_allocation_resolutions (source_id, source_date, active);

CREATE TABLE IF NOT EXISTS contact_allocation_resolution_history (
  id TEXT PRIMARY KEY,
  resolution_id TEXT NOT NULL,
  revision INTEGER NOT NULL,
  action TEXT NOT NULL,
  doctor_key TEXT NOT NULL DEFAULT '',
  display_name TEXT NOT NULL DEFAULT '',
  actor_email TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_contact_allocation_resolution_history_resolution
  ON contact_allocation_resolution_history (resolution_id, revision);
