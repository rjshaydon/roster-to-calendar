ALTER TABLE roster_files ADD COLUMN source_id TEXT NOT NULL DEFAULT '';

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
);

CREATE TABLE IF NOT EXISTS roster_sync_runs (
  id TEXT PRIMARY KEY,
  source_id TEXT NOT NULL,
  trigger_type TEXT NOT NULL DEFAULT '',
  provider_version TEXT NOT NULL DEFAULT '',
  content_hash TEXT NOT NULL DEFAULT '',
  file_id TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'started',
  message TEXT NOT NULL DEFAULT '',
  doctor_count INTEGER NOT NULL DEFAULT 0,
  event_count INTEGER NOT NULL DEFAULT 0,
  started_at TEXT NOT NULL DEFAULT '',
  completed_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_started
ON roster_sync_runs (source_id, started_at DESC);

CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_hash
ON roster_sync_runs (source_id, content_hash, status);
