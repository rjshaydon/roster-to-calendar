-- Compact, incrementally maintained Creator roster-storage status.
-- No backfill is performed: existing files remain unknown until a controlled
-- one-file bootstrap or a subsequent import writes their summary.

CREATE TABLE IF NOT EXISTS roster_file_status_summaries (
  file_id TEXT PRIMARY KEY,
  source_type TEXT NOT NULL DEFAULT '',
  source_id TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 0 CHECK (active IN (0, 1)),
  derived_state TEXT NOT NULL DEFAULT 'retained'
    CHECK (derived_state IN ('retained', 'building', 'ready', 'error', 'removed')),
  expected_doctor_count INTEGER NOT NULL DEFAULT 0 CHECK (expected_doctor_count >= 0),
  indexed_doctor_count INTEGER NOT NULL DEFAULT 0 CHECK (indexed_doctor_count >= 0),
  event_count INTEGER NOT NULL DEFAULT 0 CHECK (event_count >= 0),
  raw_source_available INTEGER NOT NULL DEFAULT 0 CHECK (raw_source_available IN (0, 1)),
  size INTEGER NOT NULL DEFAULT 0 CHECK (size >= 0),
  last_modified INTEGER NOT NULL DEFAULT 0 CHECK (last_modified >= 0),
  uploaded_at TEXT NOT NULL DEFAULT '',
  content_revision TEXT NOT NULL DEFAULT '',
  status_revision TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_roster_file_status_active_source_updated
ON roster_file_status_summaries (active, source_type, updated_at, file_id);

-- Status refreshes must not sort growing automation history. These indexes
-- support one latest-run probe per bounded source and one latest dispatch.
CREATE INDEX IF NOT EXISTS idx_roster_sources_label_id
ON roster_sources (label, id);

CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_started_id
ON roster_sync_runs (source_id, started_at DESC, id DESC);

CREATE INDEX IF NOT EXISTS idx_roster_dispatches_requested_id
ON roster_dispatches (requested_at DESC, id DESC);
