-- Durable fencing record for shared ED/day publication. R2 objects remain
-- immutable; this row coordinates the small manifest pointer update.

CREATE TABLE IF NOT EXISTS facility_day_publications (
  source_type TEXT PRIMARY KEY,
  generation INTEGER NOT NULL DEFAULT 0,
  operation_id TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'idle',
  base_revision TEXT NOT NULL DEFAULT '',
  candidate_revision TEXT NOT NULL DEFAULT '',
  lease_expires_at TEXT NOT NULL DEFAULT '',
  last_error TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
