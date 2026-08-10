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
);

CREATE INDEX IF NOT EXISTS idx_facility_staff_designations_active
  ON facility_staff_designations (source_type, doctor_key, active, term_start);
