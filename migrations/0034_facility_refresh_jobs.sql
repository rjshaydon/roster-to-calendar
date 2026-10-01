CREATE TABLE IF NOT EXISTS facility_refresh_jobs (
  source_type TEXT NOT NULL,
  term_start TEXT NOT NULL,
  dates_json TEXT NOT NULL DEFAULT '[]',
  content_signature TEXT NOT NULL,
  request_revision TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending',
  plan_json TEXT NOT NULL DEFAULT '',
  next_batch INTEGER NOT NULL DEFAULT 0,
  next_month INTEGER NOT NULL DEFAULT 0,
  last_error TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL,
  PRIMARY KEY (source_type, term_start)
);
CREATE INDEX IF NOT EXISTS idx_facility_refresh_source_status_term ON facility_refresh_jobs (source_type, status, term_start);
