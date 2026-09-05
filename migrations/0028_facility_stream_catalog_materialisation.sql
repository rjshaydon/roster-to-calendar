-- Per-file stream signatures used to publish compact At a glance metadata
-- without grouping roster_events on a reader request.

CREATE TABLE IF NOT EXISTS facility_stream_catalog_contributions (
  source_type TEXT NOT NULL,
  term_start TEXT NOT NULL,
  file_id TEXT NOT NULL,
  catalog_key TEXT NOT NULL,
  seniority TEXT NOT NULL DEFAULT '',
  title TEXT NOT NULL DEFAULT '',
  raw_value TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT '',
  all_day INTEGER NOT NULL DEFAULT 0,
  time_label TEXT NOT NULL DEFAULT '',
  start_time TEXT NOT NULL DEFAULT '',
  end_time TEXT NOT NULL DEFAULT '',
  first_date TEXT NOT NULL,
  last_date TEXT NOT NULL,
  fact_digest TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (source_type, term_start, file_id, catalog_key)
);

CREATE INDEX IF NOT EXISTS idx_facility_stream_catalog_lookup
ON facility_stream_catalog_contributions (source_type, term_start, catalog_key);
