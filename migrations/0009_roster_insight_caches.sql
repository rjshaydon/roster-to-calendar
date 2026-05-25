CREATE TABLE IF NOT EXISTS roster_day_coworkers (
  term_key TEXT NOT NULL,
  date TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  event_id TEXT NOT NULL,
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (term_key, date, source_type, doctor_key, event_id)
);

CREATE INDEX IF NOT EXISTS idx_roster_day_coworkers_doctor_date
ON roster_day_coworkers (doctor_key, date);

CREATE INDEX IF NOT EXISTS idx_roster_day_coworkers_date_source
ON roster_day_coworkers (date, source_type);

CREATE TABLE IF NOT EXISTS roster_term_overlap_doctors (
  term_key TEXT NOT NULL,
  date TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  overlap_doctor_key TEXT NOT NULL,
  overlap_display_name TEXT NOT NULL,
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (term_key, date, source_type, doctor_key, overlap_doctor_key)
);

CREATE INDEX IF NOT EXISTS idx_roster_term_overlap_doctor_date
ON roster_term_overlap_doctors (doctor_key, date);

CREATE INDEX IF NOT EXISTS idx_roster_term_overlap_date_source
ON roster_term_overlap_doctors (date, source_type);
