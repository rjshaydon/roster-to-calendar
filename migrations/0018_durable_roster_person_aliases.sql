-- Forward-only durable identity registry. Historical roster_events retain the
-- original doctor_key; aliases are expanded at calendar-query time.
CREATE TABLE IF NOT EXISTS roster_people (
  person_id TEXT PRIMARY KEY,
  preferred_display_name TEXT NOT NULL DEFAULT '',
  provenance TEXT NOT NULL DEFAULT 'automatic',
  review_state TEXT NOT NULL DEFAULT 'approved',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  approved_by TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS roster_person_aliases (
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL DEFAULT '',
  person_id TEXT NOT NULL,
  provenance TEXT NOT NULL DEFAULT 'automatic',
  confidence TEXT NOT NULL DEFAULT 'high',
  review_state TEXT NOT NULL DEFAULT 'approved',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  approved_by TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, doctor_key),
  FOREIGN KEY (person_id) REFERENCES roster_people(person_id)
);
CREATE INDEX IF NOT EXISTS idx_roster_person_aliases_person ON roster_person_aliases (person_id);

CREATE TABLE IF NOT EXISTS account_people (
  email TEXT PRIMARY KEY,
  person_id TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  FOREIGN KEY (person_id) REFERENCES roster_people(person_id)
);
CREATE INDEX IF NOT EXISTS idx_account_people_person ON account_people (person_id);
