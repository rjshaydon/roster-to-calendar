-- Every observed roster name participates in duplicate discovery, whether or
-- not a clinician has claimed it. Source rows remain immutable evidence.
CREATE TABLE IF NOT EXISTS roster_source_identities (
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL DEFAULT '',
  first_seen_date TEXT NOT NULL DEFAULT '',
  last_seen_date TEXT NOT NULL DEFAULT '',
  event_count INTEGER NOT NULL DEFAULT 0,
  source_watermark TEXT NOT NULL DEFAULT '',
  person_id TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  feature_version INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, doctor_key)
);
CREATE INDEX IF NOT EXISTS idx_roster_source_identities_person ON roster_source_identities (person_id, active);
CREATE INDEX IF NOT EXISTS idx_roster_source_identities_updated ON roster_source_identities (updated_at, source_type, doctor_key);
CREATE INDEX IF NOT EXISTS idx_roster_source_identities_audit ON roster_source_identities (source_type, active, updated_at, doctor_key);

INSERT OR IGNORE INTO roster_source_identities (
  source_type, doctor_key, display_name, first_seen_date, last_seen_date,
  event_count, source_watermark, person_id, active, updated_at
)
SELECT d.source_type, d.doctor_key, d.display_name,
  COALESCE((SELECT MIN(e.start_date) FROM roster_events e WHERE e.source_type = d.source_type AND e.doctor_key = d.doctor_key), ''),
  COALESCE((SELECT MAX(e.start_date) FROM roster_events e WHERE e.source_type = d.source_type AND e.doctor_key = d.doctor_key), ''),
  COALESCE((SELECT COUNT(*) FROM roster_events e WHERE e.source_type = d.source_type AND e.doctor_key = d.doctor_key), 0),
  d.updated_at,
  COALESCE((SELECT a.person_id FROM roster_person_aliases a WHERE a.source_type = d.source_type AND a.doctor_key = d.doctor_key), ''),
  1, d.updated_at
FROM roster_doctors d;

-- Readable references are kept independently of the immutable relationship
-- key. Existing human-readable person IDs become current references.
CREATE TABLE IF NOT EXISTS roster_person_references (
  person_reference TEXT PRIMARY KEY,
  person_id TEXT NOT NULL,
  state TEXT NOT NULL DEFAULT 'current',
  operation_id TEXT NOT NULL DEFAULT '',
  created_by TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_person_references_person ON roster_person_references (person_id, state);

INSERT OR IGNORE INTO roster_person_references (person_reference, person_id, state, created_at, updated_at)
SELECT REPLACE(person_id, 'person:', ''), person_id, 'current', created_at, updated_at
FROM roster_people
WHERE person_id LIKE 'person:%';
