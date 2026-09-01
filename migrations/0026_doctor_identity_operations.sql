-- Canonical doctor identities are durable.  Roster source names and events are
-- deliberately not referenced by these mutations and remain immutable.
ALTER TABLE roster_people ADD COLUMN status TEXT NOT NULL DEFAULT 'active';
ALTER TABLE roster_people ADD COLUMN merged_into_person_id TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_people ADD COLUMN version INTEGER NOT NULL DEFAULT 1;

UPDATE roster_people
SET status = 'active', merged_into_person_id = '', version = CASE WHEN version < 1 THEN 1 ELSE version END
WHERE status = '' OR status IS NULL OR version < 1;

CREATE TABLE IF NOT EXISTS roster_person_redirects (
  old_person_id TEXT PRIMARY KEY,
  current_person_id TEXT NOT NULL,
  operation_id TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  created_by TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_person_redirects_current
ON roster_person_redirects (current_person_id, active);

CREATE TABLE IF NOT EXISTS roster_identity_operations (
  operation_id TEXT PRIMARY KEY,
  operation_type TEXT NOT NULL,
  target_person_id TEXT NOT NULL DEFAULT '',
  affected_person_ids_json TEXT NOT NULL DEFAULT '[]',
  status TEXT NOT NULL DEFAULT 'committed',
  administrator_email TEXT NOT NULL DEFAULT '',
  reason TEXT NOT NULL DEFAULT '',
  expected_versions_json TEXT NOT NULL DEFAULT '{}',
  before_summary_json TEXT NOT NULL DEFAULT '{}',
  after_summary_json TEXT NOT NULL DEFAULT '{}',
  reversed_operation_id TEXT NOT NULL DEFAULT '',
  reversed_by_operation_id TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_identity_operations_target
ON roster_identity_operations (target_person_id, created_at DESC);

CREATE TABLE IF NOT EXISTS roster_identity_operation_items (
  item_id TEXT PRIMARY KEY,
  operation_id TEXT NOT NULL,
  entity_type TEXT NOT NULL,
  entity_key TEXT NOT NULL,
  before_json TEXT NOT NULL DEFAULT 'null',
  after_json TEXT NOT NULL DEFAULT 'null',
  created_at TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_identity_operation_items_operation
ON roster_identity_operation_items (operation_id, entity_type);

CREATE TABLE IF NOT EXISTS roster_identity_candidates (
  candidate_id TEXT PRIMARY KEY,
  candidate_fingerprint TEXT NOT NULL UNIQUE,
  left_person_id TEXT NOT NULL DEFAULT '',
  right_person_id TEXT NOT NULL DEFAULT '',
  left_alias_json TEXT NOT NULL DEFAULT '{}',
  right_alias_json TEXT NOT NULL DEFAULT '{}',
  score REAL NOT NULL DEFAULT 0,
  reasons_json TEXT NOT NULL DEFAULT '[]',
  warnings_json TEXT NOT NULL DEFAULT '[]',
  status TEXT NOT NULL DEFAULT 'pending',
  evidence_fingerprint TEXT NOT NULL DEFAULT '',
  first_seen_at TEXT NOT NULL DEFAULT '',
  last_seen_at TEXT NOT NULL DEFAULT '',
  audit_run_id TEXT NOT NULL DEFAULT '',
  reviewed_by TEXT NOT NULL DEFAULT '',
  reviewed_at TEXT NOT NULL DEFAULT '',
  rejection_reason TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_identity_candidates_status
ON roster_identity_candidates (status, score DESC, last_seen_at DESC);

CREATE TABLE IF NOT EXISTS roster_identity_features (
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  person_id TEXT NOT NULL DEFAULT '',
  compact_name TEXT NOT NULL DEFAULT '',
  surname_key TEXT NOT NULL DEFAULT '',
  surname_prefix TEXT NOT NULL DEFAULT '',
  given_key TEXT NOT NULL DEFAULT '',
  given_initial TEXT NOT NULL DEFAULT '',
  phonetic_surname TEXT NOT NULL DEFAULT '',
  token_count INTEGER NOT NULL DEFAULT 0,
  alias_fingerprint TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, doctor_key)
);
CREATE INDEX IF NOT EXISTS idx_roster_identity_features_block
ON roster_identity_features (surname_key, given_key);
CREATE INDEX IF NOT EXISTS idx_roster_identity_features_prefix
ON roster_identity_features (surname_prefix, given_initial);

CREATE TABLE IF NOT EXISTS roster_identity_audit_runs (
  audit_run_id TEXT PRIMARY KEY,
  trigger_type TEXT NOT NULL,
  scope_json TEXT NOT NULL DEFAULT '{}',
  status TEXT NOT NULL DEFAULT 'queued',
  cursor_value TEXT NOT NULL DEFAULT '',
  rows_examined INTEGER NOT NULL DEFAULT 0,
  comparisons_made INTEGER NOT NULL DEFAULT 0,
  suggestions_changed INTEGER NOT NULL DEFAULT 0,
  deferral_reason TEXT NOT NULL DEFAULT '',
  error_text TEXT NOT NULL DEFAULT '',
  started_at TEXT NOT NULL DEFAULT '',
  completed_at TEXT NOT NULL DEFAULT '',
  created_by TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_roster_identity_audit_runs_status
ON roster_identity_audit_runs (status, updated_at);

CREATE TABLE IF NOT EXISTS roster_identity_audit_state (
  state_key TEXT PRIMARY KEY,
  cursor_value TEXT NOT NULL DEFAULT '',
  lease_owner TEXT NOT NULL DEFAULT '',
  lease_expires_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);
