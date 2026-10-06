-- Identity-only persistence. No roster/event backfill and no runtime DDL.
ALTER TABLE roster_people ADD COLUMN status TEXT NOT NULL DEFAULT 'active';
ALTER TABLE roster_people ADD COLUMN merged_into_person_id TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_people ADD COLUMN version INTEGER NOT NULL DEFAULT 1;
CREATE INDEX idx_identity_person_name ON roster_people(preferred_display_name COLLATE NOCASE,person_id);
CREATE TABLE roster_identity_revision (id INTEGER PRIMARY KEY CHECK(id=1), revision INTEGER NOT NULL);
INSERT INTO roster_identity_revision VALUES(1,0);
CREATE TABLE roster_person_redirects (
 old_person_id TEXT PRIMARY KEY, person_id TEXT NOT NULL,
 operation_id TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1,
 created_at TEXT NOT NULL, FOREIGN KEY(person_id) REFERENCES roster_people(person_id)
);
CREATE INDEX idx_identity_redirect_target ON roster_person_redirects(person_id,active);
CREATE TABLE roster_identity_operations (
 operation_id TEXT PRIMARY KEY, kind TEXT NOT NULL, status TEXT NOT NULL,
 actor TEXT NOT NULL, reason TEXT NOT NULL, created_at TEXT NOT NULL,
 reversed_operation_id TEXT NOT NULL DEFAULT '', reversed_by TEXT NOT NULL DEFAULT '',
 before_json TEXT NOT NULL, after_json TEXT NOT NULL, affected_json TEXT NOT NULL
);
CREATE INDEX idx_identity_operation_created ON roster_identity_operations(created_at,operation_id);
CREATE INDEX idx_identity_operation_actor ON roster_identity_operations(actor,created_at);
CREATE TABLE roster_identity_operation_people (
 person_id TEXT NOT NULL, operation_id TEXT NOT NULL,
 created_at TEXT NOT NULL, PRIMARY KEY(person_id,operation_id)
);
CREATE INDEX idx_identity_person_history ON roster_identity_operation_people(person_id,created_at,operation_id);
CREATE TABLE roster_identity_jobs (
 operation_id TEXT PRIMARY KEY, affected_json TEXT NOT NULL,
 status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
 last_error TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL
);
CREATE TABLE roster_identity_candidates (
 pair_key TEXT PRIMARY KEY, left_json TEXT NOT NULL, right_json TEXT NOT NULL,
 evidence_json TEXT NOT NULL, fingerprint TEXT NOT NULL, status TEXT NOT NULL,
 reviewer TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL
);
CREATE INDEX idx_identity_candidate_status ON roster_identity_candidates(status,pair_key);
CREATE TABLE roster_identity_features (
 source_type TEXT NOT NULL, doctor_key TEXT NOT NULL, display_name TEXT NOT NULL,
 normalized_key TEXT NOT NULL, given_block TEXT NOT NULL, surname_block TEXT NOT NULL,
 fingerprint TEXT NOT NULL, audited_fingerprint TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL, PRIMARY KEY(source_type,doctor_key)
);
CREATE INDEX idx_identity_feature_exact ON roster_identity_features(normalized_key,source_type,doctor_key);
CREATE INDEX idx_identity_feature_given ON roster_identity_features(given_block,source_type,doctor_key);
CREATE INDEX idx_identity_feature_surname ON roster_identity_features(surname_block,source_type,doctor_key);
CREATE TABLE roster_identity_audit_runs (
 run_id TEXT PRIMARY KEY, cursor TEXT NOT NULL DEFAULT '', status TEXT NOT NULL,
 mode TEXT NOT NULL DEFAULT 'audit', scope_json TEXT NOT NULL DEFAULT '[]', week_key TEXT NOT NULL DEFAULT '', lease_token TEXT NOT NULL DEFAULT '', lease_until TEXT NOT NULL DEFAULT '',
 examined INTEGER NOT NULL DEFAULT 0, candidates INTEGER NOT NULL DEFAULT 0,
 created_at TEXT NOT NULL, updated_at TEXT NOT NULL, last_error TEXT NOT NULL DEFAULT ''
);
CREATE INDEX idx_identity_audit_status ON roster_identity_audit_runs(status,created_at);
CREATE INDEX idx_identity_audit_week ON roster_identity_audit_runs(week_key,status,mode);

CREATE TABLE roster_identity_registrations (
 source_type TEXT NOT NULL, doctor_key TEXT NOT NULL, person_id TEXT NOT NULL,
 actor TEXT NOT NULL, reason TEXT NOT NULL, created_at TEXT NOT NULL,
 PRIMARY KEY(source_type,doctor_key)
);

CREATE INDEX idx_identity_audit_scope ON roster_identity_audit_runs(mode,scope_json,status,created_at);

CREATE INDEX idx_identity_feature_pending ON roster_identity_features(source_type,doctor_key) WHERE audited_fingerprint<>fingerprint;
