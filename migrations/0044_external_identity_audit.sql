-- Inactive external-audit checkpoints only. No historical backfill.
CREATE TABLE external_identity_audit_jobs (
 run_id TEXT PRIMARY KEY, plan_revision TEXT NOT NULL, identity_revision INTEGER NOT NULL,
 manifest_json TEXT NOT NULL, next_candidate_batch INTEGER NOT NULL DEFAULT 0,
 next_feature_batch INTEGER NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'running',
 created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE external_identity_audit_receipts (
 run_id TEXT NOT NULL, kind TEXT NOT NULL, batch_index INTEGER NOT NULL,
 payload_hash TEXT NOT NULL, PRIMARY KEY(run_id,kind,batch_index)
);
CREATE TABLE external_identity_audit_features (
 run_id TEXT NOT NULL, source_type TEXT NOT NULL, doctor_key TEXT NOT NULL,
 fingerprint TEXT NOT NULL, PRIMARY KEY(run_id,source_type,doctor_key)
);
