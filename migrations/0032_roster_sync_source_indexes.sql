-- Bound automatic ingestion idempotency and queue probes to one exact source.
-- These indexes contain only compact sync-run metadata; no roster-event
-- backfill or application data rewrite is performed.

CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_version_status_started
ON roster_sync_runs (source_id, provider_version, status, started_at DESC);

CREATE INDEX IF NOT EXISTS idx_roster_sync_runs_source_status_started
ON roster_sync_runs (source_id, status, started_at, id);

CREATE INDEX IF NOT EXISTS idx_account_claims_source_doctor_email
ON account_claims (source_type, doctor_key, email);
