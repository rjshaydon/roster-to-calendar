-- Inactive staged imports only. No backfill or live feature enablement.
CREATE TABLE IF NOT EXISTS roster_import_jobs (
  run_id TEXT PRIMARY KEY,
  file_id TEXT NOT NULL,
  source_id TEXT NOT NULL,
  plan_revision TEXT NOT NULL,
  manifest_json TEXT NOT NULL,
  initialized INTEGER NOT NULL DEFAULT 0,
  initialize_started_at TEXT NOT NULL DEFAULT '',
  prepared_batch INTEGER NOT NULL DEFAULT 0,
  compact_ready INTEGER NOT NULL DEFAULT 0,
  activated INTEGER NOT NULL DEFAULT 0,
  next_batch INTEGER NOT NULL DEFAULT 0,
  event_count INTEGER NOT NULL DEFAULT 0,
  issue_count INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS roster_import_batch_receipts (
  run_id TEXT NOT NULL,
  batch_index INTEGER NOT NULL,
  plan_revision TEXT NOT NULL,
  payload_hash TEXT NOT NULL,
  PRIMARY KEY (run_id, batch_index)
);
CREATE TABLE IF NOT EXISTS roster_import_daily_budget (
  utc_day TEXT PRIMARY KEY,
  reserved_writes INTEGER NOT NULL DEFAULT 0,
  reserved_reads INTEGER NOT NULL DEFAULT 0
);
