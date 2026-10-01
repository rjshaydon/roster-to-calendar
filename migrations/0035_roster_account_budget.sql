-- No roster-history backfill. Carry old conservative reservations until UTC reset.
CREATE TABLE IF NOT EXISTS roster_account_budget (
  utc_day TEXT PRIMARY KEY,
  allocated_reads INTEGER NOT NULL DEFAULT 0,
  allocated_writes INTEGER NOT NULL DEFAULT 0,
  maximum_reads INTEGER NOT NULL DEFAULT 0,
  maximum_writes INTEGER NOT NULL DEFAULT 0,
  valid_until TEXT NOT NULL DEFAULT '',
  stop_reason TEXT NOT NULL DEFAULT '',
  observed_until TEXT NOT NULL DEFAULT '',
  account_reads INTEGER NOT NULL DEFAULT 0,
  account_writes INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS roster_maintenance_receipts (
  request_id TEXT PRIMARY KEY,
  utc_day TEXT NOT NULL,
  reserved_reads INTEGER NOT NULL DEFAULT 0,
  reserved_writes INTEGER NOT NULL DEFAULT 0,
  actual_reads INTEGER NOT NULL DEFAULT 0,
  actual_writes INTEGER NOT NULL DEFAULT 0,
  finished_at TEXT NOT NULL DEFAULT '',
  metadata_complete INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_roster_maintenance_receipts_day ON roster_maintenance_receipts(utc_day);
INSERT OR IGNORE INTO roster_maintenance_receipts(request_id,utc_day,reserved_reads,reserved_writes)
SELECT 'legacy-' || utc_day,utc_day,reserved_reads,reserved_writes FROM roster_import_daily_budget;
