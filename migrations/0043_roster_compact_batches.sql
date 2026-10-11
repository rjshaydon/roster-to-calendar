-- Additive staging receipts. Inactive roster metadata can resume without
-- re-reading and hashing the entire term in one Pages request.
CREATE TABLE roster_import_compact_receipts (
  run_id TEXT NOT NULL,
  batch_index INTEGER NOT NULL,
  plan_revision TEXT NOT NULL,
  payload_hash TEXT NOT NULL,
  PRIMARY KEY (run_id,batch_index)
);
CREATE TABLE roster_patch_jobs (
  run_id TEXT PRIMARY KEY,
  file_id TEXT NOT NULL,
  source_id TEXT NOT NULL,
  source_type TEXT NOT NULL,
  baseline_revision TEXT NOT NULL,
  baseline_parsed_at TEXT NOT NULL,
  manifest_json TEXT NOT NULL DEFAULT '',
  plan_revision TEXT NOT NULL DEFAULT '',
  next_batch INTEGER NOT NULL DEFAULT 0,
  activated INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE roster_patch_batches (
  run_id TEXT NOT NULL,
  batch_index INTEGER NOT NULL,
  payload_hash TEXT NOT NULL,
  entries_json TEXT NOT NULL,
  PRIMARY KEY (run_id,batch_index)
);
