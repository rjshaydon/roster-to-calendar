-- Reuse inactive staged facts without exposing an unfinished roster.
ALTER TABLE roster_import_jobs ADD COLUMN rebase_owner TEXT NOT NULL DEFAULT '';
CREATE INDEX idx_roster_import_reuse ON roster_import_jobs(source_id,activated,initialized,
 json_extract(manifest_json,'$.startDate'),json_extract(manifest_json,'$.rosterEndDate'),event_count DESC);
CREATE INDEX idx_roster_import_file_reuse ON roster_import_jobs(file_id,event_count DESC);
CREATE TABLE roster_rebase_jobs (
 run_id TEXT PRIMARY KEY, baseline_run_id TEXT NOT NULL, file_id TEXT NOT NULL,
 source_id TEXT NOT NULL, baseline_revision TEXT NOT NULL, baseline_parsed_at TEXT NOT NULL,
 manifest_json TEXT NOT NULL DEFAULT '', plan_revision TEXT NOT NULL DEFAULT '',
 next_batch INTEGER NOT NULL DEFAULT 0, prepared INTEGER NOT NULL DEFAULT 0
);
