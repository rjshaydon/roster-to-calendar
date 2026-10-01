-- Pins the active roster set before staging. No roster-history backfill.
ALTER TABLE roster_import_jobs ADD COLUMN promotion_fence_json TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_import_jobs ADD COLUMN activation_token TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_import_jobs ADD COLUMN retired_file_ids_json TEXT NOT NULL DEFAULT '[]';
