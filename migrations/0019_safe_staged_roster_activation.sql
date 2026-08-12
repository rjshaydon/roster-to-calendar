-- Existing derived rows pre-date the shared parser rollout. Do not represent
-- them as having been successfully processed by the new parser.
ALTER TABLE roster_sync_runs ADD COLUMN source_file_id TEXT NOT NULL DEFAULT '';

UPDATE roster_sync_runs
SET source_file_id = file_id
WHERE source_file_id = '';

UPDATE roster_files
SET parser_version = 'legacy-unverified'
WHERE parser_version = '' OR parser_version = 'manual-core-v1';
