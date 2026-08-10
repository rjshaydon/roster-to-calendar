ALTER TABLE roster_file_doctors ADD COLUMN seniority TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_file_doctors ADD COLUMN membership_source TEXT NOT NULL DEFAULT 'roster';
CREATE INDEX IF NOT EXISTS idx_roster_file_doctors_source_file ON roster_file_doctors (source_type, file_id);
