-- Exact per-file indexes used by the bounded facility bootstrap. These tables
-- are populated only after the controlled compact-fact bootstrap is approved.

CREATE INDEX IF NOT EXISTS idx_facility_term_staff_file
ON facility_term_staff_contributions (file_id, term_start, doctor_key);

CREATE INDEX IF NOT EXISTS idx_facility_stream_catalog_file
ON facility_stream_catalog_contributions (file_id, term_start, catalog_key);

CREATE INDEX IF NOT EXISTS idx_facility_staff_designations_term
ON facility_staff_designations (source_type, active, term_start, doctor_key);

CREATE INDEX IF NOT EXISTS idx_facility_staff_seniority_overrides_term
ON facility_staff_seniority_overrides (source_type, active, term_start, doctor_key);

CREATE INDEX IF NOT EXISTS idx_roster_files_source_active
ON roster_files (source_type, active, id);
