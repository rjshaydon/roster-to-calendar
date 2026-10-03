-- Bounded lookup of a linked clinician's present and former hospital terms.
CREATE INDEX IF NOT EXISTS idx_facility_staff_identity_terms
ON facility_term_staff_contributions (source_type, doctor_key, term_start);

-- Cached scope must be recomputed after roster or grade corrections. Deleting
-- the small access cache once leaves subsequent row triggers with no writes.
CREATE TRIGGER IF NOT EXISTS facility_access_staff_insert AFTER INSERT ON facility_term_staff_contributions
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_staff_update AFTER UPDATE ON facility_term_staff_contributions
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_staff_delete AFTER DELETE ON facility_term_staff_contributions
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_grade_insert AFTER INSERT ON facility_staff_seniority_overrides
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_grade_update AFTER UPDATE ON facility_staff_seniority_overrides
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_grade_delete AFTER DELETE ON facility_staff_seniority_overrides
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_sms_insert AFTER INSERT ON facility_sms_memberships
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_sms_update AFTER UPDATE ON facility_sms_memberships
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_sms_delete AFTER DELETE ON facility_sms_memberships
BEGIN DELETE FROM facility_access_sessions; END;
CREATE TRIGGER IF NOT EXISTS facility_access_file_activation AFTER UPDATE OF active ON roster_files
WHEN OLD.active <> NEW.active
BEGIN DELETE FROM facility_access_sessions; END;
