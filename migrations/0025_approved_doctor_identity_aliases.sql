-- Durable, audited roster identity associations. Roster event rows keep their
-- original source spelling and doctor_key; calendar reads expand approved
-- aliases through roster_person_aliases.
CREATE TABLE IF NOT EXISTS roster_person_alias_audit (
  audit_id TEXT PRIMARY KEY,
  person_id TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL DEFAULT '',
  action TEXT NOT NULL,
  provenance TEXT NOT NULL,
  approved_by TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  details_json TEXT NOT NULL DEFAULT '{}'
);

CREATE INDEX IF NOT EXISTS idx_roster_person_alias_audit_person
ON roster_person_alias_audit (person_id, created_at DESC);

-- Confirmed cross-hospital identity: Jay WEERARATNE at VHH is Jayantha
-- WEERARATNE at DDH. This changes identity expansion only; it does not update
-- or delete any roster doctor/event row.
INSERT OR IGNORE INTO roster_people (
  person_id, preferred_display_name, provenance, review_state,
  created_at, updated_at, approved_by
) VALUES (
  'person:weeraratne-jayantha', 'Jayantha WEERARATNE',
  'admin-approved', 'approved', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP,
  'confirmed-mapping:2026-08-31'
);

UPDATE account_people
SET person_id = 'person:weeraratne-jayantha', updated_at = CURRENT_TIMESTAMP
WHERE person_id IN (
  SELECT person_id FROM roster_person_aliases
  WHERE (source_type = 'vhh' AND doctor_key = 'JAY WEERARATNE')
     OR (source_type = 'ddh' AND doctor_key = 'JAYANTHA WEERARATNE')
);

UPDATE roster_person_aliases
SET person_id = 'person:weeraratne-jayantha',
    provenance = 'admin-approved', confidence = 'confirmed',
    review_state = 'approved', updated_at = CURRENT_TIMESTAMP,
    approved_by = 'confirmed-mapping:2026-08-31'
WHERE (source_type = 'vhh' AND doctor_key = 'JAY WEERARATNE')
   OR (source_type = 'ddh' AND doctor_key = 'JAYANTHA WEERARATNE');

INSERT OR IGNORE INTO roster_person_aliases (
  source_type, doctor_key, display_name, person_id, provenance, confidence,
  review_state, created_at, updated_at, approved_by
) VALUES
  ('vhh', 'JAY WEERARATNE', 'Jay WEERARATNE',
   'person:weeraratne-jayantha', 'admin-approved', 'confirmed', 'approved',
   CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, 'confirmed-mapping:2026-08-31'),
  ('ddh', 'JAYANTHA WEERARATNE', 'Jayantha WEERARATNE',
   'person:weeraratne-jayantha', 'admin-approved', 'confirmed', 'approved',
   CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, 'confirmed-mapping:2026-08-31');

INSERT OR IGNORE INTO roster_person_alias_audit (
  audit_id, person_id, source_type, doctor_key, display_name, action,
  provenance, approved_by, created_at, details_json
) VALUES
  ('confirmed-weeraratne-vhh-2026-08-31', 'person:weeraratne-jayantha',
   'vhh', 'JAY WEERARATNE', 'Jay WEERARATNE', 'link', 'admin-approved',
   'confirmed-mapping:2026-08-31', CURRENT_TIMESTAMP,
   '{"canonicalSourceType":"ddh","canonicalDoctorKey":"JAYANTHA WEERARATNE"}'),
  ('confirmed-weeraratne-ddh-2026-08-31', 'person:weeraratne-jayantha',
   'ddh', 'JAYANTHA WEERARATNE', 'Jayantha WEERARATNE', 'canonical',
   'admin-approved', 'confirmed-mapping:2026-08-31', CURRENT_TIMESTAMP,
   '{"aliasSourceType":"vhh","aliasDoctorKey":"JAY WEERARATNE"}');
