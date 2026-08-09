CREATE TABLE IF NOT EXISTS facility_sms_memberships (
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL DEFAULT '',
  first_seen_date TEXT NOT NULL DEFAULT '',
  last_seen_date TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, doctor_key)
);

CREATE INDEX IF NOT EXISTS idx_facility_sms_memberships_first_seen
  ON facility_sms_memberships (source_type, first_seen_date);

INSERT INTO facility_sms_memberships (
  source_type, doctor_key, display_name, first_seen_date, last_seen_date, created_at, updated_at
)
SELECT roster_file_doctors.source_type, roster_file_doctors.doctor_key,
  MAX(roster_file_doctors.display_name), MIN(roster_events.start_date), MAX(roster_events.start_date), datetime('now'), datetime('now')
FROM roster_file_doctors
INNER JOIN roster_events ON roster_events.file_id = roster_file_doctors.file_id
  AND roster_events.doctor_key = roster_file_doctors.doctor_key
WHERE UPPER(roster_file_doctors.seniority) = 'SMS'
GROUP BY roster_file_doctors.source_type, roster_file_doctors.doctor_key
ON CONFLICT(source_type, doctor_key) DO UPDATE SET
  display_name = excluded.display_name,
  first_seen_date = MIN(facility_sms_memberships.first_seen_date, excluded.first_seen_date),
  last_seen_date = MAX(facility_sms_memberships.last_seen_date, excluded.last_seen_date),
  updated_at = excluded.updated_at
WHERE facility_sms_memberships.display_name <> excluded.display_name
  OR facility_sms_memberships.first_seen_date > excluded.first_seen_date
  OR facility_sms_memberships.last_seen_date < excluded.last_seen_date;
