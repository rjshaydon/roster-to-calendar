CREATE TABLE IF NOT EXISTS doctor_profiles (
  profile_id TEXT PRIMARY KEY,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  source_types_json TEXT NOT NULL DEFAULT '[]',
  state_json TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_doctor_profiles_doctor ON doctor_profiles (doctor_key);
