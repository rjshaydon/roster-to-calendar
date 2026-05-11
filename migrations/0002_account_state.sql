CREATE TABLE IF NOT EXISTS account_profiles (
  email TEXT PRIMARY KEY,
  real_name TEXT NOT NULL DEFAULT '',
  role TEXT NOT NULL DEFAULT 'user',
  insights_enabled INTEGER NOT NULL DEFAULT 0,
  subscription_token TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS account_claims (
  email TEXT NOT NULL,
  source_type TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  display_name TEXT NOT NULL,
  matched_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (email, source_type, doctor_key)
);

CREATE INDEX IF NOT EXISTS idx_account_claims_doctor ON account_claims (source_type, doctor_key);

CREATE TABLE IF NOT EXISTS account_states (
  email TEXT PRIMARY KEY,
  session_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);
