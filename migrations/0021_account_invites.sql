CREATE TABLE IF NOT EXISTS account_invites (
  token_hash TEXT PRIMARY KEY,
  email TEXT NOT NULL,
  created_by TEXT NOT NULL DEFAULT '',
  expires_at TEXT NOT NULL,
  accepted_at TEXT NOT NULL DEFAULT '',
  revoked_at TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_account_invites_email ON account_invites (email);
