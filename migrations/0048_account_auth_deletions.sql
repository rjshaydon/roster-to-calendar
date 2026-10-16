-- Preserve provider cleanup identity after local FK cascades remove its mapping.
CREATE TABLE account_auth_deletions (
  provider_subject TEXT PRIMARY KEY,
  email TEXT NOT NULL,
  phase TEXT NOT NULL DEFAULT 'queued' CHECK(phase IN ('queued','complete')),
  attempts INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE INDEX account_auth_deletions_email ON account_auth_deletions(email,phase);
CREATE INDEX account_auth_deletions_queue ON account_auth_deletions(phase,attempts,created_at);
CREATE TRIGGER account_authentication_cleanup BEFORE DELETE ON account_profiles
BEGIN
  INSERT INTO account_auth_deletions(provider_subject,email,phase,attempts,created_at,updated_at)
    SELECT provider_subject,email,'queued',0,strftime('%Y-%m-%dT%H:%M:%fZ','now'),strftime('%Y-%m-%dT%H:%M:%fZ','now')
    FROM account_authentication_links WHERE email=OLD.email
    ON CONFLICT(provider_subject) DO NOTHING;
END;
