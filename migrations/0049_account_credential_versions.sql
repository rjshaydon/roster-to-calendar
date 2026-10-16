-- Credential operations fence password proof obtained before or during a reset.
-- The pending target is a server-keyed HMAC, never a reusable password hash.
ALTER TABLE account_authentication_links ADD COLUMN credential_version INTEGER NOT NULL DEFAULT 0;
ALTER TABLE account_authentication_links ADD COLUMN credential_operation_id TEXT NOT NULL DEFAULT '';
ALTER TABLE account_authentication_links ADD COLUMN credential_target_digest TEXT NOT NULL DEFAULT '';
ALTER TABLE account_authentication_links ADD COLUMN credential_operation_started_at TEXT NOT NULL DEFAULT '';

CREATE TRIGGER account_sessions_credential_version_changed
AFTER UPDATE OF credential_version ON account_authentication_links
WHEN OLD.credential_version <> NEW.credential_version
BEGIN
  DELETE FROM account_sessions WHERE email=NEW.email;
END;

-- Outstanding reset/invitation links must not outlive a password change.
CREATE TRIGGER account_credential_version_retires_email_proofs
AFTER UPDATE OF credential_version ON account_authentication_links
WHEN OLD.credential_version <> NEW.credential_version
BEGIN
  UPDATE account_email_challenges SET phase='complete' WHERE email=NEW.email AND phase='issued';
  UPDATE account_invites SET revoked_at=strftime('%Y-%m-%dT%H:%M:%fZ','now'),updated_at=strftime('%Y-%m-%dT%H:%M:%fZ','now')
    WHERE email=NEW.email AND accepted_at='' AND revoked_at='';
END;
