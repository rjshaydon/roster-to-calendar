-- Dormant until the explicit session rollout gate is enabled. Never store tokens.
CREATE TABLE account_sessions (
  token_hash TEXT PRIMARY KEY,
  email TEXT NOT NULL REFERENCES account_profiles(email) ON DELETE CASCADE,
  authenticated_at INTEGER NOT NULL,
  idle_expires_at INTEGER NOT NULL,
  absolute_expires_at INTEGER NOT NULL,
  renew_after INTEGER NOT NULL,
  persistent INTEGER NOT NULL CHECK (persistent IN (0, 1))
);
CREATE INDEX account_sessions_email ON account_sessions(email, authenticated_at DESC, token_hash);
CREATE INDEX account_sessions_expiry ON account_sessions(idle_expires_at, token_hash);

-- Old/cached clients and invitation/admin password paths must never restore a
-- fast legacy credential after verified retirement, even outside the new route.
CREATE TRIGGER account_retired_password_guard BEFORE UPDATE OF password_hash, password_salt ON account_profiles
WHEN (NEW.password_hash <> '' OR NEW.password_salt <> '') AND EXISTS
  (SELECT 1 FROM account_authentication_links WHERE email=OLD.email AND phase='ready')
BEGIN
  SELECT RAISE(ABORT, 'Retired password cannot be restored.');
END;

-- Existing password/invitation paths must also revoke cookies. Account deletion
-- cascades; authority/claims are read from the current account on each request.
CREATE TRIGGER account_sessions_password_changed AFTER UPDATE OF password_hash, password_salt ON account_profiles
WHEN OLD.password_hash <> NEW.password_hash OR OLD.password_salt <> NEW.password_salt
BEGIN
  DELETE FROM account_sessions WHERE email = NEW.email;
END;

CREATE TRIGGER account_sessions_role_changed AFTER UPDATE OF role ON account_profiles
WHEN OLD.role <> NEW.role
BEGIN
  DELETE FROM account_sessions WHERE email = NEW.email;
END;
