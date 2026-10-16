-- Dormant, bounded email proofs. Store only a digest of the emailed token.
CREATE TABLE account_email_challenges (
  id TEXT PRIMARY KEY,
  email TEXT NOT NULL REFERENCES account_profiles(email) ON DELETE CASCADE,
  provider_subject TEXT NOT NULL,
  purpose TEXT NOT NULL CHECK(purpose IN ('confirm','recovery')),
  verification_type TEXT NOT NULL CHECK(verification_type IN ('magiclink','recovery')),
  token_digest TEXT NOT NULL,
  expires_at INTEGER NOT NULL,
  phase TEXT NOT NULL CHECK(phase IN ('issued','using','complete')),
  created_at INTEGER NOT NULL
);
CREATE INDEX account_email_challenges_email ON account_email_challenges(email,purpose,created_at DESC);
CREATE INDEX account_email_challenges_expiry ON account_email_challenges(expires_at,id);
CREATE TABLE account_auth_email_budget (
  month TEXT PRIMARY KEY,
  reserved INTEGER NOT NULL
);
CREATE TABLE account_email_delivery_locks (
  email TEXT NOT NULL REFERENCES account_profiles(email) ON DELETE CASCADE,
  purpose TEXT NOT NULL,
  next_at INTEGER NOT NULL,
  PRIMARY KEY(email,purpose)
);
