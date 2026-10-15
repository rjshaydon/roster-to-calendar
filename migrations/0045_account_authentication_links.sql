-- Additive, opt-in authentication preparation. No account backfill or activation.
-- A reserved UUID survives an uncertain provider response; never overwrite an
-- existing external account simply because its email matches.
CREATE TABLE account_authentication_links (
  email TEXT PRIMARY KEY REFERENCES account_profiles(email) ON DELETE CASCADE,
  provider TEXT NOT NULL CHECK (provider = 'supabase'),
  provider_subject TEXT NOT NULL UNIQUE,
  phase TEXT NOT NULL CHECK (phase IN ('pending', 'ready')),
  credential_revision TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
