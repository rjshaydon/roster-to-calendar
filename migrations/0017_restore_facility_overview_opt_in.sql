-- The At a glance rollout is opt-in for standard users. Migration 0016
-- unintentionally enabled it for every existing account, so reset standard
-- accounts to the secure default. Creator access is unconditional in the API,
-- but keeping the stored value enabled makes the database state explicit too.
UPDATE account_profiles
SET facility_overview_enabled = CASE
  WHEN role IN ('creator', 'owner') THEN 1
  ELSE 0
END;
