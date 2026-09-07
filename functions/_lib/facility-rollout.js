const ENABLED = new Set(["1", "true", "yes", "on"]);
const DISABLED = new Set(["0", "false", "no", "off"]);

export function facilityRolloutPaused(env = {}) {
  return !DISABLED.has(String(env.FACILITY_SHARED_EMERGENCY_PAUSED || "").trim().toLowerCase());
}

export function facilitySharedRolloutActive(env = {}) {
  return ENABLED.has(String(env.FACILITY_SHARED_ROLLOUT_ACTIVE || "").trim().toLowerCase());
}

export function facilityLegacyReadsPaused(env = {}) {
  return !DISABLED.has(String(env.FACILITY_LEGACY_READS_PAUSED || "").trim().toLowerCase());
}

export function facilityBuildSources(env = {}, sources = []) {
  if (facilityRolloutPaused(env)) return [];
  const allowed = configuredSet(env.FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST);
  return normalizeSources(sources).filter((source) => allowed.has(source));
}

export function facilitySharedReaderAllowed(env = {}, { actorRole = "", actorEmail = "", subjectEmail = "", sources = [] } = {}) {
  return facilityReadRoute(env, { actorRole, actorEmail, subjectEmail, sources }) === "shared";
}

export function facilityRolloutCohortEligible(env = {}, { actorRole = "", actorEmail = "", subjectEmail = "" } = {}) {
  if (!facilitySharedRolloutActive(env) || facilityRolloutPaused(env)) return false;
  const cohort = String(env.FACILITY_SHARED_READER_COHORT || "").trim().toLowerCase();
  if (cohort === "all") return true;
  if (cohort !== "creator") return false;
  const role = String(actorRole || "").toLowerCase();
  return (role === "creator" || role === "owner")
    && String(actorEmail || "").trim().toLowerCase() === String(subjectEmail || actorEmail || "").trim().toLowerCase();
}

export function facilityReadRoute(env = {}, { actorRole = "", actorEmail = "", subjectEmail = "", sources = [] } = {}) {
  if (!facilitySharedRolloutActive(env)) return facilityLegacyReadsPaused(env) ? "blocked" : "legacy";
  if (facilityRolloutPaused(env)) return "blocked";
  if (!facilityRolloutCohortEligible(env, { actorRole, actorEmail, subjectEmail })) {
    return facilityLegacyReadsPaused(env) ? "blocked" : "legacy";
  }
  const allowed = configuredSet(env.FACILITY_SHARED_READER_SOURCE_ALLOWLIST);
  const requested = normalizeSources(sources);
  return requested.length && requested.every((source) => allowed.has(source)) ? "shared" : "blocked";
}

function configuredSet(value) {
  return new Set(normalizeSources(String(value || "").split(",")));
}

function normalizeSources(values) {
  return [...new Set((values || []).map((value) => String(value || "").trim().toLowerCase()).filter((value) => /^(mmc|mch|ddh|vhh|casey)$/.test(value)))];
}
