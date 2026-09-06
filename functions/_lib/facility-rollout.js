const ENABLED = new Set(["1", "true", "yes", "on"]);

export function facilityRolloutPaused(env = {}) {
  return ENABLED.has(String(env.FACILITY_SHARED_EMERGENCY_PAUSED || "").trim().toLowerCase());
}

export function facilityBuildSources(env = {}, sources = []) {
  if (facilityRolloutPaused(env)) return [];
  const allowed = configuredSet(env.FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST);
  return normalizeSources(sources).filter((source) => allowed.has(source));
}

export function facilitySharedReaderAllowed(env = {}, { actorRole = "", actorEmail = "", subjectEmail = "", sources = [] } = {}) {
  if (facilityRolloutPaused(env)) return false;
  const allowed = configuredSet(env.FACILITY_SHARED_READER_SOURCE_ALLOWLIST);
  const requested = normalizeSources(sources);
  if (!requested.length || requested.some((source) => !allowed.has(source))) return false;
  const cohort = String(env.FACILITY_SHARED_READER_COHORT || "").trim().toLowerCase();
  if (cohort === "all") return true;
  if (cohort !== "creator") return false;
  const role = String(actorRole || "").toLowerCase();
  return (role === "creator" || role === "owner")
    && String(actorEmail || "").trim().toLowerCase() === String(subjectEmail || actorEmail || "").trim().toLowerCase();
}

function configuredSet(value) {
  return new Set(normalizeSources(String(value || "").split(",")));
}

function normalizeSources(values) {
  return [...new Set((values || []).map((value) => String(value || "").trim().toLowerCase()).filter((value) => /^(mmc|mch|ddh|vhh|casey)$/.test(value)))];
}
