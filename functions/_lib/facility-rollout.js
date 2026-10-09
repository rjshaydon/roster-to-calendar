const ENABLED = new Set(["1", "true", "yes", "on"]);
const DISABLED = new Set(["0", "false", "no", "off"]);

export function facilityOverviewMaintenanceMode(env = {}) {
  return !DISABLED.has(String(env.FACILITY_OVERVIEW_MAINTENANCE_MODE || "").trim().toLowerCase());
}

export function facilityRolloutPaused(env = {}) {
  return !DISABLED.has(String(env.FACILITY_SHARED_EMERGENCY_PAUSED || "").trim().toLowerCase());
}

export function facilitySharedRolloutActive(env = {}) {
  return ENABLED.has(String(env.FACILITY_SHARED_ROLLOUT_ACTIVE || "").trim().toLowerCase());
}

export function facilityOverviewAutomaticLaunchEnabled(env = {}) {
  return ENABLED.has(String(env.FACILITY_OVERVIEW_AUTOMATIC_LAUNCH_ENABLED || "").trim().toLowerCase());
}

export function onShiftForAllEnabled(env = {}) {
  return String(env.ON_SHIFT_FOR_ALL_ENABLED || "").toLowerCase() === "true";
}

export function facilityLegacyReadsPaused(env = {}) {
  return !DISABLED.has(String(env.FACILITY_LEGACY_READS_PAUSED || "").trim().toLowerCase());
}

export function facilityBuildSources(env = {}, sources = []) {
  if (facilityRolloutPaused(env)) return [];
  const allowed = configuredSet(env.FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST);
  return normalizeSources(sources).filter((source) => allowed.has(source));
}

export function facilityReaderSources(env = {}) {
  if (!facilitySharedRolloutActive(env) || facilityRolloutPaused(env)) return [];
  return normalizeSources(String(env.FACILITY_SHARED_READER_SOURCE_ALLOWLIST || "").split(","));
}

export function facilityContactReaderSources(env = {}) {
  if (String(env.FACILITY_SHARED_CONTACTS_ENABLED || "").trim().toLowerCase() !== "true") return [];
  const sharedSources = new Set(facilityReaderSources(env));
  return normalizeSources(String(env.FACILITY_SHARED_CONTACTS_SOURCE_ALLOWLIST || "").split(","))
    .filter((source) => sharedSources.has(source));
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

export function facilityOverviewMaintenanceForViewer(env = {}, viewer = {}) {
  return facilityOverviewMaintenanceMode(env)
    || !facilityRolloutCohortEligible(env, viewer);
}

export function facilityReadRoute(env = {}, { actorRole = "", actorEmail = "", subjectEmail = "", sources = [] } = {}) {
  if (!facilitySharedRolloutActive(env)) return facilityLegacyReadsPaused(env) ? "blocked" : "legacy";
  if (facilityRolloutPaused(env)) return "blocked";
  if (!facilityRolloutCohortEligible(env, { actorRole, actorEmail, subjectEmail })) {
    return facilityLegacyReadsPaused(env) ? "blocked" : "legacy";
  }
  const allowed = new Set(facilityReaderSources(env));
  const requested = normalizeSources(sources);
  return requested.length && requested.every((source) => allowed.has(source)) ? "shared" : "blocked";
}

// Call only after authorising every requested hospital. Combined views may
// retain enabled hospitals without weakening explicit single-site requests.
export function facilityReaderSelection(env, viewer, sources, allowPartial = false) {
  const requested = normalizeSources(sources);
  const route = facilityReadRoute(env, { ...viewer, sources: requested });
  if (route !== "blocked") return { route, sources: requested, missing: [] };
  const enabled = new Set(facilityReaderSources(env));
  const available = requested.filter(source => enabled.has(source));
  const missing = requested.filter(source => !enabled.has(source)).map(sourceType => ({ sourceType, reason: "reader-unavailable" }));
  if (allowPartial && facilityReadRoute(env, { ...viewer, sources: available }) === "shared") {
    return { route: "shared", sources: available, missing };
  }
  return { route: "blocked", sources: [], missing };
}

function configuredSet(value) {
  return new Set(normalizeSources(String(value || "").split(",")));
}

function normalizeSources(values) {
  return [...new Set((values || []).map((value) => String(value || "").trim().toLowerCase()).filter((value) => /^(mmc|mch|ddh|vhh|casey)$/.test(value)))];
}
