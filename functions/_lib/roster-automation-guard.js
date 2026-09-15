const ENABLED_VALUES = new Set(["1", "true", "yes", "on"]);

export function automatedRosterWritesEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.ROSTER_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase());
}

export function manualRosterWritesEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.MANUAL_ROSTER_WRITES_ENABLED || "").trim().toLowerCase());
}

export function rosterWritesExplicitlyPaused(env = {}) {
  return !manualRosterWritesEnabled(env);
}

export function rosterStatusSummaryEnabled(env = {}) {
  return String(env.ROSTER_STATUS_SUMMARY_ENABLED || "").trim().toLowerCase() === "true";
}

export function automatedRosterSourceEnabled(env = {}, sourceId = "") {
  if (!automatedRosterWritesEnabled(env)) return false;
  const allowed = new Set(String(env.ROSTER_AUTOMATION_SOURCE_ALLOWLIST || "").split(",").map((value) => value.trim()).filter(Boolean));
  return allowed.has(String(sourceId || "").trim());
}

export function automatedRosterQueueEnabled(env = {}) {
  return automatedRosterWritesEnabled(env) && ENABLED_VALUES.has(String(env.ROSTER_AUTOMATION_QUEUE_ENABLED || "").trim().toLowerCase());
}

export function reviewedRosterFactLimit(env = {}, sourceId = "", contentHash = "") {
  if (!automatedRosterSourceEnabled(env, sourceId) || !automatedRosterQueueEnabled(env)) return 0;
  const limit = Math.max(0, Math.min(Math.floor(Number(env.ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT || 0)), 5000));
  const reviewedHash = String(env.ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256 || "").trim().toLowerCase();
  const suppliedHash = String(contentHash || "").trim().toLowerCase();
  if (!limit) return 0;
  // A hash pins a one-shot canary to one reviewed workbook. Routine ingestion
  // deliberately leaves it blank and relies on the exact source allowlist plus
  // this hard incremental-fact ceiling for future provider versions.
  if (reviewedHash && (!/^[a-f0-9]{64}$/.test(reviewedHash) || suppliedHash !== reviewedHash)) return 0;
  return limit;
}

export function advancedRosterMaintenanceEnabled(env = {}) {
  return manualRosterWritesEnabled(env) && ENABLED_VALUES.has(String(env.ROSTER_ADVANCED_MAINTENANCE_ENABLED || "").trim().toLowerCase());
}

export function facilityMaterializationMaintenanceEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.ROSTER_ADVANCED_MAINTENANCE_ENABLED || "").trim().toLowerCase());
}

export function facilityBootstrapInspectionEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.FACILITY_BOOTSTRAP_INSPECTION_ENABLED || "").trim().toLowerCase());
}

export function facilityBootstrapExecutionEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.FACILITY_BOOTSTRAP_EXECUTION_ENABLED || "").trim().toLowerCase());
}

export function facilityBootstrapFileAllowed(env = {}, fileId = "") {
  const normalized = String(fileId || "").trim();
  if (!normalized) return false;
  const entries = String(env.FACILITY_BOOTSTRAP_FILE_ALLOWLIST || "")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (entries.length !== 1) return false;
  const allowed = new Set(entries);
  return allowed.has(normalized);
}

export function rosterWritePausedResponse() {
  return Response.json({
    ok: false,
    status: "paused",
    error: "Roster updates are temporarily paused to protect the D1 free-tier quota. Existing calendars remain available.",
  }, {
    status: 503,
    headers: { "Retry-After": "3600" },
  });
}
