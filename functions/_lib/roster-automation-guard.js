const ENABLED_VALUES = new Set(["1", "true", "yes", "on"]);

export function automatedRosterWritesEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.ROSTER_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase());
}

export function rosterWritesExplicitlyPaused(env = {}) {
  return String(env.ROSTER_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase() === "false";
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

export function advancedRosterMaintenanceEnabled(env = {}) {
  return automatedRosterWritesEnabled(env) && ENABLED_VALUES.has(String(env.ROSTER_ADVANCED_MAINTENANCE_ENABLED || "").trim().toLowerCase());
}

export function facilityMaterializationMaintenanceEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.ROSTER_ADVANCED_MAINTENANCE_ENABLED || "").trim().toLowerCase());
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
