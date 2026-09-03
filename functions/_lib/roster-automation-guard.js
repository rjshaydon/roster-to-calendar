const ENABLED_VALUES = new Set(["1", "true", "yes", "on"]);

export function automatedRosterWritesEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.ROSTER_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase());
}

export function rosterWritesExplicitlyPaused(env = {}) {
  return String(env.ROSTER_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase() === "false";
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
