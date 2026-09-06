const ENABLED = new Set(["1", "true", "yes", "on"]);

export function contactAutomationSourceEnabled(env = {}, sourceId = "") {
  if (!ENABLED.has(String(env.CONTACT_AUTOMATION_WRITES_ENABLED || "").trim().toLowerCase())) return false;
  const allowed = new Set(String(env.CONTACT_AUTOMATION_SOURCE_ALLOWLIST || "")
    .split(",").map((value) => value.trim()).filter(Boolean));
  return allowed.has(String(sourceId || "").trim());
}

export function contactAutomationPausedResponse() {
  return Response.json({
    ok: false,
    status: "paused",
    error: "Contact updates are temporarily paused to protect the D1 free-tier quota. Existing valid contact information remains available.",
  }, { status: 503, headers: { "Retry-After": "3600" } });
}
