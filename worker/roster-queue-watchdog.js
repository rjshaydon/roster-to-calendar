const defaultAutomationUrl = "https://roster-to-calendar.pages.dev/api/automation/dispatch";
const defaultFindMyShiftUrl = "https://roster-to-calendar.pages.dev/api/automation/findmyshift-check";

export default {
  async scheduled(_controller, env, ctx) {
    if (automationPaused(env)) {
      console.warn(JSON.stringify({ event: "roster-watchdog", status: "paused", reason: "d1-write-quota-protection" }));
      return;
    }
    ctx.waitUntil(kickRosterProcessor(env));
    ctx.waitUntil(checkFindMyShift(env));
  },

  async fetch(request, env) {
    if (new URL(request.url).pathname !== "/health") return new Response("Not found", { status: 404 });
    return Response.json({
      ok: true,
      service: "roster-queue-watchdog",
      configured: Boolean(String(env.ROSTER_WATCHDOG_TOKEN || "").trim()),
      paused: automationPaused(env),
    });
  },
};

function automationPaused(env = {}) {
  return !["1", "true", "yes", "on"].includes(String(env.ROSTER_AUTOMATION_ENABLED || "").trim().toLowerCase());
}

async function kickRosterProcessor(env) {
  const token = String(env.ROSTER_WATCHDOG_TOKEN || "").trim();
  if (!token) return { ok: false, error: "ROSTER_WATCHDOG_TOKEN is not configured." };
  const response = await fetch(String(env.ROSTER_AUTOMATION_URL || defaultAutomationUrl), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "User-Agent": "roster-queue-watchdog",
    },
    body: JSON.stringify({ action: "kick", reason: "cloudflare-watchdog" }),
  });
  const text = await response.text();
  let result = {};
  try {
    result = text ? JSON.parse(text) : {};
  } catch {
    result = { error: text.slice(0, 300) };
  }
  if (!response.ok) throw new Error(result.error || `Roster dispatch returned HTTP ${response.status}.`);
  console.log(JSON.stringify({ event: "roster-watchdog", dispatched: result.dispatched === true, reason: result.reason || "" }));
  return { ok: true, ...result };
}

async function checkFindMyShift(env) {
  const token = String(env.ROSTER_WATCHDOG_TOKEN || "").trim();
  if (!token) return { ok: false, error: "ROSTER_WATCHDOG_TOKEN is not configured." };
  const response = await fetch(String(env.FINDMYSHIFT_CHECK_URL || defaultFindMyShiftUrl), {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "User-Agent": "roster-queue-watchdog" },
  });
  const result = await response.json().catch(() => ({}));
  console.log(JSON.stringify({ event: "findmyshift-check", ok: response.ok, status: result.status || "" }));
  return { ok: response.ok, ...result };
}
