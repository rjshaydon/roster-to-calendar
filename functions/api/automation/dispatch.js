import { hasCalendarDb } from "../../_lib/d1-calendar.js";
import { recordRosterDispatchLifecycle, requestQueuedRosterProcessing } from "../../_lib/automation-dispatch.js";

export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN, context.env.ROSTER_WATCHDOG_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  try {
    const body = await context.request.json().catch(() => ({}));
    const action = String(body?.action || "kick").trim().toLowerCase();
    if (action === "kick") {
      const result = await requestQueuedRosterProcessing(context.env, { reason: String(body?.reason || "watchdog").slice(0, 80) });
      return Response.json(result, { status: result.ok ? 200 : 202 });
    }
    if (action === "lifecycle") {
      const result = await recordRosterDispatchLifecycle(context.env, body);
      return Response.json(result, { status: result.ok ? 200 : 400 });
    }
    return Response.json({ error: "Unsupported dispatch action." }, { status: 400 });
  } catch (error) {
    console.error("Roster dispatch endpoint failed", error);
    return Response.json({ error: "Roster dispatch could not be updated." }, { status: 422 });
  }
}

function hasValidAutomationToken(request, ...configuredTokens) {
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!provided) return false;
  return configuredTokens.some((configuredToken) => {
    const token = String(configuredToken || "");
    if (!token || token.length !== provided.length) return false;
    let mismatch = 0;
    for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
    return mismatch === 0;
  });
}
