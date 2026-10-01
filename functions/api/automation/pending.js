import { hasCalendarDb, listQueuedRosterSyncRuns } from "../../_lib/d1-calendar.js";
import { localFeatureDisabledResponse } from "../../_lib/outbound-network.js";
import { automatedRosterQueueEnabled, automatedRosterSourceEnabled, automatedRosterWritesEnabled, rosterWritePausedResponse } from "../../_lib/roster-automation-guard.js";

export async function onRequestGet(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  const localDisabled = localFeatureDisabledResponse(context.env, "Roster automation queue polling");
  if (localDisabled) return localDisabled;
  if (!automatedRosterWritesEnabled(context.env)) return rosterWritePausedResponse();
  if (!automatedRosterQueueEnabled(context.env)) return rosterWritePausedResponse();
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  const url = new URL(context.request.url);
  const sourceId = String(url.searchParams.get("sourceId") || "").trim();
  if (!automatedRosterSourceEnabled(context.env, sourceId)) return rosterWritePausedResponse();
  const limit = Number(url.searchParams.get("limit") || 1);
  const boundedImportEnabled = String(context.env.ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED || "") === "true";
  const allowance = boundedImportEnabled ? await context.env.ROSTER_DB.prepare("SELECT reserved_writes, reserved_reads FROM roster_import_daily_budget WHERE utc_day = ?").bind(new Date().toISOString().slice(0, 10)).first() : null;
  const maintenanceDeferred = Number(allowance?.reserved_writes || 0) >= 9500 || Number(allowance?.reserved_reads || 0) >= 490000;
  const runs = await listQueuedRosterSyncRuns(context.env.ROSTER_DB, sourceId, limit);
  return Response.json({
    ok: true,
    boundedImportEnabled,
    maintenanceDeferred,
    runs: runs.map(({ objectKey: _objectKey, ...run }) => run),
  });
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
