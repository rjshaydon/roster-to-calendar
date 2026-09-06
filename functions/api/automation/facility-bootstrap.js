import { inspectFacilityOverviewBootstrap, refreshFacilityOverviewMaterializationForFile } from "../../_lib/d1-calendar.js";
import { facilityBuildSources } from "../../_lib/facility-rollout.js";
import { advancedRosterMaintenanceEnabled, rosterWritePausedResponse } from "../../_lib/roster-automation-guard.js";

const MAX_EVENT_ROWS = 25000;
const MAX_COMPACT_WRITES = 750;

export async function onRequestPost(context) {
  if (!validToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!advancedRosterMaintenanceEnabled(context.env)) return rosterWritePausedResponse();
  const body = await context.request.json().catch(() => ({}));
  const sourceType = facilityBuildSources(context.env, [body?.sourceType])[0] || "";
  const fileId = String(body?.fileId || "").trim();
  if (!sourceType || !fileId) return rosterWritePausedResponse();
  const maximumEventRows = Math.max(1, Math.min(Number(body?.maximumEventRows || MAX_EVENT_ROWS), MAX_EVENT_ROWS));
  const maximumWrites = Math.max(1, Math.min(Number(body?.maximumWrites || MAX_COMPACT_WRITES), MAX_COMPACT_WRITES));
  const inspection = await inspectFacilityOverviewBootstrap(context.env.ROSTER_DB, { fileId, sourceType });
  if (!inspection.ok) return Response.json(inspection, { status: 400 });
  const estimate = {
    maximumEventRowsExamined: maximumEventRows + 1,
    maximumCompactMutationStatements: maximumWrites,
    maximumEstimatedD1RowsWrittenIncludingIndexes: maximumWrites * 3,
    broadSourceScans: 0,
    automaticContinuation: false,
  };
  if (body?.execute !== true) return Response.json({ ...inspection, dryRun: true, estimate });
  if (inspection.compactReady) return Response.json({ ...inspection, dryRun: false, unchanged: true, writes: 0, estimate });
  if (!body?.planRevision || String(body.planRevision) !== inspection.planRevision) {
    return Response.json({ ok: false, stalePlan: true, reason: "bootstrap-plan-changed" }, { status: 409 });
  }
  const result = await refreshFacilityOverviewMaterializationForFile(context.env.ROSTER_DB, fileId, { maximumEventRows, maximumWrites });
  return Response.json({ ...inspection, dryRun: false, estimate, result }, { status: result?.overBudget ? 409 : 200 });
}

function validToken(request, configured) {
  const token = String(configured || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
