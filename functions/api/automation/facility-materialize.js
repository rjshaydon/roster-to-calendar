import { runFacilityPublicationStep } from "../../_lib/facility-overview-cache.js";
import { facilityBuildSources } from "../../_lib/facility-rollout.js";
import { facilityMaterializationMaintenanceEnabled, rosterWritePausedResponse } from "../../_lib/roster-automation-guard.js";

export async function onRequestPost(context) {
  if (!validToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!facilityMaterializationMaintenanceEnabled(context.env)) return rosterWritePausedResponse();
  const body = await context.request.json().catch(() => ({}));
  const sourceType = facilityBuildSources(context.env, [body?.sourceType])[0] || "";
  if (!sourceType) return rosterWritePausedResponse();
  const result = await runFacilityPublicationStep(context, sourceType, {
    mode: body?.mode || "plan",
    termStart: body?.termStart,
    operationRevision: body?.operationRevision,
    batchIndex: body?.batchIndex,
    month: body?.month,
  });
  return Response.json(result, { status: result.ok ? 200 : result.overBudget || result.stalePlan ? 409 : 400 });
}

function validToken(request, configured) {
  const token = String(configured || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
