import { FACILITY_BOOTSTRAP_LIMITS, inspectFacilityOverviewBootstrap, refreshFacilityOverviewMaterializationForFile, upsertRosterFileStatusSummary } from "../../_lib/d1-calendar.js";
import { facilityBuildSources } from "../../_lib/facility-rollout.js";
import { facilityMaterializationMaintenanceEnabled, rosterWritePausedResponse } from "../../_lib/roster-automation-guard.js";

export async function onRequestPost(context) {
  if (!validToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!facilityMaterializationMaintenanceEnabled(context.env)) return rosterWritePausedResponse();
  const body = await context.request.json().catch(() => ({}));
  const sourceType = facilityBuildSources(context.env, [body?.sourceType])[0] || "";
  const fileId = String(body?.fileId || "").trim();
  if (!sourceType || !fileId) return rosterWritePausedResponse();
  const maximumEventRows = Math.max(1, Math.min(Number(body?.maximumEventRows || FACILITY_BOOTSTRAP_LIMITS.eventRows), FACILITY_BOOTSTRAP_LIMITS.eventRows));
  const maximumWrites = Math.max(1, Math.min(Number(body?.maximumWrites || FACILITY_BOOTSTRAP_LIMITS.mutationStatements), FACILITY_BOOTSTRAP_LIMITS.mutationStatements));
  const inspection = await inspectFacilityOverviewBootstrap(context.env.ROSTER_DB, { fileId, sourceType });
  if (!inspection.ok) return Response.json(inspection, { status: 400 });
  const estimate = {
    d1ReadStatements: 10,
    maximumRowsReturned: {
      file: 1, doctors: FACILITY_BOOTSTRAP_LIMITS.doctorRows + 1, events: maximumEventRows + 1,
      coverage: 1, rawSource: 1, existingStaff: FACILITY_BOOTSTRAP_LIMITS.existingStaffRows + 1,
      existingCatalog: FACILITY_BOOTSTRAP_LIMITS.existingCatalogRows + 1,
    },
    maximumEstimatedD1RowsExamined: 5 + (FACILITY_BOOTSTRAP_LIMITS.doctorRows + 1) + (maximumEventRows + 1)
      + (FACILITY_BOOTSTRAP_LIMITS.existingStaffRows + 1) + (FACILITY_BOOTSTRAP_LIMITS.existingCatalogRows + 1),
    maximumCompactMutationStatements: maximumWrites,
    maximumEstimatedD1RowsWrittenIncludingIndexes: maximumWrites * 3,
    maximumStatusSummaryRowsWritten: 1,
    maximumEstimatedStatusIndexRowsWritten: 1,
    maximumR2Operations: 0,
    broadSourceScans: 0,
    automaticContinuation: false,
  };
  if (body?.execute !== true) return Response.json({ ...inspection, dryRun: true, estimate });
  if (inspection.compactReady && inspection.statusReady) return Response.json({ ...inspection, dryRun: false, unchanged: true, writes: 0, estimate });
  if (!body?.planRevision || String(body.planRevision) !== inspection.planRevision) {
    return Response.json({ ok: false, stalePlan: true, reason: "bootstrap-plan-changed" }, { status: 409 });
  }
  const result = await refreshFacilityOverviewMaterializationForFile(context.env.ROSTER_DB, fileId, {
    sourceType, maximumEventRows, maximumDoctorRows: FACILITY_BOOTSTRAP_LIMITS.doctorRows,
    maximumExistingStaffRows: FACILITY_BOOTSTRAP_LIMITS.existingStaffRows,
    maximumExistingCatalogRows: FACILITY_BOOTSTRAP_LIMITS.existingCatalogRows, maximumWrites,
  });
  if (!result?.overBudget && Number.isFinite(result?.doctorCount) && Number.isFinite(result?.eventCount)) {
    await upsertRosterFileStatusSummary(context.env.ROSTER_DB, {
      fileId,
      sourceType,
      sourceId: inspection.sourceId,
      name: inspection.name,
      active: true,
      derivedState: "ready",
      expectedDoctorCount: result.doctorCount,
      indexedDoctorCount: result.doctorCount,
      eventCount: result.eventCount,
      size: inspection.size,
      lastModified: inspection.lastModified,
      contentRevision: result.contentRevision,
      rawSourceAvailable: inspection.rawSourceAvailable,
    });
  }
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
