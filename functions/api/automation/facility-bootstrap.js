import { FACILITY_BOOTSTRAP_LIMITS, inspectFacilityOverviewBootstrap, refreshFacilityOverviewMaterializationForFile, upsertRosterFileStatusSummary } from "../../_lib/d1-calendar.js";
import { facilityRolloutPaused } from "../../_lib/facility-rollout.js";
import {
  facilityBootstrapExecutionEnabled,
  facilityBootstrapFileAllowed,
  facilityBootstrapInspectionEnabled,
  facilityMaterializationMaintenanceEnabled,
  rosterWritePausedResponse,
} from "../../_lib/roster-automation-guard.js";

const SOURCE_TYPES = new Set(["mmc", "mch", "ddh", "vhh", "casey"]);
const BOOTSTRAP_PLAN_MAX_AGE_MS = 10 * 60 * 1000;

export async function onRequestPost(context) {
  if (!validToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  const body = await context.request.json().catch(() => ({}));
  const sourceType = String(body?.sourceType || "").trim().toLowerCase();
  const fileId = String(body?.fileId || "").trim();
  const sourceAllowed = configuredSet(context.env.FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST).has(sourceType);
  const fileAllowed = facilityBootstrapFileAllowed(context.env, fileId);
  const execute = body?.execute === true;
  const inspectionAllowed = facilityBootstrapInspectionEnabled(context.env) && sourceAllowed && fileAllowed;
  const executionAllowed = facilityBootstrapExecutionEnabled(context.env)
    && facilityMaterializationMaintenanceEnabled(context.env)
    && !facilityRolloutPaused(context.env)
    && sourceAllowed
    && fileAllowed;
  if (!SOURCE_TYPES.has(sourceType) || !fileId || (execute ? !executionAllowed : !inspectionAllowed)) {
    return rosterWritePausedResponse();
  }
  const planGeneratedAt = execute ? validPlanGeneratedAt(body?.planGeneratedAt) : new Date().toISOString();
  if (!planGeneratedAt) return Response.json({ ok: false, stalePlan: true, reason: "bootstrap-plan-expired" }, { status: 409 });
  const maximumEventRows = boundedLimit(body?.maximumEventRows, FACILITY_BOOTSTRAP_LIMITS.eventRows);
  const maximumWrites = boundedLimit(body?.maximumWrites, FACILITY_BOOTSTRAP_LIMITS.mutationStatements);
  if (!maximumEventRows || !maximumWrites) return rosterWritePausedResponse();
  const inspection = await inspectFacilityOverviewBootstrap(context.env.ROSTER_DB, { fileId, sourceType, planGeneratedAt });
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
  const evidence = {
    deploymentCommit: String(context.env.CF_PAGES_COMMIT_SHA || "unknown"),
    capability: { mode: execute ? "execution" : "inspection", sourceType, fileId },
  };
  if (!execute) return Response.json({ ...inspection, ...evidence, dryRun: true, estimate });
  if (inspection.compactReady && inspection.statusReady) return Response.json({ ...inspection, ...evidence, dryRun: false, unchanged: true, writes: 0, estimate });
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
  return Response.json({ ...inspection, ...evidence, dryRun: false, estimate, result }, { status: result?.overBudget ? 409 : 200 });
}

function boundedLimit(value, fallback) {
  if (value == null || value === "") return fallback;
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 1 && number <= fallback ? number : null;
}

function configuredSet(value) {
  return new Set(String(value || "").split(",").map((item) => item.trim().toLowerCase()).filter(Boolean));
}

function validPlanGeneratedAt(value) {
  const normalized = String(value || "").trim();
  const timestamp = Date.parse(normalized);
  const now = Date.now();
  if (!normalized || !Number.isFinite(timestamp) || timestamp > now + 30_000 || now - timestamp > BOOTSTRAP_PLAN_MAX_AGE_MS) return "";
  return new Date(timestamp).toISOString();
}

function validToken(request, configured) {
  const token = String(configured || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
