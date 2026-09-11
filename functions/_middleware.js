const TRUE_VALUES = new Set(["1", "true", "yes", "on"]);
const STATE_ACTIONS = new Set([
  "acceptInvite", "adminCreateUser", "adminLoadUser", "adminSendInvite", "appendConsoleMessage",
  "calendarStoreStatus", "claimRosterName", "clearFacilityStaffDesignation", "clearUserError",
  "consoleMessages", "decideParserRuleSuggestion", "deleteAccount", "deleteParserExtensionRule",
  "downloadFindmyshiftExceptions", "fetchRawRosterFile", "ignoreUserErrorForever", "listRosterDoctors",
  "listUnresolvedShiftCodes", "listUsers", "loadAccountContext", "loadCalendarEvents", "loadDoctorProfile",
  "loadDoctorProfileImports", "loadImportRefs", "loadImports", "loadInsightImports", "login",
  "queryDoctorProfileFacilityOverviewAccess", "queryFacilityOverviewByStream", "queryFacilityOverviewContactList",
  "queryFacilityOverviewMetadata", "queryFacilityOverviewOnShift", "queryFacilityOverviewStaff",
  "queryFacilityOverviewWorkingTogether", "queryRosterInsights", "queryRosterOverlapDoctors",
  "refreshAutomatedRosterSource", "removeRosterClaim", "removeRosterImports", "repairRosterDailyPresence",
  "replaceActiveRosterFiles", "reportRosterIdentityIssue", "reportUserError", "resetDerivedCalendarFile",
  "resolveAccountClaims", "resolveDoctorAccount", "save", "saveDerivedCalendarFile", "saveDoctorProfile",
  "saveLocalParserExtensionRule", "saveParserExtensionRule", "setAccountRosterClaims",
  "setContactAllocationResolution", "setFacilityStaffDesignation", "setFacilityStaffSeniorityOverride",
  "setFacilityStaffSeniorityOverrides", "setUserDirectorViewEnabled", "setUserFacilityOverviewEnabled",
  "setUserInsightsEnabled", "syncFindmyshift", "syncRosterRepository", "testFindmyshiftConnection",
  "updateAccount", "uploadRawRosterFile",
]);
const DEFAULT_D1_STATEMENT_LIMIT = 64;
const FACILITY_BOOTSTRAP_INSPECTION_D1_STATEMENT_LIMIT = 16;
const FACILITY_BOOTSTRAP_EXECUTION_D1_STATEMENT_LIMIT = 768;
const FACILITY_MATERIALIZATION_INSPECTION_D1_STATEMENT_LIMIT = 16;
const FACILITY_MATERIALIZATION_EXECUTION_D1_STATEMENT_LIMIT = 112;
const ACTION_D1_STATEMENT_LIMITS = Object.freeze({
  login: 32,
  loadAccountContext: 48,
  loadCalendarEvents: 48,
  save: 64,
  updateAccount: 64,
});

export async function onRequest(context) {
  const startedAt = Date.now();
  const url = new URL(context.request.url);
  const requestId = context.request.headers.get("cf-ray") || crypto.randomUUID();
  const action = await safeStateAction(context.request, url.pathname);
  const limit = await requestD1StatementLimit(context.request, url.pathname, action);
  let contained = automationContainmentReason(url.pathname, context.env);
  const d1 = contained ? emptyD1Meter() : createD1Meter(context.env.ROSTER_DB, limit);
  let response;

  if (d1.binding) context.env.ROSTER_DB = d1.binding;

  try {
    response = contained
      ? pausedResponse(contained)
      : await context.next();
    return response;
  } catch (error) {
    if (error?.code === "d1-statement-budget-exceeded") {
      contained = "d1-statement-budget";
      response = Response.json({ error: "This request was stopped by the database safety limit." }, { status: 503 });
      return response;
    }
    throw error;
  } finally {
    const record = {
      event: "api-invocation",
      deployment: String(context.env.CF_PAGES_COMMIT_SHA || "unknown").slice(0, 40),
      requestId: String(requestId).slice(0, 80),
      method: context.request.method,
      route: url.pathname,
      action,
      status: Number(response?.status || 500),
      durationMs: Date.now() - startedAt,
      contained,
      callerClass: callerClass(context.request, url.pathname),
      d1Statements: d1.statementCount,
      d1RowsRead: d1.rowsRead,
      d1RowsWritten: d1.rowsWritten,
      d1MetadataComplete: d1.metadataComplete,
      d1Limit: limit,
    };
    console.log(JSON.stringify(record));
    writeAnalytics(context.env.REQUEST_ANALYTICS, record);
  }
}

async function requestD1StatementLimit(request, pathname, action) {
  if (["/api/automation/facility-bootstrap", "/api/automation/facility-materialize"].includes(pathname) && request.method === "POST") {
    try {
      const body = await request.clone().json();
      if (pathname === "/api/automation/facility-materialize") {
        return body?.execute === true
          ? FACILITY_MATERIALIZATION_EXECUTION_D1_STATEMENT_LIMIT
          : FACILITY_MATERIALIZATION_INSPECTION_D1_STATEMENT_LIMIT;
      }
      return body?.execute === true
        ? FACILITY_BOOTSTRAP_EXECUTION_D1_STATEMENT_LIMIT
        : FACILITY_BOOTSTRAP_INSPECTION_D1_STATEMENT_LIMIT;
    } catch {
      return pathname === "/api/automation/facility-materialize"
        ? FACILITY_MATERIALIZATION_INSPECTION_D1_STATEMENT_LIMIT
        : FACILITY_BOOTSTRAP_INSPECTION_D1_STATEMENT_LIMIT;
    }
  }
  return ACTION_D1_STATEMENT_LIMITS[action] || DEFAULT_D1_STATEMENT_LIMIT;
}

function writeAnalytics(dataset, record) {
  if (!dataset?.writeDataPoint) return;
  try {
    dataset.writeDataPoint({
      indexes: [record.deployment],
      blobs: [record.route, record.action, String(record.status), record.contained, record.callerClass,
        record.requestId, String(record.d1MetadataComplete)],
      doubles: [record.durationMs, record.d1Statements, record.d1RowsRead, record.d1RowsWritten, record.d1Limit],
    });
  } catch (error) {
    console.warn(JSON.stringify({ event: "api-attribution-write-failed", error: String(error?.message || error).slice(0, 160) }));
  }
}

function emptyD1Meter() {
  return { binding: null, statementCount: 0, rowsRead: 0, rowsWritten: 0, metadataComplete: true };
}

async function safeStateAction(request, pathname) {
  if (pathname !== "/api/state" || request.method !== "POST") return "";
  try {
    const body = await request.clone().json();
    const action = String(body?.action || "login");
    return STATE_ACTIONS.has(action) ? action : "unknown-action";
  } catch {
    return "invalid-json";
  }
}

function automationContainmentReason(pathname, env = {}) {
  if (!pathname.startsWith("/api/automation/")) return "";
  if (["/api/automation/contact-list", "/api/automation/contact-list-binary", "/api/automation/contact-list-extract"].includes(pathname)) {
    return enabled(env.CONTACT_AUTOMATION_WRITES_ENABLED) ? "" : "contact-automation-paused";
  }
  if (pathname === "/api/automation/facility-bootstrap") {
    return enabled(env.FACILITY_BOOTSTRAP_INSPECTION_ENABLED) || enabled(env.FACILITY_BOOTSTRAP_EXECUTION_ENABLED)
      ? "" : "facility-bootstrap-paused";
  }
  if (pathname === "/api/automation/facility-materialize") {
    return enabled(env.ROSTER_ADVANCED_MAINTENANCE_ENABLED) ? "" : "facility-materialization-paused";
  }
  return enabled(env.ROSTER_AUTOMATION_WRITES_ENABLED) ? "" : "roster-automation-paused";
}

function enabled(value) {
  return TRUE_VALUES.has(String(value || "").trim().toLowerCase());
}

function pausedResponse(reason) {
  return Response.json({ ok: false, status: "paused", reason }, {
    status: 503,
    headers: { "Cache-Control": "no-store", "Retry-After": "3600" },
  });
}

function callerClass(request, pathname) {
  if (pathname.startsWith("/api/automation/")) return "automation";
  const agent = String(request.headers.get("user-agent") || "").toLowerCase();
  if (agent.includes("roster-queue-watchdog")) return "watchdog";
  if (agent.includes("github-actions") || agent.includes("actions/")) return "github-actions";
  return "interactive-or-unknown";
}

function createD1Meter(database, limit) {
  const state = { statementCount: 0, rowsRead: 0, rowsWritten: 0, metadataComplete: true };
  const originals = new WeakMap();
  const before = (count = 1) => {
    if (state.statementCount + count > limit) {
      const error = new Error("D1 statement budget exceeded.");
      error.code = "d1-statement-budget-exceeded";
      throw error;
    }
    state.statementCount += count;
  };
  const record = (result) => {
    const meta = result?.meta;
    if (!meta || !Number.isFinite(Number(meta.rows_read)) || !Number.isFinite(Number(meta.rows_written))) {
      state.metadataComplete = false;
      return result;
    }
    state.rowsRead += Number(meta.rows_read);
    state.rowsWritten += Number(meta.rows_written);
    return result;
  };
  const wrapStatement = (statement) => {
    if (!statement || typeof statement !== "object") return statement;
    const wrapped = new Proxy(statement, {
      get(target, property) {
        if (property === "bind") return (...values) => wrapStatement(target.bind(...values));
        if (["all", "run", "raw"].includes(property)) return async (...args) => {
          before();
          return record(await target[property](...args));
        };
        if (property === "first") return async (...args) => {
          before();
          state.metadataComplete = false;
          return target.first(...args);
        };
        return Reflect.get(target, property, target);
      },
    });
    originals.set(wrapped, statement);
    return wrapped;
  };
  if (!database?.prepare) return { ...state, binding: database, get statementCount() { return state.statementCount; }, get rowsRead() { return state.rowsRead; }, get rowsWritten() { return state.rowsWritten; }, get metadataComplete() { return state.metadataComplete; } };
  const binding = new Proxy(database, {
    get(target, property) {
      if (property === "prepare") return (sql) => {
        const statement = target.prepare(sql);
        return wrapStatement(statement);
      };
      if (property === "batch") return async (statements) => {
        const items = Array.from(statements || []);
        before(items.length);
        const results = await target.batch(items.map((item) => originals.get(item) || item));
        results.forEach(record);
        return results;
      };
      if (property === "exec") return async (...args) => {
        before();
        return record(await target.exec(...args));
      };
      return Reflect.get(target, property, target);
    },
  });
  return {
    binding,
    get statementCount() { return state.statementCount; },
    get rowsRead() { return state.rowsRead; },
    get rowsWritten() { return state.rowsWritten; },
    get metadataComplete() { return state.metadataComplete; },
  };
}
