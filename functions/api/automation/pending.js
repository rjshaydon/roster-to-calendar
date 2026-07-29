import { hasCalendarDb, listQueuedRosterSyncRuns } from "../../_lib/d1-calendar.js";

export async function onRequestGet(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  const url = new URL(context.request.url);
  const limit = Number(url.searchParams.get("limit") || 4);
  const runs = await listQueuedRosterSyncRuns(context.env.ROSTER_DB, limit);
  return Response.json({
    ok: true,
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
