import { hasCalendarDb, loadRawRosterFile, loadRosterSyncRun } from "../../_lib/d1-calendar.js";

export async function onRequestGet(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env) || !context.env.ROSTER_FILES?.get) {
    return Response.json({ error: "Roster storage is unavailable." }, { status: 503 });
  }
  const runId = String(new URL(context.request.url).searchParams.get("runId") || "").trim();
  const run = await loadRosterSyncRun(context.env.ROSTER_DB, runId);
  if (!run?.fileId) return Response.json({ error: "Queued roster was not found." }, { status: 404 });
  const raw = await loadRawRosterFile(context.env.ROSTER_DB, run.fileId);
  let body = null;
  let size = raw.size;
  let contentType = raw.type || "application/octet-stream";
  if (raw?.objectKey) {
    const object = await context.env.ROSTER_FILES.get(raw.objectKey);
    if (object?.body) {
      body = object.body;
      size = raw.size || object.size;
      contentType = raw.type || object.httpMetadata?.contentType || contentType;
    }
  }
  if (!body && raw?.dataUrl) {
    const base64 = String(raw.dataUrl).split(",")[1] || "";
    if (base64) body = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
  }
  if (!body) return Response.json({ error: "Queued roster object was not found." }, { status: 404 });
  const safeName = String(raw.name || "roster.xlsx").replace(/[\r\n"]/g, "_");
  return new Response(body, {
    headers: {
      "Content-Type": contentType,
      "Content-Length": String(size || ""),
      "Content-Disposition": `attachment; filename="${safeName}"`,
      "Cache-Control": "private, no-store",
    },
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
