import { hasCalendarDb } from "../../_lib/d1-calendar.js";

export async function onRequestGet(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });

  const rows = await context.env.ROSTER_DB
    .prepare("SELECT rule_json FROM parser_rules WHERE scope = 'global' ORDER BY source_type, seniority, code")
    .all()
    .catch(() => ({ results: [] }));
  const parserExtensions = { mmc: [], ddh: [], casey: [], mch: [] };
  for (const row of rows.results || []) {
    try {
      const rule = JSON.parse(String(row.rule_json || "{}"));
      const source = String(rule?.source || "").trim().toLowerCase();
      if (parserExtensions[source] && rule?.code) parserExtensions[source].push(rule);
    } catch {
      // A malformed stored rule must not prevent an otherwise valid roster sync.
    }
  }
  return Response.json({ ok: true, parserExtensions });
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
