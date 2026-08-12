// Normalisation deliberately excludes parser-internal ids so browser and
// automatic parsing can be compared by calendar behaviour.
export function normalizeParserResult(result = {}) {
  const normalise = (value) => JSON.parse(JSON.stringify(value || null));
  const sortByJson = (items) => (Array.isArray(items) ? items : [])
    .map(normalise)
    .sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
  return {
    doctors: sortByJson(result.doctors),
    aliases: sortByJson(result.aliases),
    memberships: sortByJson(result.memberships),
    events: sortByJson(result.events),
    issues: sortByJson(result.issues),
  };
}

export function parserResultDelta(browserResult, serverResult) {
  const browser = normalizeParserResult(browserResult);
  const server = normalizeParserResult(serverResult);
  const delta = {};
  for (const field of Object.keys(browser)) {
    const left = JSON.stringify(browser[field]);
    const right = JSON.stringify(server[field]);
    if (left !== right) delta[field] = { browser: browser[field], server: server[field] };
  }
  return delta;
}

export function unresolvedCodeSummary(issues = []) {
  const occurrences = (Array.isArray(issues) ? issues : []).filter((issue) => String(issue?.status || "").toLowerCase() === "unknown");
  const distinct = new Set(occurrences.map((issue) => [
    String(issue?.source || issue?.sourceType || "").toLowerCase(),
    String(issue?.seniority || "Unknown"),
    String(issue?.rawValue || "").trim().toUpperCase(),
  ].join("|")));
  return { occurrences: occurrences.length, distinctCodes: distinct.size };
}
