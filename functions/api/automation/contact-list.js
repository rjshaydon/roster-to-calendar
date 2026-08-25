// Compatibility route for the former full-workbook automation. Production
// and Preview accept contact allocations only through the small JSON endpoint.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(
    context.request,
    context.env.ROSTER_AUTOMATION_TOKEN,
    context.env.DDH_CONTACT_AUTOMATION_TOKEN,
  )) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  return Response.json({
    error: "Full-workbook contact uploads are disabled. Send the clinicians-only JSON extract instead.",
    endpoint: "/api/automation/contact-list-extract",
  }, { status: 410 });
}

function hasValidAutomationToken(request, ...configuredTokens) {
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!provided) return false;
  return configuredTokens.some((configuredToken) => {
    const token = String(configuredToken || "");
    if (!token || token.length !== provided.length) return false;
    let mismatch = 0;
    for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
    return mismatch === 0;
  });
}
