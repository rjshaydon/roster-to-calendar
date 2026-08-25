// The full MMC workbook is large and changes frequently around handover.
// Contact allocations must arrive through the doctors-only Office Script JSON
// endpoint instead of downloading and retaining the complete workbook.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(
    context.request,
    context.env.ROSTER_AUTOMATION_TOKEN,
    context.env.DDH_CONTACT_AUTOMATION_TOKEN,
  )) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  return Response.json({
    error: "Full-workbook MMC contact uploads are disabled. Send the doctors-only JSON extract instead.",
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
