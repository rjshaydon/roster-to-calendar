const ENABLED_VALUES = new Set(["1", "true", "yes", "on"]);
const LOOPBACK_HOSTS = new Set(["127.0.0.1", "localhost", "[::1]"]);

export function localOnlyEnabled(env = {}) {
  return ENABLED_VALUES.has(String(env.LOCAL_ONLY || "").trim().toLowerCase());
}

export function localFeatureDisabledResponse(env, feature = "This operation") {
  if (!localOnlyEnabled(env)) return null;
  return Response.json({
    ok: false,
    status: "local-disabled",
    code: "LOCAL_ONLY",
    error: `${String(feature || "This operation").trim()} is disabled in local development. No external service was contacted.`,
  }, { status: 503 });
}

export function assertLocalOutboundAllowed(env, input, label = "External request") {
  if (!localOnlyEnabled(env)) return;
  const rawUrl = input instanceof Request ? input.url : input instanceof URL ? input.href : String(input || "");
  let url;
  try {
    url = new URL(rawUrl);
  } catch {
    throw localOutboundError(label, "an invalid or relative URL");
  }
  if (!LOOPBACK_HOSTS.has(url.hostname.toLowerCase())) {
    throw localOutboundError(label, url.hostname || "a non-loopback destination");
  }
}

export async function guardedFetch(env, input, init, options = {}) {
  assertLocalOutboundAllowed(env, input, options.label || "External request");
  return await fetch(input, init);
}

function localOutboundError(label, destination) {
  const error = new Error(`${String(label || "External request").trim()} is disabled in local development (${destination}). No external service was contacted.`);
  error.code = "LOCAL_OUTBOUND_BLOCKED";
  return error;
}
