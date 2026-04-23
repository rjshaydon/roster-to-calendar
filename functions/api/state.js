const CREATOR_EMAIL = "rhaydon@gmail.com";

export async function onRequestGet(context) {
  const email = normalizeEmail(new URL(context.request.url).searchParams.get("email"));
  if (!email) {
    return Response.json({ error: "Email address is required." }, { status: 400 });
  }
  if (!context.env.ROSTER_STORE) {
    return Response.json({
      cloudAvailable: false,
      role: roleForEmail(email),
      state: null,
    });
  }
  const state = await context.env.ROSTER_STORE.get(storageKey(email), "json");
  return Response.json({
    cloudAvailable: true,
    role: roleForEmail(email),
    state: state || null,
  });
}

export async function onRequestPost(context) {
  const body = await context.request.json().catch(() => null);
  const email = normalizeEmail(body?.email);
  if (!email) {
    return Response.json({ error: "Email address is required." }, { status: 400 });
  }
  if (!context.env.ROSTER_STORE) {
    return Response.json({ error: "Cloud storage is not configured." }, { status: 503 });
  }
  const state = sanitizeState(body?.state);
  await context.env.ROSTER_STORE.put(storageKey(email), JSON.stringify({
    ...state,
    email,
    role: roleForEmail(email),
    updatedAt: new Date().toISOString(),
  }));
  return Response.json({ ok: true, role: roleForEmail(email) });
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function roleForEmail(email) {
  return email === CREATOR_EMAIL ? "creator" : "user";
}

function storageKey(email) {
  return `account:${email}`;
}

function sanitizeState(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    version: 1,
    imports: Array.isArray(input.imports) ? input.imports : [],
    session: input.session && typeof input.session === "object" ? input.session : {},
  };
}
