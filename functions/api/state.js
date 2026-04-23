const CREATOR_EMAIL = "rhaydon@gmail.com";

export async function onRequestGet(context) {
  return Response.json({ error: "Use POST for account requests." }, { status: 405 });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json().catch(() => null);
    const email = normalizeEmail(body?.email);
    const password = String(body?.password || "");
    const action = String(body?.action || "login");
    if (!email) {
      return Response.json({ error: "Email address is required." }, { status: 400 });
    }
    if (!context.env.ROSTER_STORE) {
      return Response.json({ error: "Cloud storage is not configured." }, { status: 503 });
    }
    if (!password) {
      return Response.json({ error: "Password is required." }, { status: 400 });
    }

    if (action === "login") {
      const account = await loadOrCreateAccount(context.env.ROSTER_STORE, email, password);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: account.created,
        role: account.role,
        state: account.state,
      });
    }

    const account = await verifyAccount(context.env.ROSTER_STORE, email, password);
    if (action === "listUsers") {
      if (account.role !== "creator") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      return Response.json({ ok: true, users: await listUsers(context.env.ROSTER_STORE) });
    }

    if (action === "save") {
      const state = sanitizeState(body?.state);
      await context.env.ROSTER_STORE.put(storageKey(email), JSON.stringify({
        ...account.record,
        email,
        role: account.role,
        state,
        updatedAt: new Date().toISOString(),
      }));
      return Response.json({ ok: true, role: account.role });
    }

    return Response.json({ error: "Unsupported account action." }, { status: 400 });
  } catch (error) {
    const message = error.message || "Account request failed.";
    const status = message === "Incorrect password." || message === "Account not found." ? 401 : 400;
    return Response.json({ error: message }, { status });
  }
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

async function loadOrCreateAccount(store, email, password) {
  const existing = await store.get(storageKey(email), "json");
  if (!existing) {
    const passwordRecord = await hashPassword(password);
    const record = {
      email,
      role: roleForEmail(email),
      ...passwordRecord,
      state: sanitizeState(null),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    await store.put(storageKey(email), JSON.stringify(record));
    return {
      created: true,
      role: record.role,
      state: record.state,
    };
  }
  if (!existing.passwordHash || !existing.passwordSalt) {
    const passwordRecord = await hashPassword(password);
    const role = existing.role || roleForEmail(email);
    const state = role === "creator" ? sanitizeState(existing.state) : sanitizeState(null);
    const upgraded = {
      ...existing,
      state,
      ...passwordRecord,
      updatedAt: new Date().toISOString(),
    };
    await store.put(storageKey(email), JSON.stringify(upgraded));
    return {
      created: false,
      role,
      state,
    };
  }
  const ok = await verifyPassword(password, existing.passwordSalt, existing.passwordHash);
  if (!ok) {
    throw new Error("Incorrect password.");
  }
  return {
    created: false,
    role: existing.role || roleForEmail(email),
    state: sanitizeState(existing.state),
  };
}

async function verifyAccount(store, email, password) {
  const record = await store.get(storageKey(email), "json");
  if (!record?.passwordHash || !record?.passwordSalt) {
    throw new Error("Account not found.");
  }
  const ok = await verifyPassword(password, record.passwordSalt, record.passwordHash);
  if (!ok) {
    throw new Error("Incorrect password.");
  }
  return {
    record,
    role: record.role || roleForEmail(email),
  };
}

async function listUsers(store) {
  const result = await store.list({ prefix: "account:" });
  return (result.keys || [])
    .map((item) => item.name.replace(/^account:/, ""))
    .sort();
}

async function hashPassword(password, salt = randomSalt()) {
  const passwordSalt = salt;
  const passwordHash = await sha256(`${passwordSalt}:${password}`);
  return { passwordSalt, passwordHash };
}

async function verifyPassword(password, salt, expectedHash) {
  const { passwordHash } = await hashPassword(password, salt);
  return passwordHash === expectedHash;
}

function randomSalt() {
  const values = new Uint8Array(16);
  crypto.getRandomValues(values);
  return [...values].map((value) => value.toString(16).padStart(2, "0")).join("");
}

async function sha256(value) {
  const bytes = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(hash)].map((item) => item.toString(16).padStart(2, "0")).join("");
}

function sanitizeState(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    version: 1,
    imports: Array.isArray(input.imports) ? input.imports : [],
    session: input.session && typeof input.session === "object" ? input.session : {},
  };
}
