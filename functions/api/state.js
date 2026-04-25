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
    const mode = String(body?.mode || "login");
    const realName = String(body?.realName || "").trim();
    const targetEmail = normalizeEmail(body?.targetEmail);
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
      const account = await loadOrCreateAccount(context.env.ROSTER_STORE, email, password, { mode, realName });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: account.created,
        role: account.role,
        realName: account.realName,
        state: account.state,
      });
    }

    const account = await verifyAccount(context.env.ROSTER_STORE, email, password);
    if (action === "adminLoadUser") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const target = await loadAccountRecord(context.env.ROSTER_STORE, targetEmail);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: target.role || roleForEmail(targetEmail),
        realName: target.realName || "",
        state: sanitizeState(target.state),
      });
    }

    if (action === "listUsers") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      return Response.json({ ok: true, users: await listUsers(context.env.ROSTER_STORE) });
    }

    if (action === "deleteAccount") {
      const deleteEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      if (deleteEmail === CREATOR_EMAIL) {
        return Response.json({ error: "The creator account cannot be deleted." }, { status: 400 });
      }
      if (deleteEmail !== email && account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      await context.env.ROSTER_STORE.delete(storageKey(deleteEmail));
      return Response.json({ ok: true, deletedEmail: deleteEmail });
    }

    if (action === "save") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, saveEmail);
      const targetRole = targetRecord.role || roleForEmail(saveEmail);
      const state = sanitizeState(body?.state);
      await context.env.ROSTER_STORE.put(storageKey(saveEmail), JSON.stringify({
        ...targetRecord,
        email: saveEmail,
        role: targetRole,
        realName: targetRecord.realName || "",
        state,
        updatedAt: new Date().toISOString(),
      }));
      return Response.json({ ok: true, role: targetRole });
    }

    return Response.json({ error: "Unsupported account action." }, { status: 400 });
  } catch (error) {
    const message = error.message || "Account request failed.";
    const status = message === "Incorrect password." || message.startsWith("Account not found") ? 401 : 400;
    return Response.json({ error: message }, { status });
  }
}

async function loadAccountRecord(store, email) {
  if (!email) {
    throw new Error("Target account is required.");
  }
  const record = await store.get(storageKey(email), "json");
  if (!record) {
    throw new Error("Account not found.");
  }
  return record;
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

async function loadOrCreateAccount(store, email, password, options = {}) {
  const mode = options.mode || "login";
  const realName = String(options.realName || "").trim();
  const existing = await store.get(storageKey(email), "json");
  if (!existing) {
    if (mode !== "create" || !realName) {
      throw new Error("Account not found. Create an account first.");
    }
    const passwordRecord = await hashPassword(password);
    const record = {
      email,
      realName,
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
      realName: record.realName,
      state: record.state,
    };
  }
  if (mode === "create") {
    throw new Error("An account already exists for that email. Use log in.");
  }
  if (!existing.passwordHash || !existing.passwordSalt) {
    const passwordRecord = await hashPassword(password);
    const role = existing.role || roleForEmail(email);
    const state = role === "creator" ? sanitizeState(existing.state) : sanitizeState(null);
    const upgraded = {
      ...existing,
      realName: existing.realName || realName || "",
      state,
      ...passwordRecord,
      updatedAt: new Date().toISOString(),
    };
    await store.put(storageKey(email), JSON.stringify(upgraded));
    return {
      created: false,
      role,
      realName: upgraded.realName || "",
      state,
    };
  }
  const ok = await verifyPassword(password, existing.passwordSalt, existing.passwordHash);
  if (!ok) {
    throw new Error("Incorrect password.");
  }
  let updated = existing;
  if (realName && !existing.realName) {
    updated = {
      ...existing,
      realName,
      updatedAt: new Date().toISOString(),
    };
    await store.put(storageKey(email), JSON.stringify(updated));
  }
  return {
    created: false,
    role: updated.role || roleForEmail(email),
    realName: updated.realName || "",
    state: sanitizeState(updated.state),
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
  const users = await Promise.all((result.keys || []).map(async (item) => {
    const email = item.name.replace(/^account:/, "");
    const record = await store.get(item.name, "json").catch(() => null);
    return {
      email,
      realName: String(record?.realName || "").trim(),
      role: record?.role || roleForEmail(email),
      createdAt: record?.createdAt || "",
      updatedAt: record?.updatedAt || "",
    };
  }));
  return users.sort((a, b) => a.email.localeCompare(b.email));
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
