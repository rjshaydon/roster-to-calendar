import { inspectImportRecord, normalizeRosterName } from "../_lib/roster.js";

const CREATOR_EMAIL = "rhaydon@gmail.com";
const REPOSITORY_INDEX_KEY = "repository:index";
const REPOSITORY_FILE_PREFIX = "repository:file:";
const DOCTOR_PROFILE_PREFIX = "doctor-profile:";

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
      await hydrateRepositoryFromExistingAccounts(context.env.ROSTER_STORE);
      const account = await loadOrCreateAccount(context.env.ROSTER_STORE, email, password, { mode, realName });
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, account.record);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: account.created,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
      });
    }

    const account = await verifyAccount(context.env.ROSTER_STORE, email, password);
    if (action === "adminCreateUser") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const targetPassword = String(body?.targetPassword || "");
      const targetRealName = String(body?.targetRealName || body?.realName || "").trim();
      if (!targetEmail) {
        return Response.json({ error: "New account email is required." }, { status: 400 });
      }
      if (!targetRealName) {
        return Response.json({ error: "New account real name is required." }, { status: 400 });
      }
      if (!targetPassword) {
        return Response.json({ error: "New account password is required." }, { status: 400 });
      }
      await hydrateRepositoryFromExistingAccounts(context.env.ROSTER_STORE);
      const created = await loadOrCreateAccount(context.env.ROSTER_STORE, targetEmail, targetPassword, {
        mode: "create",
        realName: targetRealName,
      });
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, created.record);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: true,
        user: {
          email: targetEmail,
          realName: prepared.realName,
          role: prepared.role,
          sites: [...new Set(sanitizeClaims(prepared.claims).map((claim) => claim.sourceType.toUpperCase()))].sort(),
          claims: prepared.claims,
          createdAt: created.record.createdAt || "",
          updatedAt: created.record.updatedAt || "",
        },
      });
    }

    if (action === "adminLoadUser") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      await hydrateRepositoryFromExistingAccounts(context.env.ROSTER_STORE);
      const target = await loadAccountRecord(context.env.ROSTER_STORE, targetEmail);
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, target);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
      });
    }

    if (action === "claimRosterName") {
      const claimEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = claimEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, claimEmail);
      const index = await loadRepositoryIndex(context.env.ROSTER_STORE);
      const claim = findRepositoryDoctor(index, body?.claim);
      if (!claim) {
        return Response.json({ error: "Roster name was not found in the repository." }, { status: 400 });
      }
      const claims = mergeClaims(targetRecord.claims, [{ ...claim, matchedAt: new Date().toISOString() }]);
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: (await repositoryImportsForClaims(context.env.ROSTER_STORE, index, claims)).map(repositoryImportRef),
      };
      const updated = {
        ...targetRecord,
        email: claimEmail,
        claims,
        state,
        updatedAt: new Date().toISOString(),
      };
      await context.env.ROSTER_STORE.put(storageKey(claimEmail), JSON.stringify(updated));
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, updated);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
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
      const repository = await upsertStateImports(context.env.ROSTER_STORE, state.imports, saveEmail);
      state.imports = repository.refs;
      const claims = targetRole === "creator" || targetRole === "owner"
        ? sanitizeClaims(targetRecord.claims)
        : mergeClaims(targetRecord.claims, matchRepositoryClaims(repository.index, targetRecord.realName || ""));
      await context.env.ROSTER_STORE.put(storageKey(saveEmail), JSON.stringify({
        ...targetRecord,
        email: saveEmail,
        role: targetRole,
        realName: targetRecord.realName || "",
        claims,
        state,
        updatedAt: new Date().toISOString(),
      }));
      return Response.json({ ok: true, role: targetRole, claims });
    }

    if (action === "loadDoctorProfile") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const profileId = String(body?.profileId || "").trim();
      if (!profileId) {
        return Response.json({ error: "Doctor profile is required." }, { status: 400 });
      }
      return Response.json({
        ok: true,
        cloudAvailable: true,
        profile: await loadDoctorProfileRecord(context.env.ROSTER_STORE, profileId),
      });
    }

    if (action === "saveDoctorProfile") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const profileId = String(body?.profileId || "").trim();
      const doctorKey = normalizeRosterName(body?.doctorKey || "");
      const displayName = String(body?.displayName || "").trim();
      const sourceTypes = sanitizeSourceTypes(body?.sourceTypes);
      const state = sanitizeState(body?.state);
      if (!profileId || !doctorKey || !displayName || !sourceTypes.length) {
        return Response.json({ error: "Doctor profile details are incomplete." }, { status: 400 });
      }
      if (!hasDoctorProfileState(state)) {
        await context.env.ROSTER_STORE.delete(doctorProfileKey(profileId));
        return Response.json({ ok: true, deleted: true });
      }
      const existing = await loadDoctorProfileRecord(context.env.ROSTER_STORE, profileId);
      const next = {
        profileId,
        doctorKey,
        displayName,
        sourceTypes,
        state,
        createdAt: existing?.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      await context.env.ROSTER_STORE.put(doctorProfileKey(profileId), JSON.stringify(next));
      return Response.json({ ok: true, profile: next });
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

async function loadDoctorProfileRecord(store, profileId) {
  if (!profileId) return null;
  return sanitizeDoctorProfile(await store.get(doctorProfileKey(profileId), "json").catch(() => null));
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
    const canBootstrapCreator = email === CREATOR_EMAIL && mode === "login";
    if (!canBootstrapCreator && (mode !== "create" || !realName)) {
      throw new Error("Account not found. Create an account first.");
    }
    const passwordRecord = await hashPassword(password);
    const record = {
      email,
      realName: realName || (email === CREATOR_EMAIL ? "Richard Haydon" : ""),
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
      record,
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
      record: upgraded,
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
    record: updated,
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
    const claims = sanitizeClaims(record?.claims);
    return {
      email,
      realName: String(record?.realName || "").trim(),
      role: record?.role || roleForEmail(email),
      sites: [...new Set(claims.map((claim) => claim.sourceType.toUpperCase()))].sort(),
      claims,
      createdAt: record?.createdAt || "",
      updatedAt: record?.updatedAt || "",
    };
  }));
  return users.sort((a, b) => a.email.localeCompare(b.email));
}

async function prepareAccountResponse(store, record) {
  const role = record.role || roleForEmail(record.email);
  const index = await loadRepositoryIndex(store);
  let claims = sanitizeClaims(record.claims);
  let nameMatches = [];
  let state = sanitizeState(record.state);

  if (role !== "creator" && role !== "owner") {
    const matchedClaims = matchRepositoryClaims(index, record.realName || "");
    const merged = mergeClaims(claims, matchedClaims);
    nameMatches = matchedClaims.filter((claim) => !claims.some((existing) => sameClaim(existing, claim)));
    claims = merged;
    state = {
      ...state,
      imports: await repositoryImportsForClaims(store, index, claims),
    };
    state = await mergeLinkedDoctorProfiles(store, state, claims, record.email);
    if (nameMatches.length || JSON.stringify(claims) !== JSON.stringify(sanitizeClaims(record.claims))) {
      await store.put(storageKey(record.email), JSON.stringify({
        ...record,
        claims,
        state: {
          ...sanitizeState(record.state),
          imports: state.imports.map(repositoryImportRef),
        },
        updatedAt: new Date().toISOString(),
      }));
    }
  } else {
    const imported = await upsertStateImports(store, state.imports, record.email);
    const stateWithRefs = { ...state, imports: imported.refs };
    if (imported.changed || importsChanged(state.imports, imported.refs)) {
      state = stateWithRefs;
      await store.put(storageKey(record.email), JSON.stringify({
        ...record,
        state,
        updatedAt: new Date().toISOString(),
      }));
    } else {
      state = stateWithRefs;
    }
    state = {
      ...state,
      imports: await resolveStateImports(store, state.imports),
    };
  }

  return {
    role,
    realName: record.realName || "",
    state,
    claims,
    nameMatches,
    availableDoctors: await repositoryDoctorCandidates(store, index),
  };
}

async function hydrateRepositoryFromExistingAccounts(store) {
  let index = await loadRepositoryIndex(store);
  const result = await store.list({ prefix: "account:" });
  let changed = false;
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    if (!record?.state?.imports?.some((importItem) => importItem?.dataUrl)) continue;
    const upserted = await upsertImportsIntoRepository(store, index, record.state.imports, record.email || item.name.replace(/^account:/, ""));
    index = upserted.index;
    changed = changed || upserted.changed;
    const refs = record.state.imports.map((importItem) => {
      const repoId = importItem.repoId || importItem.repositoryId || upserted.idByOriginalId.get(importItem.id) || upserted.idByDataUrl.get(importItem.dataUrl);
      return repoId ? repositoryImportRef(index.files.find((file) => file.id === repoId) || { ...importItem, id: repoId }) : repositoryImportRef(importItem);
    });
    if (importsChanged(record.state.imports, refs)) {
      await store.put(item.name, JSON.stringify({
        ...record,
        state: {
          ...sanitizeState(record.state),
          imports: refs,
        },
        updatedAt: new Date().toISOString(),
      }));
    }
  }
  if (changed) await saveRepositoryIndex(store, index);
}

async function upsertStateImports(store, imports, uploadedBy) {
  let index = await loadRepositoryIndex(store);
  const upserted = await upsertImportsIntoRepository(store, index, imports, uploadedBy);
  index = upserted.index;
  if (upserted.changed) await saveRepositoryIndex(store, index);
  return {
    index,
    refs: (imports || []).map((item) => {
      const repoId = item.repoId || item.repositoryId || upserted.idByOriginalId.get(item.id) || upserted.idByDataUrl.get(item.dataUrl);
      return repoId ? repositoryImportRef(index.files.find((file) => file.id === repoId) || { ...item, id: repoId }) : repositoryImportRef(item);
    }),
    changed: upserted.changed,
  };
}

async function upsertImportsIntoRepository(store, index, imports = [], uploadedBy = "") {
  const idByOriginalId = new Map();
  const idByDataUrl = new Map();
  let changed = false;
  for (const item of imports || []) {
    if (!item?.dataUrl) {
      const repoId = item?.repoId || item?.repositoryId || item?.id || "";
      if (repoId) idByOriginalId.set(item.id, repoId);
      continue;
    }
    const contentHash = await sha256(item.dataUrl);
    const repoId = `sha256-${contentHash}`;
    idByOriginalId.set(item.id, repoId);
    idByDataUrl.set(item.dataUrl, repoId);
    const existing = index.files.find((file) => file.id === repoId);
    let inspected = {
      sourceType: String(item.sourceType || "").toLowerCase(),
      doctors: sanitizeRepositoryDoctors(item.doctors),
    };
    if (!inspected.sourceType || !inspected.doctors.length) {
      try {
        inspected = await inspectImportRecord(item);
      } catch {
        inspected = { sourceType: item.sourceType || "unknown", doctors: [] };
      }
    }
    const meta = {
      id: repoId,
      name: String(item.name || existing?.name || "roster.xlsx"),
      size: Number(item.size || existing?.size || 0),
      lastModified: Number(item.lastModified || existing?.lastModified || 0),
      addedAt: String(item.addedAt || existing?.addedAt || new Date().toISOString()),
      uploadedAt: existing?.uploadedAt || new Date().toISOString(),
      uploadedBy: existing?.uploadedBy || uploadedBy,
      sourceType: inspected.sourceType || item.sourceType || existing?.sourceType || "unknown",
      doctors: inspected.doctors?.length ? inspected.doctors : sanitizeRepositoryDoctors(existing?.doctors),
      active: existing?.active !== false,
    };
    if (!existing || JSON.stringify(existing) !== JSON.stringify(meta)) {
      if (existing) {
        index.files = index.files.map((file) => file.id === repoId ? meta : file);
      } else {
        index.files.push(meta);
      }
      changed = true;
    }
    if (!existing) {
      await store.put(repositoryFileKey(repoId), JSON.stringify({
        ...meta,
        type: item.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        dataUrl: item.dataUrl,
      }));
    }
  }
  index.files.sort((left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || left.name.localeCompare(right.name));
  return { index, changed, idByOriginalId, idByDataUrl };
}

async function loadRepositoryIndex(store) {
  const raw = await store.get(REPOSITORY_INDEX_KEY, "json").catch(() => null);
  return {
    version: 1,
    files: Array.isArray(raw?.files) ? raw.files.map(sanitizeRepositoryFile).filter(Boolean) : [],
  };
}

async function saveRepositoryIndex(store, index) {
  await store.put(REPOSITORY_INDEX_KEY, JSON.stringify({
    version: 1,
    files: (index.files || []).map(sanitizeRepositoryFile).filter(Boolean),
    updatedAt: new Date().toISOString(),
  }));
}

function sanitizeRepositoryFile(file) {
  if (!file?.id) return null;
  return {
    id: String(file.id),
    name: String(file.name || "roster.xlsx"),
    size: Number(file.size || 0),
    lastModified: Number(file.lastModified || 0),
    addedAt: String(file.addedAt || ""),
    uploadedAt: String(file.uploadedAt || ""),
    uploadedBy: normalizeEmail(file.uploadedBy || ""),
    sourceType: String(file.sourceType || "unknown").toLowerCase(),
    doctors: sanitizeRepositoryDoctors(file.doctors),
    active: file.active !== false,
  };
}

function sanitizeRepositoryDoctors(doctors) {
  if (!Array.isArray(doctors)) return [];
  return doctors
    .map((doctor) => ({
      key: normalizeRosterName(doctor?.key || ""),
      displayName: String(doctor?.displayName || "").trim(),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
    }))
    .filter((doctor) => doctor.key && doctor.displayName);
}

function repositoryFileKey(id) {
  return `${REPOSITORY_FILE_PREFIX}${id}`;
}

function doctorProfileKey(profileId) {
  return `${DOCTOR_PROFILE_PREFIX}${profileId}`;
}

function repositoryImportRef(item) {
  return {
    repoId: item.repoId || item.repositoryId || item.id,
    id: item.repoId || item.repositoryId || item.id,
    name: item.name || "roster.xlsx",
    size: Number(item.size || 0),
    lastModified: Number(item.lastModified || 0),
    addedAt: item.addedAt || "",
    sourceType: item.sourceType || "pending",
  };
}

function importsChanged(current = [], next = []) {
  return JSON.stringify((current || []).map(repositoryImportRef)) !== JSON.stringify((next || []).map(repositoryImportRef));
}

async function resolveStateImports(store, imports = []) {
  const resolved = [];
  for (const ref of imports || []) {
    const repoId = ref.repoId || ref.repositoryId || ref.id;
    const stored = repoId ? await store.get(repositoryFileKey(repoId), "json").catch(() => null) : null;
    if (stored?.dataUrl) {
      resolved.push({
        id: repoId,
        repoId,
        name: stored.name || ref.name,
        size: stored.size || ref.size || 0,
        lastModified: stored.lastModified || ref.lastModified || 0,
        addedAt: ref.addedAt || stored.addedAt || "",
        sourceType: stored.sourceType || ref.sourceType || "pending",
        type: stored.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        dataUrl: stored.dataUrl,
      });
    } else if (ref?.dataUrl) {
      resolved.push(ref);
    }
  }
  return resolved;
}

async function repositoryImportsForClaims(store, index, claims) {
  const claimSet = new Set(sanitizeClaims(claims).map((claim) => `${claim.sourceType}:${claim.key}`));
  const refs = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    const hasClaim = sanitizeRepositoryDoctors(file.doctors).some((doctor) => claimSet.has(`${doctor.sourceType}:${doctor.key}`));
    if (hasClaim) refs.push(repositoryImportRef(file));
  }
  return resolveStateImports(store, refs);
}

async function mergeLinkedDoctorProfiles(store, state, claims, ownerEmail = "") {
  const profileResult = await store.list({ prefix: DOCTOR_PROFILE_PREFIX });
  if (!(profileResult.keys || []).length) return state;
  const claimSourcesByKey = new Map();
  for (const claim of sanitizeClaims(claims)) {
    if (!claimSourcesByKey.has(claim.key)) claimSourcesByKey.set(claim.key, new Set());
    claimSourcesByKey.get(claim.key).add(claim.sourceType);
  }
  const session = state?.session && typeof state.session === "object" ? { ...state.session } : {};
  const mergedOverrides = { ...(session.overrides && typeof session.overrides === "object" ? session.overrides : {}) };
  const mergedConflictSelections = { ...(session.conflictSelections && typeof session.conflictSelections === "object" ? session.conflictSelections : {}) };
  const mergedCustomEvents = Array.isArray(session.customEvents) ? [...session.customEvents] : [];

  for (const item of profileResult.keys || []) {
    const profile = sanitizeDoctorProfile(await store.get(item.name, "json").catch(() => null));
    if (!profile) continue;
    const allowedSources = claimSourcesByKey.get(profile.doctorKey);
    if (!allowedSources) continue;
    if (!profile.sourceTypes.every((sourceType) => allowedSources.has(sourceType))) continue;
    const profileSession = profile.state?.session && typeof profile.state.session === "object" ? profile.state.session : {};
    Object.assign(mergedOverrides, profileSession.overrides && typeof profileSession.overrides === "object" ? profileSession.overrides : {});
    Object.assign(mergedConflictSelections, profileSession.conflictSelections && typeof profileSession.conflictSelections === "object" ? profileSession.conflictSelections : {});
    for (const event of Array.isArray(profileSession.customEvents) ? profileSession.customEvents : []) {
      const reassigned = {
        ...event,
        ownerEmail: normalizeEmail(ownerEmail || event.ownerEmail || ""),
      };
      if (!mergedCustomEvents.some((existing) => existing.id === reassigned.id)) {
        mergedCustomEvents.push(reassigned);
      }
    }
  }

  return {
    ...state,
    session: {
      ...session,
      overrides: mergedOverrides,
      conflictSelections: mergedConflictSelections,
      customEvents: mergedCustomEvents,
    },
  };
}

function matchRepositoryClaims(index, realName) {
  const claims = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    for (const doctor of sanitizeRepositoryDoctors(file.doctors)) {
      if (!doctorMatchesRealName(doctor, realName)) continue;
      claims.push({
        key: doctor.key,
        displayName: doctor.displayName,
        sourceType: doctor.sourceType,
        matchedAt: new Date().toISOString(),
      });
    }
  }
  return mergeClaims([], claims);
}

async function repositoryDoctorCandidates(store, index) {
  const claimed = await claimedRosterNames(store);
  const seen = new Set();
  const candidates = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    for (const doctor of sanitizeRepositoryDoctors(file.doctors)) {
      const marker = `${doctor.sourceType}:${doctor.key}`;
      if (seen.has(marker)) continue;
      seen.add(marker);
      candidates.push({
        key: doctor.key,
        displayName: doctor.displayName,
        sourceType: doctor.sourceType,
        claimedBy: claimed.get(marker)?.email || "",
        claimedByName: claimed.get(marker)?.realName || "",
      });
    }
  }
  return candidates.sort((left, right) => {
    const leftClaimed = left.claimedBy ? 1 : 0;
    const rightClaimed = right.claimedBy ? 1 : 0;
    if (leftClaimed !== rightClaimed) return leftClaimed - rightClaimed;
    return left.displayName.localeCompare(right.displayName) || left.sourceType.localeCompare(right.sourceType);
  });
}

async function claimedRosterNames(store) {
  const claimed = new Map();
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    const email = normalizeEmail(record?.email || item.name.replace(/^account:/, ""));
    for (const claim of sanitizeClaims(record?.claims)) {
      const marker = `${claim.sourceType}:${claim.key}`;
      if (!claimed.has(marker)) {
        claimed.set(marker, {
          email,
          realName: String(record?.realName || "").trim(),
        });
      }
    }
  }
  return claimed;
}

function findRepositoryDoctor(index, rawClaim) {
  const claim = {
    key: normalizeRosterName(rawClaim?.key || ""),
    sourceType: String(rawClaim?.sourceType || "").toLowerCase(),
  };
  if (!claim.key || !claim.sourceType) return null;
  const seen = new Set();
  for (const file of index.files || []) {
    if (file.active === false) continue;
    for (const doctor of sanitizeRepositoryDoctors(file.doctors)) {
      const marker = `${doctor.sourceType}:${doctor.key}`;
      if (seen.has(marker)) continue;
      seen.add(marker);
      if (doctor.key === claim.key && doctor.sourceType === claim.sourceType) return doctor;
    }
  }
  return null;
}

function sanitizeClaims(claims) {
  if (!Array.isArray(claims)) return [];
  return claims
    .map((claim) => ({
      key: normalizeRosterName(claim?.key || ""),
      displayName: String(claim?.displayName || "").trim(),
      sourceType: String(claim?.sourceType || "").toLowerCase(),
      matchedAt: String(claim?.matchedAt || ""),
    }))
    .filter((claim) => claim.key && claim.displayName && (claim.sourceType === "mmc" || claim.sourceType === "ddh"));
}

function sanitizeSourceTypes(items) {
  if (!Array.isArray(items)) return [];
  return [...new Set(items.map((item) => String(item || "").toLowerCase()).filter((item) => item === "mmc" || item === "ddh"))];
}

function sanitizeDoctorProfile(value) {
  if (!value || typeof value !== "object") return null;
  const profileId = String(value.profileId || "").trim();
  const doctorKey = normalizeRosterName(value.doctorKey || "");
  const displayName = String(value.displayName || "").trim();
  const sourceTypes = sanitizeSourceTypes(value.sourceTypes);
  if (!profileId || !doctorKey || !displayName || !sourceTypes.length) return null;
  return {
    profileId,
    doctorKey,
    displayName,
    sourceTypes,
    state: sanitizeState(value.state),
    createdAt: String(value.createdAt || ""),
    updatedAt: String(value.updatedAt || ""),
  };
}

function hasDoctorProfileState(state) {
  const session = state?.session && typeof state.session === "object" ? state.session : {};
  const overrides = session.overrides && typeof session.overrides === "object" ? session.overrides : {};
  const conflictSelections = session.conflictSelections && typeof session.conflictSelections === "object" ? session.conflictSelections : {};
  const customEvents = Array.isArray(session.customEvents) ? session.customEvents : [];
  return Boolean(Object.keys(overrides).length || Object.keys(conflictSelections).length || customEvents.length);
}

function mergeClaims(existing, incoming) {
  const claims = [];
  for (const claim of [...sanitizeClaims(existing), ...sanitizeClaims(incoming)]) {
    if (claims.some((item) => sameClaim(item, claim))) continue;
    claims.push(claim);
  }
  return claims.sort((left, right) => left.sourceType.localeCompare(right.sourceType) || left.displayName.localeCompare(right.displayName));
}

function sameClaim(left, right) {
  return left?.sourceType === right?.sourceType && left?.key === right?.key;
}

function doctorMatchesRealName(doctor, realName) {
  const realKey = normalizeRosterName(realName);
  if (!realKey) return false;
  if (doctor.key === realKey) return true;
  if (nameTokenMatch(realName, doctor.displayName)) return true;
  return likelySameRosterName(realName, doctor.displayName);
}

function nameTokenMatch(left, right) {
  const leftTokens = rosterNameTokens(left);
  const rightTokens = rosterNameTokens(right);
  if (leftTokens.length < 2 || rightTokens.length < 2) return false;
  const rightSet = new Set(rightTokens);
  return leftTokens.every((token) => rightSet.has(token));
}

function likelySameRosterName(left, right) {
  const leftTokens = rosterNameTokens(left);
  const rightTokens = rosterNameTokens(right);
  if (leftTokens.length < 2 || rightTokens.length < 2) return false;
  if (leftTokens[leftTokens.length - 1] !== rightTokens[rightTokens.length - 1]) return false;
  const leftFirst = leftTokens[0] || "";
  const rightFirst = rightTokens[0] || "";
  return leftFirst.length >= 3 && rightFirst.length >= 3 && (leftFirst.startsWith(rightFirst) || rightFirst.startsWith(leftFirst));
}

function rosterNameTokens(value) {
  return normalizeRosterName(value).split(" ").filter(Boolean);
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
