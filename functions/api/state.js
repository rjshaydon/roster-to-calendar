import { inspectImportRecord, normalizeRosterName } from "../_lib/roster.js";

const CREATOR_EMAIL = "rhaydon@gmail.com";
const REPOSITORY_INDEX_KEY = "repository:index";
const REPOSITORY_FILE_PREFIX = "repository:file:";
const DOCTOR_PROFILE_PREFIX = "doctor-profile:";
const SUBSCRIPTION_TOKEN_PREFIX = "subscription:token:";
const SNAPSHOT_PREFIX = "snapshot:";
const SNAPSHOT_SCHEMA_VERSION = 1;
const ADMIN_ISSUE_DISMISS_PREFIX = "admin-issue-dismiss:";
const ADMIN_ISSUE_IGNORE_PREFIX = "admin-issue-ignore:";
const PARSER_EXTENSION_RULES_KEY = "parser-extension-rules:v1";

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
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, account.record, { includeAvailableDoctors: account.record.role !== "creator" && account.record.role !== "owner" });
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
        subscription: prepared.subscription,
        snapshot: prepared.snapshot,
        snapshotAvailable: prepared.snapshotAvailable,
        snapshotStale: prepared.snapshotStale,
        snapshotBuiltAt: prepared.snapshotBuiltAt,
        issueConfig: prepared.issueConfig,
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
      const target = await loadAccountRecord(context.env.ROSTER_STORE, targetEmail);
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, target, { includeAvailableDoctors: false });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
        subscription: prepared.subscription,
        snapshot: prepared.snapshot,
        snapshotAvailable: prepared.snapshotAvailable,
        snapshotStale: prepared.snapshotStale,
        snapshotBuiltAt: prepared.snapshotBuiltAt,
        issueConfig: prepared.issueConfig,
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
        subscription: prepared.subscription,
        issueConfig: prepared.issueConfig,
      });
    }

    if (action === "listUsers") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      return Response.json({ ok: true, users: await listUsers(context.env.ROSTER_STORE) });
    }

    if (action === "reportUserError") {
      const reportEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = reportEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, reportEmail);
      if ((targetRecord.role || roleForEmail(targetRecord.email)) === "creator") {
        return Response.json({ ok: true, ignored: true });
      }
      const issue = sanitizeAdminIssues([{
        id: String(body?.errorId || "").trim(),
        message: body?.message,
        source: body?.issue?.source,
        date: body?.issue?.date || body?.issue?.startDay,
        rawValue: body?.issue?.rawValue,
        timeLabel: body?.issue?.timeLabel,
        suggestedTitle: body?.issue?.suggestedTitle,
        fingerprint: body?.issue?.fingerprint,
        firstSeenAt: new Date().toISOString(),
        lastSeenAt: new Date().toISOString(),
        count: 1,
      }])[0];
      if (!issue) {
        return Response.json({ error: "Structured issue details are required." }, { status: 400 });
      }
      const dismissed = new Set(await loadDismissedIssueFingerprints(context.env.ROSTER_STORE, reportEmail));
      const ignored = new Set(await loadIgnoredIssueFingerprints(context.env.ROSTER_STORE));
      if (dismissed.has(issue.fingerprint) || ignored.has(issue.fingerprint)) {
        return Response.json({ ok: true, ignored: true });
      }
      const nextIssues = mergeAdminIssues(targetRecord.adminIssues, [{
        ...issue,
      }]);
      await context.env.ROSTER_STORE.put(storageKey(reportEmail), JSON.stringify({
        ...targetRecord,
        adminIssues: nextIssues,
        updatedAt: new Date().toISOString(),
      }));
      return Response.json({ ok: true, issuesCount: nextIssues.length });
    }

    if (action === "clearUserError") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const clearEmail = normalizeEmail(targetEmail);
      if (!clearEmail) {
        return Response.json({ error: "Target account is required." }, { status: 400 });
      }
      const targetRecord = await loadAccountRecord(context.env.ROSTER_STORE, clearEmail);
      const errorId = String(body?.errorId || "").trim();
      const existingIssues = sanitizeAdminIssues(targetRecord.adminIssues);
      const fingerprintsToDismiss = errorId
        ? existingIssues.filter((issue) => issue.id === errorId || issue.fingerprint === errorId).map((issue) => issue.fingerprint)
        : existingIssues.map((issue) => issue.fingerprint);
      const nextDismissed = [...new Set([...(await loadDismissedIssueFingerprints(context.env.ROSTER_STORE, clearEmail)), ...fingerprintsToDismiss])];
      await saveDismissedIssueFingerprints(context.env.ROSTER_STORE, clearEmail, nextDismissed);
      const nextIssues = errorId
        ? existingIssues.filter((issue) => issue.id !== errorId && issue.fingerprint !== errorId)
        : [];
      await context.env.ROSTER_STORE.put(storageKey(clearEmail), JSON.stringify({
        ...targetRecord,
        adminIssues: nextIssues,
        updatedAt: new Date().toISOString(),
      }));
      return Response.json({ ok: true });
    }

    if (action === "ignoreUserErrorForever") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const ignoreFingerprint = sanitizeIssueFingerprint(body?.fingerprint || issueFingerprint(body?.source, body?.rawValue));
      if (!ignoreFingerprint) {
        return Response.json({ error: "Issue fingerprint is required." }, { status: 400 });
      }
      const ignored = new Set(await loadIgnoredIssueFingerprints(context.env.ROSTER_STORE));
      ignored.add(ignoreFingerprint);
      await saveIgnoredIssueFingerprints(context.env.ROSTER_STORE, [...ignored]);
      await clearIssueFromAllUsers(context.env.ROSTER_STORE, ignoreFingerprint);
      return Response.json({ ok: true, fingerprint: ignoreFingerprint });
    }

    if (action === "saveParserExtensionRule") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const rule = sanitizeParserExtensionRule(body?.rule);
      if (!rule) {
        return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      }
      const previousCode = String(body?.previousCode || "").trim().toUpperCase();
      let parserExtensions = await loadParserExtensionRules(context.env.ROSTER_STORE);
      if (previousCode && previousCode !== rule.code) {
        const sourceKey = rule.source.toLowerCase();
        parserExtensions = {
          ...parserExtensions,
          [sourceKey]: (parserExtensions[sourceKey] || []).filter((item) => item.code !== previousCode),
        };
      }
      parserExtensions = upsertParserExtensionRule(parserExtensions, rule);
      await saveParserExtensionRules(context.env.ROSTER_STORE, parserExtensions);
      const ignoreFingerprint = sanitizeIssueFingerprint(body?.fingerprint || issueFingerprint(body?.source, body?.rawValue));
      if (ignoreFingerprint) {
        await clearIssueFromAllUsers(context.env.ROSTER_STORE, ignoreFingerprint);
      }
      return Response.json({ ok: true, parserExtensions });
    }

    if (action === "deleteAccount") {
      const deleteEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      if (deleteEmail === CREATOR_EMAIL) {
        return Response.json({ error: "The creator account cannot be deleted." }, { status: 400 });
      }
      if (deleteEmail !== email && account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const record = await loadAccountRecord(context.env.ROSTER_STORE, deleteEmail).catch(() => null);
      if (record?.subscriptionToken) {
        await context.env.ROSTER_STORE.delete(subscriptionTokenKey(record.subscriptionToken));
      }
      if (record?.email) {
        const owner = accountSnapshotOwner(record.email, record.role || roleForEmail(record.email));
        await context.env.ROSTER_STORE.delete(snapshotKey(owner.ownerType, owner.ownerId));
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
      if (targetRole === "creator" || targetRole === "owner") {
        await reconcileRepositoryActiveFiles(context.env.ROSTER_STORE, repository.index, state.imports);
      }
      await context.env.ROSTER_STORE.put(storageKey(saveEmail), JSON.stringify({
        ...targetRecord,
        email: saveEmail,
        role: targetRole,
        realName: targetRecord.realName || "",
        claims,
        state,
        updatedAt: new Date().toISOString(),
      }));
      await storeSnapshotForAccount(context.env.ROSTER_STORE, {
        email: saveEmail,
        role: targetRole,
        claims,
        state,
        record: { ...targetRecord, email: saveEmail, role: targetRole, claims, state },
        snapshot: body?.snapshot,
      });
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
      const profile = await loadDoctorProfileRecord(context.env.ROSTER_STORE, profileId) || sanitizeDoctorProfile({
        profileId,
        doctorKey: body?.doctorKey,
        displayName: body?.displayName,
        sourceTypes: body?.sourceTypes,
        state: sanitizeState(null),
      });
      const snapshotInfo = await loadDoctorProfileSnapshotInfo(context.env.ROSTER_STORE, profile);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        profile,
        snapshot: snapshotInfo.snapshot,
        snapshotAvailable: snapshotInfo.snapshotAvailable,
        snapshotStale: snapshotInfo.snapshotStale,
        snapshotBuiltAt: snapshotInfo.snapshotBuiltAt,
        issueConfig: await buildIssueConfig(context.env.ROSTER_STORE, ""),
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
        await context.env.ROSTER_STORE.delete(snapshotKey("doctor-profile", profileId));
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
      await storeSnapshotForDoctorProfile(context.env.ROSTER_STORE, next, body?.snapshot);
      return Response.json({ ok: true, profile: next });
    }

    if (action === "loadImports") {
      const loadEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = loadEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, loadEmail);
      const imports = await resolveAccountImports(context.env.ROSTER_STORE, targetRecord);
      return Response.json({ ok: true, imports });
    }

    if (action === "loadDoctorProfileImports") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const profileId = String(body?.profileId || "").trim();
      const profile = await loadDoctorProfileRecord(context.env.ROSTER_STORE, profileId) || sanitizeDoctorProfile({
        profileId,
        doctorKey: body?.doctorKey,
        displayName: body?.displayName,
        sourceTypes: body?.sourceTypes,
        state: sanitizeState(null),
      });
      if (!profile) {
        return Response.json({ error: "Doctor profile was not found." }, { status: 404 });
      }
      const imports = await repositoryImportsForDoctorProfile(context.env.ROSTER_STORE, profile);
      return Response.json({ ok: true, imports });
    }

    if (action === "loadInsightImports") {
      const index = await loadRepositoryIndex(context.env.ROSTER_STORE);
      const refs = (index.files || []).filter((file) => file.active !== false).map((file) => repositoryImportRef(file));
      const imports = await resolveStateImports(context.env.ROSTER_STORE, refs);
      return Response.json({ ok: true, imports });
    }

    return Response.json({ error: "Unsupported account action." }, { status: 400 });
  } catch (error) {
    const message = error.message || "Account request failed.";
    const status = message === "Incorrect password." || message.startsWith("Account not found") ? 401 : 400;
    return Response.json({ error: message }, { status });
  }
}

export async function loadAccountRecord(store, email) {
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

export function normalizeEmail(value) {
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
      adminIssues: sanitizeAdminIssues(record?.adminIssues),
      issuesCount: sanitizeAdminIssues(record?.adminIssues).length,
      createdAt: record?.createdAt || "",
      updatedAt: record?.updatedAt || "",
    };
  }));
  return users.sort((a, b) => a.email.localeCompare(b.email));
}

export async function prepareAccountResponse(store, rawRecord, options = {}) {
  let record = await ensureAccountSubscriptionToken(store, rawRecord);
  const role = record.role || roleForEmail(record.email);
  const index = await loadRepositoryIndex(store);
  let claims = sanitizeClaims(record.claims);
  let nameMatches = [];
  let state = sanitizeState(record.state);
  let linkedProfiles = [];

  if (role !== "creator" && role !== "owner") {
    const matchedClaims = matchRepositoryClaims(index, record.realName || "");
    const merged = mergeClaims(claims, matchedClaims);
    nameMatches = matchedClaims.filter((claim) => !claims.some((existing) => sameClaim(existing, claim)));
    claims = merged;
    linkedProfiles = await linkedDoctorProfilesForClaims(store, claims);
    state = {
      ...state,
      imports: repositoryImportRefsForClaims(index, claims),
    };
    state = mergeProfileSessionIntoState(state, linkedProfiles, record.email);
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
    const hasEmbeddedImports = Array.isArray(state.imports) && state.imports.some((item) => item?.dataUrl);
    const imported = hasEmbeddedImports ? await upsertStateImports(store, state.imports, record.email) : {
      index,
      refs: (index.files || []).filter((file) => file.active !== false).map(repositoryImportRef),
      changed: false,
    };
    const creatorRepositoryRefs = (imported.index.files || []).filter((file) => file.active !== false).map(repositoryImportRef);
    const stateWithRefs = { ...state, imports: creatorRepositoryRefs };
    if (hasEmbeddedImports && (imported.changed || importsChanged(state.imports, creatorRepositoryRefs))) {
      state = stateWithRefs;
      await store.put(storageKey(record.email), JSON.stringify({
        ...record,
        state,
        updatedAt: new Date().toISOString(),
      }));
    } else {
      state = stateWithRefs;
    }
  }

  const owner = accountSnapshotOwner(record.email, role);
  const buildStamp = await buildAccountSnapshotStamp(store, {
    role,
    email: record.email,
    claims,
    state,
    linkedProfiles,
    index,
  });
  const snapshot = await loadSnapshotRecord(store, owner.ownerType, owner.ownerId);
  const snapshotAvailable = Boolean(snapshot);
  const snapshotStale = !snapshot || snapshot.buildStamp !== buildStamp;
  const issueConfig = await buildIssueConfig(store, record.email);

  return {
    role,
    realName: record.realName || "",
    state,
    claims,
    nameMatches,
    availableDoctors: options.includeAvailableDoctors === false ? [] : await repositoryDoctorCandidates(store, index),
    subscription: {
      token: String(record.subscriptionToken || ""),
      enabled: Boolean(snapshot?.subscriptionFeeds?.full?.ics),
    },
    adminIssues: sanitizeAdminIssues(record.adminIssues),
    issueConfig,
    snapshot,
    snapshotAvailable,
    snapshotStale,
    snapshotBuiltAt: snapshot?.builtAt || "",
    snapshotBuildStamp: buildStamp,
  };
}

export async function loadAccountBySubscriptionToken(store, token) {
  const normalizedToken = String(token || "").trim();
  if (!normalizedToken) return null;
  const email = await store.get(subscriptionTokenKey(normalizedToken), "text").catch(() => "");
  if (!email) return null;
  return await loadAccountRecord(store, normalizeEmail(email)).catch(() => null);
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

async function buildIssueConfig(store, email = "") {
  return {
    parserExtensions: await loadParserExtensionRules(store),
    dismissedFingerprints: await loadDismissedIssueFingerprints(store, email),
    ignoredFingerprints: await loadIgnoredIssueFingerprints(store),
  };
}

async function clearIssueFromAllUsers(store, fingerprint) {
  const normalizedFingerprint = sanitizeIssueFingerprint(fingerprint);
  if (!normalizedFingerprint) return;
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    if (!record?.adminIssues?.length) continue;
    const nextIssues = sanitizeAdminIssues(record.adminIssues).filter((issue) => issue.fingerprint !== normalizedFingerprint);
    if (nextIssues.length === sanitizeAdminIssues(record.adminIssues).length) continue;
    await store.put(item.name, JSON.stringify({
      ...record,
      adminIssues: nextIssues,
      updatedAt: new Date().toISOString(),
    }));
  }
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

async function reconcileRepositoryActiveFiles(store, index, activeImports = []) {
  const activeIds = new Set((activeImports || [])
    .map((item) => item?.repoId || item?.repositoryId || item?.id)
    .filter(Boolean));
  let changed = false;
  const files = (index.files || []).map((file) => {
    const active = activeIds.has(file.id);
    if ((file.active !== false) === active) return file;
    changed = true;
    return { ...file, active };
  });
  if (!changed) return index;
  const next = { ...index, files };
  await saveRepositoryIndex(store, next);
  return next;
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

function subscriptionTokenKey(token) {
  return `${SUBSCRIPTION_TOKEN_PREFIX}${token}`;
}

function snapshotKey(ownerType, ownerId) {
  return `${SNAPSHOT_PREFIX}${ownerType}:${ownerId}`;
}

export function accountSnapshotOwner(email, role) {
  return {
    ownerType: role === "creator" || role === "owner" ? "creator-account" : "claimed-account",
    ownerId: normalizeEmail(email),
  };
}

function doctorProfileSnapshotOwner(profile) {
  return {
    ownerType: "doctor-profile",
    ownerId: String(profile?.profileId || "").trim(),
  };
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

function sanitizeAvailableDoctors(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((doctor) => ({
      key: normalizeRosterName(doctor?.key || ""),
      displayName: String(doctor?.displayName || "").trim(),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
      claimedBy: normalizeEmail(doctor?.claimedBy || ""),
      claimedByName: String(doctor?.claimedByName || "").trim(),
      accountEmail: normalizeEmail(doctor?.accountEmail || doctor?.claimedBy || ""),
      aliases: Array.isArray(doctor?.aliases)
        ? doctor.aliases.map((alias) => ({
            key: normalizeRosterName(alias?.key || ""),
            displayName: String(alias?.displayName || "").trim(),
            sourceType: String(alias?.sourceType || "").toLowerCase(),
          })).filter((alias) => alias.key && alias.displayName)
        : [],
    }))
    .filter((doctor) => doctor.key && doctor.displayName);
}

function sanitizeSnapshotPreview(value) {
  if (!value || typeof value !== "object") return null;
  return JSON.parse(JSON.stringify(value));
}

function sanitizeDetectedSources(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    mmc: Array.isArray(input.mmc) ? input.mmc.map((item) => String(item || "")).filter(Boolean) : [],
    ddh: Array.isArray(input.ddh) ? input.ddh.map((item) => String(item || "")).filter(Boolean) : [],
  };
}

function sanitizeSnapshotFileRefs(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => repositoryImportRef(item))
    .filter((item) => item.id);
}

function sanitizeSnapshotRecord(value) {
  if (!value || typeof value !== "object") return null;
  const ownerType = String(value.ownerType || "").trim();
  const ownerId = String(value.ownerId || "").trim();
  const preview = sanitizeSnapshotPreview(value.preview);
  if (!ownerType || !ownerId || !preview) return null;
  return {
    ownerType,
    ownerId,
    schemaVersion: Number(value.schemaVersion || SNAPSHOT_SCHEMA_VERSION) || SNAPSHOT_SCHEMA_VERSION,
    buildStamp: String(value.buildStamp || "").trim(),
    builtAt: String(value.builtAt || ""),
    preview,
    session: value.session && typeof value.session === "object" ? JSON.parse(JSON.stringify(value.session)) : {},
    doctorOptions: sanitizeAvailableDoctors(value.doctorOptions),
    detectedSources: sanitizeDetectedSources(value.detectedSources),
    fileRefs: sanitizeSnapshotFileRefs(value.fileRefs),
    subscriptionFeeds: sanitizeSubscriptionFeeds(value.subscriptionFeeds),
  };
}

export async function loadSnapshotRecord(store, ownerType, ownerId) {
  if (!ownerType || !ownerId) return null;
  return sanitizeSnapshotRecord(await store.get(snapshotKey(ownerType, ownerId), "json").catch(() => null));
}

async function persistSnapshotRecord(store, ownerType, ownerId, snapshot, buildStamp) {
  const sanitizedInput = sanitizeSnapshotRecord({
    ownerType,
    ownerId,
    ...snapshot,
    ownerType,
    ownerId,
    buildStamp,
    builtAt: new Date().toISOString(),
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
  });
  if (!sanitizedInput) return null;
  const persisted = {
    ...sanitizedInput,
    buildStamp,
    builtAt: new Date().toISOString(),
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
  };
  await store.put(snapshotKey(ownerType, ownerId), JSON.stringify(persisted));
  return persisted;
}

async function buildStateRevision(state) {
  const session = state?.session && typeof state.session === "object" ? state.session : {};
  return await sha256(JSON.stringify({
    imports: sanitizeSnapshotFileRefs(state?.imports || []),
    settings: session.settings || {},
    exportRange: session.exportRange || {},
    overrides: session.overrides || {},
    customEvents: session.customEvents || [],
    conflictSelections: session.conflictSelections || {},
    doctorKey: session.doctorKey || "",
  }));
}

async function buildAccountSnapshotStamp(store, context) {
  const role = context?.role || "user";
  const refs = role === "creator" || role === "owner"
    ? sanitizeSnapshotFileRefs(context?.state?.imports || [])
    : repositoryImportRefsForClaims(context?.index || await loadRepositoryIndex(store), context?.claims || []);
  const fileMarkers = refs.map((ref) => ({
    id: ref.id,
    sourceType: ref.sourceType,
    size: ref.size,
    lastModified: ref.lastModified,
    addedAt: ref.addedAt,
  }));
  const linkedProfileMarkers = Array.isArray(context?.linkedProfiles)
    ? context.linkedProfiles.map((profile) => ({
        profileId: profile.profileId,
        doctorKey: profile.doctorKey,
        sourceTypes: profile.sourceTypes,
        updatedAt: profile.updatedAt,
        stateRevision: profile.state ? JSON.stringify(profile.state.session || {}) : "",
      }))
    : [];
  const stateRevision = await buildStateRevision(context?.state || {});
  const parserExtensions = await loadParserExtensionRules(store);
  const ignoredFingerprints = await loadIgnoredIssueFingerprints(store);
  const dismissedFingerprints = await loadDismissedIssueFingerprints(store, context?.email || "");
  return await sha256(JSON.stringify({
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    ownerType: role === "creator" || role === "owner" ? "creator-account" : "claimed-account",
    ownerId: normalizeEmail(context?.email || ""),
    claims: sanitizeClaims(context?.claims || []),
    files: fileMarkers,
    linkedProfiles: linkedProfileMarkers,
    stateRevision,
    parserExtensions,
    ignoredFingerprints,
    dismissedFingerprints,
  }));
}

async function buildDoctorProfileSnapshotStamp(store, profile) {
  const refs = await repositoryImportRefsForDoctorProfile(store, profile);
  const fileMarkers = refs.map((ref) => ({
    id: ref.id,
    sourceType: ref.sourceType,
    size: ref.size,
    lastModified: ref.lastModified,
    addedAt: ref.addedAt,
  }));
  const stateRevision = await buildStateRevision(profile?.state || {});
  const parserExtensions = await loadParserExtensionRules(store);
  const ignoredFingerprints = await loadIgnoredIssueFingerprints(store);
  return await sha256(JSON.stringify({
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    ownerType: "doctor-profile",
    ownerId: String(profile?.profileId || "").trim(),
    doctorKey: normalizeRosterName(profile?.doctorKey || ""),
    displayName: String(profile?.displayName || "").trim(),
    sourceTypes: sanitizeSourceTypes(profile?.sourceTypes),
    files: fileMarkers,
    stateRevision,
    parserExtensions,
    ignoredFingerprints,
  }));
}

function repositoryImportRefsForClaims(index, claims) {
  const claimSet = new Set(sanitizeClaims(claims).map((claim) => `${claim.sourceType}:${claim.key}`));
  const refs = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    const hasClaim = sanitizeRepositoryDoctors(file.doctors).some((doctor) => claimSet.has(`${doctor.sourceType}:${doctor.key}`));
    if (hasClaim) refs.push(repositoryImportRef(file));
  }
  return refs;
}

async function repositoryImportsForClaims(store, index, claims) {
  return resolveStateImports(store, repositoryImportRefsForClaims(index, claims));
}

async function resolveAccountImports(store, record) {
  const role = record?.role || roleForEmail(record?.email || "");
  const state = sanitizeState(record?.state);
  if (role === "creator" || role === "owner") {
    return resolveStateImports(store, state.imports || []);
  }
  const index = await loadRepositoryIndex(store);
  return repositoryImportsForClaims(store, index, sanitizeClaims(record?.claims));
}

async function linkedDoctorProfilesForClaims(store, claims) {
  const profileResult = await store.list({ prefix: DOCTOR_PROFILE_PREFIX });
  if (!(profileResult.keys || []).length) return [];
  const claimSourcesByKey = new Map();
  for (const claim of sanitizeClaims(claims)) {
    if (!claimSourcesByKey.has(claim.key)) claimSourcesByKey.set(claim.key, new Set());
    claimSourcesByKey.get(claim.key).add(claim.sourceType);
  }
  const profiles = [];
  for (const item of profileResult.keys || []) {
    const profile = sanitizeDoctorProfile(await store.get(item.name, "json").catch(() => null));
    if (!profile) continue;
    const allowedSources = claimSourcesByKey.get(profile.doctorKey);
    if (!allowedSources) continue;
    if (!profile.sourceTypes.every((sourceType) => allowedSources.has(sourceType))) continue;
    profiles.push(profile);
  }
  return profiles;
}

function mergeProfileSessionIntoState(state, profiles, ownerEmail = "") {
  const session = state?.session && typeof state.session === "object" ? { ...state.session } : {};
  const mergedOverrides = { ...(session.overrides && typeof session.overrides === "object" ? session.overrides : {}) };
  const mergedConflictSelections = { ...(session.conflictSelections && typeof session.conflictSelections === "object" ? session.conflictSelections : {}) };
  const mergedCustomEvents = Array.isArray(session.customEvents) ? [...session.customEvents] : [];
  for (const profile of profiles || []) {
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

async function repositoryImportRefsForDoctorProfile(store, profile) {
  const index = await loadRepositoryIndex(store);
  const refs = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    const hasProfileDoctor = sanitizeRepositoryDoctors(file.doctors).some((doctor) => doctor.key === profile.doctorKey && profile.sourceTypes.includes(doctor.sourceType));
    if (hasProfileDoctor) refs.push(repositoryImportRef(file));
  }
  return refs;
}

async function repositoryImportsForDoctorProfile(store, profile) {
  return resolveStateImports(store, await repositoryImportRefsForDoctorProfile(store, profile));
}

async function loadDoctorProfileSnapshotInfo(store, profile) {
  const owner = doctorProfileSnapshotOwner(profile);
  const buildStamp = await buildDoctorProfileSnapshotStamp(store, profile);
  const snapshot = await loadSnapshotRecord(store, owner.ownerType, owner.ownerId);
  return {
    snapshot,
    snapshotAvailable: Boolean(snapshot),
    snapshotStale: !snapshot || snapshot.buildStamp !== buildStamp,
    snapshotBuiltAt: snapshot?.builtAt || "",
    snapshotBuildStamp: buildStamp,
  };
}

async function storeSnapshotForAccount(store, context) {
  if (!context?.snapshot) return null;
  const role = context.role || roleForEmail(context.email || "");
  const owner = accountSnapshotOwner(context.email, role);
  const index = await loadRepositoryIndex(store);
  const claims = sanitizeClaims(context.claims);
  const linkedProfiles = role === "creator" || role === "owner" ? [] : await linkedDoctorProfilesForClaims(store, claims);
  const buildStamp = await buildAccountSnapshotStamp(store, {
    role,
    email: context.email,
    claims,
    state: context.state,
    linkedProfiles,
    index,
  });
  return persistSnapshotRecord(store, owner.ownerType, owner.ownerId, context.snapshot, buildStamp);
}

async function storeSnapshotForDoctorProfile(store, profile, snapshot) {
  if (!snapshot || !profile?.profileId) return null;
  const owner = doctorProfileSnapshotOwner(profile);
  const buildStamp = await buildDoctorProfileSnapshotStamp(store, profile);
  return persistSnapshotRecord(store, owner.ownerType, owner.ownerId, snapshot, buildStamp);
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

function randomSubscriptionToken() {
  const values = new Uint8Array(24);
  crypto.getRandomValues(values);
  return [...values].map((value) => value.toString(16).padStart(2, "0")).join("");
}

async function ensureAccountSubscriptionToken(store, record) {
  if (!record?.email) return record;
  if (record.subscriptionToken) {
    await store.put(subscriptionTokenKey(record.subscriptionToken), normalizeEmail(record.email));
    return record;
  }
  const updated = {
    ...record,
    subscriptionToken: randomSubscriptionToken(),
    updatedAt: new Date().toISOString(),
  };
  await store.put(storageKey(updated.email), JSON.stringify(updated));
  await store.put(subscriptionTokenKey(updated.subscriptionToken), normalizeEmail(updated.email));
  return updated;
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
    subscriptionFeeds: sanitizeSubscriptionFeeds(input.subscriptionFeeds),
  };
}

function sanitizeSubscriptionFeeds(value) {
  if (!value || typeof value !== "object") return {};
  const next = {};
  for (const key of ["full", "range"]) {
    const item = value[key];
    if (!item || typeof item !== "object" || typeof item.ics !== "string" || !item.ics.trim()) continue;
    next[key] = {
      doctorKey: normalizeRosterName(item.doctorKey || ""),
      doctorDisplay: String(item.doctorDisplay || "").trim(),
      startDate: String(item.startDate || "").trim(),
      endDate: String(item.endDate || "").trim(),
      allFuture: item.allFuture !== false,
      generatedAt: String(item.generatedAt || ""),
      ics: String(item.ics || ""),
    };
  }
  return next;
}

function sanitizeAdminIssues(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => ({
      id: String(item?.id || "").trim(),
      message: String(item?.message || "").trim(),
      source: sanitizeIssueSource(item?.source),
      date: String(item?.date || item?.startDay || "").trim(),
      rawValue: String(item?.rawValue || "").trim(),
      timeLabel: String(item?.timeLabel || "").trim(),
      suggestedTitle: String(item?.suggestedTitle || "").trim(),
      fingerprint: sanitizeIssueFingerprint(item?.fingerprint || issueFingerprint(item?.source, item?.rawValue)),
      firstSeenAt: String(item?.firstSeenAt || ""),
      lastSeenAt: String(item?.lastSeenAt || ""),
      count: Number(item?.count || 1),
    }))
    .filter((item) => item.message && item.fingerprint)
    .map((item) => ({
      ...item,
      id: item.id || item.fingerprint,
      count: Number.isFinite(item.count) && item.count > 0 ? Math.floor(item.count) : 1,
    }));
}

function mergeAdminIssues(existing, incoming) {
  const issues = sanitizeAdminIssues(existing);
  for (const item of sanitizeAdminIssues(incoming)) {
    const match = issues.find((issue) => issue.fingerprint === item.fingerprint);
    if (match) {
      match.lastSeenAt = item.lastSeenAt || new Date().toISOString();
      match.count = Math.max(match.count, item.count || 1);
      match.message = item.message || match.message;
      match.source = item.source || match.source;
      match.date = item.date || match.date;
      match.rawValue = item.rawValue || match.rawValue;
      match.timeLabel = item.timeLabel || match.timeLabel;
      match.suggestedTitle = item.suggestedTitle || match.suggestedTitle;
      continue;
    }
    issues.unshift(item);
  }
  return issues
    .sort((left, right) => (right.lastSeenAt || "").localeCompare(left.lastSeenAt || ""))
    .slice(0, 50);
}

function sanitizeIssueSource(value) {
  const source = String(value || "").trim().toUpperCase();
  return source === "MMC" || source === "DDH" ? source : "";
}

function issueFingerprint(source, rawValue) {
  const normalizedSource = sanitizeIssueSource(source);
  const normalizedRawValue = String(rawValue || "").trim();
  return normalizedSource && normalizedRawValue ? `${normalizedSource}::${normalizedRawValue}` : "";
}

function sanitizeIssueFingerprint(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const [source, ...rest] = raw.split("::");
  return issueFingerprint(source, rest.join("::"));
}

function adminIssueDismissKey(email) {
  return `${ADMIN_ISSUE_DISMISS_PREFIX}${normalizeEmail(email)}`;
}

function adminIssueIgnoreKey() {
  return ADMIN_ISSUE_IGNORE_PREFIX;
}

async function loadDismissedIssueFingerprints(store, email) {
  if (!email) return [];
  const values = await store.get(adminIssueDismissKey(email), "json").catch(() => []);
  return sanitizeIssueFingerprintList(values);
}

async function saveDismissedIssueFingerprints(store, email, values) {
  const next = sanitizeIssueFingerprintList(values);
  if (!email) return;
  if (!next.length) {
    await store.delete(adminIssueDismissKey(email));
    return;
  }
  await store.put(adminIssueDismissKey(email), JSON.stringify(next));
}

async function loadIgnoredIssueFingerprints(store) {
  const values = await store.get(adminIssueIgnoreKey(), "json").catch(() => []);
  return sanitizeIssueFingerprintList(values);
}

async function saveIgnoredIssueFingerprints(store, values) {
  const next = sanitizeIssueFingerprintList(values);
  if (!next.length) {
    await store.delete(adminIssueIgnoreKey());
    return;
  }
  await store.put(adminIssueIgnoreKey(), JSON.stringify(next));
}

function sanitizeIssueFingerprintList(values) {
  if (!Array.isArray(values)) return [];
  return [...new Set(values.map((value) => sanitizeIssueFingerprint(value)).filter(Boolean))].sort();
}

function sanitizeParserExtensionRules(value) {
  const source = value && typeof value === "object" ? value : {};
  return {
    mmc: sanitizeParserExtensionRuleList(source.mmc, "MMC"),
    ddh: sanitizeParserExtensionRuleList(source.ddh, "DDH"),
  };
}

function sanitizeParserExtensionRuleList(items, source) {
  if (!Array.isArray(items)) return [];
  return items
    .map((item) => sanitizeParserExtensionRule(item, source))
    .filter(Boolean)
    .sort((left, right) => left.code.localeCompare(right.code));
}

function sanitizeParserExtensionRule(item, forcedSource = "") {
  if (!item || typeof item !== "object") return null;
  const source = sanitizeIssueSource(forcedSource || item.source);
  const code = String(item.code || item.rawCode || "").trim().toUpperCase();
  const kind = String(item.kind || "shift").trim().toLowerCase();
  const base = String(item.base || item.titleParts?.base || "").trim();
  const period = String(item.period || item.titleParts?.period || "").trim().toUpperCase();
  const suffix = String(item.suffix || item.titleParts?.suffix || "").trim();
  const location = String(item.location || "").trim();
  const allDay = item.allDay === true;
  const startTime = String(item.startTime || "").trim();
  const endTime = String(item.endTime || "").trim();
  if (!source || !code || !base) return null;
  if (!allDay && (!isClockString(startTime) || !isClockString(endTime))) return null;
  return {
    source,
    code,
    kind,
    base,
    period,
    suffix,
    allDay,
    startTime: allDay ? "" : startTime,
    endTime: allDay ? "" : endTime,
    location,
  };
}

async function loadParserExtensionRules(store) {
  const value = await store.get(PARSER_EXTENSION_RULES_KEY, "json").catch(() => null);
  return sanitizeParserExtensionRules(value);
}

async function saveParserExtensionRules(store, value) {
  const sanitized = sanitizeParserExtensionRules(value);
  await store.put(PARSER_EXTENSION_RULES_KEY, JSON.stringify(sanitized));
  return sanitized;
}

function upsertParserExtensionRule(existing, rule) {
  const sanitized = sanitizeParserExtensionRules(existing);
  const nextRule = sanitizeParserExtensionRule(rule);
  if (!nextRule) return sanitized;
  const key = nextRule.source.toLowerCase();
  const items = sanitized[key] || [];
  const nextItems = items.filter((item) => item.code !== nextRule.code);
  nextItems.push(nextRule);
  return {
    ...sanitized,
    [key]: nextItems.sort((left, right) => left.code.localeCompare(right.code)),
  };
}

function isClockString(value) {
  return /^\d{2}:\d{2}$/.test(String(value || "").trim());
}
