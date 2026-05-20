import { applyEventOverrides, customEventsToEvents, defaultSettings, inspectImportRecord, normalizeRosterName } from "../_lib/roster.js";
import {
  buildPreviewFromDerivedEvents,
  accountMirrorStatus,
  appendConsoleMessage,
  countDerivedDoctorsByFile,
  countDerivedEventsByFile,
  countDerivedEventsByFileDoctorPairs,
  deleteAccountMirror,
  deleteDerivedRosterFile,
  deleteDoctorProfileMirror,
  deleteRawRosterFile,
  hasCalendarDb,
  applyAccountHospitalLocations,
  loadAccountHospitalLocations,
  mergeHospitalLocationsIntoSettings,
  listAccountMirrors,
  listConsoleMessages,
  loadAccountMirror,
  loadAccountStateMirror,
  loadDoctorProfileMirror,
  loadRawRosterFile,
  queryCoworkerEvents,
  queryOverlapDoctors,
  queryClaimedAccounts,
  queryDoctorProfileMirrors,
  queryDoctorEvents,
  queryAccountCustomEvents,
  queryCanonicalDoctors,
  queryDoctorSeniorities,
  queryDoctorEventsForFileDoctorPairs,
  queryRosterFileDoctors,
  queryRosterFileDoctorsForKeys,
  queryRosterFileRefsForDoctors,
  queryRosterFiles,
  queryRosterFileRanges,
  queryRosterDoctors,
  replaceDerivedRosterFile,
  replaceAccountCustomEvents,
  replaceCanonicalDoctors,
  upsertAccountHospitalLocations,
  setDerivedRosterFileActive,
  upsertAccountMirror,
  upsertDoctorProfileMirror,
  upsertDerivedRosterFile,
  upsertRawRosterFile,
} from "../_lib/d1-calendar.js";

const CREATOR_EMAIL = "rhaydon@gmail.com";
const OWNER_DOCTOR_KEY = "RICHARD HAYDON";
const SNAPSHOT_SCHEMA_VERSION = 5;

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
    if (!hasCalendarDb(context.env)) {
      return Response.json({ error: "D1 database is not configured." }, { status: 503 });
    }
    if (!password) {
      return Response.json({ error: "Password is required." }, { status: 400 });
    }
    if (action === "login") {
      const account = await loadOrCreateD1Account(context.env.ROSTER_DB, email, password, { mode, realName });
      const loginRecord = account.created
        ? await autoClaimMatchedCanonicalDoctors(account.record, context.env.ROSTER_DB)
        : account.record;
      if (loginRecord !== account.record) await upsertAccountMirror(context.env.ROSTER_DB, loginRecord);
      const prepared = await prepareLightweightAccountResponse(loginRecord, { db: context.env.ROSTER_DB });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: account.created,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: [],
        suggestedClaims: [],
        availableDoctors: [],
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
        snapshot: null,
        snapshotAvailable: false,
        snapshotStale: false,
        snapshotBuiltAt: "",
        issueConfig: null,
      });
    }

    const account = await verifyD1Account(context.env.ROSTER_DB, email, password);
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
      const created = await loadOrCreateD1Account(context.env.ROSTER_DB, targetEmail, targetPassword, {
        mode: "create",
        realName: targetRealName,
      });
      const createdRecord = await autoClaimMatchedCanonicalDoctors(created.record, context.env.ROSTER_DB);
      await upsertAccountMirror(context.env.ROSTER_DB, createdRecord);
      const createdClaims = sanitizeClaims(createdRecord.claims);
      const createdRole = createdRecord.role || roleForEmail(targetEmail);
      const createdSeniorities = await queryDoctorSeniorities(context.env.ROSTER_DB, createdClaims.map((claim) => claim.key)).catch(() => []);
      return Response.json({
        ok: true,
        cloudAvailable: true,
        created: true,
        user: {
          email: targetEmail,
          realName: createdRecord.realName || "",
          role: createdRole,
          sites: [...new Set(createdClaims.map((claim) => claim.sourceType.toUpperCase()))].sort(),
          seniorities: createdSeniorities,
          claims: createdClaims,
          suggestedClaims: [],
          insightsEnabled: insightsEnabledForRecord({ ...createdRecord, role: createdRole }),
          createdAt: created.record.createdAt || "",
          updatedAt: created.record.updatedAt || "",
        },
      });
    }

    if (action === "resolveAccountClaims") {
      const targetRecord = targetEmail && (account.role === "creator" || account.role === "owner")
        ? await loadAccountMirror(context.env.ROSTER_DB, targetEmail)
        : account.record;
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const resolved = await autoClaimMatchedRosterNames(null, targetRecord, context.env.ROSTER_DB);
      const prepared = await prepareAccountResponse(null, resolved, {
        db: context.env.ROSTER_DB,
        includeAvailableDoctors: resolved.role !== "creator" && resolved.role !== "owner",
      });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: [],
        suggestedClaims: [],
        availableDoctors: prepared.availableDoctors,
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
        issueConfig: prepared.issueConfig,
      });
    }

    if (action === "adminLoadUser") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const target = await loadAccountMirror(context.env.ROSTER_DB, targetEmail);
      if (!target) return Response.json({ error: "Account not found." }, { status: 404 });
      const targetClaims = sanitizeClaims(target.claims);
      const prepared = await prepareLightweightAccountResponse(target, { db: context.env.ROSTER_DB });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: [],
        suggestedClaims: [],
        availableDoctors: targetClaims.length
          ? []
          : await repositoryDoctorCandidates(null, null, context.env.ROSTER_DB, { hideZeroEventStandalone: true }),
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
        snapshot: null,
        snapshotAvailable: false,
        snapshotStale: false,
        snapshotBuiltAt: "",
        issueConfig: null,
      });
    }

    if (action === "claimRosterName") {
      const claimEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = claimEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, claimEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const claim = await resolveDoctorClaimCandidate(context.env.ROSTER_DB, body?.claim);
      if (!claim) {
        return Response.json({ error: "Roster name was not found." }, { status: 400 });
      }
      const claims = mergeClaims(targetRecord.claims, [{ ...claim, matchedAt: new Date().toISOString() }]);
      const updatedAdminIssues = claimMatchesAccountIdentity(claim, targetRecord.realName || "")
        ? targetRecord.adminIssues
        : mergeAdminIssues(targetRecord.adminIssues, [manualRosterClaimIssue(targetRecord, claim)]);
      const d1Refs = await d1RepositoryImportRefsForClaims(context.env.ROSTER_DB, claims);
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: d1Refs,
      };
      const updated = {
        ...targetRecord,
        email: claimEmail,
        claims,
        adminIssues: updatedAdminIssues,
        state,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      const prepared = await prepareAccountResponse(null, updated, {
        db: context.env.ROSTER_DB,
        includeAvailableDoctors: false,
      });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        role: prepared.role,
        realName: prepared.realName,
        state: prepared.state,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        suggestedClaims: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
        issueConfig: prepared.issueConfig,
      });
    }

    if (action === "listUsers") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      return Response.json({
        ok: true,
        users: await listD1Users(context.env.ROSTER_DB),
        availableDoctors: await repositoryDoctorCandidates(null, null, context.env.ROSTER_DB, { hideZeroEventStandalone: true }),
        issueConfig: await buildIssueConfig(null, email, context.env.ROSTER_DB),
      });
    }

    if (action === "calendarStoreStatus") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!hasCalendarDb(context.env)) {
        return Response.json({ ok: false, unavailable: true, total: 0, populated: 0, remaining: 0 });
      }
      const status = await calendarStoreStatus(null, context.env.ROSTER_DB, {
        doctorKey: body?.selectedDoctorKey || body?.doctorKey || OWNER_DOCTOR_KEY,
        expectedFileIds: sanitizeRepositoryFileIds(body?.expectedFileIds),
      });
      const accounts = await accountMirrorStatus(context.env.ROSTER_DB).catch(() => ({ unavailable: true, profiles: 0, claims: 0, states: 0 }));
      return Response.json({ ok: true, ...status, accounts });
    }

    if (action === "appendConsoleMessage") {
      if (!hasCalendarDb(context.env)) return Response.json({ ok: false, unavailable: true });
      await appendConsoleMessage(context.env.ROSTER_DB, {
        actorEmail: email,
        message: body?.message,
        isError: body?.isError === true,
      });
      return Response.json({ ok: true });
    }

    if (action === "consoleMessages") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!hasCalendarDb(context.env)) return Response.json({ ok: false, unavailable: true, messages: [] });
      return Response.json({ ok: true, messages: await listConsoleMessages(context.env.ROSTER_DB, 50) });
    }

    if (action === "saveDerivedCalendarFile") {
      if (!hasCalendarDb(context.env)) {
        return Response.json({ ok: false, unavailable: true });
      }
      const derivedPayloadIssue = validateDerivedCalendarPayload(body?.doctors, body?.eventsByDoctor);
      if (derivedPayloadIssue) {
        return Response.json({ error: derivedPayloadIssue }, { status: 422 });
      }
      const selectedDoctorKey = normalizeRosterName(body?.selectedDoctorKey || body?.doctorKey || OWNER_DOCTOR_KEY);
      const filePayload = {
        ...(body?.file || {}),
        uploadedBy: email,
        uploadedAt: new Date().toISOString(),
      };
      const result = await replaceDerivedRosterFile(
        context.env.ROSTER_DB,
        filePayload,
        body?.doctors || [],
        body?.eventsByDoctor || {},
      );
      const supersession = await reconcileRosterFileSupersession(context.env.ROSTER_DB, filePayload, { uploaderEmail: email });
      await refreshCanonicalDoctors(context.env.ROSTER_DB);
      if (body?.skipStatus === true) {
        return Response.json({
          ok: true,
          result,
          supersession,
          fileStatus: {
            id: String(filePayload.id || ""),
            name: String(filePayload.name || "roster.xlsx"),
            sourceType: String(filePayload.sourceType || "").toLowerCase(),
            indexedDoctors: Number(result.doctors || 0),
            eventCount: Number(result.events || 0),
            status: Number(result.events || 0) > 0 ? "populated" : "missing",
          },
        });
      }
      const status = await calendarStoreStatus(null, context.env.ROSTER_DB, {
        doctorKey: selectedDoctorKey,
        expectedFileIds: sanitizeRepositoryFileIds(body?.expectedFileIds),
      });
      return Response.json({ ok: true, result, supersession, ...status });
    }

    if (action === "uploadRawRosterFile") {
      if (!hasCalendarDb(context.env)) return Response.json({ ok: false, unavailable: true });
      const file = body?.file || {};
      const dataUrl = String(body?.dataUrl || "");
      if (!file?.id || !dataUrl) return Response.json({ error: "Roster source file is required." }, { status: 400 });
      const objectKey = rawRosterObjectKey(file.id);
      if (context.env.ROSTER_FILES?.put) {
        await context.env.ROSTER_FILES.put(objectKey, dataUrlToBytes(dataUrl), {
          httpMetadata: { contentType: String(body?.type || file.type || "application/octet-stream") },
        });
      }
      await upsertRawRosterFile(context.env.ROSTER_DB, file, {
        objectKey: context.env.ROSTER_FILES?.put ? objectKey : "",
        dataUrl: context.env.ROSTER_FILES?.put ? "" : dataUrl,
        type: body?.type || file.type,
        uploadedAt: new Date().toISOString(),
      });
      return Response.json({ ok: true, objectKey, retainedIn: context.env.ROSTER_FILES?.put ? "r2" : "d1" });
    }

    if (action === "fetchRawRosterFile") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const fileId = String(body?.fileId || "").trim();
      if (!fileId) return Response.json({ error: "Roster file is required." }, { status: 400 });
      const raw = await readRetainedRawRosterFile(context.env, fileId);
      if (!raw?.dataUrl) return Response.json({ error: "Retained roster source was not found." }, { status: 404 });
      return Response.json({ ok: true, fileId, type: raw.type, dataUrl: raw.dataUrl });
    }

    if (action === "resetDerivedCalendarFile") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!hasCalendarDb(context.env)) {
        return Response.json({ ok: false, unavailable: true });
      }
      const fileId = String(body?.fileId || "").trim();
      if (!fileId) {
        return Response.json({ error: "Roster file is required." }, { status: 400 });
      }
      await deleteDerivedRosterFile(context.env.ROSTER_DB, fileId);
      await refreshCanonicalDoctors(context.env.ROSTER_DB);
      await deleteRawRosterFile(context.env.ROSTER_DB, fileId);
      const status = await calendarStoreStatus(null, context.env.ROSTER_DB, { doctorKey: OWNER_DOCTOR_KEY });
      return Response.json({ ok: true, reset: fileId, ...status });
    }

    if (action === "replaceActiveRosterFiles") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const keepFileIds = sanitizeRepositoryFileIds(body?.keepFileIds);
      const activeFiles = await queryRosterFileRanges(context.env.ROSTER_DB, { includeInactive: false });
      const removedFileIds = activeFiles
        .map((file) => file.id)
        .filter((id) => id && !keepFileIds.includes(id));
      await Promise.all(removedFileIds.map((id) => deleteDerivedRosterFile(context.env.ROSTER_DB, id).catch(() => null)));
      await refreshCanonicalDoctors(context.env.ROSTER_DB);
      await Promise.all(removedFileIds.map((id) => deleteRawRosterFile(context.env.ROSTER_DB, id).catch(() => null)));
      const status = await calendarStoreStatus(null, context.env.ROSTER_DB, {
        doctorKey: normalizeRosterName(body?.selectedDoctorKey || OWNER_DOCTOR_KEY),
        expectedFileIds: keepFileIds,
      });
      return Response.json({ ok: true, keptFileIds: keepFileIds, removedFileIds, ...status });
    }

    if (action === "updateAccount") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, saveEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const nextRealName = String(body?.realName ?? targetRecord.realName ?? "").trim();
      const nextPassword = String(body?.newPassword || "");
      const passwordRecord = nextPassword ? await hashPassword(nextPassword) : {};
      const updated = {
        ...targetRecord,
        email: saveEmail,
        realName: nextRealName,
        ...passwordRecord,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      const prepared = await prepareAccountResponse(null, updated, { db: context.env.ROSTER_DB, includeAvailableDoctors: false });
      return Response.json({
        ok: true,
        realName: prepared.realName,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        suggestedClaims: prepared.nameMatches,
        user: await userSummaryFromRecord(saveEmail, { ...updated, claims: prepared.claims }, { db: context.env.ROSTER_DB }),
      });
    }

    if (action === "setAccountRosterClaims") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!targetEmail) {
        return Response.json({ error: "Target account is required." }, { status: 400 });
      }
      const targetRecord = await loadAccountMirror(context.env.ROSTER_DB, targetEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const canonicalDoctors = await loadSqlDoctorCandidates(context.env.ROSTER_DB);
      const claims = sanitizeClaims((body?.claims || [])
        .map((claim) => findDoctorClaimCandidate(canonicalDoctors, claim))
        .filter(Boolean)
        .map((claim) => ({ ...claim, matchedAt: new Date().toISOString() })));
      const d1Refs = await d1RepositoryImportRefsForClaims(context.env.ROSTER_DB, claims);
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: d1Refs,
      };
      const updated = {
        ...targetRecord,
        email: targetEmail,
        claims,
        state,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      return Response.json({
        ok: true,
        user: await userSummaryFromRecord(targetEmail, updated, { db: context.env.ROSTER_DB }),
        claims,
      });
    }

    if (action === "removeRosterClaim") {
      const claimEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = claimEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, claimEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const rawClaim = {
        sourceType: String(body?.claim?.sourceType || "").toLowerCase(),
        key: normalizeRosterName(body?.claim?.key || ""),
      };
      const claims = sanitizeClaims(targetRecord.claims).filter((claim) => !(claim.sourceType === rawClaim.sourceType && claim.key === rawClaim.key));
      const d1Refs = await d1RepositoryImportRefsForClaims(context.env.ROSTER_DB, claims);
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: d1Refs,
      };
      const updated = {
        ...targetRecord,
        email: claimEmail,
        claims,
        state,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      return Response.json({ ok: true, claims, user: await userSummaryFromRecord(claimEmail, updated, { db: context.env.ROSTER_DB }) });
    }

    if (action === "reportRosterIdentityIssue") {
      const reportEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = reportEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, reportEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const issue = sanitizeAdminIssues([{
        id: `identity:${Date.now()}`,
        message: String(body?.message || "Roster name match needs review.").trim(),
        source: "MMC",
        seniority: "Unknown",
        rawValue: String(body?.rawValue || `Identity review: ${targetRecord.realName || reportEmail}`),
        fingerprint: `MMC::Unknown::Identity review: ${reportEmail}::${Date.now()}`,
        firstSeenAt: new Date().toISOString(),
        lastSeenAt: new Date().toISOString(),
        count: 1,
      }])[0];
      const updated = {
        ...targetRecord,
        adminIssues: mergeAdminIssues(targetRecord.adminIssues, [issue]),
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      return Response.json({ ok: true, user: await userSummaryFromRecord(reportEmail, updated, { db: context.env.ROSTER_DB }) });
    }

    if (action === "resolveDoctorAccount") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const resolution = await resolveDoctorAccount(null, body?.doctor, context.env.ROSTER_DB);
      return Response.json({
        ok: true,
        ...resolution,
      });
    }

    if (action === "setUserInsightsEnabled") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!targetEmail) {
        return Response.json({ error: "Target account is required." }, { status: 400 });
      }
      const targetRecord = await loadAccountMirror(context.env.ROSTER_DB, targetEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const targetRole = targetRecord.role || roleForEmail(targetEmail);
      if (targetRole === "creator" || targetRole === "owner") {
        return Response.json({ error: "Creator insights cannot be disabled." }, { status: 400 });
      }
      const updated = {
        ...targetRecord,
        insightsEnabled: body?.insightsEnabled === true,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
      return Response.json({
        ok: true,
        user: await userSummaryFromRecord(targetEmail, updated, { db: context.env.ROSTER_DB }),
      });
    }

    if (action === "reportUserError") {
      const reportEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = reportEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, reportEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const issue = sanitizeAdminIssues([{
        id: String(body?.errorId || "").trim(),
        message: body?.message,
        source: body?.issue?.source,
        seniority: body?.issue?.seniority,
        date: body?.issue?.date || body?.issue?.startDay,
        rawValue: body?.issue?.rawValue,
        code: body?.issue?.code,
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
      const dismissed = new Set();
      const ignored = new Set();
      if (dismissed.has(issue.fingerprint) || ignored.has(issue.fingerprint) || await isIssueResolvedByParserRules(null, reportEmail, issue, context.env.ROSTER_DB)) {
        return Response.json({ ok: true, ignored: true });
      }
      const nextIssues = mergeAdminIssues(targetRecord.adminIssues, [{
        ...issue,
      }]);
      const updated = {
        ...targetRecord,
        adminIssues: nextIssues,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updated);
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
      const targetRecord = await loadAccountMirror(context.env.ROSTER_DB, clearEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const errorId = String(body?.errorId || "").trim();
      const existingIssues = sanitizeAdminIssues(targetRecord.adminIssues);
      const fingerprintsToDismiss = errorId
        ? existingIssues.filter((issue) => issue.id === errorId || issue.fingerprint === errorId).map((issue) => issue.fingerprint)
        : existingIssues.map((issue) => issue.fingerprint);
      const nextIssues = errorId
        ? existingIssues.filter((issue) => issue.id !== errorId && issue.fingerprint !== errorId)
        : [];
      await upsertAccountMirror(context.env.ROSTER_DB, {
        ...targetRecord,
        adminIssues: nextIssues,
        updatedAt: new Date().toISOString(),
      });
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
      await clearIssueFromAllUsers(context.env.ROSTER_DB, ignoreFingerprint);
      return Response.json({ ok: true, fingerprint: ignoreFingerprint });
    }

    if (action === "saveParserExtensionRule") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const rules = (Array.isArray(body?.rules) ? body.rules : [body?.rule])
        .map((item) => sanitizeParserExtensionRule(item))
        .filter(Boolean);
      if (!rules.length) {
        return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      }
      const rule = rules[0];
      const previousCode = String(body?.previousCode || "").trim().toUpperCase();
      const previousSeniority = sanitizeRuleSeniority(body?.previousSeniority || rule.seniority);
      let parserExtensions = await loadD1ParserExtensionRules(context.env.ROSTER_DB);
      const replacementTargets = Array.isArray(body?.replacementTargets)
        ? body.replacementTargets.map((item) => sanitizeParserRuleRemoval(item)).filter(Boolean)
        : [];
      const nextKeys = new Set(rules.map((item) => `${item.source}|${item.seniority}|${item.code}`));
      for (const target of replacementTargets) {
        if (nextKeys.has(`${target.source}|${target.seniority}|${target.code}`)) continue;
        parserExtensions = removeParserExtensionRuleByKey(parserExtensions, target);
        await deleteD1ParserExtensionRule(context.env.ROSTER_DB, target);
      }
      if (!replacementTargets.length && previousCode && (previousCode !== rule.code || previousSeniority !== rule.seniority)) {
        parserExtensions = removeParserExtensionRuleByKey(parserExtensions, {
          source: rule.source,
          seniority: previousSeniority,
          code: previousCode,
        });
        await deleteD1ParserExtensionRule(context.env.ROSTER_DB, {
          source: rule.source,
          seniority: previousSeniority,
          code: previousCode,
        });
      }
      for (const item of rules) {
        parserExtensions = upsertParserExtensionRule(parserExtensions, item);
        await saveD1ParserExtensionRule(context.env.ROSTER_DB, item);
      }
      return Response.json({ ok: true, parserExtensions });
    }

    if (action === "deleteParserExtensionRule") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const target = sanitizeParserRuleRemoval(body?.rule || body);
      if (!target) {
        return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      }
      const parserExtensions = removeParserExtensionRuleByKey(await loadD1ParserExtensionRules(context.env.ROSTER_DB), target);
      await deleteD1ParserExtensionRule(context.env.ROSTER_DB, target);
      return Response.json({ ok: true, parserExtensions });
    }

    if (action === "saveLocalParserExtensionRule") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, saveEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const rule = sanitizeParserExtensionRule(body?.rule);
      if (!rule) {
        return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      }
      const localParserExtensions = upsertParserExtensionRule(targetRecord.localParserExtensions, rule);
      const suggestion = sanitizeParserRuleSuggestion({
        email: saveEmail,
        realName: targetRecord.realName || "",
        fingerprint: body?.fingerprint || issueFingerprint(rule.source, rule.code, rule.seniority),
        rawValue: body?.rawValue || rule.code,
        rule,
        status: "pending",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      await upsertAccountMirror(context.env.ROSTER_DB, {
        ...targetRecord,
        localParserExtensions,
        updatedAt: new Date().toISOString(),
      });
      if (suggestion) {
        const suggestions = await loadParserRuleSuggestions(null, context.env.ROSTER_DB);
        await saveParserRuleSuggestions(null, upsertParserRuleSuggestion(suggestions, suggestion), context.env.ROSTER_DB);
      }
      return Response.json({
        ok: true,
        issueConfig: await buildIssueConfig(null, saveEmail, context.env.ROSTER_DB),
      });
    }

    if (action === "decideParserRuleSuggestion") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const suggestionId = String(body?.suggestionId || "").trim();
      const decision = String(body?.decision || "").trim();
      if (!["approveGlobal", "approveUser", "reject"].includes(decision)) {
        return Response.json({ error: "Unsupported suggestion decision." }, { status: 400 });
      }
      const suggestions = await loadParserRuleSuggestions(null, context.env.ROSTER_DB);
      const suggestion = suggestions.find((item) => item.id === suggestionId);
      if (!suggestion) return Response.json({ error: "Suggestion not found." }, { status: 404 });
      const rule = sanitizeParserExtensionRule(body?.rule || suggestion.rule);
      if (!rule) return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      if (decision === "approveGlobal") {
        await saveD1ParserExtensionRule(context.env.ROSTER_DB, rule);
      } else if (decision === "approveUser") {
        const target = await loadAccountMirror(context.env.ROSTER_DB, suggestion.email);
        if (target) {
          await upsertAccountMirror(context.env.ROSTER_DB, {
            ...target,
            localParserExtensions: upsertParserExtensionRule(target.localParserExtensions, rule),
            updatedAt: new Date().toISOString(),
          });
        }
      }
      await saveParserRuleSuggestions(null, suggestions.filter((item) => item.id !== suggestionId), context.env.ROSTER_DB);
      return Response.json({
        ok: true,
        parserExtensions: await loadD1ParserExtensionRules(context.env.ROSTER_DB),
        suggestions: await loadParserRuleSuggestions(null, context.env.ROSTER_DB),
      });
    }

    if (action === "deleteAccount") {
      const deleteEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      if (deleteEmail === CREATOR_EMAIL) {
        return Response.json({ error: "The creator account cannot be deleted." }, { status: 400 });
      }
      if (deleteEmail !== email && account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const record = await loadAccountMirror(context.env.ROSTER_DB, deleteEmail).catch(() => null);
      await deleteAccountMirror(context.env.ROSTER_DB, deleteEmail).catch(() => null);
      await clearDeletedAccountClaimMetadata(context.env.ROSTER_DB, deleteEmail, record);
      return Response.json({ ok: true, deletedEmail: deleteEmail });
    }

    if (action === "save") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountMirror(context.env.ROSTER_DB, saveEmail);
      if (!targetRecord) return Response.json({ error: "Account not found." }, { status: 404 });
      const targetRole = targetRecord.role || roleForEmail(saveEmail);
      const state = sanitizeState(body?.state);
      if (!null && state.imports.some((item) => item?.dataUrl)) {
        return Response.json({ error: "Raw roster file storage is not configured. Upload persistence requires object storage." }, { status: 503 });
      }
      state.imports = state.imports.map(repositoryImportRef);
      const claims = sanitizeClaims(targetRecord.claims);
      const removedImportIds = sanitizeRepositoryFileIds(body?.removedImportIds);
      if ((targetRole === "creator" || targetRole === "owner") && saveEmail === email) {
        const keepFileIds = sanitizeRepositoryFileIds(state.imports.map((item) => item.repoId || item.repositoryId || item.id));
        const allFiles = await queryRosterFiles(context.env.ROSTER_DB, { includeInactive: true }).catch(() => []);
        const persistedKeepFileIds = keepFileIds.filter((id) => allFiles.some((file) => file.id === id));
        const canReconcileToFullSet = keepFileIds.length > 0 && persistedKeepFileIds.length === keepFileIds.length;
        const activeFiles = await queryRosterFileRanges(context.env.ROSTER_DB, { includeInactive: false }).catch(() => []);
        const staleFileIds = canReconcileToFullSet
          ? activeFiles.map((file) => file.id).filter((id) => id && !keepFileIds.includes(id))
          : [];
        const removedIds = [...new Set([...removedImportIds, ...staleFileIds])];
        await Promise.all(removedIds.map((id) => deleteDerivedRosterFile(context.env.ROSTER_DB, id).catch(() => null)));
        await Promise.all(removedIds.map((id) => deleteRawRosterFile(context.env.ROSTER_DB, id).catch(() => null)));
        if (removedIds.length) await refreshCanonicalDoctors(context.env.ROSTER_DB);
        state.imports = state.imports.filter((item) => {
          const repoId = item.repoId || item.repositoryId || item.id;
          return keepFileIds.includes(repoId);
        });
      }
      await replaceAccountCustomEvents(context.env.ROSTER_DB, saveEmail, sanitizeSnapshotCustomEvents(state.session?.customEvents, saveEmail));
      const durableState = {
        ...state,
        session: stripRelationalCustomEventsFromSession(state.session),
      };
      const updatedRecord = {
        ...targetRecord,
        email: saveEmail,
        role: targetRole,
        realName: targetRecord.realName || "",
        claims,
        state: durableState,
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(context.env.ROSTER_DB, updatedRecord);
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
      const profile = await loadDoctorProfileState(null, context.env.ROSTER_DB, profileId) || sanitizeDoctorProfile({
        profileId,
        doctorKey: body?.doctorKey,
        displayName: body?.displayName,
        sourceTypes: body?.sourceTypes,
        state: sanitizeState(null),
      });
      const requestedAliases = sanitizeDoctorAccountResolutionInput({ aliases: body?.aliases }).aliases;
      const snapshotInfo = await loadDoctorProfileSnapshotInfo(null, {
        ...profile,
        aliases: requestedAliases,
      }, context.env.ROSTER_DB, email);
      const diagnostics = snapshotInfo.snapshotAvailable ? null : await doctorProfileLoadDiagnostics(context.env.ROSTER_DB, {
        ...profile,
        aliases: requestedAliases,
      });
      return Response.json({
        ok: true,
        cloudAvailable: true,
        profile,
        snapshot: snapshotInfo.snapshot,
        snapshotAvailable: snapshotInfo.snapshotAvailable,
        snapshotStale: snapshotInfo.snapshotStale,
        snapshotBuiltAt: snapshotInfo.snapshotBuiltAt,
        diagnostics,
        issueConfig: await buildIssueConfig(null, ""),
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
        await deleteDoctorProfileMirror(context.env.ROSTER_DB, profileId).catch(() => null);
        return Response.json({ ok: true, deleted: true });
      }
      const existing = await loadDoctorProfileState(null, context.env.ROSTER_DB, profileId);
      const next = {
        profileId,
        doctorKey,
        displayName,
        sourceTypes,
        state,
        createdAt: existing?.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      await upsertDoctorProfileMirror(context.env.ROSTER_DB, next).catch(() => null);
      return Response.json({ ok: true, profile: next });
    }

    if (action === "loadImports") {
      return Response.json({ error: "Raw roster import loading has been removed. Calendar data is served from D1; re-upload roster files if D1 has no rows." }, { status: 410 });
    }

    if (action === "loadDoctorProfileImports") {
      return Response.json({ error: "Raw doctor profile import loading has been removed. Doctor profile calendars are served from D1." }, { status: 410 });
    }

    if (action === "loadImportRefs") {
      return Response.json({ error: "Raw roster import refs have been removed. Calendar data is served from D1." }, { status: 410 });
    }

    if (action === "queryRosterInsights") {
      if (!hasCalendarDb(context.env)) {
        return Response.json({ ok: false, unavailable: true, coworkers: [] });
      }
      const startDate = String(body?.startDate || body?.date || "").slice(0, 10);
      const endDate = String(body?.endDate || body?.date || startDate).slice(0, 10);
      const sourceTypes = sanitizeSourceTypes(body?.sourceTypes || []);
      const excludeDoctorKeys = (Array.isArray(body?.excludeDoctorKeys) ? body.excludeDoctorKeys : [])
        .map((key) => normalizeRosterName(key))
        .filter(Boolean);
      const doctorKeys = (Array.isArray(body?.doctorKeys) ? body.doctorKeys : [])
        .map((key) => normalizeRosterName(key))
        .filter(Boolean);
      const overlapDoctorKeys = (Array.isArray(body?.overlapDoctorKeys) ? body.overlapDoctorKeys : [])
        .map((key) => normalizeRosterName(key))
        .filter(Boolean);
      const startedAt = Date.now();
      try {
        const coworkers = await queryCoworkerEvents(context.env.ROSTER_DB, {
          startDate,
          endDate,
          sourceTypes,
          excludeDoctorKeys,
          doctorKeys,
          overlapDoctorKeys,
        });
        return Response.json({ ok: true, coworkers, queryMs: Date.now() - startedAt });
      } catch (error) {
        console.error("queryRosterInsights failed", {
          startDate,
          endDate,
          sourceTypes,
          doctorKeyCount: doctorKeys.length,
          overlapDoctorKeyCount: overlapDoctorKeys.length,
          excludeDoctorKeyCount: excludeDoctorKeys.length,
          queryMs: Date.now() - startedAt,
          error: error?.message || String(error),
        });
        return Response.json({ ok: false, unavailable: true, coworkers: [] }, { status: 503 });
      }
    }

    if (action === "queryRosterOverlapDoctors") {
      const startDate = String(body?.startDate || body?.date || "").slice(0, 10);
      const endDate = String(body?.endDate || body?.date || startDate).slice(0, 10);
      const sourceTypes = sanitizeSourceTypes(body?.sourceTypes || []);
      const excludeDoctorKeys = (Array.isArray(body?.excludeDoctorKeys) ? body.excludeDoctorKeys : [])
        .map((key) => normalizeRosterName(key))
        .filter(Boolean);
      const overlapDoctorKeys = (Array.isArray(body?.overlapDoctorKeys) ? body.overlapDoctorKeys : [])
        .map((key) => normalizeRosterName(key))
        .filter(Boolean);
      const startedAt = Date.now();
      try {
        const doctors = await queryOverlapDoctors(context.env.ROSTER_DB, {
          startDate,
          endDate,
          sourceTypes,
          excludeDoctorKeys,
          overlapDoctorKeys,
        });
        return Response.json({ ok: true, doctors, queryMs: Date.now() - startedAt });
      } catch (error) {
        console.error("queryRosterOverlapDoctors failed", {
          startDate,
          endDate,
          sourceTypes,
          overlapDoctorKeyCount: overlapDoctorKeys.length,
          excludeDoctorKeyCount: excludeDoctorKeys.length,
          queryMs: Date.now() - startedAt,
          error: error?.message || String(error),
        });
        return Response.json({ ok: false, unavailable: true, doctors: [] }, { status: 503 });
      }
    }

    if (action === "loadCalendarEvents") {
      if (!hasCalendarDb(context.env)) {
        return Response.json({ ok: false, unavailable: true, snapshot: null });
      }
      const targetRecord = targetEmail && (account.role === "creator" || account.role === "owner")
        ? await loadAccountMirror(context.env.ROSTER_DB, targetEmail)
        : account.record;
      const prepared = await prepareLightweightAccountResponse(targetRecord, { db: context.env.ROSTER_DB });
      const requestedRange = boundedCalendarEventRange({
        startDate: body?.startDate,
        endDate: body?.endDate,
      });
      const diagnostics = {};
      const snapshot = await buildDerivedAccountSnapshot(context.env.ROSTER_DB, {
        role: prepared.role,
        record: targetRecord,
        state: prepared.state,
        claims: prepared.claims,
        index: null,
        startDate: requestedRange.startDate,
        endDate: requestedRange.endDate,
        doctorKey: normalizeRosterName(body?.doctorKey || ""),
        diagnostics,
      });
      return Response.json({
        ok: true,
        snapshot,
        snapshotAvailable: Boolean(snapshot),
        snapshotStale: false,
        snapshotBuiltAt: snapshot?.builtAt || "",
        diagnostics,
      });
    }

    if (action === "loadInsightImports") {
      return Response.json({ error: "Insight import hydration has been removed. Insights query D1 directly." }, { status: 410 });
    }

    return Response.json({ error: "Unsupported account action." }, { status: 400 });
  } catch (error) {
    const message = error.message || "Account request failed.";
    const status = message === "Incorrect password." || message.startsWith("Account not found") ? 401 : 400;
    return Response.json({ error: message }, { status });
  }
}

function validateDerivedCalendarPayload(doctors, eventsByDoctor) {
  const safeDoctors = Array.isArray(doctors) ? doctors.filter((doctor) => doctor?.key) : [];
  if (!safeDoctors.length) {
    return "Roster indexing produced no doctors. The uploaded file was not saved to D1.";
  }
  const eventCount = safeDoctors.reduce((count, doctor) => (
    count + (Array.isArray(eventsByDoctor?.[doctor.key]) ? eventsByDoctor[doctor.key].length : 0)
  ), 0);
  if (!eventCount) {
    return "Roster indexing produced no events. The uploaded file was not saved to D1.";
  }
  if (eventCount < safeDoctors.length) {
    return `Roster indexing produced only ${eventCount} events for ${safeDoctors.length} doctors. The uploaded file was not saved to D1.`;
  }
  return "";
}

async function loadDoctorProfileState(store, db, profileId) {
  return sanitizeDoctorProfile(await loadDoctorProfileMirror(db, profileId).catch(() => null));
}

export function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function roleForEmail(email) {
  return email === CREATOR_EMAIL ? "creator" : "user";
}

async function loadOrCreateD1Account(db, email, password, options = {}) {
  const mode = options.mode || "login";
  const realName = String(options.realName || "").trim();
  const existing = await loadAccountMirror(db, email);
  if (!existing) {
    const canBootstrapCreator = email === CREATOR_EMAIL && mode === "login";
    if (!canBootstrapCreator && (mode !== "create" || !realName)) {
      throw new Error("Account not found. Create an account first.");
    }
    const passwordRecord = await hashPassword(password);
    const now = new Date().toISOString();
    const record = {
      email,
      realName: realName || (email === CREATOR_EMAIL ? "Richard Haydon" : ""),
      role: roleForEmail(email),
      ...passwordRecord,
      state: sanitizeState(null),
      claims: [],
      adminIssues: [],
      localParserExtensions: [],
      subscriptionToken: randomSubscriptionToken(),
      createdAt: now,
      updatedAt: now,
    };
    await upsertAccountMirror(db, record);
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
    const updated = {
      ...existing,
      realName: existing.realName || realName || "",
      ...passwordRecord,
      updatedAt: new Date().toISOString(),
    };
    await upsertAccountMirror(db, updated, { preserveExistingState: false });
    return {
      created: false,
      role: updated.role || roleForEmail(email),
      realName: updated.realName || "",
      state: sanitizeState(updated.state),
      record: updated,
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
    await upsertAccountMirror(db, updated, { preserveExistingState: false });
  }
  return {
    created: false,
    role: updated.role || roleForEmail(email),
    realName: updated.realName || "",
    state: sanitizeState(updated.state),
    record: updated,
  };
}

async function verifyD1Account(db, email, password) {
  const record = await loadAccountMirror(db, email);
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

async function listD1Users(db) {
  const records = await listAccountMirrors(db);
  const users = [];
  for (const record of records) {
    users.push(await userSummaryFromRecord(record.email, record, { db }));
  }
  return users.sort((a, b) => a.email.localeCompare(b.email));
}

async function autoClaimMatchedRosterNames(store, record, db = null) {
  const role = record?.role || roleForEmail(record?.email || "");
  if (!record?.email || role === "creator" || role === "owner") return record;
  const matchedClaims = matchDoctorClaims(await loadSqlDoctorCandidates(db), record.realName || "");
  const claims = mergeClaims(sanitizeClaims(record.claims), matchedClaims);
  if (!claims.length || JSON.stringify(claims) === JSON.stringify(sanitizeClaims(record.claims))) return record;
  const d1Refs = await d1RepositoryImportRefsForClaims(db, claims);
  const state = {
    ...sanitizeState(record.state),
    imports: d1Refs,
  };
  const updated = {
    ...record,
    claims,
    state,
    updatedAt: new Date().toISOString(),
  };
  await upsertAccountMirror(db, updated).catch(() => null);
  return updated;
}

async function autoClaimMatchedCanonicalDoctors(record, db = null) {
  const role = record?.role || roleForEmail(record?.email || "");
  if (!record?.email || role === "creator" || role === "owner") return record;
  const canonicalDoctors = await queryCanonicalDoctors(db).catch(() => []);
  const indexedDoctors = canonicalDoctors.length ? canonicalDoctors : await queryRosterDoctors(db).catch(() => []);
  const claims = mergeClaims(sanitizeClaims(record.claims), matchDoctorClaims(indexedDoctors, record.realName || ""));
  if (!claims.length || JSON.stringify(claims) === JSON.stringify(sanitizeClaims(record.claims))) return record;
  return {
    ...record,
    claims,
    updatedAt: new Date().toISOString(),
  };
}

async function calendarStoreStatus(store, db, options = {}) {
  const allD1Files = await queryRosterFiles(db, { includeInactive: true }).catch(() => []);
  const d1Files = allD1Files.filter((file) => file.active !== false);
  const activeFiles = d1Files;
  const counts = await countDerivedEventsByFile(db, activeFiles.map((file) => file.id));
  const doctorCounts = await countDerivedDoctorsByFile(db, activeFiles.map((file) => file.id));
  const selectedDoctorKey = normalizeRosterName(options.doctorKey || "");
  const selectedDoctorRows = selectedDoctorKey
    ? await resolveSelectedRosterFileDoctorRows(db, selectedDoctorKey)
    : [];
  const selectedPairs = selectedDoctorRows.map((row) => ({ fileId: row.fileId, doctorKey: row.doctorKey }));
  const selectedCounts = await countDerivedEventsByFileDoctorPairs(db, selectedPairs);
  const selectedCountsByFile = new Map();
  for (const row of selectedDoctorRows) {
    selectedCountsByFile.set(row.fileId, (selectedCountsByFile.get(row.fileId) || 0) + Number(selectedCounts.get(`${row.fileId}:${row.doctorKey}`) || 0));
  }
  const rawAvailability = new Map(await Promise.all(activeFiles.map(async (file) => {
    const raw = await loadRawRosterFile(db, file.id).catch(() => null);
    return [file.id, Boolean(raw?.objectKey || raw?.dataUrl)];
  })));
  const files = activeFiles.map((file) => ({
    id: file.id,
    name: file.name,
    sourceType: file.sourceType,
    expectedDoctors: Number(file.expectedDoctors || 0) || sanitizeRepositoryDoctors(file.doctors).length,
    indexedDoctors: Number(file.indexedDoctors || 0) || doctorCounts.get(file.id) || 0,
    eventCount: Number(file.eventCount || 0) || counts.get(file.id) || 0,
    selectedDoctorEventCount: selectedCountsByFile.get(file.id) || 0,
    rawSourceAvailable: rawAvailability.get(file.id) === true,
  })).map((file) => ({
    ...file,
    status: file.eventCount <= 0
      ? "missing"
      : file.expectedDoctors > 0 && file.indexedDoctors < file.expectedDoctors
        ? "partial"
        : "populated",
  }));
  const populated = files.filter((file) => file.status === "populated").length;
  const partial = files.filter((file) => file.status === "partial").length;
  const expectedFiles = summarizeExpectedRosterFiles(allD1Files, options.expectedFileIds);
  return {
    total: files.length,
    populated,
    partial,
    remaining: Math.max(0, files.length - populated),
    eventCount: files.reduce((total, file) => total + file.eventCount, 0),
    selectedDoctorKey,
    selectedDoctorEventCount: files.reduce((total, file) => total + file.selectedDoctorEventCount, 0),
    selectedDoctorFiles: selectedDoctorRows.map(rosterFileDoctorDiagnostic),
    expectedFiles,
    nextFile: files.find((file) => file.status !== "populated") || null,
    files,
  };
}

function summarizeExpectedRosterFiles(allFiles = [], expectedFileIds = []) {
  const expectedIds = sanitizeRepositoryFileIds(expectedFileIds);
  const filesById = new Map((allFiles || []).map((file) => [file.id, file]));
  const persistedFileIds = expectedIds.filter((id) => filesById.has(id));
  const populatedFileIds = persistedFileIds.filter((id) => Number(filesById.get(id)?.eventCount || 0) > 0);
  const activeFileIds = populatedFileIds.filter((id) => filesById.get(id)?.active !== false);
  return {
    expectedCount: expectedIds.length,
    expectedFileIds: expectedIds,
    persistedCount: persistedFileIds.length,
    populatedCount: populatedFileIds.length,
    activeCount: activeFileIds.length,
    persistedFileIds,
    populatedFileIds,
    activeFileIds,
    missingFileIds: expectedIds.filter((id) => !filesById.has(id)),
  };
}

async function reconcileRosterFileSupersession(db, savedFile = {}, options = {}) {
  if (!db?.prepare) return { deactivated: [], ambiguous: [] };
  const files = (await queryRosterFileRanges(db, { includeInactive: false }).catch(() => []))
    .filter((file) => file.active && file.eventCount > 0 && file.startDate && file.endDate);
  const savedId = String(savedFile?.id || "").trim();
  const affected = files.filter((file) => !savedId || file.id === savedId || file.sourceType === String(savedFile?.sourceType || "").toLowerCase());
  const deactivated = [];
  const ambiguous = [];
  for (let leftIndex = 0; leftIndex < affected.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < affected.length; rightIndex += 1) {
      const left = affected[leftIndex];
      const right = affected[rightIndex];
      if (left.sourceType !== right.sourceType || !dateRangesOverlap(left, right)) continue;
      const winner = chooseLatestRosterFile(left, right);
      if (!winner) {
        ambiguous.push({ left, right });
        continue;
      }
      const loser = winner.id === left.id ? right : left;
      if (deactivated.some((file) => file.id === loser.id)) continue;
      await setDerivedRosterFileActive(db, loser.id, false);
      deactivated.push(loser);
    }
  }
  if (ambiguous.length) {
    await reportSupersessionAmbiguity(db, ambiguous, options.uploaderEmail || "");
  }
  return {
    deactivated: deactivated.map((file) => ({ id: file.id, name: file.name, sourceType: file.sourceType })),
    ambiguous: ambiguous.map(({ left, right }) => ({
      sourceType: left.sourceType,
      files: [left, right].map((file) => ({ id: file.id, name: file.name, startDate: file.startDate, endDate: file.endDate })),
    })),
  };
}

function dateRangesOverlap(left, right) {
  const leftEnd = String(left.coverageEndDate || left.endDate || "");
  const rightEnd = String(right.coverageEndDate || right.endDate || "");
  return String(left.startDate || "") <= rightEnd && String(right.startDate || "") <= leftEnd;
}

function chooseLatestRosterFile(left, right) {
  if (Number(left.lastModified || 0) && Number(right.lastModified || 0) && Number(left.lastModified) !== Number(right.lastModified)) {
    return Number(left.lastModified) > Number(right.lastModified) ? left : right;
  }
  const leftNamedDate = rosterFileNameDate(left.name);
  const rightNamedDate = rosterFileNameDate(right.name);
  if (leftNamedDate && rightNamedDate && leftNamedDate !== rightNamedDate) {
    return leftNamedDate > rightNamedDate ? left : right;
  }
  return null;
}

function rosterFileNameDate(name = "") {
  const value = String(name || "");
  const rangeMatch = value.match(/(\d{2})[-_](\d{2})[-_](\d{4}).*?(?:to|_to_).*?(\d{2})[-_](\d{2})[-_](\d{4})/i);
  if (rangeMatch) return `${rangeMatch[6]}-${rangeMatch[5]}-${rangeMatch[4]}`;
  const termMatch = value.match(/term\s*([1-4])\D+(\d{4})/i);
  if (termMatch) return `${termMatch[2]}-${String(Number(termMatch[1]) * 3).padStart(2, "0")}-01`;
  return "";
}

async function reportSupersessionAmbiguity(db, ambiguousPairs = [], uploaderEmail = "") {
  const creator = await loadAccountMirror(db, CREATOR_EMAIL).catch(() => null);
  if (!creator) return;
  const now = new Date().toISOString();
  const issues = ambiguousPairs.map(({ left, right }) => ({
    id: `supersession:${left.sourceType}:${left.id}:${right.id}`,
    message: `Could not determine the latest ${left.sourceType.toUpperCase()} roster between ${left.name} and ${right.name}.`,
    source: left.sourceType.toUpperCase(),
    seniority: "Unknown",
    rawValue: `Roster supersession review: ${left.name} vs ${right.name}`,
    fingerprint: `${left.sourceType.toUpperCase()}::Unknown::Roster supersession review::${left.id}::${right.id}`,
    firstSeenAt: now,
    lastSeenAt: now,
    count: 1,
  }));
  await upsertAccountMirror(db, {
    ...creator,
    adminIssues: mergeAdminIssues(creator.adminIssues, issues),
    updatedAt: now,
  });
}

function manualRosterClaimIssue(record, claim) {
  const now = new Date().toISOString();
  const displayName = String(claim?.displayName || claim?.key || "").trim();
  const source = String(claim?.sourceType || "").toUpperCase() || "MMC";
  const accountName = String(record?.realName || record?.email || "").trim();
  return {
    id: `identity-claim:${normalizeEmail(record?.email)}:${claim?.sourceType || ""}:${claim?.key || ""}`,
    message: `${accountName || record?.email || "User"} manually linked roster name ${displayName || claim?.key || "Unknown"}.`,
    source,
    seniority: "Unknown",
    rawValue: `Manual roster claim review: ${accountName || record?.email || ""} -> ${displayName || claim?.key || ""}`,
    fingerprint: `${source}::Unknown::Manual roster claim review::${normalizeEmail(record?.email)}::${claim?.sourceType || ""}:${claim?.key || ""}`,
    firstSeenAt: now,
    lastSeenAt: now,
    count: 1,
  };
}

async function userSummaryFromRecord(email, record, options = {}) {
  const claims = sanitizeClaims(record?.claims);
  const adminIssues = sanitizeAdminIssues(record?.adminIssues);
  const seniorities = options.db
    ? await queryDoctorSeniorities(options.db, claims.map((claim) => claim.key)).catch(() => [])
    : [];
  return {
    email,
    realName: String(record?.realName || "").trim(),
    role: record?.role || roleForEmail(email),
    sites: [...new Set(claims.map((claim) => claim.sourceType.toUpperCase()))].sort(),
    seniorities,
    claims,
    insightsEnabled: insightsEnabledForRecord(record),
    adminIssues,
    issuesCount: adminIssues.length,
    createdAt: record?.createdAt || "",
    updatedAt: record?.updatedAt || "",
  };
}

function insightsEnabledForRecord(record) {
  const role = record?.role || roleForEmail(normalizeEmail(record?.email));
  if (role === "creator" || role === "owner") return true;
  return record?.insightsEnabled === true;
}

async function prepareLightweightAccountResponse(rawRecord, options = {}) {
  const record = rawRecord;
  const role = record.role || roleForEmail(record.email);
  let state = sanitizeState(record.state);
  const mirroredState = await loadAccountStateMirror(options.db, record.email).catch(() => null);
  if (mirroredState?.session) {
    state = {
      ...state,
      session: {
        ...state.session,
        ...mirroredState.session,
      },
    };
  }
  state = await applySqlHospitalLocationSettings(options.db, record.email, state);
  const d1CustomEvents = await queryAccountCustomEvents(options.db, record.email).catch(() => []);
  if (d1CustomEvents.length) {
    state = {
      ...state,
      session: {
        ...state.session,
        customEvents: latestCustomEventsByIdentity([
          ...sanitizeSnapshotCustomEvents(state.session?.customEvents, record.email),
          ...d1CustomEvents,
        ]),
      },
    };
  }
  const claims = sanitizeClaims(record.claims);
  if ((role === "creator" || role === "owner") && options.db) {
    const files = await queryRosterFileRanges(options.db, { includeInactive: false }).catch(() => []);
    state = {
      ...state,
      imports: files.map(repositoryImportRef),
    };
  } else if (claims.length && options.db) {
    state = {
      ...state,
      imports: await d1RepositoryImportRefsForClaims(options.db, claims),
    };
  }
  return {
    role,
    realName: record.realName || "",
    state: applyDefaultSelectedDoctorToState(state, role, claims),
    claims,
    subscription: {
      token: String(record.subscriptionToken || ""),
      enabled: Boolean(record.subscriptionToken),
    },
    insightsEnabled: insightsEnabledForRecord(record),
  };
}

async function applySqlHospitalLocationSettings(db, email, state) {
  const locations = await loadAccountHospitalLocations(db, email, state?.session).catch(() => null);
  if (!locations) return state;
  return {
    ...state,
    session: {
      ...(state.session || {}),
      settings: mergeHospitalLocationsIntoSettings(state.session?.settings || {}, locations),
    },
  };
}

export async function prepareAccountResponse(store, rawRecord, options = {}) {
  let record = await ensureAccountSubscriptionToken(store, rawRecord);
  if (record?.subscriptionToken && record.subscriptionToken !== rawRecord?.subscriptionToken) {
    await upsertAccountMirror(options.db, record, { preserveExistingState: true }).catch(() => null);
  }
  const role = record.role || roleForEmail(record.email);
  let claims = sanitizeClaims(record.claims);
  let nameMatches = [];
  let state = sanitizeState(record.state);
  const mirroredState = await loadAccountStateMirror(options.db, record.email).catch(() => null);
  if (mirroredState?.session) {
    state = {
      ...state,
      session: {
        ...state.session,
        ...mirroredState.session,
      },
    };
  }
  state = await applySqlHospitalLocationSettings(options.db, record.email, state);
  const d1CustomEvents = await queryAccountCustomEvents(options.db, record.email).catch(() => []);
  if (d1CustomEvents.length) {
    state = {
      ...state,
      session: {
        ...state.session,
        customEvents: latestCustomEventsByIdentity([
          ...sanitizeSnapshotCustomEvents(state.session?.customEvents, record.email),
          ...d1CustomEvents,
        ]),
      },
    };
  } else if (state.session?.customEvents?.length) {
    await replaceAccountCustomEvents(options.db, record.email, sanitizeSnapshotCustomEvents(state.session.customEvents, record.email)).catch(() => null);
  }
  let linkedProfiles = [];

  if (role !== "creator" && role !== "owner") {
    const originalClaims = claims;
    const matchedClaims = claims.length ? [] : matchDoctorClaims(await loadSqlDoctorCandidates(options.db), record.realName || "");
    nameMatches = matchedClaims.filter((claim) => !claims.some((existing) => sameClaim(existing, claim)));
    linkedProfiles = await linkedDoctorProfilesForClaims(store, claims, options.db);
    const d1Refs = await d1RepositoryImportRefsForClaims(options.db, claims);
    const accountImportRefs = d1Refs;
    state = {
      ...state,
      imports: accountImportRefs,
    };
    state = mergeProfileSessionIntoState(state, linkedProfiles, record.email);
    const previousState = sanitizeState(record.state);
    const importRefsChanged = importsChanged(previousState.imports, accountImportRefs);
    const claimsChanged = JSON.stringify(claims) !== JSON.stringify(originalClaims);
    if (claimsChanged || importRefsChanged) {
      const updatedRecord = {
        ...record,
        claims,
        state: {
          ...previousState,
          imports: accountImportRefs.map(repositoryImportRef),
        },
        updatedAt: new Date().toISOString(),
      };
      await upsertAccountMirror(options.db, updatedRecord).catch(() => null);
    }
  } else {
    const d1Files = await queryRosterFileRanges(options.db, { includeInactive: false }).catch(() => []);
    const imported = {
      index: { version: 1, files: d1Files },
      refs: d1Files.map(repositoryImportRef),
      changed: false,
    };
    const creatorRepositoryRefs = (imported.index.files || []).map(repositoryImportRef);
    const stateWithRefs = { ...state, imports: creatorRepositoryRefs };
    state = stateWithRefs;
  }
  state = applyDefaultSelectedDoctorToState(state, role, claims);

  const snapshotAvailable = false;
  const snapshotStale = false;
  const issueConfig = await buildIssueConfig(store, record.email, options.db);

  return {
    role,
    realName: record.realName || "",
    state,
    claims,
    nameMatches,
    availableDoctors: options.includeAvailableDoctors === false ? [] : await repositoryDoctorCandidates(store, null, options.db),
    subscription: {
      token: String(record.subscriptionToken || ""),
      enabled: Boolean(record.subscriptionToken),
    },
    insightsEnabled: insightsEnabledForRecord(record),
    adminIssues: sanitizeAdminIssues(record.adminIssues),
    issueConfig,
    snapshot: null,
    snapshotAvailable,
    snapshotStale,
    snapshotBuiltAt: "",
    snapshotBuildStamp: "",
  };
}

function applyDefaultSelectedDoctorToState(state, role, claims = []) {
  const session = state.session && typeof state.session === "object" ? state.session : {};
  const existingKey = normalizeRosterName(session.doctorKey || "");
  const defaultKey = role === "creator" || role === "owner"
    ? OWNER_DOCTOR_KEY
    : sanitizeClaims(claims)[0]?.key || "";
  return {
    ...state,
    session: {
      ...session,
      doctorKey: existingKey || defaultKey,
    },
  };
}

function boundedCalendarEventRange(input = {}) {
  const now = new Date();
  const currentYear = now.getUTCFullYear();
  const requestedStart = dateKeyOrEmpty(input.startDate);
  const requestedEnd = dateKeyOrEmpty(input.endDate);
  let startDate = requestedStart || `${currentYear}-01-01`;
  let endDate = requestedEnd || `${currentYear + 1}-01-31`;
  if (endDate < startDate) endDate = startDate;
  const maxEndDate = isoDateKey(addUtcDays(startDate, 370));
  if (endDate > maxEndDate) endDate = maxEndDate;
  return { startDate, endDate };
}

function dateKeyOrEmpty(value) {
  const key = String(value || "").slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(key) ? key : "";
}

function isoDateKey(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value || "").slice(0, 10);
}

function addUtcDays(dateKey, days) {
  const [year, month, day] = String(dateKey || "").split("-").map((part) => Number(part));
  const date = new Date(Date.UTC(year || 1970, (month || 1) - 1, day || 1));
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return date;
}

async function buildDerivedAccountSnapshot(db, context) {
  if (!hasCalendarDb({ ROSTER_DB: db })) return null;
  const role = context.role || "user";
  const state = sanitizeState(context.state);
  let doctorKeys = [];
  let doctorPairs = [];
  let doctorDiagnostics = [];
  let doctorOptions = [];
  let selectedKey = "";
  let selectedSourceTypes = [];
  if (role === "creator" || role === "owner") {
    const requestedKey = normalizeRosterName(context.doctorKey || state.session?.doctorKey || "");
    selectedKey = requestedKey || OWNER_DOCTOR_KEY;
    let resolved = await resolveSelectedDoctorForCalendarSnapshot(db, selectedKey);
    if (requestedKey && selectedKey !== OWNER_DOCTOR_KEY && !resolved.doctor) {
      selectedKey = OWNER_DOCTOR_KEY;
      resolved = await resolveSelectedDoctorForCalendarSnapshot(db, selectedKey);
    }
    const doctor = resolved.doctor;
    if (!doctor) return null;
    doctorKeys = doctorKeysForOption(doctor);
    doctorDiagnostics = resolved.rows;
    doctorPairs = doctorDiagnostics.map((row) => ({ fileId: row.fileId, doctorKey: row.doctorKey }));
    doctorOptions = [doctor];
    selectedSourceTypes = doctor?.sourceTypes || [doctor?.sourceType].filter(Boolean);
  } else {
    const claims = sanitizeClaims(context.claims);
    if (!claims.length) return null;
    const groupedClaims = buildCreatorDoctorOptions(claims.map((claim) => ({
      key: claim.key,
      displayName: claim.displayName,
      sourceType: claim.sourceType,
    })));
    const requestedKey = normalizeRosterName(context.doctorKey || state.session?.doctorKey || "");
    const selectedClaimOption = findDoctorOptionByKey(groupedClaims, requestedKey) || groupedClaims[0];
    selectedKey = selectedClaimOption?.key || claims[0].key;
    doctorKeys = doctorKeysForOption(selectedClaimOption);
    doctorOptions = groupedClaims;
    selectedSourceTypes = selectedClaimOption?.sourceTypes || [selectedClaimOption?.sourceType].filter(Boolean);
  }
  const session = state.session && typeof state.session === "object" ? state.session : {};
  const normalizedSession = {
    ...session,
    doctorKey: selectedKey,
  };
  const settings = {
    ...defaultSettings(),
    ...(normalizedSession.settings || {}),
    dateFrom: context.startDate || normalizedSession.settings?.dateFrom || "",
    dateTo: context.endDate || normalizedSession.settings?.dateTo || "",
  };
  const rosterEvents = doctorPairs.length
    ? await queryDoctorEventsForFileDoctorPairs(db, doctorPairs, {
        startDate: context.startDate || "",
        endDate: context.endDate || "",
      })
    : await queryDoctorEvents(db, doctorKeys, {
        startDate: context.startDate || "",
        endDate: context.endDate || "",
      });
  if (context.diagnostics && typeof context.diagnostics === "object") {
    context.diagnostics.selectedDoctorKey = selectedKey;
    context.diagnostics.selectedDoctorFiles = doctorDiagnostics.map(rosterFileDoctorDiagnostic);
    context.diagnostics.queryMode = doctorPairs.length ? "file-doctor-pairs" : "doctor-keys";
  }
  const d1CustomEvents = await queryAccountCustomEvents(db, context.record.email).catch(() => []);
  const hospitalLocations = await loadAccountHospitalLocations(db, context.record.email, normalizedSession).catch(() => null);
  const events = [
    ...applyEventOverrides(applyAccountHospitalLocations(rosterEvents, hospitalLocations || {}, { includeLocations: settings.includeLocations !== false }), normalizedSession.overrides || {}),
    ...customEventsToEvents(latestCustomEventsByIdentity([
      ...sanitizeSnapshotCustomEvents(normalizedSession.customEvents, context.record.email),
      ...d1CustomEvents,
    ]), settings),
  ];
  if (!events.length) return null;
  return sanitizeSnapshotRecord({
    ownerType: role === "creator" || role === "owner" ? "creator-account" : "claimed-account",
    ownerId: normalizeEmail(context.record.email),
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    builtAt: new Date().toISOString(),
    buildStamp: "d1-derived",
    preview: buildPreviewFromDerivedEvents(events, { customEventsMaterialized: true }),
    session: normalizedSession,
    doctorOptions,
    detectedSources: detectedSourcesForSnapshot(
      role === "creator" || role === "owner"
        ? sanitizeSnapshotFileRefs(state.imports)
        : selectedSourceTypes,
    ),
    fileRefs: sanitizeSnapshotFileRefs(state.imports),
    subscriptionFeeds: {},
    insightCache: null,
  });
}

async function resolveSelectedDoctorForCalendarSnapshot(db, doctorKey) {
  const normalizedKey = normalizeRosterName(doctorKey || "");
  if (!normalizedKey) return { doctor: null, rows: [] };
  const rows = await queryRosterFileDoctorsForKeys(db, selectedDoctorKeyVariants(normalizedKey)).catch(() => []);
  if (!rows.length) return { doctor: null, rows: [] };
  return {
    doctor: buildDoctorOptionFromRows([...rows]),
    rows,
  };
}

function selectedDoctorKeyVariants(doctorKey) {
  const normalized = normalizeRosterName(doctorKey || "");
  const tokens = normalized.split(" ").filter(Boolean);
  const variants = [normalized];
  if (tokens.length === 2) variants.push(`${tokens[1]} ${tokens[0]}`);
  return [...new Set(variants.filter(Boolean))];
}

function rawRosterObjectKey(fileId) {
  return `rosters/${String(fileId || "").trim()}`;
}

async function readRetainedRawRosterFile(env, fileId) {
  const raw = await loadRawRosterFile(env.ROSTER_DB, fileId);
  if (!raw) return null;
  if (raw.objectKey && env.ROSTER_FILES?.get) {
    const object = await env.ROSTER_FILES.get(raw.objectKey);
    if (object) {
      const bytes = new Uint8Array(await object.arrayBuffer());
      return {
        ...raw,
        type: raw.type || object.httpMetadata?.contentType || "application/octet-stream",
        dataUrl: bytesToDataUrl(bytes, raw.type || object.httpMetadata?.contentType || "application/octet-stream"),
      };
    }
  }
  if (raw.dataUrl && env.ROSTER_FILES?.put) {
    const objectKey = raw.objectKey || rawRosterObjectKey(fileId);
    await env.ROSTER_FILES.put(objectKey, dataUrlToBytes(raw.dataUrl), {
      httpMetadata: { contentType: raw.type || "application/octet-stream" },
    });
    await upsertRawRosterFile(env.ROSTER_DB, { id: fileId }, {
      objectKey,
      dataUrl: "",
      type: raw.type,
      uploadedAt: raw.uploadedAt,
    });
    return { ...raw, objectKey };
  }
  return raw.dataUrl ? raw : null;
}

function dataUrlToBytes(dataUrl) {
  const base64 = String(dataUrl || "").split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function bytesToDataUrl(bytes, type) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:${type || "application/octet-stream"};base64,${btoa(binary)}`;
}

function findDoctorOptionByKey(options, key) {
  const normalizedKey = normalizeRosterName(key || "");
  if (!normalizedKey) return null;
  return (options || []).find((doctor) => {
    if (normalizeRosterName(doctor?.key || "") === normalizedKey) return true;
    return (doctor?.aliases || []).some((alias) => normalizeRosterName(alias?.key || "") === normalizedKey);
  }) || null;
}

function doctorKeysForOption(doctor) {
  const keys = [
    doctor?.key,
    ...(Array.isArray(doctor?.aliases) ? doctor.aliases.map((alias) => alias.key) : []),
  ];
  return [...new Set(keys.map((key) => normalizeRosterName(key || "")).filter(Boolean))];
}

async function creatorDoctorOptionsForD1(db, index) {
  return await mergedDoctorOptionsFromD1(db, { includeZeroEventStandalone: true });
}

async function loadSqlDoctorCandidates(db) {
  const mergedDoctors = await mergedDoctorOptionsFromD1(db, { includeZeroEventStandalone: true });
  if (mergedDoctors.length) return mergedDoctors;
  const rosterDoctors = await queryRosterDoctors(db).catch(() => []);
  if (rosterDoctors.length) return rosterDoctors;
  return [];
}

async function mergedDoctorOptionsFromD1(db, options = {}) {
  const [canonicalDoctors, doctorRows] = await Promise.all([
    queryCanonicalDoctors(db, { includeZeroEventStandalone: options.includeZeroEventStandalone !== false }).catch(() => []),
    queryRosterFileDoctors(db).catch(() => []),
  ]);
  const rowDoctors = doctorRows.length
    ? await buildCanonicalDoctorOptionsFromRows(db, doctorRows, {
        includeZeroEventStandalone: options.includeZeroEventStandalone === true,
      })
    : [];
  return mergeDoctorOptions(canonicalDoctors, rowDoctors, options);
}

function mergeDoctorOptions(primary = [], fallback = [], options = {}) {
  const byIdentity = new Map();
  const includeZero = options.includeZeroEventStandalone === true;
  for (const doctor of [...primary, ...fallback]) {
    const normalized = normalizeDoctorOptionForMerge(doctor);
    if (!normalized?.key || !normalized.displayName) continue;
    const marker = `${normalized.sourceType || ""}:${normalized.key}`;
    if (byIdentity.has(marker)) {
      const existing = byIdentity.get(marker);
      byIdentity.set(marker, {
        ...existing,
        ...normalized,
        sourceTypes: [...new Set([...(existing.sourceTypes || []), ...(normalized.sourceTypes || [])])],
        aliases: mergeDoctorAliases(existing.aliases || [], normalized.aliases || []),
        hasEvents: existing.hasEvents !== false || normalized.hasEvents !== false,
      });
      continue;
    }
    byIdentity.set(marker, normalized);
  }
  return [...byIdentity.values()]
    .filter((doctor) => includeZero || doctor.hasEvents !== false)
    .sort((left, right) => left.displayName.localeCompare(right.displayName) || String(left.sourceType || "").localeCompare(String(right.sourceType || "")));
}

function normalizeDoctorOptionForMerge(doctor) {
  const sourceTypes = sanitizeSourceTypes(doctor?.sourceTypes || (doctor?.sourceType ? [doctor.sourceType] : []));
  const sourceType = String(doctor?.sourceType || sourceTypes[0] || "").toLowerCase();
  const key = normalizeRosterName(doctor?.key || doctor?.doctorKey || "");
  const displayName = formatRosterDisplayName(doctor?.displayName || doctor?.key || doctor?.doctorKey || "");
  const aliases = Array.isArray(doctor?.aliases)
    ? doctor.aliases.map((alias) => ({
        ...alias,
        key: normalizeRosterName(alias?.key || ""),
        displayName: formatRosterDisplayName(alias?.displayName || alias?.key || ""),
        sourceType: String(alias?.sourceType || "").toLowerCase(),
      })).filter((alias) => alias.key && alias.displayName && isRosterSourceType(alias.sourceType))
    : sourceTypes.map((item) => ({
        sourceType: item,
        key,
        displayName,
      })).filter((alias) => alias.key && alias.displayName && isRosterSourceType(alias.sourceType));
  return {
    ...doctor,
    key,
    displayName,
    sourceType: isRosterSourceType(sourceType) ? sourceType : sourceTypes[0] || "",
    sourceTypes,
    aliases,
    hasEvents: doctor?.hasEvents !== false,
  };
}

function mergeDoctorAliases(left = [], right = []) {
  const byMarker = new Map();
  for (const alias of [...left, ...right]) {
    const sourceType = String(alias?.sourceType || "").toLowerCase();
    const key = normalizeRosterName(alias?.key || "");
    if (!sourceType || !key) continue;
    byMarker.set(`${sourceType}:${key}`, {
      ...alias,
      sourceType,
      key,
      displayName: String(alias?.displayName || key).trim(),
    });
  }
  return [...byMarker.values()];
}

async function resolveSelectedRosterFileDoctorRows(db, doctorKey) {
  const doctorRows = await queryRosterFileDoctors(db).catch(() => []);
  if (!doctorRows.length) return [];
  const selectedOption = await resolveCanonicalDoctorOptionForKey(db, doctorRows, doctorKey);
  return await resolveRosterFileDoctorRows(db, {
    doctorKey,
    doctorRows,
    doctorOptions: selectedOption ? [selectedOption] : [],
  });
}

async function resolveRosterFileDoctorRows(db, options = {}) {
  const doctorRows = options.doctorRows || await queryRosterFileDoctors(db).catch(() => []);
  if (!doctorRows.length) return [];
  const requestedKey = normalizeRosterName(options.doctorKey || "");
  const selectedOption = findDoctorOptionByKey(options.doctorOptions || [], requestedKey);
  const candidateKeys = new Set([requestedKey, ...doctorKeysForOption(selectedOption)]);
  const candidateIdentities = new Set([
    rosterIdentityKey(options.doctorKey || ""),
    rosterIdentityKey(selectedOption?.displayName || ""),
    ...(selectedOption?.aliases || []).flatMap((alias) => [
      rosterIdentityKey(alias.displayName || ""),
      rosterIdentityKey(alias.key || ""),
    ]),
  ].filter(Boolean));
  const matched = [];
  for (const row of doctorRows) {
    const rowKey = normalizeRosterName(row.doctorKey || "");
    const rowIdentity = rosterIdentityKey(row.displayName || row.doctorKey || "");
    const directMatch = candidateKeys.has(rowKey) || candidateIdentities.has(rowIdentity);
    const fuzzyMatch = [...candidateIdentities].some((identity) => identity && (nameTokenMatch(identity, rowIdentity) || likelySameRosterName(identity, rowIdentity)));
    if (!directMatch && !fuzzyMatch) continue;
    const existing = matched.find((item) => item.fileId === row.fileId);
    if (existing && existing.eventCount >= row.eventCount) continue;
    if (existing) {
      const index = matched.indexOf(existing);
      matched.splice(index, 1);
    }
    matched.push(row);
  }
  return matched.sort((left, right) => left.fileId.localeCompare(right.fileId) || left.sourceType.localeCompare(right.sourceType));
}

function rosterFileDoctorDiagnostic(row) {
  return {
    fileId: row.fileId,
    fileName: row.fileName || "",
    sourceType: row.sourceType || row.fileSourceType || "",
    doctorKey: row.doctorKey,
    displayName: row.displayName,
    eventCount: Number(row.eventCount || 0),
  };
}

function sanitizeSnapshotCustomEvents(items, defaultOwnerEmail = "") {
  const ownerEmail = normalizeEmail(defaultOwnerEmail);
  const events = (Array.isArray(items) ? items : [])
    .filter((item) => item && item.id && item.title && item.startDate && item.endDate)
    .map((item) => ({
      id: String(item.id),
      ownerEmail: normalizeEmail(item.ownerEmail || ownerEmail),
      title: String(item.title),
      startDate: String(item.startDate).slice(0, 10),
      endDate: String(item.endDate).slice(0, 10),
      allDay: item.allDay === true,
      startTime: item.allDay ? "" : String(item.startTime || ""),
      endTime: item.allDay ? "" : String(item.endTime || ""),
      location: String(item.location || ""),
      include: item.include !== false,
    }))
    .filter((item) => !ownerEmail || !item.ownerEmail || item.ownerEmail === ownerEmail);
  return latestCustomEventsByIdentity(events);
}

function stripRelationalCustomEventsFromSession(session) {
  if (!session || typeof session !== "object") return {};
  const { customEvents, ...rest } = session;
  return rest;
}

function latestCustomEventsById(events) {
  const byId = new Map();
  for (const event of events || []) {
    byId.delete(event.id);
    byId.set(event.id, event);
  }
  return [...byId.values()];
}

function latestCustomEventsByIdentity(events) {
  const byIdentity = new Map();
  for (const event of latestCustomEventsById(events)) {
    const key = [
      normalizeEmail(event.ownerEmail),
      event.title,
      event.startDate,
      event.endDate,
      event.allDay ? "all-day" : `${event.startTime}|${event.endTime}`,
      event.location,
    ].join("|");
    byIdentity.delete(key);
    byIdentity.set(key, event);
  }
  return [...byIdentity.values()];
}

async function buildIssueConfig(store, email = "", db = null) {
  const globalParserExtensions = await loadD1ParserExtensionRules(db);
  const record = email
    ? await loadAccountMirror(db, email).catch(() => null)
    : null;
  const role = record?.role || roleForEmail(email);
  const localParserExtensions = sanitizeParserExtensionRules(record?.localParserExtensions);
  return {
    parserExtensions: mergeParserExtensionSets(globalParserExtensions, localParserExtensions),
    globalParserExtensions,
    localParserExtensions,
    parserRuleSuggestions: (role === "creator" || role === "owner") ? await loadParserRuleSuggestions(null, db) : [],
    dismissedFingerprints: [],
    ignoredFingerprints: [],
  };
}

async function clearIssueFromAllUsers(storeOrDb, fingerprint) {
  const normalizedFingerprint = sanitizeIssueFingerprint(fingerprint);
  if (!normalizedFingerprint) return;
  const sourceRecords = await listAccountMirrors(storeOrDb).catch(() => []);
  for (const record of sourceRecords) {
    if (!record?.adminIssues?.length) continue;
    const nextIssues = sanitizeAdminIssues(record.adminIssues).filter((issue) => issue.fingerprint !== normalizedFingerprint);
    if (nextIssues.length === sanitizeAdminIssues(record.adminIssues).length) continue;
    await persistAccountRecord(storeOrDb, {
      ...record,
      adminIssues: nextIssues,
      updatedAt: new Date().toISOString(),
    });
  }
}

async function clearIssuesResolvedByParserRule(storeOrDb, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedRule) return;
  const sourceRecords = await listAccountMirrors(storeOrDb).catch(() => []);
  for (const record of sourceRecords) {
    if (!record?.adminIssues?.length) continue;
    const existingIssues = sanitizeAdminIssues(record.adminIssues);
    const nextIssues = existingIssues.filter((issue) => !issueMatchesParserRule(issue, normalizedRule));
    if (nextIssues.length === existingIssues.length) continue;
    await persistAccountRecord(storeOrDb, {
      ...record,
      adminIssues: nextIssues,
      updatedAt: new Date().toISOString(),
    });
  }
}

async function clearIssuesResolvedByParserRuleForUser(storeOrDb, email, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedRule || !normalizedEmail) return;
  const record = await loadAccountMirror(storeOrDb, normalizedEmail).catch(() => null);
  if (!record?.adminIssues?.length) return;
  const existingIssues = sanitizeAdminIssues(record.adminIssues);
  const nextIssues = existingIssues.filter((issue) => !issueMatchesParserRule(issue, normalizedRule));
  if (nextIssues.length === existingIssues.length) return;
  await persistAccountRecord(storeOrDb, {
    ...record,
    adminIssues: nextIssues,
    updatedAt: new Date().toISOString(),
  });
}

function issueMatchesParserRule(issue, rule) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const issueCode = parserRuleCodeForIssue(issue);
  const fingerprint = sanitizeIssueFingerprint(issue?.fingerprint);
  const ruleFingerprint = issueFingerprint(rule.source, rule.code, rule.seniority);
  return source === rule.source
    && seniority === rule.seniority
    && (issueCode === rule.code || fingerprint === ruleFingerprint);
}

async function isIssueResolvedByParserRules(store, email, issue, db = null) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = parserRuleCodeForIssue(issue);
  if (!source || !code) return false;
  if (isKnownResolvedShiftCodeValue(source, issue?.rawValue || code)) return true;
  const config = await buildIssueConfig(store, email, db);
  const rules = sanitizeParserExtensionRules(config.parserExtensions);
  const sourceRules = rules[source.toLowerCase()] || [];
  if (seniority === "Unknown") {
    return sourceRules.some((rule) => rule.source === source && rule.code === code);
  }
  return sourceRules.some((rule) => rule.source === source && rule.seniority === seniority && rule.code === code);
}

async function clearIssuesResolvedByIssue(storeOrDb, email, issue) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return;
  const record = await loadAccountMirror(storeOrDb, normalizedEmail).catch(() => null);
  if (!record?.adminIssues?.length) return;
  const existingIssues = sanitizeAdminIssues(record.adminIssues);
  const nextIssues = existingIssues.filter((item) => !sameParserIssue(item, issue));
  if (nextIssues.length === existingIssues.length) return;
  await persistAccountRecord(storeOrDb, {
    ...record,
    adminIssues: nextIssues,
    updatedAt: new Date().toISOString(),
  });
}

async function persistAccountRecord(storeOrDb, record) {
  if (!record?.email) return;
  await upsertAccountMirror(storeOrDb, record);
}

function sameParserIssue(left, right) {
  return sanitizeIssueSource(left?.source) === sanitizeIssueSource(right?.source)
    && sanitizeRuleSeniority(left?.seniority) === sanitizeRuleSeniority(right?.seniority)
    && parserRuleCodeForIssue(left) === parserRuleCodeForIssue(right);
}

function parserRuleCodeForIssue(issue) {
  const code = String(issue?.code || "").trim().toUpperCase();
  if (code) return code;
  return parserRuleCodeFromRawValue(issue?.source, issue?.rawValue);
}

function parserRuleCodeFromRawValue(sourceValue, rawValue) {
  const source = sanitizeIssueSource(sourceValue);
  const text = String(rawValue || "").trim();
  const upper = text.toUpperCase();
  if (source === "DDH") return normalizeDdhParserRuleCodeText(text);
  if (source === "MMC" || source === "Casey") {
    const prefixMatch = upper.match(/^\s*\d{2}:?\d{2}\s*[-–]\s*\d{2}:?\d{2}\s+(.+?)\s*$/);
    if (prefixMatch) return prefixMatch[1].trim().toUpperCase();
    const suffixMatch = upper.match(/^\s*(.+?)\s+\d{2}:?\d{2}\s*[-–]\s*\d{2}:?\d{2}\s*$/);
    if (suffixMatch) return suffixMatch[1].trim().toUpperCase();
  }
  return upper;
}

function normalizeDdhParserRuleCodeText(value) {
  const text = String(value || "").trim();
  const upper = text.toUpperCase();
  const aliases = new Map([
    ["CLINICAL SUPPORT", "CS"],
    ["SSU SMS", "SSU"],
    ["ORANGE PM (ON-CALL)", "ORANGE PM"],
    ["PM FAST IC", "FAST PM"],
    ["ORANGE AM IC", "ORANGE AM"],
    ["ONSITE CS", "CS ONSITE"],
  ]);
  return aliases.get(upper) || upper;
}

function isKnownResolvedShiftCodeValue(sourceValue, rawValue) {
  const source = sanitizeIssueSource(sourceValue);
  const code = parserRuleCodeFromRawValue(source, rawValue);
  if (source === "DDH") {
    if (["CS", "CS ONSITE", "SSU"].includes(code)) return true;
    if (/^(ORANGE|SILVER|FAST|AVAO)\s+(AM|PM)$/.test(code)) return true;
  }
  return false;
}

function sanitizeRepositoryFileIds(ids = []) {
  return [...new Set((Array.isArray(ids) ? ids : [])
    .map((id) => String(id || "").trim())
    .filter(Boolean))];
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
      displayName: formatRosterDisplayName(doctor?.displayName || doctor?.key || ""),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
    }))
    .filter((doctor) => doctor.key && doctor.displayName);
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

function sanitizeAvailableDoctors(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((doctor) => ({
      key: normalizeRosterName(doctor?.key || ""),
      displayName: formatRosterDisplayName(doctor?.displayName || doctor?.key || ""),
      sourceType: String(doctor?.sourceType || "").toLowerCase(),
      sourceTypes: sanitizeSourceTypes(doctor?.sourceTypes || (doctor?.sourceType ? [doctor.sourceType] : [])),
      claimedBy: normalizeEmail(doctor?.claimedBy || ""),
      claimedByName: String(doctor?.claimedByName || "").trim(),
      accountEmail: normalizeEmail(doctor?.accountEmail || doctor?.claimedBy || ""),
      aliases: Array.isArray(doctor?.aliases)
        ? doctor.aliases.map((alias) => ({
            key: normalizeRosterName(alias?.key || ""),
            displayName: formatRosterDisplayName(alias?.displayName || alias?.key || ""),
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
    casey: Array.isArray(input.casey) ? input.casey.map((item) => String(item || "")).filter(Boolean) : [],
    mch: Array.isArray(input.mch) ? input.mch.map((item) => String(item || "")).filter(Boolean) : [],
  };
}

function detectedSourcesForSnapshot(value) {
  const imports = Array.isArray(value) && value.some((item) => item && typeof item === "object")
    ? value
    : sanitizeSourceTypes(Array.isArray(value) ? value : []).map((sourceType) => ({ sourceType, name: sourceType }));
  return {
    mmc: imports.filter((item) => item.sourceType === "mmc").map((item) => String(item.name || item.sourceType || "mmc")),
    ddh: imports.filter((item) => item.sourceType === "ddh").map((item) => String(item.name || item.sourceType || "ddh")),
    casey: imports.filter((item) => item.sourceType === "casey").map((item) => String(item.name || item.sourceType || "casey")),
    mch: imports.filter((item) => item.sourceType === "mch").map((item) => String(item.name || item.sourceType || "mch")),
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
    profileCoverage: sanitizeProfileCoverage(value.profileCoverage),
  };
}

function sanitizeProfileCoverage(value) {
  if (!value || typeof value !== "object") return null;
  return {
    matchedAliases: Array.isArray(value.matchedAliases) ? value.matchedAliases.map(rosterFileDoctorDiagnostic) : [],
    matchedFiles: Array.isArray(value.matchedFiles) ? value.matchedFiles.map((item) => String(item || "")).filter(Boolean) : [],
    zeroEventAliases: Array.isArray(value.zeroEventAliases) ? value.zeroEventAliases.map(rosterFileDoctorDiagnostic) : [],
    absentSources: sanitizeSourceTypes(value.absentSources),
  };
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

function repositoryImportRefsForAccount(index, record) {
  const claims = sanitizeClaims(record?.claims);
  return repositoryImportRefsForClaims(index, claims);
}

async function d1RepositoryImportRefsForClaims(db, claims) {
  return await queryRosterFileRefsForDoctors(db, sanitizeClaims(claims).map((claim) => claim.key)).catch(() => []);
}

async function linkedDoctorProfilesForClaims(store, claims, db = null) {
  const d1Profiles = await queryDoctorProfileMirrors(db).catch(() => []);
  return filterLinkedDoctorProfiles(d1Profiles, claims);
}

function filterLinkedDoctorProfiles(profiles, claims) {
  const claimSourcesByKey = new Map();
  for (const claim of sanitizeClaims(claims)) {
    if (!claimSourcesByKey.has(claim.key)) claimSourcesByKey.set(claim.key, new Set());
    claimSourcesByKey.get(claim.key).add(claim.sourceType);
  }
  const matches = [];
  for (const profile of profiles || []) {
    if (!profile) continue;
    const allowedSources = claimSourcesByKey.get(profile.doctorKey);
    if (!allowedSources) continue;
    if (!profile.sourceTypes.every((sourceType) => allowedSources.has(sourceType))) continue;
    matches.push(profile);
  }
  return matches;
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
      mergedCustomEvents.push(reassigned);
    }
  }
  return {
    ...state,
    session: {
      ...session,
      overrides: mergedOverrides,
      conflictSelections: mergedConflictSelections,
      customEvents: latestCustomEventsByIdentity(sanitizeSnapshotCustomEvents(mergedCustomEvents, ownerEmail)),
    },
  };
}

async function repositoryImportRefsForDoctorProfile(store, profile, db = null) {
  return await queryRosterFileRefsForDoctors(db, [profile?.doctorKey].filter(Boolean)).catch(() => []);
}

async function loadDoctorProfileSnapshotInfo(store, profile, db = null, ownerEmail = "") {
  const derivedSnapshot = await buildDerivedDoctorProfileSnapshot(store, db, profile, ownerEmail);
  return {
    snapshot: derivedSnapshot,
    snapshotAvailable: Boolean(derivedSnapshot),
    snapshotStale: false,
    snapshotBuiltAt: derivedSnapshot?.builtAt || "",
    snapshotBuildStamp: "",
  };
}

async function buildDerivedDoctorProfileSnapshot(store, db, profile, ownerEmail = "") {
  if (!hasCalendarDb({ ROSTER_DB: db }) || !profile?.profileId || !profile?.doctorKey) return null;
  const session = profile.state?.session && typeof profile.state.session === "object" ? profile.state.session : {};
  const settings = {
    ...defaultSettings(),
    ...(session.settings || {}),
  };
  let doctorDiagnostics = await queryRosterFileDoctorsForKeys(db, doctorKeysForOption(profile));
  if (!doctorDiagnostics.length) {
    const doctorRows = await queryRosterFileDoctors(db).catch(() => []);
    const fallbackDoctor = doctorRows.length ? await resolveCanonicalDoctorOptionForKey(db, doctorRows, profile.doctorKey) : null;
    doctorDiagnostics = fallbackDoctor
      ? doctorRows.filter((row) => doctorKeysForOption(fallbackDoctor).includes(normalizeRosterName(row.doctorKey)))
      : [];
  }
  const doctorKeys = doctorDiagnostics.length
    ? [...new Set(doctorDiagnostics.map((row) => normalizeRosterName(row.doctorKey)).filter(Boolean))]
    : doctorKeysForOption(profile);
  const doctorPairs = doctorDiagnostics.map((row) => ({ fileId: row.fileId, doctorKey: row.doctorKey }));
  const hospitalLocations = await loadAccountHospitalLocations(db, ownerEmail, session).catch(() => null);
  const events = [
    ...applyEventOverrides(
      applyAccountHospitalLocations(
        doctorPairs.length
          ? await queryDoctorEventsForFileDoctorPairs(db, doctorPairs)
          : await queryDoctorEvents(db, doctorKeys),
        hospitalLocations || {},
        { includeLocations: settings.includeLocations !== false },
      ),
      session.overrides || {},
    ),
    ...customEventsToEvents(sanitizeSnapshotCustomEvents(session.customEvents, ""), settings),
  ];
  if (!events.length) return null;
  const refs = await repositoryImportRefsForDoctorProfile(store, profile, db);
  const profileSources = sanitizeSourceTypes(profile.sourceTypes);
  const doctorOption = doctorDiagnostics.length ? buildDoctorOptionFromRows([...doctorDiagnostics]) : null;
  const doctorOptions = doctorOption
    ? [doctorOption]
    : [];
  return sanitizeSnapshotRecord({
    ownerType: "doctor-profile",
    ownerId: profile.profileId,
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    builtAt: new Date().toISOString(),
    buildStamp: "d1-derived",
    preview: buildPreviewFromDerivedEvents(events, { customEventsMaterialized: true }),
    session: session && Object.keys(session).length ? session : { doctorKey: profile.doctorKey },
    doctorOptions: doctorOptions.length ? doctorOptions : [{
      key: profile.doctorKey,
      displayName: profile.displayName,
      sourceTypes: profileSources,
      aliases: profileSources.map((sourceType) => ({
        sourceType,
        key: profile.doctorKey,
        displayName: profile.displayName,
      })),
    }],
    detectedSources: detectedSourcesForSnapshot(profileSources),
    fileRefs: refs,
    subscriptionFeeds: {},
    insightCache: null,
    profileCoverage: doctorProfileCoverage(doctorDiagnostics, doctorOption || profile, profileSources),
  });
}

async function doctorProfileLoadDiagnostics(db, profile) {
  if (!hasCalendarDb({ ROSTER_DB: db }) || !profile?.doctorKey) return null;
  const requestedKeys = doctorKeysForOption(profile);
  const [directRows, allRows, canonicalDoctors] = await Promise.all([
    queryRosterFileDoctorsForKeys(db, requestedKeys).catch(() => []),
    queryRosterFileDoctors(db).catch(() => []),
    queryCanonicalDoctors(db).catch(() => []),
  ]);
  const requestedKeySet = new Set(requestedKeys.map(normalizeRosterName));
  const canonicalMatches = canonicalDoctors.filter((doctor) => doctorKeysForOption(doctor).some((key) => requestedKeySet.has(normalizeRosterName(key))));
  const identity = rosterIdentityKey(profile.displayName || profile.doctorKey);
  const fuzzyRows = allRows.filter((row) => {
    const rowIdentity = rosterIdentityKey(row.displayName || row.doctorKey);
    return requestedKeySet.has(normalizeRosterName(row.doctorKey)) || (identity && rowIdentity && likelySameRosterName(identity, rowIdentity));
  });
  return {
    requestedKey: normalizeRosterName(profile.doctorKey || ""),
    requestedSourceTypes: sanitizeSourceTypes(profile.sourceTypes),
    directRosterRowCount: directRows.length,
    fuzzyRosterRowCount: fuzzyRows.length,
    canonicalMatchCount: canonicalMatches.length,
    directEventCount: directRows.reduce((total, row) => total + Number(row.eventCount || 0), 0),
    fuzzyEventCount: fuzzyRows.reduce((total, row) => total + Number(row.eventCount || 0), 0),
  };
}

function matchDoctorClaims(doctors, realName) {
  const claims = [];
  const realIdentity = rosterIdentityKey(realName);
  if (!realIdentity) return claims;
  for (const doctor of doctors || []) {
    if (!doctorMatchesRealName(doctor, realName)) continue;
    const aliases = Array.isArray(doctor.aliases) && doctor.aliases.length ? doctor.aliases : [doctor];
    for (const alias of aliases) {
      claims.push({
        key: alias.key || doctor.key,
        displayName: alias.displayName || doctor.displayName,
        sourceType: alias.sourceType || doctor.sourceType,
        matchedAt: new Date().toISOString(),
      });
    }
  }
  return mergeClaims([], claims);
}

async function repositoryDoctorCandidates(store, index, db = null, options = {}) {
  const accountIndex = await loadClaimedAccountIndex(store, db);
  const doctors = await mergedDoctorOptionsFromD1(db, {
    includeZeroEventStandalone: options.hideZeroEventStandalone !== true,
  });
  if (doctors.length) return attachClaimedAccountMetadata(doctors, accountIndex);
  return [];
}

async function refreshCanonicalDoctors(db) {
  const doctorRows = await queryRosterFileDoctors(db).catch(() => []);
  if (!doctorRows.length) {
    await replaceCanonicalDoctors(db, []);
    return [];
  }
  const doctors = await buildCanonicalDoctorOptionsFromRows(db, doctorRows, {
    includeZeroEventStandalone: true,
  });
  await replaceCanonicalDoctors(db, doctors);
  return doctors;
}

function attachClaimedAccountMetadata(doctors, accountIndex) {
  const seen = new Set();
  const candidates = [];
  for (const doctor of doctors || []) {
    const marker = `${doctor.sourceType}:${doctor.key}`;
    if (seen.has(marker)) continue;
    seen.add(marker);
    const resolved = resolveDoctorAccountFromIndex(accountIndex, doctor);
    candidates.push({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType: doctor.sourceType || doctor.sourceTypes?.[0] || "",
      sourceTypes: doctor.sourceTypes || [doctor.sourceType].filter(Boolean),
      aliases: doctor.aliases || [],
      hasEvents: doctor.hasEvents !== false,
      claimedBy: resolved.mode === "claimed-account" ? resolved.email : "",
      claimedByName: resolved.mode === "claimed-account" ? resolved.realName : "",
      accountEmail: resolved.mode === "claimed-account" ? resolved.email : "",
    });
  }
  return candidates.sort((left, right) => {
    const leftClaimed = left.claimedBy ? 1 : 0;
    const rightClaimed = right.claimedBy ? 1 : 0;
    if (leftClaimed !== rightClaimed) return leftClaimed - rightClaimed;
    return left.displayName.localeCompare(right.displayName) || left.sourceType.localeCompare(right.sourceType);
  });
}

function buildCreatorDoctorOptions(doctors) {
  const groups = new Map();
  for (const doctor of doctors || []) {
    const identity = rosterIdentityKey(doctor.displayName || doctor.key);
    if (!identity) continue;
    if (!groups.has(identity)) groups.set(identity, []);
    groups.get(identity).push({
      sourceType: doctor.sourceType,
      key: doctor.key,
      displayName: doctor.displayName,
      eventCount: Number(doctor.eventCount || 0),
    });
  }
  return [...groups.values()].map((aliases) => {
    aliases.sort((left, right) => sourcePriority(left.sourceType) - sourcePriority(right.sourceType) || left.displayName.localeCompare(right.displayName));
    const primary = aliases[0];
    return {
      key: primary.key,
      displayName: primary.displayName,
      sourceType: primary.sourceType,
      sourceTypes: [...new Set(aliases.map((alias) => alias.sourceType))],
      aliases,
      hasEvents: aliases.some((alias) => Number(alias.eventCount || 0) > 0),
    };
  }).sort((left, right) => left.displayName.localeCompare(right.displayName));
}

async function buildCanonicalDoctorOptionsFromRows(db, rows, options = {}) {
  const exactGroups = new Map();
  for (const row of rows || []) {
    const identity = rosterIdentityKey(row.displayName || row.doctorKey);
    if (!identity) continue;
    if (!exactGroups.has(identity)) exactGroups.set(identity, []);
    exactGroups.get(identity).push(row);
  }
  const groups = [...exactGroups.values()];
  const eventCache = new Map();
  const groupEvents = async (group) => {
    const key = group.map((row) => row.doctorKey).sort().join("|");
    if (!eventCache.has(key)) eventCache.set(key, queryDoctorEvents(db, [...new Set(group.map((row) => row.doctorKey))]));
    return await eventCache.get(key);
  };
  for (let leftIndex = 0; leftIndex < groups.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < groups.length; rightIndex += 1) {
      const left = groups[leftIndex];
      const right = groups[rightIndex];
      if (!isConservativeDoctorVariant(left, right)) continue;
      if (groupsShareFile(left, right)) continue;
      if (await groupsHaveConflictingWorkingEvents(await groupEvents(left), await groupEvents(right))) continue;
      left.push(...right);
      groups.splice(rightIndex, 1);
      rightIndex -= 1;
    }
  }
  return groups
    .filter((aliases) => options.includeZeroEventStandalone === true || aliases.some((alias) => Number(alias.eventCount || 0) > 0))
    .map(buildDoctorOptionFromRows)
    .sort((left, right) => left.displayName.localeCompare(right.displayName));
}

async function resolveCanonicalDoctorOptionForKey(db, rows, doctorKey) {
  const normalizedKey = normalizeRosterName(doctorKey || "");
  if (!normalizedKey) return null;
  const group = (rows || []).filter((row) => normalizeRosterName(row.doctorKey) === normalizedKey);
  if (!group.length) return null;
  const seen = new Set(group.map((row) => `${row.fileId}:${row.doctorKey}`));
  let existingEvents = await queryDoctorEvents(db, [...new Set(group.map((row) => row.doctorKey))]);
  for (const row of rows || []) {
    const marker = `${row.fileId}:${row.doctorKey}`;
    if (seen.has(marker) || !isConservativeDoctorVariant(group, [row]) || groupsShareFile(group, [row])) continue;
    const candidateEvents = await queryDoctorEvents(db, [row.doctorKey]);
    if (await groupsHaveConflictingWorkingEvents(existingEvents, candidateEvents)) continue;
    group.push(row);
    existingEvents = [...existingEvents, ...candidateEvents];
    seen.add(marker);
  }
  return buildDoctorOptionFromRows(group);
}

function buildDoctorOptionFromRows(aliases) {
  aliases.sort((left, right) => Number(right.eventCount || 0) - Number(left.eventCount || 0) || sourcePriority(left.sourceType) - sourcePriority(right.sourceType) || left.displayName.localeCompare(right.displayName));
  const primary = aliases[0];
  return {
    key: primary.doctorKey,
    displayName: primary.displayName,
    sourceType: primary.sourceType,
    sourceTypes: [...new Set(aliases.map((alias) => alias.sourceType))],
    aliases: aliases.map((alias) => ({
      sourceType: alias.sourceType,
      key: alias.doctorKey,
      displayName: alias.displayName,
      fileId: alias.fileId,
      fileName: alias.fileName,
      eventCount: Number(alias.eventCount || 0),
    })),
    hasEvents: aliases.some((alias) => Number(alias.eventCount || 0) > 0),
  };
}

function isConservativeDoctorVariant(leftRows, rightRows) {
  const left = rosterNameTokens(leftRows[0]?.displayName || leftRows[0]?.doctorKey || "");
  const right = rosterNameTokens(rightRows[0]?.displayName || rightRows[0]?.doctorKey || "");
  if (left.length < 2 || right.length < 2 || left.length !== right.length) return false;
  const leftGiven = left.slice(0, -1).join(" ");
  const rightGiven = right.slice(0, -1).join(" ");
  if (leftGiven !== rightGiven) return false;
  return levenshteinDistance(left.at(-1), right.at(-1)) === 1;
}

function groupsShareFile(left, right) {
  const files = new Set(left.map((row) => row.fileId));
  return right.some((row) => files.has(row.fileId));
}

async function groupsHaveConflictingWorkingEvents(leftEvents, rightEvents) {
  return leftEvents.some((left) => isWorkingRosterEvent(left) && rightEvents.some((right) => (
    isWorkingRosterEvent(right)
    && String(left.start || "") < String(right.end || right.start || "")
    && String(right.start || "") < String(left.end || left.start || "")
  )));
}

function isWorkingRosterEvent(event) {
  const title = String(event?.title || "").toLowerCase();
  return !["annual leave", "conference leave", "sick leave", "phnw", "public holiday"].some((label) => title.includes(label));
}

function levenshteinDistance(left, right) {
  const a = String(left || "");
  const b = String(right || "");
  const table = Array.from({ length: a.length + 1 }, (_, row) => Array.from({ length: b.length + 1 }, (_, col) => row ? (col ? 0 : row) : col));
  for (let row = 1; row <= a.length; row += 1) {
    for (let col = 1; col <= b.length; col += 1) {
      table[row][col] = Math.min(
        table[row - 1][col] + 1,
        table[row][col - 1] + 1,
        table[row - 1][col - 1] + (a[row - 1] === b[col - 1] ? 0 : 1),
      );
    }
  }
  return table[a.length][b.length];
}

function doctorProfileCoverage(rows, doctor, sourceTypes) {
  const keys = new Set(doctorKeysForOption(doctor));
  const matched = (rows || []).filter((row) => keys.has(normalizeRosterName(row.doctorKey)));
  const matchedSources = new Set(matched.map((row) => row.sourceType));
  return {
    matchedAliases: matched.map(rosterFileDoctorDiagnostic),
    matchedFiles: [...new Set(matched.map((row) => row.fileId))],
    zeroEventAliases: matched.filter((row) => Number(row.eventCount || 0) === 0).map(rosterFileDoctorDiagnostic),
    absentSources: (sourceTypes || []).filter((source) => !matchedSources.has(source)),
  };
}

function sourcePriority(sourceType) {
  return { mmc: 0, ddh: 1, casey: 2, mch: 3 }[sourceType] ?? 99;
}

function sanitizeDoctorAccountResolutionInput(value) {
  const sourceTypes = sanitizeSourceTypes(value?.sourceTypes || (value?.sourceType ? [value.sourceType] : []));
  const aliases = Array.isArray(value?.aliases)
    ? value.aliases.map((alias) => ({
        key: normalizeRosterName(alias?.key || ""),
        displayName: formatRosterDisplayName(alias?.displayName || alias?.key || ""),
        sourceType: String(alias?.sourceType || "").toLowerCase(),
      })).filter((alias) => alias.key && alias.displayName && isRosterSourceType(alias.sourceType))
    : [];
  const key = normalizeRosterName(value?.key || value?.doctorKey || "");
  const displayName = formatRosterDisplayName(value?.displayName || value?.key || value?.doctorKey || "");
  const sourceType = String(value?.sourceType || "").toLowerCase();
  if (sourceType && isRosterSourceType(sourceType) && !sourceTypes.includes(sourceType)) {
    sourceTypes.push(sourceType);
  }
  return {
    key,
    displayName,
    sourceTypes,
    aliases,
  };
}

function doctorResolutionMarkers(input) {
  const doctor = sanitizeDoctorAccountResolutionInput(input);
  const markers = new Set();
  const keys = new Set([doctor.key, ...doctor.aliases.map((alias) => alias.key)].filter(Boolean));
  const aliasesByKey = new Map();
  for (const alias of doctor.aliases) {
    if (!aliasesByKey.has(alias.key)) aliasesByKey.set(alias.key, new Set());
    aliasesByKey.get(alias.key).add(alias.sourceType);
  }
  for (const key of keys) {
    const sources = new Set([...(doctor.sourceTypes || []), ...(aliasesByKey.get(key) || [])]);
    if (!sources.size) {
      markers.add(`*:${key}`);
      continue;
    }
    for (const sourceType of sources) markers.add(`${sourceType}:${key}`);
  }
  return {
    ...doctor,
    markers,
    keys,
    identityKeys: new Set([
      rosterIdentityKey(doctor.displayName),
      rosterIdentityKey(doctor.key),
      ...doctor.aliases.flatMap((alias) => [rosterIdentityKey(alias.displayName), rosterIdentityKey(alias.key)]),
    ].filter(Boolean)),
  };
}

async function resolveDoctorAccount(store, rawDoctor, db = null) {
  return resolveDoctorAccountFromIndex(await loadClaimedAccountIndex(store, db), rawDoctor);
}

async function loadClaimedAccountIndex(store, db = null) {
  return await queryClaimedAccounts(db).catch(() => []);
}

function resolveDoctorAccountFromIndex(accounts, rawDoctor) {
  const requested = doctorResolutionMarkers(rawDoctor);
  if (!requested.key && !requested.displayName && !requested.aliases.length) {
    return { mode: "doctor-profile", email: "", realName: "" };
  }
  for (const account of accounts || []) {
    for (const claim of account.claims || []) {
      if (
        requested.markers.has(`${claim.sourceType}:${claim.key}`)
        || requested.markers.has(`*:${claim.key}`)
        || requested.keys.has(claim.key)
        || requested.identityKeys.has(rosterIdentityKey(claim.displayName || claim.key))
      ) {
        return {
          mode: "claimed-account",
          email: account.email,
          realName: account.realName,
        };
      }
    }
  }
  return { mode: "doctor-profile", email: "", realName: "" };
}

async function clearDeletedAccountClaimMetadata(store, email, record) {
  const deletedEmail = normalizeEmail(email);
  if (!deletedEmail) return;
  const deletedClaims = sanitizeClaims(record?.claims);
  if (!deletedClaims.length) return;
  await removeDeletedClaimsFromRemainingD1Accounts(store, deletedEmail, deletedClaims);
}

async function removeDeletedClaimsFromRemainingD1Accounts(db, deletedEmail, deletedClaims) {
  const records = await listAccountMirrors(db);
  for (const record of records || []) {
    const email = normalizeEmail(record?.email);
    if (!record || email === deletedEmail) continue;
    const claims = sanitizeClaims(record.claims);
    const filtered = claims.filter((claim) => !deletedClaims.some((deletedClaim) => sameClaim(claim, deletedClaim)));
    if (filtered.length === claims.length) continue;
    await upsertAccountMirror(db, {
      ...record,
      claims: filtered,
      updatedAt: new Date().toISOString(),
    }, { preserveExistingState: true });
  }
}

function findDoctorClaimCandidate(canonicalDoctors, rawClaim) {
  const claim = {
    key: normalizeRosterName(rawClaim?.key || ""),
    sourceType: String(rawClaim?.sourceType || "").toLowerCase(),
  };
  if (!claim.key || !claim.sourceType) return null;
  for (const doctor of canonicalDoctors || []) {
    const aliases = Array.isArray(doctor?.aliases) && doctor.aliases.length ? doctor.aliases : [doctor];
    const alias = aliases.find((item) => (
      normalizeRosterName(item?.key || "") === claim.key
      && String(item?.sourceType || doctor?.sourceType || "").toLowerCase() === claim.sourceType
    ));
    if (!alias) continue;
    return {
      key: normalizeRosterName(alias.key || doctor.key),
      displayName: String(alias.displayName || doctor.displayName || alias.key || doctor.key).trim(),
      sourceType: String(alias.sourceType || doctor.sourceType || "").toLowerCase(),
    };
  }
  return null;
}

async function resolveDoctorClaimCandidate(db, rawClaim) {
  const claim = {
    key: normalizeRosterName(rawClaim?.key || ""),
    sourceType: String(rawClaim?.sourceType || "").toLowerCase(),
  };
  if (!claim.key || !isRosterSourceType(claim.sourceType)) return null;
  const rows = await queryRosterFileDoctorsForKeys(db, [claim.key]).catch(() => []);
  const exact = rows.find((row) => String(row.sourceType || "").toLowerCase() === claim.sourceType);
  if (!exact) return null;
  return {
    key: normalizeRosterName(exact.doctorKey),
    displayName: String(exact.displayName || exact.doctorKey || "").trim(),
    sourceType: String(exact.sourceType || "").toLowerCase(),
  };
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
    .filter((claim) => claim.key && claim.displayName && isRosterSourceType(claim.sourceType));
}

function sanitizeSourceTypes(items) {
  if (!Array.isArray(items)) return [];
  return [...new Set(items.map((item) => String(item || "").toLowerCase()).filter(isRosterSourceType))];
}

function isRosterSourceType(value) {
  const source = String(value || "").toLowerCase();
  return source === "mmc" || source === "ddh" || source === "casey" || source === "mch";
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
  return Boolean(session.hadPreview || Object.keys(overrides).length || Object.keys(conflictSelections).length || customEvents.length);
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

function claimMatchesAccountIdentity(claim, realName) {
  if (!rosterIdentityKey(realName)) return true;
  return doctorMatchesRealName({
    key: normalizeRosterName(claim?.key || ""),
    displayName: String(claim?.displayName || claim?.key || "").trim(),
  }, realName);
}

function doctorMatchesRealName(doctor, realName) {
  const realKey = normalizeRosterName(realName);
  const realIdentityKey = rosterIdentityKey(realName);
  const doctorIdentityKey = rosterIdentityKey(doctor.displayName || doctor.key);
  if (!realKey) return false;
  if (doctor.key === realKey) return true;
  if (realIdentityKey && doctorIdentityKey && realIdentityKey === doctorIdentityKey) return true;
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
  return rosterIdentityKey(value).split(" ").filter(Boolean);
}

function rosterIdentityKey(value) {
  const stripped = String(value || "")
    .replace(/[^A-Za-z0-9,]+/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase()
    .replace(/^(DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A PROF|ASSOC PROF)\s+/, "");
  const parts = stripped.split(/\s*,\s*/).filter(Boolean);
  return parts.length === 2 ? `${parts[1]} ${parts[0]}`.trim() : stripped.replace(/,/g, "");
}

function formatRosterDisplayName(value) {
  const identity = rosterIdentityKey(value);
  const tokens = identity.split(" ").filter(Boolean);
  if (!tokens.length) return String(value || "").trim();
  return tokens.map((token, index) => index === tokens.length - 1 ? token : toDisplayNameToken(token)).join(" ");
}

function toDisplayNameToken(value) {
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
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
  if (record.subscriptionToken) return record;
  return {
    ...record,
    subscriptionToken: randomSubscriptionToken(),
    updatedAt: new Date().toISOString(),
  };
}

async function sha256(value) {
  const bytes = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(hash)].map((item) => item.toString(16).padStart(2, "0")).join("");
}

function sanitizeState(value) {
  const input = value && typeof value === "object" ? value : {};
  const session = input.session && typeof input.session === "object" ? {
    ...input.session,
    customEvents: sanitizeSnapshotCustomEvents(input.session.customEvents),
  } : {};
  return {
    version: 1,
    imports: Array.isArray(input.imports) ? input.imports : [],
    session,
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
      seniority: sanitizeRuleSeniority(item?.seniority),
      date: String(item?.date || item?.startDay || "").trim(),
      rawValue: String(item?.rawValue || "").trim(),
      code: String(item?.code || "").trim().toUpperCase(),
      timeLabel: String(item?.timeLabel || "").trim(),
      suggestedTitle: String(item?.suggestedTitle || "").trim(),
      fingerprint: sanitizeIssueFingerprint(item?.fingerprint || issueFingerprint(item?.source, item?.rawValue, item?.seniority)),
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
      match.seniority = item.seniority || match.seniority;
      match.date = item.date || match.date;
      match.rawValue = item.rawValue || match.rawValue;
      match.code = item.code || match.code;
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
  if (source === "MMC" || source === "DDH") return source;
  if (source === "CASEY") return "Casey";
  if (source === "MCH") return "MCH";
  return "";
}

function issueFingerprint(source, rawValue, seniority = "") {
  const normalizedSource = sanitizeIssueSource(source);
  const normalizedSeniority = seniority ? sanitizeRuleSeniority(seniority) : "";
  const normalizedRawValue = String(rawValue || "").trim();
  return normalizedSource && normalizedRawValue ? `${normalizedSource}::${normalizedSeniority ? `${normalizedSeniority}::` : ""}${normalizedRawValue}` : "";
}

function sanitizeIssueFingerprint(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const [source, ...rest] = raw.split("::");
  if (rest.length >= 2) {
    const seniority = sanitizeRuleSeniority(rest[0]);
    const rawValue = rest.slice(1).join("::");
    return issueFingerprint(source, rawValue, seniority);
  }
  return issueFingerprint(source, rest.join("::"));
}

function sanitizeIssueFingerprintList(values) {
  if (!Array.isArray(values)) return [];
  return [...new Set(values.map((value) => sanitizeIssueFingerprint(value)).filter(Boolean))].sort();
}

function sanitizeParserExtensionRules(value) {
  const source = value && typeof value === "object" ? value : {};
  const removed = sanitizeParserRuleRemovals(source._removed);
  return {
    mmc: applyParserRuleRemovals(sanitizeParserExtensionRuleList(source.mmc, "MMC"), removed),
    ddh: applyParserRuleRemovals(sanitizeParserExtensionRuleList(source.ddh, "DDH"), removed),
    casey: applyParserRuleRemovals(sanitizeParserExtensionRuleList(source.casey, "Casey"), removed),
    mch: applyParserRuleRemovals(sanitizeParserExtensionRuleList(source.mch, "MCH"), removed),
    _removed: removed,
  };
}

function sanitizeParserRuleRemovals(items) {
  if (!Array.isArray(items)) return [];
  const byKey = new Map();
  for (const item of items) {
    const target = sanitizeParserRuleRemoval(item);
    if (target) byKey.set(`${target.source}|${target.seniority}|${target.code}`, target);
  }
  return [...byKey.values()].sort((left, right) => left.code.localeCompare(right.code));
}

function sanitizeParserRuleRemoval(item) {
  if (!item || typeof item !== "object") return null;
  const source = sanitizeIssueSource(item.source);
  const seniority = sanitizeRuleSeniority(item.seniority);
  const code = normalizeParserExtensionRuleCode(source, item.code || item.rawCode || "");
  if (!source || !code) return null;
  return { source, seniority, code };
}

function applyParserRuleRemovals(rules, removals) {
  const removedKeys = new Set((removals || []).map((item) => `${item.source}|${item.seniority}|${item.code}`));
  return (rules || []).filter((rule) => !removedKeys.has(`${rule.source}|${rule.seniority}|${rule.code}`));
}

function sanitizeParserExtensionRuleList(items, source) {
  if (!Array.isArray(items)) return [];
  return items
    .map((item) => sanitizeParserExtensionRule(item, source))
    .filter((item) => !isObsoleteSeededParserRule(item))
    .filter(Boolean)
    .sort((left, right) => left.code.localeCompare(right.code));
}

function sanitizeParserExtensionRule(item, forcedSource = "") {
  if (!item || typeof item !== "object") return null;
  const source = sanitizeIssueSource(forcedSource || item.source);
  const seniority = sanitizeRuleSeniority(item.seniority);
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
  if (isRestrictedClinicalSupportRule({ seniority, code, base })) return null;
  if (!allDay && (!isClockString(startTime) || !isClockString(endTime))) return null;
  return {
    source,
    seniority,
    code,
    kind,
    base,
    period,
    suffix,
    allDay,
    startTime: allDay ? "" : startTime,
    endTime: allDay ? "" : endTime,
    location,
    includeAsShift: item.includeAsShift !== false,
  };
}

function isRestrictedClinicalSupportRule(rule) {
  const seniority = sanitizeRuleSeniority(rule?.seniority);
  if (seniority === "SMS" || seniority === "CMO") return false;
  const code = String(rule?.code || "").trim().toUpperCase();
  const base = String(rule?.base || "").trim().toUpperCase();
  return code === "CS"
    || code === "CSO"
    || code === "CS ONSITE"
    || code === "CLIN SUPP"
    || code === "CLINICAL SUPP"
    || base === "CS"
    || base === "CSO"
    || base === "CS ONSITE";
}

function normalizeParserExtensionRuleCode(source, value) {
  const text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (source === "MMC" || source === "Casey") {
    return parserRuleCodeFromRawValue(source, text);
  }
  return text;
}

function sanitizeRuleSeniority(value) {
  const text = String(value || "").trim();
  const upper = text.toUpperCase();
  const aliases = new Map([
    ["SR", "Senior Registrar"],
    ["SENIOR REG", "Senior Registrar"],
    ["SENIOR REGISTRAR", "Senior Registrar"],
    ["IR", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REG", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REGISTRAR", "Transitional/Intermediate Registrar"],
    ["TRANSITIONAL REGISTRAR", "Transitional/Intermediate Registrar"],
    ["JR", "Junior Registrar"],
    ["JUNIOR REG", "Junior Registrar"],
    ["JUNIOR REGISTRAR", "Junior Registrar"],
    ["I", "Intern"],
    ["INTERN", "Intern"],
  ]);
  if (aliases.has(upper)) return aliases.get(upper);
  const labels = ["SMS", "CMO", "Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "ENP", "AMP", "Intern", "Unknown"];
  return labels.find((item) => item.toUpperCase() === upper) || "Unknown";
}

function isObsoleteSeededParserRule(rule) {
  if (!rule || rule.source !== "MMC") return false;
  if (rule.code === "CS" || rule.code === "CSO") return false;
  const impossibleFloat = new Set(["ACR", "PCR", "ARR", "PRR", "ASSR", "PSSR"]);
  if (impossibleFloat.has(rule.code) && isOldDefaultMmcRule(rule)) return true;
  return false;
}

function isConsultantStyleMmcCode(code) {
  const text = String(code || "").trim().toUpperCase();
  return /^[AP][GARC][CR]$/.test(text) || /^[AP]SS[CR]$/.test(text);
}

function isOldDefaultMmcRule(rule) {
  const expected = oldDefaultMmcRuleShape(rule.code);
  if (!expected) return false;
  return rule.base === expected.base
    && rule.period === expected.period
    && rule.suffix === expected.suffix
    && rule.allDay === false
    && rule.startTime === expected.startTime
    && rule.endTime === expected.endTime
    && rule.includeAsShift !== false;
}

function oldDefaultMmcRuleShape(code) {
  const text = String(code || "").trim().toUpperCase();
  const teamMap = { G: "Green", A: "Amber", R: "Resus", C: "Clinic" };
  const teamMatch = text.match(/^([AP])([GARC])([CR])$/);
  if (teamMatch) {
    return {
      base: teamMap[teamMatch[2]],
      period: teamMatch[1] === "A" ? "AM" : "PM",
      suffix: teamMatch[3] === "R" ? "Float" : "",
      startTime: teamMatch[1] === "A" ? "08:00" : "14:30",
      endTime: teamMatch[1] === "A" ? "17:30" : "00:00",
    };
  }
  const ssuMatch = text.match(/^([AP])SS([CR])$/);
  if (!ssuMatch) return null;
  return {
    base: "SSU",
    period: ssuMatch[1] === "A" ? "AM" : "PM",
    suffix: ssuMatch[2] === "R" ? "Float" : "",
    startTime: ssuMatch[1] === "A" ? "07:30" : "14:30",
    endTime: ssuMatch[1] === "A" ? "17:30" : "00:00",
  };
}

async function loadD1ParserExtensionRules(db) {
  if (!db?.prepare) return {};
  const rows = await db.prepare("SELECT rule_json FROM parser_rules WHERE scope = 'global'").all().catch(() => ({ results: [] }));
  const rules = {};
  for (const row of rows.results || []) {
    let parsed = null;
    try {
      parsed = JSON.parse(row.rule_json || "{}");
    } catch {
      parsed = null;
    }
    const rule = sanitizeParserExtensionRule(parsed);
    if (!rule) continue;
    const sourceKey = rule.source.toLowerCase();
    rules[sourceKey] = [...(rules[sourceKey] || []), rule];
  }
  return sanitizeParserExtensionRules(rules);
}

async function saveD1ParserExtensionRule(db, rule) {
  const normalized = sanitizeParserExtensionRule(rule);
  if (!db?.prepare || !normalized) return;
  const id = `global:${normalized.source}:${normalized.seniority}:${normalized.code}`;
  const now = new Date().toISOString();
  await db.prepare(`
    INSERT INTO parser_rules (id, scope, email, source_type, seniority, code, title, rule_json, updated_at)
    VALUES (?, 'global', '', ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      source_type = excluded.source_type,
      seniority = excluded.seniority,
      code = excluded.code,
      title = excluded.title,
      rule_json = excluded.rule_json,
      updated_at = excluded.updated_at
  `).bind(id, normalized.source, normalized.seniority, normalized.code, normalized.base || normalized.title || "", JSON.stringify(normalized), now).run();
}

async function deleteD1ParserExtensionRule(db, rule) {
  const target = sanitizeParserRuleRemoval(rule);
  if (!db?.prepare || !target) return;
  await db.prepare("DELETE FROM parser_rules WHERE scope = 'global' AND source_type = ? AND seniority = ? AND code = ?")
    .bind(target.source, target.seniority, target.code)
    .run();
}

function upsertParserExtensionRule(existing, rule) {
  const sanitized = sanitizeParserExtensionRules(existing);
  const nextRule = sanitizeParserExtensionRule(rule);
  if (!nextRule) return sanitized;
  const key = nextRule.source.toLowerCase();
  const items = sanitized[key] || [];
  const nextItems = items.filter((item) => item.code !== nextRule.code || item.seniority !== nextRule.seniority);
  nextItems.push(nextRule);
  return {
    ...sanitized,
    _removed: (sanitized._removed || []).filter((item) => item.source !== nextRule.source || item.code !== nextRule.code || item.seniority !== nextRule.seniority),
    [key]: nextItems.sort((left, right) => left.code.localeCompare(right.code)),
  };
}

function removeParserExtensionRuleByKey(existing, targetRule) {
  const sanitized = sanitizeParserExtensionRules(existing);
  const target = sanitizeParserRuleRemoval(targetRule);
  if (!target) return sanitized;
  const key = target.source.toLowerCase();
  const removals = sanitizeParserRuleRemovals([...(sanitized._removed || []), target]);
  return {
    ...sanitized,
    _removed: removals,
    [key]: (sanitized[key] || []).filter((item) => item.code !== target.code || item.seniority !== target.seniority),
  };
}

function removeParserExtensionRule(existing, rule) {
  const sanitized = sanitizeParserExtensionRules(existing);
  const target = sanitizeParserExtensionRule(rule);
  if (!target) return sanitized;
  const key = target.source.toLowerCase();
  return {
    ...sanitized,
    [key]: (sanitized[key] || []).filter((item) => item.code !== target.code || item.seniority !== target.seniority),
  };
}

function mergeParserExtensionSets(globalRules, localRules) {
  const base = sanitizeParserExtensionRules(globalRules);
  const local = sanitizeParserExtensionRules(localRules);
  return {
    mmc: mergeParserExtensionRuleLists(base.mmc, local.mmc),
    ddh: mergeParserExtensionRuleLists(base.ddh, local.ddh),
    casey: mergeParserExtensionRuleLists(base.casey, local.casey),
    mch: mergeParserExtensionRuleLists(base.mch, local.mch),
  };
}

function mergeParserExtensionRuleLists(globalRules, localRules) {
  const byKey = new Map();
  for (const rule of globalRules || []) byKey.set(parserExtensionRuleKey(rule), rule);
  for (const rule of localRules || []) byKey.set(parserExtensionRuleKey(rule), rule);
  return [...byKey.values()].sort((left, right) => left.code.localeCompare(right.code));
}

function parserExtensionRuleKey(rule) {
  const normalized = sanitizeParserExtensionRule(rule);
  return normalized ? `${normalized.source}|${normalized.seniority}|${normalized.code}` : "";
}

async function loadParserRuleSuggestions(_store, db = null) {
  if (!db?.prepare) return [];
  const result = await db.prepare("SELECT suggestion_json FROM parser_rule_suggestions WHERE status = 'pending' ORDER BY updated_at DESC").all();
  return sanitizeParserRuleSuggestions((result?.results || []).map((row) => {
    try { return JSON.parse(row.suggestion_json || "{}"); } catch { return null; }
  }));
}

async function saveParserRuleSuggestions(_store, value, db = null) {
  const sanitized = sanitizeParserRuleSuggestions(value);
  if (!db?.prepare) return sanitized;
  await db.prepare("DELETE FROM parser_rule_suggestions").run();
  for (const suggestion of sanitized) {
    await db.prepare(`
      INSERT INTO parser_rule_suggestions (id, email, status, suggestion_json, created_at, updated_at)
      VALUES (?, ?, 'pending', ?, ?, ?)
    `).bind(
      suggestion.id,
      suggestion.email,
      JSON.stringify(suggestion),
      suggestion.createdAt,
      suggestion.updatedAt,
    ).run();
  }
  return sanitized;
}

function sanitizeParserRuleSuggestions(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => sanitizeParserRuleSuggestion(item))
    .filter(Boolean)
    .sort((left, right) => (right.updatedAt || "").localeCompare(left.updatedAt || ""));
}

function sanitizeParserRuleSuggestion(item) {
  if (!item || typeof item !== "object") return null;
  const email = normalizeEmail(item.email);
  const rule = sanitizeParserExtensionRule(item.rule);
  if (!email || !rule) return null;
  const fingerprint = sanitizeIssueFingerprint(item.fingerprint || issueFingerprint(rule.source, rule.code, rule.seniority));
  const id = `${email}::${parserExtensionRuleKey(rule)}`;
  return {
    id,
    email,
    realName: String(item.realName || "").trim(),
    fingerprint,
    rawValue: String(item.rawValue || rule.code || "").trim(),
    rule,
    status: "pending",
    createdAt: String(item.createdAt || item.updatedAt || new Date().toISOString()),
    updatedAt: String(item.updatedAt || item.createdAt || new Date().toISOString()),
  };
}

function upsertParserRuleSuggestion(existing, suggestion) {
  const nextSuggestion = sanitizeParserRuleSuggestion(suggestion);
  const suggestions = sanitizeParserRuleSuggestions(existing);
  if (!nextSuggestion) return suggestions;
  return [
    nextSuggestion,
    ...suggestions.filter((item) => item.id !== nextSuggestion.id),
  ].sort((left, right) => (right.updatedAt || "").localeCompare(left.updatedAt || ""));
}

function isClockString(value) {
  return /^\d{2}:\d{2}$/.test(String(value || "").trim());
}
