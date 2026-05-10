import { inspectImportRecord, normalizeRosterName } from "../_lib/roster.js";

const CREATOR_EMAIL = "rhaydon@gmail.com";
const REPOSITORY_INDEX_KEY = "repository:index";
const REPOSITORY_FILE_PREFIX = "repository:file:";
const DOCTOR_PROFILE_PREFIX = "doctor-profile:";
const SUBSCRIPTION_TOKEN_PREFIX = "subscription:token:";
const SNAPSHOT_PREFIX = "snapshot:";
const SNAPSHOT_SCHEMA_VERSION = 4;
const ADMIN_ISSUE_DISMISS_PREFIX = "admin-issue-dismiss:";
const ADMIN_ISSUE_IGNORE_PREFIX = "admin-issue-ignore:";
const PARSER_EXTENSION_RULES_KEY = "parser-extension-rules:v1";
const PARSER_RULE_SUGGESTIONS_KEY = "parser-rule-suggestions:v1";

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
        suggestedClaims: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
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
          suggestedClaims: prepared.nameMatches,
          insightsEnabled: insightsEnabledForRecord({ ...created.record, role: prepared.role }),
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
        suggestedClaims: prepared.nameMatches,
        availableDoctors: prepared.availableDoctors,
        subscription: prepared.subscription,
        insightsEnabled: prepared.insightsEnabled,
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
        users: await listUsers(context.env.ROSTER_STORE),
        availableDoctors: await repositoryDoctorCandidates(context.env.ROSTER_STORE, await loadRepositoryIndex(context.env.ROSTER_STORE)),
        issueConfig: await buildIssueConfig(context.env.ROSTER_STORE, email),
      });
    }

    if (action === "updateAccount") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, saveEmail);
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
      await context.env.ROSTER_STORE.put(storageKey(saveEmail), JSON.stringify(updated));
      const owner = accountSnapshotOwner(saveEmail, updated.role || roleForEmail(saveEmail));
      await context.env.ROSTER_STORE.delete(snapshotKey(owner.ownerType, owner.ownerId));
      const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, updated, { includeAvailableDoctors: false });
      return Response.json({
        ok: true,
        realName: prepared.realName,
        claims: prepared.claims,
        nameMatches: prepared.nameMatches,
        suggestedClaims: prepared.nameMatches,
        user: userSummaryFromRecord(saveEmail, { ...updated, claims: prepared.claims }),
      });
    }

    if (action === "setAccountRosterClaims") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      if (!targetEmail) {
        return Response.json({ error: "Target account is required." }, { status: 400 });
      }
      const targetRecord = await loadAccountRecord(context.env.ROSTER_STORE, targetEmail);
      const index = await loadRepositoryIndex(context.env.ROSTER_STORE);
      const claims = sanitizeClaims((body?.claims || []).map((claim) => findRepositoryDoctor(index, claim)).filter(Boolean).map((claim) => ({ ...claim, matchedAt: new Date().toISOString() })));
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: repositoryImportRefsForClaims(index, claims),
      };
      const updated = {
        ...targetRecord,
        email: targetEmail,
        claims,
        state,
        updatedAt: new Date().toISOString(),
      };
      await context.env.ROSTER_STORE.put(storageKey(targetEmail), JSON.stringify(updated));
      const owner = accountSnapshotOwner(targetEmail, updated.role || roleForEmail(targetEmail));
      await context.env.ROSTER_STORE.delete(snapshotKey(owner.ownerType, owner.ownerId));
      return Response.json({
        ok: true,
        user: userSummaryFromRecord(targetEmail, updated),
        claims,
      });
    }

    if (action === "removeRosterClaim") {
      const claimEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = claimEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, claimEmail);
      const rawClaim = {
        sourceType: String(body?.claim?.sourceType || "").toLowerCase(),
        key: normalizeRosterName(body?.claim?.key || ""),
      };
      const claims = sanitizeClaims(targetRecord.claims).filter((claim) => !(claim.sourceType === rawClaim.sourceType && claim.key === rawClaim.key));
      const index = await loadRepositoryIndex(context.env.ROSTER_STORE);
      const state = {
        ...sanitizeState(targetRecord.state),
        imports: repositoryImportRefsForClaims(index, claims),
      };
      const updated = {
        ...targetRecord,
        email: claimEmail,
        claims,
        state,
        updatedAt: new Date().toISOString(),
      };
      await context.env.ROSTER_STORE.put(storageKey(claimEmail), JSON.stringify(updated));
      const owner = accountSnapshotOwner(claimEmail, updated.role || roleForEmail(claimEmail));
      await context.env.ROSTER_STORE.delete(snapshotKey(owner.ownerType, owner.ownerId));
      return Response.json({ ok: true, claims, user: userSummaryFromRecord(claimEmail, updated) });
    }

    if (action === "reportRosterIdentityIssue") {
      const reportEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = reportEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, reportEmail);
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
      await context.env.ROSTER_STORE.put(storageKey(reportEmail), JSON.stringify(updated));
      return Response.json({ ok: true, user: userSummaryFromRecord(reportEmail, updated) });
    }

    if (action === "resolveDoctorAccount") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const resolution = await resolveDoctorAccount(context.env.ROSTER_STORE, body?.doctor);
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
      const targetRecord = await loadAccountRecord(context.env.ROSTER_STORE, targetEmail);
      const targetRole = targetRecord.role || roleForEmail(targetEmail);
      if (targetRole === "creator" || targetRole === "owner") {
        return Response.json({ error: "Creator insights cannot be disabled." }, { status: 400 });
      }
      const updated = {
        ...targetRecord,
        insightsEnabled: body?.insightsEnabled === true,
        updatedAt: new Date().toISOString(),
      };
      await context.env.ROSTER_STORE.put(storageKey(targetEmail), JSON.stringify(updated));
      return Response.json({
        ok: true,
        user: userSummaryFromRecord(targetEmail, updated),
      });
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
        seniority: body?.issue?.seniority,
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
      if (dismissed.has(issue.fingerprint) || ignored.has(issue.fingerprint) || await isIssueResolvedByParserRules(context.env.ROSTER_STORE, reportEmail, issue)) {
        await clearIssuesResolvedByIssue(context.env.ROSTER_STORE, reportEmail, issue);
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
      const previousSeniority = sanitizeRuleSeniority(body?.previousSeniority || rule.seniority);
      let parserExtensions = await loadParserExtensionRules(context.env.ROSTER_STORE);
      if (previousCode && (previousCode !== rule.code || previousSeniority !== rule.seniority)) {
        const sourceKey = rule.source.toLowerCase();
        parserExtensions = {
          ...parserExtensions,
          [sourceKey]: (parserExtensions[sourceKey] || []).filter((item) => item.code !== previousCode || item.seniority !== previousSeniority),
        };
      }
      parserExtensions = upsertParserExtensionRule(parserExtensions, rule);
      await saveParserExtensionRules(context.env.ROSTER_STORE, parserExtensions);
      const ignoreFingerprint = sanitizeIssueFingerprint(body?.fingerprint || issueFingerprint(body?.source, body?.rawValue));
      if (ignoreFingerprint) {
        await clearIssueFromAllUsers(context.env.ROSTER_STORE, ignoreFingerprint);
      }
      await clearIssuesResolvedByParserRule(context.env.ROSTER_STORE, rule);
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
      const parserExtensions = removeParserExtensionRuleByKey(await loadParserExtensionRules(context.env.ROSTER_STORE), target);
      await saveParserExtensionRules(context.env.ROSTER_STORE, parserExtensions);
      return Response.json({ ok: true, parserExtensions });
    }

    if (action === "saveLocalParserExtensionRule") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, saveEmail);
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
      let suggestions = await loadParserRuleSuggestions(context.env.ROSTER_STORE);
      suggestions = upsertParserRuleSuggestion(suggestions, suggestion);
      await saveParserRuleSuggestions(context.env.ROSTER_STORE, suggestions);
      await context.env.ROSTER_STORE.put(storageKey(saveEmail), JSON.stringify({
        ...targetRecord,
        localParserExtensions,
        updatedAt: new Date().toISOString(),
      }));
      await clearIssuesResolvedByParserRuleForUser(context.env.ROSTER_STORE, saveEmail, rule);
      return Response.json({
        ok: true,
        issueConfig: await buildIssueConfig(context.env.ROSTER_STORE, saveEmail),
      });
    }

    if (action === "decideParserRuleSuggestion") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const suggestionId = String(body?.suggestionId || "").trim();
      const decision = String(body?.decision || "").trim();
      let suggestions = await loadParserRuleSuggestions(context.env.ROSTER_STORE);
      const suggestion = suggestions.find((item) => item.id === suggestionId);
      if (!suggestion) {
        return Response.json({ error: "Suggestion was not found." }, { status: 404 });
      }
      const rule = sanitizeParserExtensionRule(body?.rule || suggestion.rule);
      if ((decision === "approveGlobal" || decision === "approveUser") && !rule) {
        return Response.json({ error: "A valid shift-code rule is required." }, { status: 400 });
      }
      if (decision === "approveGlobal") {
        const parserExtensions = upsertParserExtensionRule(await loadParserExtensionRules(context.env.ROSTER_STORE), rule);
        await saveParserExtensionRules(context.env.ROSTER_STORE, parserExtensions);
        await clearIssuesResolvedByParserRule(context.env.ROSTER_STORE, rule);
        await removeLocalParserRuleFromAllUsers(context.env.ROSTER_STORE, rule);
        suggestions = suggestions.filter((item) => item.id !== suggestionId);
      } else if (decision === "approveUser") {
        await upsertLocalParserRuleForUser(context.env.ROSTER_STORE, suggestion.email, rule);
        await clearIssuesResolvedByParserRuleForUser(context.env.ROSTER_STORE, suggestion.email, rule);
        suggestions = suggestions.filter((item) => item.id !== suggestionId);
      } else if (decision === "reject") {
        await removeLocalParserRuleForUser(context.env.ROSTER_STORE, suggestion.email, suggestion.rule);
        suggestions = suggestions.filter((item) => item.id !== suggestionId);
      } else {
        return Response.json({ error: "Unsupported suggestion decision." }, { status: 400 });
      }
      await saveParserRuleSuggestions(context.env.ROSTER_STORE, suggestions);
      return Response.json({
        ok: true,
        parserExtensions: await loadParserExtensionRules(context.env.ROSTER_STORE),
        suggestions,
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
      const record = await loadAccountRecord(context.env.ROSTER_STORE, deleteEmail).catch(() => null);
      if (record?.subscriptionToken) {
        await context.env.ROSTER_STORE.delete(subscriptionTokenKey(record.subscriptionToken));
      }
      if (record?.email) {
        const owner = accountSnapshotOwner(record.email, record.role || roleForEmail(record.email));
        await context.env.ROSTER_STORE.delete(snapshotKey(owner.ownerType, owner.ownerId));
      }
      await context.env.ROSTER_STORE.delete(storageKey(deleteEmail));
      await clearDeletedAccountClaimMetadata(context.env.ROSTER_STORE, deleteEmail, record);
      return Response.json({ ok: true, deletedEmail: deleteEmail });
    }

    if (action === "save") {
      const saveEmail = targetEmail && (account.role === "creator" || account.role === "owner") ? targetEmail : email;
      const targetRecord = saveEmail === email ? account.record : await loadAccountRecord(context.env.ROSTER_STORE, saveEmail);
      const targetRole = targetRecord.role || roleForEmail(saveEmail);
      const state = sanitizeState(body?.state);
      const repository = await upsertStateImports(context.env.ROSTER_STORE, state.imports, saveEmail);
      state.imports = repository.refs;
      const claims = sanitizeClaims(targetRecord.claims);
      const removedImportIds = sanitizeRepositoryFileIds(body?.removedImportIds);
      if ((targetRole === "creator" || targetRole === "owner") && saveEmail === email && removedImportIds.length) {
        repository.index = await removeRepositoryFiles(context.env.ROSTER_STORE, repository.index, removedImportIds);
        state.imports = state.imports.filter((item) => {
          const repoId = item.repoId || item.repositoryId || item.id;
          return !removedImportIds.includes(repoId);
        });
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

    if (action === "loadImportRefs") {
      if (account.role !== "creator" && account.role !== "owner") {
        return Response.json({ error: "Creator access is required." }, { status: 403 });
      }
      const refs = sanitizeSnapshotFileRefs(body?.refs || []);
      const imports = await resolveStateImports(context.env.ROSTER_STORE, refs);
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
    return userSummaryFromRecord(email, record);
  }));
  return users.sort((a, b) => a.email.localeCompare(b.email));
}

function userSummaryFromRecord(email, record) {
  const claims = sanitizeClaims(record?.claims);
  const adminIssues = sanitizeAdminIssues(record?.adminIssues);
  return {
    email,
    realName: String(record?.realName || "").trim(),
    role: record?.role || roleForEmail(email),
    sites: [...new Set(claims.map((claim) => claim.sourceType.toUpperCase()))].sort(),
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

export async function prepareAccountResponse(store, rawRecord, options = {}) {
  let record = await ensureAccountSubscriptionToken(store, rawRecord);
  const role = record.role || roleForEmail(record.email);
  const index = await loadRepositoryIndex(store);
  let claims = sanitizeClaims(record.claims);
  let nameMatches = [];
  let state = sanitizeState(record.state);
  let linkedProfiles = [];

  if (role !== "creator" && role !== "owner") {
    const originalClaims = claims;
    claims = claims.filter((claim) => claimMatchesAccountIdentity(claim, record.realName || ""));
    const matchedClaims = matchRepositoryClaims(index, record.realName || "");
    nameMatches = matchedClaims.filter((claim) => !claims.some((existing) => sameClaim(existing, claim)));
    linkedProfiles = await linkedDoctorProfilesForClaims(store, claims);
    const accountImportRefs = repositoryImportRefsForClaims(index, claims);
    state = {
      ...state,
      imports: accountImportRefs,
    };
    state = mergeProfileSessionIntoState(state, linkedProfiles, record.email);
    const previousState = sanitizeState(record.state);
    const importRefsChanged = importsChanged(previousState.imports, accountImportRefs);
    const claimsChanged = JSON.stringify(claims) !== JSON.stringify(originalClaims);
    if (claimsChanged || importRefsChanged) {
      await store.put(storageKey(record.email), JSON.stringify({
        ...record,
        claims,
        state: {
          ...previousState,
          imports: accountImportRefs.map(repositoryImportRef),
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
    realName: record.realName || "",
    claims,
    state,
    linkedProfiles,
    index,
  });
  const snapshot = await loadSnapshotRecord(store, owner.ownerType, owner.ownerId);
  const snapshotAvailable = Boolean(snapshot);
  const snapshotStale = !snapshot || snapshot.schemaVersion !== SNAPSHOT_SCHEMA_VERSION || snapshot.buildStamp !== buildStamp;
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
    insightsEnabled: insightsEnabledForRecord(record),
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
  const globalParserExtensions = await loadParserExtensionRules(store);
  const record = email ? await loadAccountRecord(store, email).catch(() => null) : null;
  const role = record?.role || roleForEmail(email);
  const localParserExtensions = sanitizeParserExtensionRules(record?.localParserExtensions);
  return {
    parserExtensions: mergeParserExtensionSets(globalParserExtensions, localParserExtensions),
    globalParserExtensions,
    localParserExtensions,
    parserRuleSuggestions: role === "creator" || role === "owner" ? await loadParserRuleSuggestions(store) : [],
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

async function clearIssuesResolvedByParserRule(store, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedRule) return;
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    if (!record?.adminIssues?.length) continue;
    const existingIssues = sanitizeAdminIssues(record.adminIssues);
    const nextIssues = existingIssues.filter((issue) => !issueMatchesParserRule(issue, normalizedRule));
    if (nextIssues.length === existingIssues.length) continue;
    await store.put(item.name, JSON.stringify({
      ...record,
      adminIssues: nextIssues,
      updatedAt: new Date().toISOString(),
    }));
  }
}

async function clearIssuesResolvedByParserRuleForUser(store, email, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedRule || !normalizedEmail) return;
  const record = await loadAccountRecord(store, normalizedEmail).catch(() => null);
  if (!record?.adminIssues?.length) return;
  const existingIssues = sanitizeAdminIssues(record.adminIssues);
  const nextIssues = existingIssues.filter((issue) => !issueMatchesParserRule(issue, normalizedRule));
  if (nextIssues.length === existingIssues.length) return;
  await store.put(storageKey(normalizedEmail), JSON.stringify({
    ...record,
    adminIssues: nextIssues,
    updatedAt: new Date().toISOString(),
  }));
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

async function isIssueResolvedByParserRules(store, email, issue) {
  const source = sanitizeIssueSource(issue?.source);
  const seniority = sanitizeRuleSeniority(issue?.seniority);
  const code = parserRuleCodeForIssue(issue);
  if (!source || !code) return false;
  const config = await buildIssueConfig(store, email);
  const rules = sanitizeParserExtensionRules(config.parserExtensions);
  const sourceRules = rules[source.toLowerCase()] || [];
  return sourceRules.some((rule) => rule.source === source && rule.seniority === seniority && rule.code === code);
}

async function clearIssuesResolvedByIssue(store, email, issue) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return;
  const record = await loadAccountRecord(store, normalizedEmail).catch(() => null);
  if (!record?.adminIssues?.length) return;
  const existingIssues = sanitizeAdminIssues(record.adminIssues);
  const nextIssues = existingIssues.filter((item) => !sameParserIssue(item, issue));
  if (nextIssues.length === existingIssues.length) return;
  await store.put(storageKey(normalizedEmail), JSON.stringify({
    ...record,
    adminIssues: nextIssues,
    updatedAt: new Date().toISOString(),
  }));
}

function sameParserIssue(left, right) {
  return sanitizeIssueSource(left?.source) === sanitizeIssueSource(right?.source)
    && sanitizeRuleSeniority(left?.seniority) === sanitizeRuleSeniority(right?.seniority)
    && parserRuleCodeForIssue(left) === parserRuleCodeForIssue(right);
}

function parserRuleCodeForIssue(issue) {
  return parserRuleCodeFromRawValue(issue?.source, issue?.rawValue);
}

function parserRuleCodeFromRawValue(sourceValue, rawValue) {
  const source = sanitizeIssueSource(sourceValue);
  const text = String(rawValue || "").trim();
  const upper = text.toUpperCase();
  if (source === "MMC" || source === "Casey") {
    const prefixMatch = upper.match(/^\s*\d{2}:?\d{2}\s*[-–]\s*\d{2}:?\d{2}\s+(.+?)\s*$/);
    if (prefixMatch) return prefixMatch[1].trim().toUpperCase();
    const suffixMatch = upper.match(/^\s*(.+?)\s+\d{2}:?\d{2}\s*[-–]\s*\d{2}:?\d{2}\s*$/);
    if (suffixMatch) return suffixMatch[1].trim().toUpperCase();
  }
  return upper;
}

async function upsertLocalParserRuleForUser(store, email, rule) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedEmail || !normalizedRule) return;
  const record = await loadAccountRecord(store, normalizedEmail).catch(() => null);
  if (!record) return;
  await store.put(storageKey(normalizedEmail), JSON.stringify({
    ...record,
    localParserExtensions: upsertParserExtensionRule(record.localParserExtensions, normalizedRule),
    updatedAt: new Date().toISOString(),
  }));
}

async function removeLocalParserRuleForUser(store, email, rule) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedEmail || !normalizedRule) return;
  const record = await loadAccountRecord(store, normalizedEmail).catch(() => null);
  if (!record) return;
  const localParserExtensions = removeParserExtensionRule(record.localParserExtensions, normalizedRule);
  await store.put(storageKey(normalizedEmail), JSON.stringify({
    ...record,
    localParserExtensions,
    updatedAt: new Date().toISOString(),
  }));
}

async function removeLocalParserRuleFromAllUsers(store, rule) {
  const normalizedRule = sanitizeParserExtensionRule(rule);
  if (!normalizedRule) return;
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    if (!record?.localParserExtensions) continue;
    const existing = JSON.stringify(sanitizeParserExtensionRules(record.localParserExtensions));
    const localParserExtensions = removeParserExtensionRule(record.localParserExtensions, normalizedRule);
    if (JSON.stringify(localParserExtensions) === existing) continue;
    await store.put(item.name, JSON.stringify({
      ...record,
      localParserExtensions,
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
  const rawFiles = Array.isArray(raw?.files) ? raw.files : [];
  const files = rawFiles.map(sanitizeRepositoryFile).filter((file) => file && file.active !== false);
  return {
    version: 1,
    files,
  };
}

function sanitizeRepositoryFileIds(ids = []) {
  return [...new Set((Array.isArray(ids) ? ids : [])
    .map((id) => String(id || "").trim())
    .filter(Boolean))];
}

async function removeRepositoryFiles(store, index, ids = []) {
  const removedIds = new Set(sanitizeRepositoryFileIds(ids));
  if (!removedIds.size) return index;
  const files = (index.files || []).filter((file) => !removedIds.has(file.id));
  await Promise.all([...removedIds].map((id) => store.delete(repositoryFileKey(id))));
  const next = { ...index, files };
  await saveRepositoryIndex(store, next);
  return next;
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
      displayName: formatRosterDisplayName(doctor?.displayName || doctor?.key || ""),
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
    : repositoryImportRefsForAccount(context?.index || await loadRepositoryIndex(store), {
        email: context?.email || "",
        realName: context?.realName || "",
        claims: context?.claims || [],
        state: context?.state || {},
      });
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

function repositoryImportRefsForAccount(index, record) {
  const claims = sanitizeClaims(record?.claims);
  return repositoryImportRefsForClaims(index, claims);
}

async function repositoryImportsForClaims(store, index, claims) {
  return resolveStateImports(store, repositoryImportRefsForClaims(index, claims));
}

async function resolveAccountImports(store, record) {
  const role = record?.role || roleForEmail(record?.email || "");
  const state = sanitizeState(record?.state);
  if (role === "creator" || role === "owner") {
    const index = await loadRepositoryIndex(store);
    const activeRefs = (index.files || []).filter((file) => file.active !== false).map(repositoryImportRef);
    return resolveStateImports(store, activeRefs.length ? activeRefs : state.imports || []);
  }
  const index = await loadRepositoryIndex(store);
  return resolveStateImports(store, repositoryImportRefsForAccount(index, record));
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
    const hasProfileDoctor = sanitizeRepositoryDoctors(file.doctors).some((doctor) => (
      doctor.key === profile.doctorKey
    ));
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
    realName: context.realName || context.record?.realName || "",
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
  const realIdentity = rosterIdentityKey(realName);
  if (!realIdentity) return claims;
  for (const file of index.files || []) {
    if (file.active === false) continue;
    for (const doctor of sanitizeRepositoryDoctors(file.doctors)) {
      if (rosterIdentityKey(doctor.displayName || doctor.key) !== realIdentity) continue;
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
  const accountIndex = await loadClaimedAccountIndex(store);
  const seen = new Set();
  const candidates = [];
  for (const file of index.files || []) {
    if (file.active === false) continue;
    for (const doctor of sanitizeRepositoryDoctors(file.doctors)) {
      const marker = `${doctor.sourceType}:${doctor.key}`;
      if (seen.has(marker)) continue;
      seen.add(marker);
      const resolved = resolveDoctorAccountFromIndex(accountIndex, doctor);
      candidates.push({
        key: doctor.key,
        displayName: doctor.displayName,
        sourceType: doctor.sourceType,
        claimedBy: resolved.mode === "claimed-account" ? resolved.email : "",
        claimedByName: resolved.mode === "claimed-account" ? resolved.realName : "",
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

async function resolveDoctorAccount(store, rawDoctor) {
  return resolveDoctorAccountFromIndex(await loadClaimedAccountIndex(store), rawDoctor);
}

async function loadClaimedAccountIndex(store) {
  const accounts = [];
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    const email = normalizeEmail(record?.email || item.name.replace(/^account:/, ""));
    const role = record?.role || roleForEmail(email);
    if (!record || !email || role === "creator" || role === "owner") continue;
    accounts.push({
      email,
      realName: String(record.realName || "").trim(),
      claims: sanitizeClaims(record.claims),
    });
  }
  return accounts;
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
  await removeDeletedClaimsFromRemainingAccounts(store, deletedEmail, deletedClaims);
}

async function removeDeletedClaimsFromRemainingAccounts(store, deletedEmail, deletedClaims) {
  const result = await store.list({ prefix: "account:" });
  for (const item of result.keys || []) {
    const record = await store.get(item.name, "json").catch(() => null);
    const email = normalizeEmail(record?.email || item.name.replace(/^account:/, ""));
    if (!record || email === deletedEmail) continue;
    const claims = sanitizeClaims(record.claims);
    const filtered = claims.filter((claim) => !deletedClaims.some((deletedClaim) => sameClaim(claim, deletedClaim)));
    if (filtered.length === claims.length) continue;
    await store.put(item.name, JSON.stringify({
      ...record,
      claims: filtered,
      updatedAt: new Date().toISOString(),
    }));
    const owner = accountSnapshotOwner(email, record.role || roleForEmail(email));
    await store.delete(snapshotKey(owner.ownerType, owner.ownerId));
  }
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
  const realIdentity = rosterIdentityKey(realName);
  if (!realIdentity) return true;
  return rosterIdentityKey(claim?.displayName || claim?.key || "") === realIdentity;
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
  return normalizeRosterName(value).replace(/^(DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A PROF|ASSOC PROF)\s+/, "");
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
      seniority: sanitizeRuleSeniority(item?.seniority),
      date: String(item?.date || item?.startDay || "").trim(),
      rawValue: String(item?.rawValue || "").trim(),
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

async function loadParserRuleSuggestions(store) {
  const value = await store.get(PARSER_RULE_SUGGESTIONS_KEY, "json").catch(() => []);
  return sanitizeParserRuleSuggestions(value);
}

async function saveParserRuleSuggestions(store, value) {
  const sanitized = sanitizeParserRuleSuggestions(value);
  if (!sanitized.length) {
    await store.delete(PARSER_RULE_SUGGESTIONS_KEY);
    return sanitized;
  }
  await store.put(PARSER_RULE_SUGGESTIONS_KEY, JSON.stringify(sanitized));
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
