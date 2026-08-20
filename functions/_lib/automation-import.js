import {
  applyRosterEventSeniorities,
  buildRosterView,
  defaultSettings,
  doctorOptions,
  attachFindmyshiftStaffIds,
  findmyshiftProviderStaffOptions,
  findmyshiftRosteredStaffOptions,
  mergeMembershipDoctors,
  normalizeRosterName,
  parseUploadForm,
  serializeEvent,
  setParserExtensions,
} from "./roster.js";

export const AUTOMATION_SOURCES = {
  "monash-adults": { provider: "sharepoint", sourceType: "mmc", label: "Monash Adults" },
  "monash-paeds": { provider: "sharepoint", sourceType: "mch", label: "Monash Paediatrics" },
  "dandenong-findmyshift": { provider: "findmyshift", sourceType: "ddh", label: "Dandenong (Findmyshift)" },
};

const REPARSE_ONLY_SOURCES = {
  "casey-manual": { provider: "manual", sourceType: "casey", label: "Casey" },
};

export function automationSourceDefinition(sourceId) {
  return AUTOMATION_SOURCES[String(sourceId || "").trim()] || REPARSE_ONLY_SOURCES[String(sourceId || "").trim()] || null;
}

export async function sha256Hex(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

export async function buildAutomatedDerivedRosterPayload({ file, sourceId, contentHash, fileId = "", providerVersion = "", parserExtensions = null }) {
  const source = automationSourceDefinition(sourceId);
  if (!source) throw new Error("Unknown automation source.");
  if (!(file instanceof File)) throw new Error("A roster file is required.");
  // Automated parsing must use the same global rules as an interactive upload.
  // The worker supplies a fresh set for every run so rule changes take effect
  // without relying on a long-lived Node process.
  if (parserExtensions && typeof parserExtensions === "object") setParserExtensions(parserExtensions);

  const resolvedFileId = String(fileId || `automation:${sourceId}:${String(contentHash || "").slice(0, 24)}`);
  const addedAt = new Date().toISOString();
  const formData = new FormData();
  formData.append("rosterFiles", file, file.name);
  formData.append("rosterFileId", resolvedFileId);
  formData.append("rosterFileAddedAt", addedAt);
  const parsed = await parseUploadForm(new Request("https://automation.invalid/import", {
    method: "POST",
    body: formData,
  }));
  const detected = Object.entries(parsed.sources || {}).find(([, entries]) => Array.isArray(entries) && entries.length);
  const detectedType = String(detected?.[0] || "").toLowerCase();
  if (detectedType !== source.sourceType) {
    throw new Error(`${source.label} must be an ${source.sourceType.toUpperCase()} roster file.`);
  }

  const sourceEntries = {
    mmc: source.sourceType === "mmc" ? parsed.sources.mmc || [] : [],
    ddh: source.sourceType === "ddh" ? parsed.sources.ddh || [] : [],
    casey: [],
    mch: source.sourceType === "mch" ? parsed.sources.mch || [] : [],
  };
  const doctors = doctorOptions(sourceEntries.mmc, sourceEntries.ddh, sourceEntries.casey, sourceEntries.mch)
    .flatMap((doctor) => {
      const aliases = Array.isArray(doctor.aliases) && doctor.aliases.length
        ? doctor.aliases.filter((alias) => String(alias.sourceType || "").toLowerCase() === source.sourceType)
        : [{ key: doctor.key, displayName: doctor.displayName, sourceType: source.sourceType }];
      return aliases.map((alias) => ({
        key: normalizeRosterName(alias.key || doctor.key),
        displayName: alias.displayName || doctor.displayName,
        sourceType: source.sourceType,
      }));
    });
  const rosteredProviderStaff = source.sourceType === "ddh" ? findmyshiftRosteredStaffOptions(sourceEntries.ddh) : [];
  const providerDoctors = source.sourceType === "ddh" ? findmyshiftProviderStaffOptions(sourceEntries.ddh) : [];
  const uniqueDoctors = uniqueRosterDoctors(mergeMembershipDoctors(attachFindmyshiftStaffIds(doctors, rosteredProviderStaff), providerDoctors));
  const settings = {
    ...defaultSettings(),
    hospitalFilter: "all",
    dateFrom: "",
    dateTo: "",
    includeLocations: true,
  };
  const eventsByDoctor = {};
  const issuesByDoctor = {};
  let eventCount = 0;
  for (const doctor of uniqueDoctors) {
    const view = buildRosterView(
      sourceEntries.mmc,
      sourceEntries.ddh,
      doctor.key,
      settings,
      {},
      {},
      [],
      sourceEntries.casey,
      sourceEntries.mch,
    );
    const events = view.events.map(serializeEvent);
    eventsByDoctor[doctor.key] = events;
    issuesByDoctor[doctor.key] = (view.issues || []).map(serializeIssue);
    eventCount += events.length;
  }
  if (!uniqueDoctors.length || !eventCount) {
    throw new Error(`${source.label} could not be indexed reliably (${uniqueDoctors.length} doctors, ${eventCount} events).`);
  }
  return {
    file: {
      id: resolvedFileId,
      name: file.name || `${sourceId}.xlsx`,
      sourceType: source.sourceType,
      sourceId,
      size: Number(file.size || 0),
      lastModified: Number(file.lastModified || Date.now()),
      addedAt,
      uploadedAt: addedAt,
      uploadedBy: `automation:${sourceId}`,
      providerVersion: String(providerVersion || ""),
    },
    doctors: applyRosterEventSeniorities(uniqueDoctors, eventsByDoctor),
    eventsByDoctor,
    issuesByDoctor,
    eventCount,
  };
}

function uniqueRosterDoctors(doctors = []) {
  const seen = new Set();
  return doctors.filter((doctor) => {
    const marker = `${doctor.sourceType}:${doctor.key}`;
    if (!doctor.key || seen.has(marker)) return false;
    seen.add(marker);
    return true;
  });
}

function serializeIssue(issue = {}) {
  return {
    id: issue.id,
    source: issue.source,
    seniority: issue.seniority,
    startDay: issue.startDay,
    rawValue: issue.rawValue,
    status: issue.status,
    message: issue.message,
    resolutionType: issue.resolutionType,
    suggestedTitle: issue.suggestedTitle,
    timeLabel: issue.timeLabel,
  };
}
