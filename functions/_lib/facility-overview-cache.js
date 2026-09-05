import {
  loadCachedSnapshot,
  queryFacilityStaffDesignations,
  queryFacilityStaffSeniorityOverrides,
  queryMaterializedFacilityMetadata,
  queryMaterializedFacilityTermStaff,
  storeCachedSnapshot,
} from "./d1-calendar.js";

const SCHEMA_VERSION = 1;

export function facilityMetadataManifestKey(sourceType) {
  return `facility-overview/v1/${safeSource(sourceType)}/manifest.json`;
}

export function facilityStaffObjectKey(sourceType, termStart, revision) {
  return `facility-overview/v1/${safeSource(sourceType)}/staff/${String(termStart).slice(0, 10)}/${revision}.json.gz`;
}

export async function publishFacilityStaffMetadata(context, sourceTypes = []) {
  const db = context?.env?.ROSTER_DB;
  const r2 = context?.env?.ROSTER_FILES;
  if (!db?.prepare || !r2?.put || !r2?.get) return { ok: false, unavailable: true };
  const results = [];
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const metadata = await queryMaterializedFacilityMetadata(db, { sourceType });
    const currentManifest = await loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType));
    const terms = [];
    for (const term of metadata.terms || []) {
      const termEnd = addDays(term.termStart, 90);
      const members = (await queryMaterializedFacilityTermStaff(db, { sourceType, termStart: term.termStart, termEnd }))
        .map((member) => ({
          ...member,
          coverageStart: member.firstApplicableDate || "",
          coverageEnd: member.membershipSource === "sms-continuity" ? "" : member.lastApplicableDate || "",
        }));
      const [designations, seniorityOverrides] = await Promise.all([
        queryFacilityStaffDesignations(db, { sourceType, termStart: term.termStart, termEnd }),
        queryFacilityStaffSeniorityOverrides(db, { sourceType, termStart: term.termStart }),
      ]);
      const staff = { schemaVersion: SCHEMA_VERSION, sourceType, termStart: term.termStart, termEnd, members, events: [], coverage: metadata.coverage, designations, seniorityOverrides };
      const staffRevision = await digest(staff);
      const staffKey = facilityStaffObjectKey(sourceType, term.termStart, staffRevision);
      const currentTerm = (currentManifest?.terms || []).find((entry) => entry.termStart === term.termStart);
      if (currentTerm?.staffRevision !== staffRevision) {
        await storeCachedSnapshot(r2, staffKey, staff, { revision: staffRevision, ownerType: "facility-staff", ownerId: sourceType, rangeKey: term.termStart });
      }
      terms.push({ ...term, termEnd, staffKey, staffRevision });
    }
    const stableManifest = { schemaVersion: SCHEMA_VERSION, sourceType, coverage: metadata.coverage, terms };
    const revision = await digest(stableManifest);
    if (currentManifest?.revision !== revision) {
      await storeCachedSnapshot(r2, facilityMetadataManifestKey(sourceType), { ...stableManifest, revision, publishedAt: new Date().toISOString() }, { revision, ownerType: "facility-metadata", ownerId: sourceType, rangeKey: "manifest" });
      results.push({ sourceType, changed: true, revision });
    } else {
      results.push({ sourceType, changed: false, revision });
    }
  }
  return { ok: true, results };
}

export async function loadPublishedFacilityMetadata(r2, sourceTypes, today) {
  if (!r2?.get) return { preparing: true, facilities: [], catalogEvents: [] };
  const manifests = (await Promise.all([...new Set(sourceTypes.map(safeSource).filter(Boolean))]
    .map((sourceType) => loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType))))).filter(Boolean);
  if (!manifests.length) return { preparing: true, facilities: [], catalogEvents: [] };
  const facilities = manifests.flatMap((manifest) => manifest.coverage || []);
  const catalogEvents = [];
  for (const manifest of manifests) {
    for (const term of manifest.terms || []) {
      if (term.visibleFrom > today || term.termEnd < today) continue;
      for (const [index, item] of (term.catalog || []).entries()) {
        for (const date of [...new Set([item.firstDate, item.lastDate].filter(Boolean))]) {
          const start = item.startTime ? `${date}T${item.startTime}` : date;
          const end = item.endTime ? `${date}T${item.endTime}` : date;
          catalogEvents.push({
            doctorKey: `FACILITY_OVERVIEW_CATALOG_${item.catalogKey || index}`,
            displayName: "Roster catalogue",
            sourceType: manifest.sourceType,
            seniority: item.seniority || "",
            date,
            event: { id: `facility-overview-catalog:${item.catalogKey || index}:${date}`, title: item.title || "", rawValue: item.rawValue || "", location: item.location || "", source: manifest.sourceType, sources: [manifest.sourceType], seniority: item.seniority || "", start, end, allDay: item.allDay === true, timeLabel: item.timeLabel || "" },
          });
        }
      }
    }
  }
  return { preparing: false, facilities, catalogEvents };
}

export async function loadPublishedFacilityStaff(r2, sourceTypes, termStart, today) {
  if (!r2?.get) return { preparing: true };
  const payloads = [];
  for (const sourceType of [...new Set(sourceTypes.map(safeSource).filter(Boolean))]) {
    const manifest = await loadCachedSnapshot(r2, facilityMetadataManifestKey(sourceType));
    if (!manifest) continue;
    const term = (manifest.terms || []).find((entry) => entry.termStart === termStart);
    if (!term || term.visibleFrom > today || !term.staffKey) continue;
    const staff = await loadCachedSnapshot(r2, term.staffKey);
    if (staff) payloads.push(staff);
  }
  if (!payloads.length) return { preparing: true };
  return {
    preparing: false,
    members: payloads.flatMap((item) => item.members || []),
    events: payloads.flatMap((item) => item.events || []),
    coverage: payloads.flatMap((item) => item.coverage || []),
    designations: payloads.flatMap((item) => item.designations || []),
    seniorityOverrides: payloads.flatMap((item) => item.seniorityOverrides || []),
  };
}

function safeSource(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9-]+/g, "");
}

function addDays(value, days) {
  const date = new Date(`${String(value).slice(0, 10)}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

async function digest(value) {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  const result = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(result)].map((part) => part.toString(16).padStart(2, "0")).join("");
}
