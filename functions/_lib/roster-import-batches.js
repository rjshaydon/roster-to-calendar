// Pure planning: no D1/R2 work. The execution protocol must persist the plan
// revision and a batch receipt before allowing activation.
const SOURCES = new Set(["monash-adults", "monash-paeds", "dandenong-findmyshift", "vhh-active-medical-roster"]);
export const ROSTER_BATCH_FACT_LIMIT = 1250;
export const ROSTER_BATCH_BYTE_LIMIT = 512 * 1024;
// Conservative table + index-write envelopes, checked against all migrations
// in local tests. D1 bills indexed writes as well as the underlying row.
export const ROSTER_IMPORT_WRITE_COST = { event: 8, issue: 4, presence: 5, doctor: 6, fileDoctor: 4, compact: 6, sms: 4, control: 64 };

export async function planRosterImportBatches(payload, options = {}) {
  const sourceId = String(payload?.file?.sourceId || "");
  if (!SOURCES.has(sourceId) || !payload.file.id) throw new Error("A known source and queued file id are required.");
  const limit = Math.min(ROSTER_BATCH_FACT_LIMIT, Number(options.maximumFacts || ROSTER_BATCH_FACT_LIMIT));
  if (!Number.isInteger(limit) || limit < 10) throw new Error("Invalid roster batch fact budget.");
  const doctors = [...(payload.doctors || [])].sort((a, b) => String(a.key).localeCompare(String(b.key)));
  const keys = new Set(doctors.map((doctor) => doctor.key));
  if (!doctors.length || doctors.length > 512 || keys.size !== doctors.length || doctors.some((doctor) => !doctor.key)) throw new Error("Invalid or duplicate roster doctors.");
  for (const map of [payload.eventsByDoctor || {}, payload.issuesByDoctor || {}]) {
    if (Object.keys(map).some((key) => !keys.has(key) || !Array.isArray(map[key]))) throw new Error("Roster facts refer to an unknown doctor or invalid list.");
  }
  const batches = [];
  let eventCount = 0;
  let issueCount = 0;
  let indexedEventCount = 0;
  let startDate = "";
  let endDate = "";
  let rosterEndDate = "";
  let batch = emptyBatch();
  const finishBatch = () => {
    if (!batch.facts) return;
    indexedEventCount += batch.eventCount;
    batches.push({ ...batch, index: batches.length, indexedEventCount });
    batch = emptyBatch();
  };
  for (const doctor of doctors) {
    for (const [kind, input] of [["eventsByDoctor", payload.eventsByDoctor], ["issuesByDoctor", payload.issuesByDoctor]]) {
      const items = [...(input?.[doctor.key] || [])].sort((a, b) => String(a.id).localeCompare(String(b.id)));
      const identities = new Set();
      for (const item of items) {
        if (!item?.id || identities.has(String(item.id))) throw new Error("Missing or duplicate roster occurrence id.");
        identities.add(String(item.id));
        const isEvent = kind === "eventsByDoctor";
        const presenceRows = isEvent ? presenceRowCount(item) : 0;
        if (isEvent) {
          const from = String(item.start).slice(0, 10);
          const to = String(item.end || item.start).slice(0, 10);
          if (!startDate || from < startDate) startDate = from;
          if (!endDate || to > endDate) endDate = to;
          if (!rosterEndDate || from > rosterEndDate) rosterEndDate = from;
        }
        // Include presence expansion and two doctor/membership rows; staging
        // bookkeeping/receipts get a fixed reserve, separate from fact rows.
        const weight = 1 + presenceRows;
        const bytes = new TextEncoder().encode(canonical(item)).length;
        const doctorBytes = new TextEncoder().encode(canonical(doctor)).length;
        const newDoctor = !batch.doctors.some((entry) => entry.key === doctor.key);
        if (batch.facts + weight + (newDoctor ? 2 : 0) + 16 > limit || batch.bytes + bytes + (newDoctor ? doctorBytes : 0) + 4096 > ROSTER_BATCH_BYTE_LIMIT) finishBatch();
        const addDoctor = !batch.doctors.some((entry) => entry.key === doctor.key);
        if (weight + (addDoctor ? 2 : 0) + 16 > limit || bytes + doctorBytes + 4096 > ROSTER_BATCH_BYTE_LIMIT) throw new Error("One roster fact exceeds the batch safety budget.");
        if (addDoctor) { batch.doctors.push(doctor); batch.facts += 2; batch.bytes += doctorBytes; }
        (batch[kind][doctor.key] ||= []).push(item);
        batch.facts += weight;
        batch.bytes += bytes;
        batch.presenceRows += presenceRows;
        if (isEvent) { batch.eventCount++; eventCount++; } else { batch.issueCount++; issueCount++; }
      }
    }
  }
  finishBatch();
  if (batches.some((entry) => new TextEncoder().encode(JSON.stringify(entry)).length > ROSTER_BATCH_BYTE_LIMIT)) throw new Error("Encoded roster batch exceeds its payload safety budget.");
  if (!eventCount || eventCount > 25000 || issueCount > 5000) throw new Error("Roster exceeds the reviewed total event/issue bounds.");
  const stable = { schemaVersion: 1, sourceId, fileId: payload.file.id, contentHash: payload.file.contentHash || "", doctors, batches, eventCount, issueCount, startDate, endDate, rosterEndDate, maximumFacts: limit };
  const manifest = { ...stable, batches: await Promise.all(batches.map(async (batch) => ({
    index: batch.index, hash: await rosterImportDigest(batch), eventCount: batch.eventCount,
    issueCount: batch.issueCount, indexedEventCount: batch.indexedEventCount, presenceRows: batch.presenceRows,
  }))) };
  validateRosterImportManifest(manifest);
  return { ...stable, manifest, revision: await rosterImportDigest(manifest) };
}

export async function rosterImportDigest(value) {
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(canonical(value)));
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function validateRosterImportManifest(manifest) {
  if (!manifest || !SOURCES.has(manifest.sourceId) || manifest.maximumFacts !== 1250 || !manifest.fileId || manifest.schemaVersion !== 1) throw new Error("Invalid bounded-import manifest.");
  if (!Array.isArray(manifest.doctors) || !manifest.doctors.length || manifest.doctors.length > 512 || new Set(manifest.doctors.map((entry) => entry.key)).size !== manifest.doctors.length || manifest.doctors.some((entry) => !entry.key)) throw new Error("Invalid manifest doctors.");
  if (!Array.isArray(manifest.batches) || !manifest.batches.length || manifest.batches.length > 500 || new TextEncoder().encode(JSON.stringify(manifest)).length > ROSTER_BATCH_BYTE_LIMIT) throw new Error("Manifest exceeds its bounded payload limits.");
  let events = 0;
  let issues = 0;
  for (const [index, batch] of manifest.batches.entries()) {
    if (batch.index !== index || !/^[a-f0-9]{64}$/.test(batch.hash) || !Number.isInteger(batch.eventCount) || batch.eventCount < 0 || !Number.isInteger(batch.issueCount) || batch.issueCount < 0 || !Number.isInteger(batch.presenceRows) || batch.presenceRows < batch.eventCount || batch.eventCount + batch.issueCount + batch.presenceRows + 16 > 1250) throw new Error("Invalid pinned batch descriptor.");
    events += batch.eventCount;
    issues += batch.issueCount;
    if (batch.indexedEventCount !== events) throw new Error("Invalid batch completion cursor.");
  }
  if (events !== manifest.eventCount || issues !== manifest.issueCount || !events || events > 25000 || issues > 5000) throw new Error("Invalid manifest total counts.");
  presenceRowCount({ start: manifest.startDate, end: manifest.endDate }, 180);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(manifest.rosterEndDate) || manifest.rosterEndDate < manifest.startDate || manifest.rosterEndDate > manifest.endDate) throw new Error("Invalid shift-start ownership range.");
}

export function validateRosterImportBatch(batch, manifest) {
  if (!batch || !Array.isArray(batch.doctors) || new TextEncoder().encode(JSON.stringify(batch)).length > ROSTER_BATCH_BYTE_LIMIT) throw new Error("Invalid bounded batch payload.");
  const doctors = new Map(manifest.doctors.map((entry) => [entry.key, canonical(entry)]));
  const keys = new Set(batch.doctors.map((entry) => entry.key));
  if (keys.size !== batch.doctors.length || batch.doctors.some((entry) => doctors.get(entry.key) !== canonical(entry))) throw new Error("Batch doctor metadata differs from the manifest.");
  let events = 0;
  let issues = 0;
  let presence = 0;
  for (const [kind, map] of [["event", batch.eventsByDoctor], ["issue", batch.issuesByDoctor]]) {
    for (const [key, rows] of Object.entries(map || {})) {
      if (!keys.has(key) || !Array.isArray(rows) || new Set(rows.map((entry) => entry.id)).size !== rows.length || rows.some((entry) => !entry.id)) throw new Error("Invalid batch occurrence list.");
      for (const row of rows) {
        if (kind === "event") { if (String(row.start).slice(0, 10) < manifest.startDate || String(row.start).slice(0, 10) > manifest.rosterEndDate) throw new Error("Batch shift start is outside its pinned ownership range."); events++; presence += presenceRowCount(row); } else issues++;
      }
    }
  }
  if (events !== batch.eventCount || issues !== batch.issueCount || presence !== batch.presenceRows || events + issues + presence + keys.size * 2 + 16 > 1250) throw new Error("Batch exceeds its recomputed fact budget.");
}

function emptyBatch() {
  return { doctors: [], eventsByDoctor: {}, issuesByDoctor: {}, facts: 0, bytes: 0, presenceRows: 0, eventCount: 0, issueCount: 0 };
}

function presenceRowCount(event, maximum = 120) {
  const from = String(event.start || "").slice(0, 10);
  const to = String(event.end || event.start || "").slice(0, 10);
  const valid = (value) => /^\d{4}-\d{2}-\d{2}$/.test(value) && Number.isFinite(Date.parse(`${value}T00:00:00Z`)) && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
  if (!valid(from) || !valid(to)) throw new Error("Invalid roster event date.");
  // Match the existing daily-presence index's inclusive end-date semantics.
  const rows = (Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / 86400000 + 1;
  if (rows < 1 || rows > maximum) throw new Error(`Roster event span exceeds the ${maximum}-day safety budget.`);
  return rows;
}

function canonical(value) {
  if (Array.isArray(value)) return `[${value.map(canonical).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}
