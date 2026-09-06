const DB_NAME = "roster-facility-overview";
const STORE_NAME = "snapshots";
export const FACILITY_SNAPSHOT_SCHEMA_VERSION = 1;

export function facilitySnapshotKey({ ownerKey, scopeKey, kind, query }) {
  return [FACILITY_SNAPSHOT_SCHEMA_VERSION, ownerKey, scopeKey, kind, stableJson(query)].join("|");
}

export function facilitySnapshotRecordIsUsable(record, context = {}) {
  return Boolean(record
    && record.schemaVersion === FACILITY_SNAPSHOT_SCHEMA_VERSION
    && record.ownerKey === context.ownerKey
    && record.scopeKey === context.scopeKey
    && Number.isFinite(Date.parse(context.accessExpiresAt || ""))
    && Date.parse(context.accessExpiresAt) > Number(context.now || Date.now()));
}

export async function loadFacilitySnapshot(context, kind, query) {
  if (!facilitySnapshotRecordIsUsable({ schemaVersion: FACILITY_SNAPSHOT_SCHEMA_VERSION, ownerKey: context.ownerKey, scopeKey: context.scopeKey }, context)) return null;
  const db = await openDatabase();
  if (!db) return null;
  const key = facilitySnapshotKey({ ...context, kind, query });
  const record = await requestResult(db.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).get(key)).catch(() => null);
  return facilitySnapshotRecordIsUsable(record, context) ? record.payload : null;
}

export async function storeFacilitySnapshot(context, kind, query, payload) {
  if (!facilitySnapshotRecordIsUsable({ schemaVersion: FACILITY_SNAPSHOT_SCHEMA_VERSION, ownerKey: context.ownerKey, scopeKey: context.scopeKey }, context)) return false;
  const db = await openDatabase();
  if (!db) return false;
  const key = facilitySnapshotKey({ ...context, kind, query });
  await requestResult(db.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).put({
    key, schemaVersion: FACILITY_SNAPSHOT_SCHEMA_VERSION, ownerKey: context.ownerKey, scopeKey: context.scopeKey,
    payload, savedAt: new Date().toISOString(),
  }));
  return true;
}

export async function clearFacilitySnapshots() {
  const db = await openDatabase();
  if (!db) return false;
  await requestResult(db.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).clear());
  return true;
}

function openDatabase() {
  if (typeof indexedDB === "undefined") return Promise.resolve(null);
  return new Promise((resolve) => {
    const request = indexedDB.open(DB_NAME, FACILITY_SNAPSHOT_SCHEMA_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (db.objectStoreNames.contains(STORE_NAME)) db.deleteObjectStore(STORE_NAME);
      db.createObjectStore(STORE_NAME, { keyPath: "key" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
  });
}

function requestResult(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function stableJson(value) {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stableJson(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}
