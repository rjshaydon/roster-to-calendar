import assert from "node:assert/strict";
import { contactOperationalDate } from "../public/static/contact-allocations.js";
import {
  loadPublishedFacilityContacts,
  publishFacilityContactExtract,
  publishFacilityContactResolutions,
} from "../functions/_lib/facility-contact-cache.js";

class LocalR2 {
  constructor() { this.objects = new Map(); this.gets = 0; this.puts = 0; this.version = 0; }
  async get(key) {
    this.gets += 1;
    const item = this.objects.get(key);
    if (!item) return null;
    return { etag: item.etag, text: async () => item.text };
  }
  async put(key, value, options = {}) {
    const current = this.objects.get(key);
    if (options.onlyIf?.etagMatches && current?.etag !== options.onlyIf.etagMatches) throw new Error("Precondition failed");
    if (options.onlyIf?.etagDoesNotMatch === "*" && current) throw new Error("Precondition failed");
    this.puts += 1;
    this.version += 1;
    this.objects.set(key, { text: typeof value === "string" ? value : new TextDecoder().decode(value), etag: `local-${this.version}` });
  }
}

const r2 = new LocalR2();
const sourceDate = contactOperationalDate();
const extract = {
  sourceId: "mmc-shift-allocations",
  sourceDate,
  providerModifiedAt: new Date().toISOString(),
  contacts: [{ area: "Adult Emergency", shift: "AM", role: "Consultant", name: "Alex Example", phone: "555-0100", isPopulated: true }],
};
const first = await publishFacilityContactExtract(r2, extract, { receivedAt: new Date().toISOString() });
assert.equal(first.changed, true);
const firstWrites = r2.puts;
const repeated = await publishFacilityContactExtract(r2, extract);
assert.equal(repeated.unchanged, true);
assert.equal(r2.puts, firstWrites, "an unchanged contact extract must write no R2 objects");

const loaded = await loadPublishedFacilityContacts(r2, { date: sourceDate, facilityKeys: ["MMC"] });
assert.equal(loaded.status, "available");
assert.equal(loaded.sourceId, "mmc-shift-allocations");
assert.ok(loaded.revision);

await publishFacilityContactResolutions(r2, loaded.sourceId, sourceDate, [{ contactKey: "test", doctorKey: "ALEX EXAMPLE", active: true, revision: 1 }]);
const resolved = await loadPublishedFacilityContacts(r2, { date: sourceDate, facilityKeys: ["MMC"] });
assert.equal(resolved.resolutions.length, 1);
assert.notEqual(resolved.revision, loaded.revision, "a correction must independently change the contact overlay revision");

const readsBeforeRefreshes = r2.gets;
const visiblePageRefreshes = 50 * 12 * 60;
for (let refresh = 0; refresh < visiblePageRefreshes; refresh += 1) {
  const response = await loadPublishedFacilityContacts(r2, { date: sourceDate, facilityKeys: ["MMC"] });
  assert.equal(response.revision, resolved.revision);
}
assert.equal(r2.gets - readsBeforeRefreshes, visiblePageRefreshes * 3,
  "50 visible pages refreshing every minute for 12 hours should remain three bounded R2 reads per refresh");
assert.equal(r2.puts, firstWrites + 1, "36,000 unchanged refreshes must not write storage");
console.log("Shared contact publication passed: 36,000 unchanged refreshes used zero D1 reads and zero writes.");
