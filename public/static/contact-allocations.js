export const MMC_CONTACT_LIST_SOURCE_ID = "mmc-shift-allocations";

const VALID_AREAS = new Set(["Adult Emergency", "Paediatric Emergency"]);
const VALID_SHIFTS = new Set(["AM", "PM", "Night"]);
const NAME_ALIASES = new Map([
  ["pat", new Set(["pat", "patrick", "patricia"])],
  ["patrick", new Set(["pat", "patrick"])],
  ["jacqui", new Set(["jacqui", "jacqueline"])],
  ["jacqueline", new Set(["jacqui", "jacqueline"])],
]);

export function normaliseContactListExtract(payload) {
  if (String(payload?.sourceId || "").trim() !== MMC_CONTACT_LIST_SOURCE_ID || !Array.isArray(payload?.contacts)) return null;
  const sourceDate = String(payload?.sourceDate || "").trim();
  if (!isIsoDate(sourceDate) || payload.contacts.length > 160) return null;
  const contacts = payload.contacts.map((entry) => {
    const name = String(entry?.name || "").trim();
    return {
      area: String(entry?.area || "").trim(),
      shift: String(entry?.shift || "").trim(),
      role: String(entry?.role || "").trim(),
      name,
      phone: String(entry?.phone || "").trim(),
      // A fixed phone left in an empty row must not be treated as an allocation.
      isPopulated: Boolean(entry?.isPopulated) && Boolean(name),
    };
  });
  if (contacts.some((entry) => !VALID_AREAS.has(entry.area)
    || !VALID_SHIFTS.has(entry.shift)
    || !entry.role
    || /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(entry.role))) return null;
  return {
    sourceId: MMC_CONTACT_LIST_SOURCE_ID,
    fileName: "SHIFT ALLOCATIONS doctors.json",
    sourceDate,
    providerModifiedAt: String(payload?.providerModifiedAt || "").trim(),
    contacts,
  };
}

export function contactExtractStatus(extract, { date = "", now = new Date() } = {}) {
  if (!extract?.sourceDate) return "unavailable";
  if (contactExtractHasExpired(extract.sourceDate, now)) return "expired";
  return extract.sourceDate === date ? "available" : "not-current";
}

export function contactExtractHasExpired(sourceDate, now = new Date()) {
  const nextDate = addDays(sourceDate, 1);
  if (!nextDate) return true;
  const melbourne = melbourneDateTime(now);
  return melbourne.date > nextDate || (melbourne.date === nextDate && melbourne.hour >= 10);
}

export function contactAreaForSource(source) {
  const code = String(source || "").trim().toUpperCase();
  if (code === "MMC") return "Adult Emergency";
  if (code === "MCH") return "Paediatric Emergency";
  return "";
}

export function attachContactAllocations(assignments = [], contacts = []) {
  const available = (contacts || [])
    .filter((contact) => contact?.isPopulated && contact.name)
    .map((contact, index) => ({ ...contact, contactKey: `${contact.area}|${contact.shift}|${contact.role}|${contact.name}|${contact.phone}|${index}` }));
  const used = new Set();
  const enriched = assignments.map((assignment) => ({ ...assignment }));
  const orderedContacts = [...available].sort((left, right) => contactSpecificity(right) - contactSpecificity(left));

  for (const contact of orderedContacts) {
    const candidates = enriched.filter((assignment, index) => !used.has(index) && assignmentMatchesContactContext(assignment, contact));
    const named = candidates
      .map((assignment) => ({ assignment, method: personMatchMethod(contact.name, assignment?.person?.displayName || assignment?.doctorName || "") }))
      .filter((candidate) => candidate.method);
    if (named.length !== 1) continue;
    const candidate = named[0];
    const index = enriched.indexOf(candidate.assignment);
    if (index < 0) continue;
    used.add(index);
    enriched[index] = {
      ...candidate.assignment,
      contactAllocation: {
        role: contact.role,
        phone: contact.phone,
        sourceName: contact.name,
        contactKey: contact.contactKey,
        matchMethod: candidate.method,
      },
    };
  }

  const matchedContacts = new Set(enriched.map((assignment) => assignment.contactAllocation?.contactKey).filter(Boolean));
  return {
    assignments: enriched,
    matchedCount: enriched.filter((assignment) => assignment.contactAllocation).length,
    unmatched: available.filter((contact) => !matchedContacts.has(contact.contactKey)),
  };
}

function assignmentMatchesContactContext(assignment, contact) {
  const source = String(assignment?.source || assignment?.person?.sourceType || "").trim().toUpperCase();
  if (contact.area !== contactAreaForSource(source) || String(assignment?.period || "") !== contact.shift) return false;
  const contactStream = contactStreamKey(contact.role);
  return !contactStream || contactStream === assignmentStreamKey(assignment);
}

function contactSpecificity(contact) {
  return contactStreamKey(contact?.role) ? 1 : 0;
}

function contactStreamKey(role) {
  const text = simplify(role);
  if (/\bgreen\b/.test(text)) return "green";
  if (/\bamber\b/.test(text)) return "amber";
  if (/\bresus\b/.test(text)) return "resus";
  if (/\bclinic\b/.test(text)) return "clinic";
  if (/\bhub\b/.test(text)) return "hub";
  if (/\bssu\b/.test(text)) return "ssu";
  if (/\bsepsis\b/.test(text)) return "sepsis";
  if (/\bgeriatric/.test(text)) return "geriatrics";
  if (/\bcart\b/.test(text)) return "cart";
  return "";
}

function assignmentStreamKey(assignment) {
  const text = simplify(`${assignment?.team || ""} ${assignment?.suggestedTitle || ""} ${assignment?.event?.title || ""}`);
  if (/\bgreen\b/.test(text)) return "green";
  if (/\bamber\b/.test(text)) return "amber";
  if (/\bresus\b/.test(text)) return "resus";
  if (/\bclinic\b/.test(text)) return "clinic";
  if (/\bhub\b/.test(text)) return "hub";
  if (/\bssu\b/.test(text)) return "ssu";
  if (/\bsepsis\b/.test(text)) return "sepsis";
  if (/\bgeriatric/.test(text)) return "geriatrics";
  if (/\bcart\b/.test(text)) return "cart";
  return "";
}

function personMatchMethod(contactName, rosterName) {
  const contact = nameTokens(contactName);
  const roster = nameTokens(rosterName);
  if (!contact.length || !roster.length) return "";
  if (contact.join(" ") === roster.join(" ")) return "exact";
  if (contact[0] === roster[0]) return "first-name";
  if (NAME_ALIASES.get(contact[0])?.has(roster[0]) || NAME_ALIASES.get(roster[0])?.has(contact[0])) return "alias";
  return "";
}

function nameTokens(value) {
  return simplify(value).split(" ").filter(Boolean);
}

function simplify(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function addDays(date, days) {
  if (!isIsoDate(date)) return "";
  const value = new Date(`${date}T12:00:00Z`);
  if (Number.isNaN(value.valueOf())) return "";
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

function isIsoDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(`${value}T12:00:00Z`);
  return !Number.isNaN(date.valueOf()) && date.toISOString().slice(0, 10) === value;
}

function melbourneDateTime(now) {
  const values = Object.fromEntries(new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", hourCycle: "h23",
  }).formatToParts(now).filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  return { date: `${values.year}-${values.month}-${values.day}`, hour: Number(values.hour) };
}
