export const MMC_CONTACT_LIST_SOURCE_ID = "mmc-shift-allocations";
export const DDH_CONTACT_LIST_SOURCE_ID = "ddh-daily-contact-sheet";

const SOURCE_AREAS = new Map([
  [MMC_CONTACT_LIST_SOURCE_ID, new Set(["Adult Emergency", "Paediatric Emergency"])],
  [DDH_CONTACT_LIST_SOURCE_ID, new Set(["Dandenong Emergency"])],
]);
const SOURCE_FILE_NAMES = new Map([
  [MMC_CONTACT_LIST_SOURCE_ID, "SHIFT ALLOCATIONS doctors.json"],
  [DDH_CONTACT_LIST_SOURCE_ID, "Daily Contact Sheet clinicians.json"],
]);
const VALID_SHIFTS = new Set(["AM", "PM", "Night"]);
const NAME_ALIASES = new Map([
  ["ian", new Set(["ian", "yiran"])],
  ["pat", new Set(["pat", "patrick", "patricia"])],
  ["patrick", new Set(["pat", "patrick"])],
  ["mel", new Set(["mel", "melanie"])],
  ["melanie", new Set(["mel", "melanie"])],
  ["michael", new Set(["michael", "mickey"])],
  ["mickey", new Set(["michael", "mickey"])],
  ["jacqui", new Set(["jacqui", "jacqueline"])],
  ["jacqueline", new Set(["jacqui", "jacqueline"])],
  ["steve", new Set(["steve", "stephen"])],
  ["stephen", new Set(["steve", "stephen"])],
  ["yiran", new Set(["ian", "yiran"])],
]);

export function normaliseContactListExtract(payload) {
  const sourceId = String(payload?.sourceId || "").trim();
  const validAreas = SOURCE_AREAS.get(sourceId);
  if (!validAreas || !Array.isArray(payload?.contacts)) return null;
  const sourceDate = String(payload?.sourceDate || "").trim();
  if (!isIsoDate(sourceDate) || payload.contacts.length > 240) return null;
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
  }).filter((entry) => !isTemporarilyExcludedContactRole(sourceId, entry));
  if (contacts.some((entry) => !validAreas.has(entry.area)
    || !VALID_SHIFTS.has(entry.shift)
    || !entry.role
    || /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(entry.role))) return null;
  const occurrences = new Map();
  const keyedContacts = contacts.map((contact) => {
    const base = contactKeyBase(sourceId, sourceDate, contact);
    const occurrence = occurrences.get(base) || 0;
    occurrences.set(base, occurrence + 1);
    return { ...contact, contactKey: `${base}|${occurrence}` };
  });
  return {
    sourceId,
    fileName: SOURCE_FILE_NAMES.get(sourceId),
    sourceDate,
    providerModifiedAt: String(payload?.providerModifiedAt || "").trim(),
    contacts: keyedContacts,
  };
}

export function contactResolutionKey(sourceId, sourceDate, contact, occurrence = 0) {
  return `${contactKeyBase(sourceId, sourceDate, contact)}|${Math.max(0, Number(occurrence) || 0)}`;
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
  if (!melbourne.date) return true;
  return melbourne.date > nextDate
    || (melbourne.date === nextDate && (melbourne.hour * 60 + melbourne.minute) >= 10 * 60);
}

// The clinical operational day rolls over at the first morning handover,
// rather than at midnight. A Tuesday Night allocation therefore remains on
// Tuesday until 07:30 on Wednesday.
export function contactOperationalDate(now = new Date()) {
  const melbourne = melbourneDateTime(now);
  if (!melbourne.date) return "";
  return (melbourne.hour * 60 + melbourne.minute) < (7 * 60 + 30)
    ? addDays(melbourne.date, -1)
    : melbourne.date;
}

// Once the operational day turns over at 07:30, retain the preceding
// Night handset allocations until 10:00 during live testing. This also lets the server use the
// last Night extract when a new morning extract has not arrived yet.
export function shouldCarryPreviousNightContacts(sourceDate, requestedDate, now = new Date()) {
  const melbourne = melbourneDateTime(now);
  if (!melbourne.date || requestedDate !== melbourne.date) return false;
  const minuteOfDay = melbourne.hour * 60 + melbourne.minute;
  return minuteOfDay >= 7 * 60 + 30
    && minuteOfDay < 10 * 60
    && sourceDate === addDays(requestedDate, -1);
}

export function contactAreaForSource(source) {
  const code = String(source || "").trim().toUpperCase();
  if (code === "MMC") return "Adult Emergency";
  if (code === "MCH") return "Paediatric Emergency";
  if (code === "DDH") return "Dandenong Emergency";
  return "";
}

// Contact sheets are commonly populated in advance with the team that is
// still carrying the phones. For today's roster, expose a period only after
// the latest normal handover time so an outgoing team cannot be attached to
// the incoming roster. Past dates remain complete; future dates remain hidden.
export function contactPeriodsAfterShiftChange(date, now = new Date()) {
  const selectedDate = String(date || "").trim();
  if (!isIsoDate(selectedDate)) return new Set();
  const melbourne = melbourneDateTime(now);
  if (selectedDate < melbourne.date) return new Set(VALID_SHIFTS);
  if (selectedDate > melbourne.date) return new Set();
  const minuteOfDay = melbourne.hour * 60 + melbourne.minute;
  const periods = new Set();
  if (minuteOfDay >= 7 * 60 + 30) periods.add("AM");
  // The new operational day starts at 07:30, but the outgoing Night team
  // remains responsible for its handsets until 10:00 during live testing.
  if (minuteOfDay >= 7 * 60 + 30 && minuteOfDay < 10 * 60) periods.add("Night");
  if (minuteOfDay >= 15 * 60) periods.add("PM");
  if (minuteOfDay >= 23 * 60) periods.add("Night");
  return periods;
}

export function contactsAfterShiftChange(contacts = [], { date = "", now = new Date() } = {}) {
  const periods = contactPeriodsAfterShiftChange(date, now);
  return (contacts || []).filter((contact) => periods.has(String(contact?.shift || "")));
}

export function attachContactAllocations(assignments = [], contacts = [], resolutions = []) {
  const available = (contacts || [])
    .filter((contact) => (contact?.isPopulated && contact.name) || isRoleOnlyServiceContact(contact))
    .map((contact, index) => ({ ...contact, contactKey: String(contact.contactKey || contactResolutionKey("legacy", "", contact, index)) }));
  const used = new Set();
  const enriched = assignments.map((assignment) => ({ ...assignment }));
  const orderedContacts = [...available].sort((left, right) => Number(isRoleOnlyServiceContact(right)) - Number(isRoleOnlyServiceContact(left))
    || contactSpecificity(right) - contactSpecificity(left)
    || String(left.name).localeCompare(String(right.name)));
  const unmatchedReasons = new Map();

  for (const contact of orderedContacts) {
    const contextCandidates = enriched
      .map((assignment, index) => ({ assignment, index }))
      .filter(({ assignment, index }) => !used.has(index) && assignmentMatchesContactContext(assignment, contact));
    if (!contextCandidates.length) {
      unmatchedReasons.set(contact.contactKey, "No roster candidate in this period");
      continue;
    }
    const roleOnlyServiceKey = !contact.name && isRoleOnlyServiceContact(contact) ? contactStreamKey(contact.role) : "";
    if (roleOnlyServiceKey) {
      const serviceCandidates = contextCandidates.filter(({ assignment }) => assignmentStreamKey(assignment) === roleOnlyServiceKey);
      if (serviceCandidates.length !== 1) {
        unmatchedReasons.set(contact.contactKey, serviceCandidates.length ? "Ambiguous service allocation" : "No rostered service allocation");
        continue;
      }
      const candidate = serviceCandidates[0];
      used.add(candidate.index);
      const stream = contactStream(contact.role);
      enriched[candidate.index] = {
        ...candidate.assignment,
        contactAllocation: {
          role: contact.role,
          phone: contact.phone,
          sourceName: "",
          contactKey: contact.contactKey,
          matchMethod: "service-role",
          streamKey: stream.key,
          streamLabel: stream.label,
          rosterStreamKey: assignmentStreamKey(candidate.assignment),
        },
      };
      continue;
    }
    const named = contextCandidates
      .map(({ assignment, index }) => {
        const nameMatch = personMatch(contact.name, assignment?.person?.displayName || assignment?.doctorName || "");
        return nameMatch ? {
          assignment,
          index,
          ...nameMatch,
          streamAligned: Boolean(contactStreamKey(contact.role)) && contactStreamKey(contact.role) === assignmentStreamKey(assignment),
        } : null;
      })
      .filter(Boolean)
      .sort((left, right) => right.score - left.score || Number(right.streamAligned) - Number(left.streamAligned));
    if (!named.length) {
      unmatchedReasons.set(contact.contactKey, "No safe name match");
      continue;
    }
    const candidate = named[0];
    // One-word entries are deliberately conservative.  A clerk's "Pat",
    // "Qing" or "Tara" can only be resolved after other specific entries
    // have consumed every other plausible roster candidate.
    if ((nameTokens(contact.name).length === 1 && named.length > 1)
      || (named.length > 1 && named[1].score === candidate.score && named[1].streamAligned === candidate.streamAligned)) {
      unmatchedReasons.set(contact.contactKey, "Ambiguous name");
      continue;
    }
    const index = candidate.index;
    used.add(index);
    const stream = contactStream(contact.role);
    const rosterStreamKey = assignmentStreamKey(candidate.assignment);
    enriched[index] = {
      ...candidate.assignment,
      contactAllocation: {
        role: contact.role,
        phone: contact.phone,
        sourceName: contact.name,
        contactKey: contact.contactKey,
        matchMethod: candidate.method,
        streamKey: stream.key,
        streamLabel: stream.label,
        rosterStreamKey,
      },
    };
  }

  // A temporary resolution is deliberately applied only after the conservative
  // automatic matcher. It can connect a review row to one rostered person but
  // must never displace a safe automatic allocation or alter roster streams.
  const unresolvedKeys = new Set(available.filter((contact) => !matchedContactsForAssignments(enriched).has(contact.contactKey)).map((contact) => contact.contactKey));
  const manualTargets = new Set();
  for (const resolution of resolutions || []) {
    if (resolution?.active === false || !unresolvedKeys.has(String(resolution?.contactKey || ""))) continue;
    const contact = available.find((item) => item.contactKey === String(resolution.contactKey));
    const targetIndex = enriched.findIndex((assignment) => assignmentMatchesContactContext(assignment, contact)
      && String(assignment?.person?.doctorKey || "") === String(resolution?.doctorKey || "")
      && !assignment.contactAllocation);
    if (!contact || targetIndex < 0 || manualTargets.has(targetIndex)) continue;
    manualTargets.add(targetIndex);
    enriched[targetIndex] = {
      ...enriched[targetIndex],
      contactAllocation: {
        role: contact.role, phone: contact.phone, sourceName: contact.name,
        contactKey: contact.contactKey, matchMethod: "manual", streamKey: contactStream(contact.role).key,
        streamLabel: contactStream(contact.role).label, rosterStreamKey: assignmentStreamKey(enriched[targetIndex]),
        resolutionId: String(resolution.id || ""), resolutionRevision: Number(resolution.revision || 0),
      },
    };
  }

  const matchedContacts = matchedContactsForAssignments(enriched);
  const serviceContacts = available.filter((contact) => !matchedContacts.has(contact.contactKey)
    && isStandaloneServiceContact(contact));
  const standaloneServiceKeys = new Set(serviceContacts.map((contact) => contact.contactKey));
  return {
    assignments: enriched,
    matchedCount: enriched.filter((assignment) => assignment.contactAllocation).length,
    unmatched: available.filter((contact) => !matchedContacts.has(contact.contactKey)
      && !standaloneServiceKeys.has(contact.contactKey)
      && (!isRoleOnlyServiceContact(contact) || unmatchedReasons.get(contact.contactKey) === "Ambiguous service allocation")).map((contact) => ({
      ...contact,
      reviewReason: unmatchedReasons.get(contact.contactKey) || "Not matched",
    })),
    serviceContacts,
  };
}

function matchedContactsForAssignments(assignments) {
  return new Set((assignments || []).map((assignment) => assignment.contactAllocation?.contactKey).filter(Boolean));
}

function contactKeyBase(sourceId, sourceDate, contact) {
  return [sourceId, sourceDate, contact?.area, contact?.shift, contact?.role, contact?.name, contact?.phone]
    .map((value) => encodeURIComponent(String(value || "").trim().toLowerCase()))
    .join("|");
}

function isTemporarilyExcludedContactRole(sourceId, contact) {
  if (sourceId !== DDH_CONTACT_LIST_SOURCE_ID || String(contact?.shift || "") !== "AM") return false;
  const role = String(contact?.role || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  return role.startsWith("geriatrician in ed") || role.startsWith("cart np npc") || role.startsWith("miprep hmo");
}

function isRoleOnlyServiceContact(contact) {
  const phoneDigits = String(contact?.phone || "").replace(/\D/g, "");
  if (phoneDigits.length < 5 || phoneDigits.length > 10 || contact?.name) return false;
  const key = contactStreamKey(contact?.role);
  if (contact?.area === "Adult Emergency") return ["sepsis", "geriatrics", "cart"].includes(key);
  if (contact?.area === "Dandenong Emergency") return ["care-co", "gap", "clinical-support-onsite"].includes(key);
  return false;
}

function isStandaloneServiceContact(contact) {
  const key = contactStreamKey(contact?.role);
  const phoneDigits = String(contact?.phone || "").replace(/\D/g, "");
  if (phoneDigits.length < 5 || phoneDigits.length > 10) return false;
  if (contact?.area === "Adult Emergency") return ["geriatrics", "cart"].includes(key);
  if (contact?.area === "Dandenong Emergency") return ["care-co", "gap"].includes(key);
  return false;
}

function assignmentMatchesContactContext(assignment, contact) {
  const source = String(assignment?.source || assignment?.person?.sourceType || "").trim().toUpperCase();
  return contact.area === contactAreaForSource(source) && String(assignment?.period || "") === contact.shift;
}

function contactSpecificity(contact) {
  const tokens = nameTokens(contact?.name);
  // A surname initial (for example "Tara K") must be resolved before the
  // generic first name that follows it (for example "Tara").
  const hasSurnameInitial = tokens.length > 1 && tokens[tokens.length - 1].length === 1;
  return tokens.length * 10 + (hasSurnameInitial ? 5 : 0) + (contactStreamKey(contact?.role) ? 1 : 0);
}

export function contactStream(role) {
  const text = simplify(role);
  if (/\bclinical support on site\b|\bclinical support onsite\b/.test(text)) return { key: "clinical-support-onsite", label: "Clinical Support on-site" };
  if (/\bed care co\b/.test(text)) return { key: "care-co", label: "ED Care-Co" };
  if (/\bgap\b|\bgeriatric ah\b/.test(text)) return { key: "gap", label: "GAP / Geriatric AH" };
  if (/\borange\b/.test(text)) return { key: "orange", label: "Orange" };
  if (/\bsilver\b/.test(text)) return { key: "silver", label: "Silver" };
  if (/\bgreen\b/.test(text)) return { key: "green", label: "Green" };
  if (/\bamber\b/.test(text)) return { key: "amber", label: "Amber" };
  if (/\bresus\b/.test(text)) return { key: "resus", label: "Resus" };
  if (/\bclinic\b/.test(text)) return { key: "clinic", label: "Clinic" };
  if (/\bhub\b/.test(text)) return { key: "hub", label: "Hub" };
  if (/\bssu\b/.test(text)) return { key: "ssu", label: "SSU" };
  if (/\bsepsis\b/.test(text)) return { key: "sepsis", label: "Sepsis" };
  if (/\bgeriatric/.test(text)) return { key: "geriatrics", label: "Geriatrics" };
  if (/\bcart\b/.test(text)) return { key: "cart", label: "CART" };
  if (/\bavao\b/.test(text)) return { key: "avao", label: "AVAO" };
  if (/\bfast track\b|\bft\b/.test(text)) return { key: "fast track", label: "Fast Track" };
  return { key: "", label: "" };
}

function contactStreamKey(role) {
  return contactStream(role).key;
}

function assignmentStreamKey(assignment) {
  const text = simplify(`${assignment?.team || ""} ${assignment?.suggestedTitle || ""} ${assignment?.rawValue || ""} ${assignment?.event?.title || ""} ${assignment?.event?.rawValue || ""}`);
  if (/\bclinical support on site\b|\bclinical support onsite\b|\bcs onsite\b|\bonsite cs\b/.test(text)
    && !/\bnot onsite\b/.test(text)) return "clinical-support-onsite";
  if (/\bed care co\b/.test(text)) return "care-co";
  if (/\bgap\b|\bgeriatric ah\b/.test(text)) return "gap";
  if (/\borange\b/.test(text)) return "orange";
  if (/\bsilver\b/.test(text)) return "silver";
  if (/\bgreen\b/.test(text)) return "green";
  if (/\bamber\b/.test(text)) return "amber";
  if (/\bresus\b/.test(text)) return "resus";
  if (/\bclinic\b/.test(text)) return "clinic";
  if (/\bhub\b/.test(text)) return "hub";
  if (/\bssu\b/.test(text)) return "ssu";
  if (/\bsepsis\b/.test(text)) return "sepsis";
  if (/\bgeriatric/.test(text)) return "geriatrics";
  if (/\bcart\b/.test(text)) return "cart";
  if (/\bavao\b/.test(text)) return "avao";
  if (/\bfast track\b|\bft\b/.test(text)) return "fast track";
  return "";
}

function personMatch(contactName, rosterName) {
  const contact = nameTokens(contactName);
  const roster = nameTokens(rosterName);
  if (!contact.length || !roster.length) return null;
  if (contact.join(" ") === roster.join(" ")) return { method: "exact", score: 100 };

  const firstNamesMatch = namesMatch(contact[0], roster[0]);
  const surnameInitial = contact.length > 1 ? contact[contact.length - 1] : "";
  if (firstNamesMatch && surnameInitial.length === 1 && roster.some((token) => token.startsWith(surnameInitial))) {
    return { method: firstNamesMatch === "alias" ? "alias-surname-initial" : "surname-initial", score: 90 };
  }
  if (contact[0] === roster[0]) return { method: "first-name", score: 70 };
  if (firstNamesMatch === "alias") return { method: "alias", score: 65 };
  if (contact[0].length >= 4 && roster[0].startsWith(contact[0])) return { method: "first-name-prefix", score: 60 };
  if (contact.length === 1 && roster.slice(1).includes(contact[0])) return { method: "internal-given-name", score: 55 };
  return null;
}

function namesMatch(left, right) {
  if (left === right) return "exact";
  if (NAME_ALIASES.get(left)?.has(right) || NAME_ALIASES.get(right)?.has(left)) return "alias";
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
  const date = now instanceof Date ? now : new Date(now);
  if (Number.isNaN(date.valueOf())) return { date: "", hour: 0, minute: 0 };
  const values = Object.fromEntries(new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
  }).formatToParts(date).filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  return { date: `${values.year}-${values.month}-${values.day}`, hour: Number(values.hour), minute: Number(values.minute) };
}
