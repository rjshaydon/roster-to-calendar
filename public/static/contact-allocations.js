export const MMC_CONTACT_LIST_SOURCE_ID = "mmc-shift-allocations";
export const DDH_CONTACT_LIST_SOURCE_ID = "ddh-daily-contact-sheet";
export const VHH_CONTACT_LIST_SOURCE_ID = "vhh-shift-phone-allocations";

const SOURCE_AREAS = new Map([
  [MMC_CONTACT_LIST_SOURCE_ID, new Set(["Adult Emergency", "Paediatric Emergency"])],
  [DDH_CONTACT_LIST_SOURCE_ID, new Set(["Dandenong Emergency"])],
  [VHH_CONTACT_LIST_SOURCE_ID, new Set(["Victorian Heart Hospital Emergency"])],
]);
const SOURCE_FILE_NAMES = new Map([
  [MMC_CONTACT_LIST_SOURCE_ID, "SHIFT ALLOCATIONS doctors.json"],
  [DDH_CONTACT_LIST_SOURCE_ID, "Daily Contact Sheet clinicians.json"],
  [VHH_CONTACT_LIST_SOURCE_ID, "Zebra Allocations clinicians.json"],
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
  ["ollie", new Set(["oliver"])],
  ["meg", new Set(["megha", "megan", "meghan", "margaret"])],
  ["ben", new Set(["benjamin", "benedict", "bennett"])],
  ["rosie", new Set(["rosemary", "rose", "rosalind", "rosalyn"])],
]);

export function normaliseContactListExtract(payload) {
  const sourceId = String(payload?.sourceId || "").trim();
  if (sourceId === VHH_CONTACT_LIST_SOURCE_ID && Array.isArray(payload?.doctors)) {
    if (payload.doctors.length > 12) return null;
    payload = { ...payload, contacts: [
      ...(payload.cic ? [{ ...payload.cic, role: "CIC" }] : []), ...payload.doctors,
    ].map((entry) => ({ ...entry, area: "Victorian Heart Hospital Emergency", shift: "Current", isPopulated: Boolean(String(entry?.name || "").trim()) })) };
  }
  const validAreas = SOURCE_AREAS.get(sourceId);
  if (!validAreas || !Array.isArray(payload?.contacts)) return null;
  const sourceDate = String(payload?.sourceDate || "").trim();
  if (!isIsoDate(sourceDate) || payload.contacts.length > (sourceId === VHH_CONTACT_LIST_SOURCE_ID ? 13 : 240)) return null;
  const contacts = payload.contacts.map((entry) => {
    const name = String(entry?.name || "").trim();
    return {
      ...(sourceId === VHH_CONTACT_LIST_SOURCE_ID ? { sourceDate } : {}),
      area: String(entry?.area || "").trim(),
      shift: String(entry?.shift || "").trim(),
      role: String(entry?.role || "").trim(),
      name,
      phone: String(entry?.phone || "").trim(),
      // A fixed phone left in an empty row must not be treated as an allocation.
      isPopulated: Boolean(entry?.isPopulated) && /\p{L}/u.test(name),
    };
  }).filter((entry) => !isTemporarilyExcludedContactRole(sourceId, entry) && !isContactWorksheetHeader(entry));
  if (contacts.some((entry) => !validAreas.has(entry.area)
    || !(sourceId === VHH_CONTACT_LIST_SOURCE_ID ? entry.shift === "Current" : VALID_SHIFTS.has(entry.shift))
    || (sourceId === VHH_CONTACT_LIST_SOURCE_ID && (!/^(CIC|(?:Swing|PM) Consultant(?: \d{4})?|ED Doctor|Sepsis Doctor|SSU Dr)$/i.test(entry.role) || !/^120(?:0[8]|1[03456789]|20)$/.test(entry.phone)))
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
    || (melbourne.date === nextDate && (melbourne.hour * 60 + melbourne.minute) >= 9 * 60);
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
// Night handset allocations until 09:00. This also lets the server use the
// last Night extract when a new morning extract has not arrived yet.
export function shouldCarryPreviousNightContacts(sourceDate, requestedDate, now = new Date()) {
  const melbourne = melbourneDateTime(now);
  if (!melbourne.date || requestedDate !== melbourne.date) return false;
  const minuteOfDay = melbourne.hour * 60 + melbourne.minute;
  return minuteOfDay >= 7 * 60 + 30
    && minuteOfDay < 9 * 60
    && sourceDate === addDays(requestedDate, -1);
}

// A DDH workbook refreshed after the 07:30 rollover is dated today, although
// its Night rows still belong to yesterday's shift. Allow those Night rows to
// be viewed against their actual start date during the carryover window.
export function shouldUseCurrentExtractForPreviousNight(sourceDate, requestedDate, now = new Date()) {
  const melbourne = melbourneDateTime(now);
  if (!melbourne.date || sourceDate !== melbourne.date) return false;
  const minuteOfDay = melbourne.hour * 60 + melbourne.minute;
  return minuteOfDay >= 7 * 60 + 30
    && minuteOfDay < 9 * 60
    && requestedDate === addDays(melbourne.date, -1);
}

export function contactAreaForSource(source) {
  const code = String(source || "").trim().toUpperCase();
  if (code === "MMC") return "Adult Emergency";
  if (code === "MCH") return "Paediatric Emergency";
  if (code === "DDH") return "Dandenong Emergency";
  if (code === "VHH") return "Victorian Heart Hospital Emergency";
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
  // remains responsible for its handsets until 09:00.
  if (minuteOfDay >= 7 * 60 + 30 && minuteOfDay < 9 * 60) periods.add("Night");
  if (minuteOfDay >= 15 * 60) periods.add("PM");
  if (minuteOfDay >= 23 * 60) periods.add("Night");
  return periods;
}

export function contactsAfterShiftChange(contacts = [], { date = "", now = new Date() } = {}) {
  const periods = contactPeriodsAfterShiftChange(date, now);
  return (contacts || []).filter((contact) => periods.has(String(contact?.shift || "")));
}

// A night belongs to its 23:00 start date, including the following morning.
// The 07:30 operational-date rollover does not end the outgoing team at 09:00.
export function ddhNightReviewWindow(date, now = new Date()) {
  const local = melbourneDateTime(now);
  const minutes = local.hour * 60 + local.minute;
  const nightDate = minutes >= 23 * 60 ? local.date : minutes < 9 * 60 ? addDays(local.date, -1) : "";
  const outgoingMorning = minutes >= 7 * 60 + 30 && minutes < 9 * 60 && date === local.date;
  const active = Boolean(nightDate && (date === nightDate || outgoingMorning));
  return { hideNight: date === local.date && !active, nightDate: active ? nightDate : "", previousNightDate: active ? addDays(nightDate, -1) : "" };
}

export function partitionDdhNightReview(matches, assignments, { date, previousNightRoster, now = new Date() } = {}) {
  const window = ddhNightReviewWindow(date, now);
  const night = (contact) => contact.area === "Dandenong Emergency" && contact.shift === "Night";
  const unresolved = (matches.unmatched || []).filter((contact) => !(window.hideNight && night(contact)));
  // Never reinterpret the 07:30–09:00 carryover as tonight's incoming shift.
  if (!window.nightDate || window.nightDate !== date || !previousNightRoster?.available
    || previousNightRoster.nightDate !== window.nightDate || previousNightRoster.previousNightDate !== window.previousNightDate) {
    return { unresolved, previousNight: [] };
  }
  const contacts = unresolved.filter(night);
  const previous = attachContactAllocations(contactRosterAssignments(ddhWorkingNightRows(previousNightRoster.rows || [], window.previousNightDate)), contacts, [], { now });
  const oldMatches = new Map(previous.assignments.filter((assignment) => assignment.contactAllocation)
    .map((assignment) => [assignment.contactAllocation.contactKey, assignment.person.displayName]));
  const previousNight = contacts.filter((contact) => oldMatches.has(contact.contactKey)
    && !contactAllocationCandidates(assignments, contact, { now }).some((candidate) => candidate.nameScore >= CONTACT_MATCH_POLICY.minimumNameScore && candidate.score >= CONTACT_MATCH_POLICY.minimumScore))
    .map((contact) => ({ ...contact, reviewReason: `Likely leftover from night starting ${window.previousNightDate} · ${oldMatches.get(contact.contactKey)}` }));
  const keys = new Set(previousNight.map((contact) => contact.contactKey));
  return { unresolved: unresolved.filter((contact) => !keys.has(contact.contactKey)), previousNight };
}

export function ddhWorkingNightRows(rows, date) {
  return rows.filter((row) => {
    const event = row.event;
    const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
    return String(row.sourceType).toLowerCase() === "ddh" && event?.allDay !== true
      && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(event?.start || "")
      && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(event?.end || "")
      && Date.parse(event.end) > Date.parse(event.start)
      && String(event.start).slice(0, 10) === date && contactRosterPeriod(event) === "Night"
      && ![event.status, event.kind].some((value) => String(value || "").toLowerCase() === "unknown")
      && !/\b(?:leave|conference|cme|annual|sick|personal|study|exam|sabbatical|parental|long service|hith|vhh|phnw|public holiday|clinical support)\b/.test(text);
  });
}

// Scores rank evidence; they are not calibrated probabilities. Context can add
// at most five points and cannot make a name below the evidence floor eligible.
export const CONTACT_MATCH_POLICY = Object.freeze({ minimumNameScore: 88, minimumScore: 90, minimumLead: 12 });

// Approved social names belong to a specific roster identity, not to everybody
// sharing a given name. Add entries only after a clinician's identity is confirmed.
export const APPROVED_CONTACT_NAME_ALIASES = Object.freeze([
  // Confirmed by the roster owner: DDH's Craig Jirayut is rostered as Craig PROMPEN.
  { sourceType: "ddh", doctorKey: "CRAIG PROMPEN", names: ["Craig Jirayut"] },
]);

// Keep an acknowledged local decision while its shared publication is pending.
// A refresh may acknowledge it or replace it with a newer revision, but cannot
// roll it back. Never carry it to a different sheet/date or contact key.
export function mergeContactResolutionRefresh(previous, next) {
  if (!next || previous?.sourceId !== next.sourceId || previous?.sourceDate !== next.sourceDate) return next;
  const resolutions = new Map((next.resolutions || []).map((resolution) => [resolution.contactKey, resolution]));
  const keys = new Set((next.contacts || []).map((contact) => contact.contactKey));
  for (const resolution of previous.resolutions || []) {
    if (resolution.pendingPublication && keys.has(resolution.contactKey)
      && Number(resolutions.get(resolution.contactKey)?.revision || 0) < Number(resolution.revision || 0)) resolutions.set(resolution.contactKey, resolution);
  }
  return { ...next, resolutions: [...resolutions.values()] };
}

export function contactRosterAssignments(rows = [], fallbackSource = "") {
  return rows.map((row) => ({
    source: String(row.sourceType || fallbackSource).toUpperCase(),
    period: contactRosterPeriod(row.event), event: row.event,
    person: { doctorKey: row.doctorKey, displayName: row.displayName, sourceType: row.sourceType || fallbackSource, seniority: row.seniority },
  }));
}

// Used by the save handler as well as local tests. Correction eligibility is
// identical to matching eligibility, including VHH's active timed events.
export function validateContactResolutionSelection(assignments, contacts, resolutions, { contact, doctorKey = "", decision = "assigned", now = new Date() }) {
  if (!["assigned", "cleared", "rejected"].includes(decision) || (decision === "assigned" ? !doctorKey : Boolean(doctorKey))) {
    return { error: "Choose a valid contact allocation decision.", status: 400 };
  }
  if (decision === "cleared") return { target: null };
  const automatic = attachContactAllocations(assignments, contacts, resolutions, { now });
  if (automatic.assignments.some((assignment) => assignment.contactAllocation?.contactKey === contact.contactKey
    && !assignment.contactAllocation.uncertain && assignment.contactAllocation.matchMethod !== "manual")) {
    return { error: "This number already has a safe automatic match.", status: 409 };
  }
  if (decision === "rejected") return { target: null };
  if (automatic.unmatched.some((entry) => entry.contactKey === contact.contactKey && entry.reviewReason === "Conflicting entries in the contact sheet")) {
    return { error: "This contact sheet contains conflicting names or phone numbers. Correct the sheet before assigning this entry.", status: 409 };
  }
  const target = assignments.find((assignment) => assignment.person?.doctorKey === doctorKey && assignmentMatchesContactContext(assignment, contact, now));
  if (!target) return { error: "Choose a clinician rostered in the same ED and shift period, with an active shift for current allocations.", status: 400 };
  if (automatic.assignments.some((assignment) => assignment.person?.doctorKey === doctorKey
    && assignmentMatchesContactContext(assignment, contact, now) && assignment.contactAllocation
    && assignment.contactAllocation.contactKey !== contact.contactKey && !assignment.contactAllocation.uncertain)) {
    return { error: "That clinician already has a confirmed contact allocation.", status: 409 };
  }
  if (resolutions.some((resolution) => resolution.active !== false && resolution.doctorKey === doctorKey && resolution.contactKey !== contact.contactKey
    && contacts.some((other) => other.contactKey === resolution.contactKey && assignmentMatchesContactContext(target, other, now)))) {
    return { error: "That clinician already has a temporary contact allocation.", status: 409 };
  }
  return { target };
}

export function contactAllocationCandidates(assignments = [], contact, { now = new Date() } = {}) {
  const groups = new Map();
  assignments.forEach((assignment, index) => {
    if (!assignmentMatchesContactContext(assignment, contact, now)) return;
    const source = String(assignment.source || assignment.person?.sourceType || "").trim().toUpperCase();
    const doctorKey = String(assignment.person?.doctorKey || assignment.doctorKey || assignment.person?.displayName || assignment.doctorName || "");
    if (!doctorKey) return;
    const identity = `${source}|${doctorKey}|${contact.shift}`;
    const group = groups.get(identity) || { identity, sourceType: source.toLowerCase(), doctorKey, displayName: assignment.person?.displayName || assignment.doctorName || doctorKey, assignments: [], indexes: [] };
    group.assignments.push(assignment); group.indexes.push(index); groups.set(identity, group);
  });
  return [...groups.values()].map((group) => {
    const matchingName = contactMatchName(contact);
    const evidence = personMatch(matchingName, group.displayName, group);
    const annotated = matchingName !== String(contact.name || "").trim();
    const streamAligned = Boolean(contactStreamKey(contact.role)) && group.assignments.some((assignment) => contactStreamKey(contact.role) === assignmentStreamKey(assignment));
    const gradeAligned = group.assignments.some((assignment) => contactGradeAligned(contact.role, assignment.event?.seniority || assignment.person?.seniority));
    return { ...group, method: evidence?.method || "", nameScore: evidence?.score || 0,
      score: evidence ? Math.min(100, evidence.score + (streamAligned ? 3 : 0) + (gradeAligned ? 2 : 0)) : 0,
      uncertain: evidence?.uncertain === true || Boolean(evidence && (annotated || contact.repeatedSheetRows)), streamAligned,
      reasons: evidence ? [evidence.reason, ...(annotated ? ["Repeated handset annotation removed from the name"] : []), ...(contact.repeatedSheetRows ? ["Repeated sheet rows agree on this name and handset"] : []), ...(streamAligned ? ["Roster stream agrees"] : []), ...(gradeAligned ? ["Roster grade agrees"] : [])] : [],
    };
  }).sort((left, right) => right.score - left.score || left.identity.localeCompare(right.identity));
}

export function attachContactAllocations(assignments = [], contacts = [], resolutions = [], { now = new Date() } = {}) {
  const available = coalesceRepeatedContactRows((contacts || []).filter((contact) => !isContactWorksheetHeader(contact)
    && ((contact?.isPopulated && /\p{L}/u.test(contact.name || "")) || isRoleOnlyServiceContact(contact)))
    .map((contact, index) => ({ ...contact, contactKey: String(contact.contactKey || contactResolutionKey("legacy", "", contact, index)) })), resolutions);
  // Never retain an allocation from a previous matching pass.
  const enriched = assignments.map(({ contactAllocation, ...assignment }) => ({ ...assignment }));
  const candidates = new Map(available.map((contact) => [contact.contactKey, contactAllocationCandidates(enriched, contact, { now })]));
  const matched = new Set();
  const usedIndexes = new Set();
  const reasons = new Map();
  const rejected = new Set(resolutions.filter((resolution) => resolution.decision === "rejected").map((resolution) => String(resolution.contactKey)));
  const hasTarget = (candidate) => candidate.indexes.some((index) => usedIndexes.has(index));
  const allocate = (contact, candidate, evidence = {}) => {
    const stream = contactStream(contact.role);
    const allocation = { role: contact.role, phone: contact.phone, sourceName: contact.name,
      contactKey: contact.contactKey, matchMethod: candidate.method, confidenceScore: candidate.score,
      uncertain: candidate.uncertain === true, matchReasons: candidate.reasons || [],
      streamKey: stream.key, streamLabel: stream.label, rosterStreamKey: assignmentStreamKey(candidate.assignments[0]), ...evidence };
    for (const index of candidate.indexes) { usedIndexes.add(index); enriched[index].contactAllocation = allocation; }
    matched.add(contact.contactKey);
  };
  const sameContext = (left, right) => left.area === right.area && left.shift === right.shift;
  // A phone cannot simultaneously belong to different people in one context.
  // Exact duplicate names with different phones are also conflicting sheet rows.
  for (const contact of available) {
    if (available.some((other) => other.contactKey !== contact.contactKey && sameContext(contact, other)
      && ((contact.phone && contact.phone.replace(/\D/g, "") === other.phone?.replace(/\D/g, ""))
        || (contact.name && simplify(contact.name) === simplify(other.name))))) reasons.set(contact.contactKey, "Conflicting entries in the contact sheet");
  }

  // Explicit daily decisions are authoritative, but still require an eligible
  // roster target. Duplicate confirmations never choose an arbitrary winner.
  const manual = resolutions.filter((resolution) => resolution.active !== false && resolution.decision !== "rejected" && resolution.doctorKey)
    .map((resolution) => {
      const contact = available.find((item) => item.contactKey === String(resolution.contactKey));
      const candidate = contact && candidates.get(contact.contactKey).find((item) => item.doctorKey === String(resolution.doctorKey)
        && (!resolution.sourceType || item.sourceType === String(resolution.sourceType).toLowerCase()));
      return { resolution, contact, candidate };
    }).filter((item) => item.candidate);
  for (const item of manual) {
    if (reasons.has(item.contact.contactKey)) continue;
    if (manual.some((other) => other !== item && (other.contact.contactKey === item.contact.contactKey || other.candidate.indexes.some((index) => item.candidate.indexes.includes(index))))) {
      reasons.set(item.contact.contactKey, "Conflicting confirmed allocations"); continue;
    }
    allocate(item.contact, item.candidate, { matchMethod: "manual", uncertain: false, confidenceScore: undefined,
      matchReasons: ["Manually confirmed for this contact sheet"], resolutionId: String(item.resolution.id || ""), resolutionRevision: Number(item.resolution.revision || 0) });
  }
  for (const contact of available) if (rejected.has(contact.contactKey) && !matched.has(contact.contactKey) && !reasons.has(contact.contactKey)) reasons.set(contact.contactKey, "Automatic suggestion rejected for this contact sheet");

  // Dedicated service phones keep their existing priority over ordinary rows.
  const serviceProposals = available.filter((contact) => !contact.name && isRoleOnlyServiceContact(contact) && !reasons.has(contact.contactKey))
    .map((contact) => {
      const eligible = candidates.get(contact.contactKey).filter((candidate) => !hasTarget(candidate) && candidate.assignments.some((assignment) => assignmentStreamKey(assignment) === contactStreamKey(contact.role)));
      if (eligible.length !== 1) reasons.set(contact.contactKey, eligible.length ? "Ambiguous service allocation" : "No rostered service allocation");
      return { contact, candidate: eligible.length === 1 ? eligible[0] : null };
    }).filter((item) => item.candidate);
  for (const proposal of serviceProposals) {
    if (serviceProposals.some((other) => other !== proposal && other.candidate.indexes.some((index) => proposal.candidate.indexes.includes(index)))) {
      reasons.set(proposal.contact.contactKey, "Ambiguous service allocation"); continue;
    }
    allocate(proposal.contact, proposal.candidate, { matchMethod: "service-role", uncertain: false, confidenceScore: undefined });
  }

  // Specific names and then unique given names may reserve identities; guesses never
  // consume candidates to make another tentative decision look unambiguous.
  for (const stage of ["specific", "given", "tentative"]) {
    const proposals = [];
    for (const contact of available) {
      if (!contact.name || matched.has(contact.contactKey) || reasons.has(contact.contactKey)) continue;
      const all = candidates.get(contact.contactKey);
      const eligible = all.filter((candidate) => !hasTarget(candidate) && candidate.nameScore >= CONTACT_MATCH_POLICY.minimumNameScore);
      const candidate = eligible[0];
      if (!candidate || candidate.score < CONTACT_MATCH_POLICY.minimumScore) continue;
      const candidateStage = candidate.uncertain ? "tentative" : candidate.method === "first-name" ? "given" : "specific";
      if (candidateStage !== stage) continue;
      if (eligible[1] && candidate.score - eligible[1].score < CONTACT_MATCH_POLICY.minimumLead) continue;
      // VHH's mutable sheet is particularly prone to retaining old names. A
      // second plausible row claiming this holder remains a conflict.
      if (contact.shift === "Current" && available.some((other) => other.contactKey !== contact.contactKey && sameContext(contact, other)
        && candidates.get(other.contactKey).some((item) => item.identity === candidate.identity && item.nameScore >= CONTACT_MATCH_POLICY.minimumNameScore))) {
        reasons.set(contact.contactKey, "Conflicting entries for this clinician"); continue;
      }
      proposals.push({ contact, candidate });
    }
    // Assess competition before applying any proposal from this batch.
    const accepted = proposals.filter((proposal) => !proposals.some((other) => other !== proposal
      && other.candidate.indexes.some((index) => proposal.candidate.indexes.includes(index))));
    for (const proposal of proposals) if (!accepted.includes(proposal)) reasons.set(proposal.contact.contactKey, "Competing allocations for this clinician");
    for (const proposal of accepted) allocate(proposal.contact, proposal.candidate);
  }

  const serviceContacts = available.filter((contact) => !matched.has(contact.contactKey) && isStandaloneServiceContact(contact));
  const standaloneKeys = new Set(serviceContacts.map((contact) => contact.contactKey));
  return { assignments: enriched, matchedCount: matched.size,
    unmatched: available.filter((contact) => !matched.has(contact.contactKey) && !standaloneKeys.has(contact.contactKey)
      && (!isRoleOnlyServiceContact(contact) || reasons.get(contact.contactKey) === "Ambiguous service allocation")).map((contact) => {
        const ranked = candidates.get(contact.contactKey).filter((candidate) => candidate.score > 0);
        const eligible = ranked.filter((candidate) => !hasTarget(candidate) && candidate.nameScore >= CONTACT_MATCH_POLICY.minimumNameScore);
        return { ...contact, reviewReason: reasons.get(contact.contactKey) || (!candidates.get(contact.contactKey).length
          ? (contact.shift === "Current" ? "No currently rostered clinician matches this entry" : "No roster candidate in this period")
          : eligible.length > 1 ? "Ambiguous name" : ranked.some(hasTarget) ? "Clinician already has a contact allocation" : "No safe name match"),
          candidates: ranked.slice(0, 3).map(({ doctorKey, sourceType, displayName, score, reasons }) => ({ doctorKey, sourceType, displayName, score, reasons })) };
      }), serviceContacts };
}

function isContactWorksheetHeader(contact) {
  return simplify(contact?.role) === "role" && simplify(contact?.name) === "name"
    && ["", "phone"].includes(simplify(contact?.phone));
}

function contactMatchName(contact) {
  const raw = String(contact?.name || "").trim();
  const annotation = raw.match(/^(.*?)\s*(?:[-–—:,]\s*|\s+)([\d\s()+-]{3,})$/u);
  const digits = String(contact?.phone || "").replace(/\D/g, "");
  return annotation && digits.length >= 3 && annotation[2].replace(/\D/g, "") === digits
    ? annotation[1].trim() : raw;
}

function coalesceRepeatedContactRows(contacts, resolutions) {
  const groups = new Map();
  for (const contact of contacts) {
    const digits = String(contact.phone || "").replace(/\D/g, "");
    const name = simplify(contactMatchName(contact));
    const key = digits.length >= 3 && /\p{L}/u.test(name)
      ? JSON.stringify([contact.area, contact.shift, digits, name]) : contact.contactKey;
    const group = groups.get(key) || [];
    group.push(contact); groups.set(key, group);
  }
  return [...groups.values()].flatMap((group) => {
    if (group.length === 1 || group[0].shift === "Current") return group;
    const decisions = resolutions.filter((resolution) => group.some((contact) => contact.contactKey === resolution.contactKey)
      && (resolution.decision === "rejected" || (resolution.active !== false && resolution.doctorKey)));
    const choices = new Set(decisions.map((resolution) => resolution.decision === "rejected" ? "rejected" : resolution.doctorKey));
    // Conflicting human choices still require review. Agreeing decisions retain
    // their original key; rejecting either repetition suppresses the whole group.
    if (choices.size > 1) return group;
    const decidedKeys = new Set(decisions.map((resolution) => resolution.contactKey));
    const ordered = [...group].sort((left, right) => Number(decidedKeys.has(right.contactKey)) - Number(decidedKeys.has(left.contactKey))
      || Number(/^dr(?:\s*\d+)?$/i.test(left.role)) - Number(/^dr(?:\s*\d+)?$/i.test(right.role))
      || left.contactKey.localeCompare(right.contactKey));
    return [{ ...ordered[0], repeatedSheetRows: group.length }];
  });
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

export function assignmentMatchesContactContext(assignment, contact, now = new Date()) {
  if (!contact) return false;
  const source = String(assignment?.source || assignment?.person?.sourceType || "").trim().toUpperCase();
  if (contact.area !== contactAreaForSource(source)) return false;
  if (contact.shift !== "Current") return contactRosterPeriod(assignment?.event, assignment?.period) === contact.shift;
  // VHH has one mutable list, not AM/PM/Night blocks. A sheet name can
  // attach only to an explicitly timed roster event that is active now.
  if (source !== "VHH" || assignment?.event?.allDay === true) return false;
  const current = melbourneDateTime(now);
  if (!current.date || contact.sourceDate !== current.date) return false;
  const start = melbourneEventMinute(assignment?.event?.start);
  const end = melbourneEventMinute(assignment?.event?.end);
  const instant = `${current.date}T${String(current.hour).padStart(2, "0")}:${String(current.minute).padStart(2, "0")}`;
  return Boolean(start && end && start < end && start <= instant && instant < end
    && Date.parse(`${end}:00Z`) - Date.parse(`${start}:00Z`) <= 24 * 60 * 60 * 1000);
}

function melbourneEventMinute(value) {
  const text = String(value || "");
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text)) return "";
  if (/(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
    const local = melbourneDateTime(new Date(text));
    return local.date ? `${local.date}T${String(local.hour).padStart(2, "0")}:${String(local.minute).padStart(2, "0")}` : "";
  }
  return /^\d{4}-\d{2}-\d{2}T(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d(?:\.\d+)?)?$/.test(text) ? text.slice(0, 16) : "";
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
  const text = simplify(assignment?.event?.title || assignment?.event?.rawValue
    ? `${assignment.event.title || ""} ${assignment.event.rawValue || ""}`
    : `${assignment?.team || ""} ${assignment?.suggestedTitle || ""} ${assignment?.rawValue || ""}`);
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

function contactRosterPeriod(event, fallback = "") {
  const text = `${event?.title || ""} ${event?.rawValue || ""}`.toLowerCase();
  if (/\bnight\b/.test(text)) return "Night";
  if (/\bpm\b/.test(text)) return "PM";
  if (/\bam\b/.test(text)) return "AM";
  const start = melbourneEventMinute(event?.start);
  if (!start) return String(fallback || "AM");
  const minutes = Number(start.slice(11, 13)) * 60 + Number(start.slice(14, 16));
  return minutes >= 20 * 60 || minutes < 6 * 60 ? "Night" : minutes >= 12 * 60 + 1 ? "PM" : "AM";
}

function contactGradeAligned(role, seniority) {
  const text = simplify(role);
  const grade = simplify(seniority);
  const code = /^(?:sr|senior registrar)$/.test(grade) ? "sr" : /^(?:tr|ir)|transitional|intermediate/.test(grade) ? "tr"
    : /^(?:jr|junior registrar)$/.test(grade) ? "jr" : /hmo/.test(grade) ? "hmo" : /^(?:sms|consultant|senior medical staff)$/.test(grade) ? "sms" : "";
  return Boolean(code && new RegExp(`\\b${code === "tr" ? "(?:tr|ir)" : code}\\b`).test(text));
}

function nameForms(value) {
  const raw = String(value || "").slice(0, 160);
  const alternatives = [];
  const main = raw.replace(/\(([^()]*)\)|["“]([^"“”]+)["”]/gu, (whole, parenthesized, quoted, offset) => {
    const name = String(parenthesized || quoted || "").trim();
    if (name && /^[\p{L}\p{M}\s'’.-]+$/u.test(name) && !/\b(?:dr|doctor|hmo|sms|sr|tr|jr|ir|am|pm|night|swing|locum|late|early|leave|onsite|offsite|on|off)\b/i.test(name)) alternatives.push({ name, tail: raw.slice(offset + whole.length).replace(/\([^()]*\)|["“][^"“”]+["”]/gu, " ") });
    return " ";
  });
  const forms = [{ tokens: nameTokens(main), alternate: false }];
  for (const alternative of alternatives) {
    const tokens = nameTokens(alternative.name);
    let tail = nameTokens(alternative.tail);
    // Keep a supplied surname when substituting an explicit alternate given
    // name. A bare nickname must not bypass contradictory remaining text.
    if (!tail.length && forms[0].tokens.length > 1 && tokens.length === 1) {
      tail = forms[0].tokens.slice(1);
      if (tail.length > 1 && tail[0] === tokens[tokens.length - 1]) tail.shift();
    }
    forms.push({ tokens: [...tokens, ...tail], alternate: true });
  }
  return forms.filter((form) => form.tokens.length);
}

function nameTokens(value) {
  const tokens = simplify(String(value || "").replace(/([\p{L}])[’'](?=[\p{L}])/gu, "$1")).split(" ").filter(Boolean);
  while (["dr", "doctor", "prof", "mr", "ms", "mrs"].includes(tokens[0])) tokens.shift();
  return tokens;
}

function personMatch(contactName, rosterName, identity) {
  const rosterForms = nameForms(rosterName);
  const forms = nameForms(contactName);
  const approved = APPROVED_CONTACT_NAME_ALIASES.filter((alias) => alias.sourceType === identity?.sourceType && alias.doctorKey === identity?.doctorKey);
  const results = [];
  const hasAlternates = forms.some((form) => form.alternate) || rosterForms.some((form) => form.alternate);
  for (const form of forms) {
    for (const roster of rosterForms) {
      const evidence = tokenMatch(form.tokens, roster.tokens);
      if (evidence) results.push(hasAlternates ? { ...evidence, method: "explicit-alternate-name", uncertain: true, reason: "Explicit alternate name agrees with roster" } : evidence);
    }
    if (approved.some((alias) => (alias.names || []).some((name) => nameTokens(name).join(" ") === form.tokens.join(" ")))) {
      results.push({ method: "approved-identity-alias", score: 98, uncertain: true, reason: "Approved alternate name for this clinician" });
    }
  }
  return results.sort((left, right) => right.score - left.score || Number(left.uncertain) - Number(right.uncertain))[0] || null;
}

function tokenMatch(contact, roster) {
  if (!contact.length || !roster.length) return null;
  const evidence = (method, score, uncertain, reason) => ({ method, score, uncertain, reason });
  if (contact.join(" ") === roster.join(" ")) return evidence("exact", 100, false, "Full name agrees");
  if (contact.length === roster.length && [...contact].sort().join(" ") === [...roster].sort().join(" ")) return evidence("reordered-name", 99, false, "Full name components agree in a different order");
  // Some rosters supply only a given name. Additional sheet components cannot
  // contradict a surname that is absent, but the identity is still inferred.
  // Only an exact given name qualifies; normal ambiguity/competition gates apply.
  if (roster.length === 1 && contact.length > 1 && contact[0] === roster[0]) {
    return evidence("roster-given-name-only", 94, true, "Given name agrees; roster omits the remaining name components");
  }
  if (contact.length === 1) {
    if (contact[0] === roster[0]) return evidence("first-name", 94, false, "Given name agrees");
    if (namesMatch(contact[0], roster[0]) === "alias") return evidence("alias", 92, true, "Recognized shortened given name");
    if (contact[0].length >= 4 && roster[0].startsWith(contact[0]) && contact[0].length / roster[0].length >= 0.5) return evidence("first-name-prefix", 90, true, "Shortened given name agrees");
    if (roster.slice(1).includes(contact[0])) return evidence("internal-given-name", 90, true, "Name component agrees with roster");
    const similarity = spellingMatch(contact[0], roster[0], { givenName: true });
    return similarity ? evidence("spelling", similarity, true, "Small spelling difference in given name") : null;
  }
  // Match an initial/given name against roster components, then account for ALL
  // remaining supplied tokens. Never discard a contradictory supplied surname.
  let best = null;
  for (let firstIndex = 0; firstIndex < Math.max(1, roster.length - 1); firstIndex += 1) {
    const first = contact[0], target = roster[firstIndex];
    const firstMethod = first === target ? "exact" : namesMatch(first, target) === "alias" ? "alias"
      : first.length === 1 && target.startsWith(first) ? "initial"
        : first.length >= 4 && target.startsWith(first) && first.length / target.length >= 0.5 ? "prefix" : spellingMatch(first, target, { givenName: true }) ? "spelling" : "";
    if (!firstMethod) continue;
    const unused = roster.map((token, index) => ({ token, index })).filter((entry) => entry.index !== firstIndex);
    let initial = false, changed = false, valid = true;
    for (const token of contact.slice(1)) {
      const index = unused.findIndex((entry) => token === entry.token || (token.length === 1 && entry.index === roster.length - 1 && entry.token.startsWith(token)));
      const fuzzyIndex = index < 0 ? unused.findIndex((entry) => spellingMatch(token, entry.token)) : -1;
      const chosen = index >= 0 ? index : fuzzyIndex;
      if (chosen < 0) { valid = false; break; }
      initial ||= token.length === 1; changed ||= fuzzyIndex >= 0;
      unused.splice(chosen, 1);
    }
    if (!valid) continue;
    const uncertain = firstMethod !== "exact" || changed || firstIndex > 0;
    const method = initial ? (firstMethod === "alias" ? "alias-surname-initial" : "surname-initial") : uncertain ? "spelling-or-alternate-full-name" : "name-components";
    const score = initial ? (uncertain ? 94 : 98) : uncertain ? 96 : 99;
    const result = evidence(method, score, uncertain, initial ? "Given name and surname initial agree" : uncertain ? "Name components agree with a small spelling or alternate-name difference" : "Supplied name components agree");
    if (!best || result.score > best.score) best = result;
  }
  return best;
}

function namesMatch(left, right) {
  if (left === right) return "exact";
  if (NAME_ALIASES.get(left)?.has(right) || NAME_ALIASES.get(right)?.has(left)) return "alias";
  return "";
}

// Bounded optimal-string-alignment distance handles common adjacent letter
// transpositions. At most two edits are considered, on short name components.
function spellingMatch(left, right, { givenName = false } = {}) {
  if (left.length < 5 || right.length < 5 || left.length > 48 || right.length > 48) return 0;
  const anchoredGivenName = givenName && Math.min(left.length, right.length) >= 7
    && left.slice(0, 4) === right.slice(0, 4) && left.slice(-1) === right.slice(-1);
  const limit = Math.min(left.length, right.length) >= 10 || anchoredGivenName ? 2 : 1;
  if (Math.abs(left.length - right.length) > limit) return 0;
  let previousPrevious = null;
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let i = 1; i <= left.length; i += 1) {
    const row = [i];
    for (let j = 1; j <= right.length; j += 1) {
      row[j] = Math.min(previous[j] + 1, row[j - 1] + 1, previous[j - 1] + (left[i - 1] === right[j - 1] ? 0 : 1));
      if (previousPrevious && i > 1 && j > 1 && left[i - 1] === right[j - 2] && left[i - 2] === right[j - 1]) row[j] = Math.min(row[j], previousPrevious[j - 2] + 1);
    }
    previousPrevious = previous; previous = row;
  }
  const distance = previous[right.length];
  const similarity = 1 - distance / Math.max(left.length, right.length);
  if (distance === 2 && anchoredGivenName && similarity >= 0.75) return 90;
  if (!distance || distance > limit || similarity < (distance === 2 ? 0.85 : 0.8)) return 0;
  return distance === 1 ? 90 : 88;
}

function simplify(value) {
  return String(value || "").normalize("NFD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
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
