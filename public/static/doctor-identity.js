const TITLE_PATTERN = /^(?:DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A\s+PROF|ASSOC\s+PROF)\b[\s.]*/i;
const DEFAULT_BUCKET_LIMIT = 80;
const DEFAULT_CANDIDATE_LIMIT = 500;

function asciiWords(value = "") {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(TITLE_PATTERN, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
}

export function normalizeDoctorIdentityName(value = "") {
  const words = asciiWords(value);
  const given = words[0] || "";
  const surnameWords = words.slice(1);
  const surname = surnameWords.join("");
  return {
    words,
    given,
    surname,
    givenInitial: given.slice(0, 1),
    exactKey: words.join(""),
    displayKey: words.join(" "),
  };
}

export function doctorSourceIdentityKey(value = {}) {
  const sourceType = String(value?.sourceType || "").trim().toLowerCase();
  const doctorKey = String(value?.key || value?.doctorKey || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
  return sourceType && doctorKey ? `${sourceType}:${doctorKey}` : "";
}

export function doctorIdentityPairKey(left, right) {
  const markers = [doctorSourceIdentityKey(left), doctorSourceIdentityKey(right)].filter(Boolean).sort();
  return markers.length === 2 && markers[0] !== markers[1] ? JSON.stringify(markers) : "";
}

function aliasListForDoctor(doctor = {}) {
  const aliases = Array.isArray(doctor.aliases) && doctor.aliases.length ? doctor.aliases : [doctor];
  return aliases.map((alias) => ({
    sourceType: String(alias?.sourceType || doctor.sourceType || "").trim().toLowerCase(),
    key: String(alias?.key || doctor.key || "").trim().toUpperCase().replace(/\s+/g, " "),
    displayName: String(alias?.displayName || alias?.key || doctor.displayName || doctor.key || "").trim().replace(/\s+/g, " "),
    personId: String(alias?.personId || doctor.personId || "").trim(),
    personReference: String(alias?.personReference || doctor.personReference || "").trim(),
    accountEmail: String(alias?.accountEmail || doctor.accountEmail || doctor.claimedBy || "").trim().toLowerCase(),
    claimedByName: String(alias?.claimedByName || doctor.claimedByName || "").trim(),
  }));
}

export function flattenDoctorIdentities(doctors = []) {
  const identities = new Map();
  for (const doctor of doctors || []) {
    for (const alias of aliasListForDoctor(doctor)) {
      const marker = doctorSourceIdentityKey(alias);
      const name = normalizeDoctorIdentityName(alias.displayName || alias.key);
      if (!marker || !name.given || !name.surname) continue;
      const existing = identities.get(marker);
      identities.set(marker, existing ? {
        ...existing,
        personId: existing.personId || alias.personId,
        personReference: existing.personReference || alias.personReference,
        accountEmail: existing.accountEmail || alias.accountEmail,
        claimedByName: existing.claimedByName || alias.claimedByName,
      } : { ...alias, marker, name });
    }
  }
  return [...identities.values()];
}

function levenshteinDistance(left = "", right = "", stopAfter = 3) {
  if (left === right) return 0;
  if (!left || !right) return Math.max(left.length, right.length);
  if (Math.abs(left.length - right.length) > stopAfter) return stopAfter + 1;
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];
    let rowMinimum = current[0];
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      const substitution = previous[rightIndex - 1] + (left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1);
      const value = Math.min(previous[rightIndex] + 1, current[rightIndex - 1] + 1, substitution);
      current.push(value);
      rowMinimum = Math.min(rowMinimum, value);
    }
    if (rowMinimum > stopAfter) return stopAfter + 1;
    previous = current;
  }
  return previous[right.length];
}

function compatibleGivenNames(left = "", right = "") {
  if (!left || !right) return false;
  if (left === right) return true;
  const shorter = left.length <= right.length ? left : right;
  const longer = left.length > right.length ? left : right;
  return shorter.length >= 3 && longer.startsWith(shorter);
}

function candidateEvidence(left, right) {
  const leftName = left.name;
  const rightName = right.name;
  if (leftName.exactKey === rightName.exactKey) {
    return { rank: 0, confidence: "formatting", reason: "Same letters with different spacing, punctuation, title, or case" };
  }
  if (leftName.surname === rightName.surname && compatibleGivenNames(leftName.given, rightName.given)) {
    return { rank: 1, confidence: "likely", reason: "Same surname with a compatible shortened first name" };
  }
  const surnameDistance = levenshteinDistance(leftName.surname, rightName.surname, 2);
  if (leftName.given === rightName.given && surnameDistance <= 2) {
    return { rank: 1, confidence: "likely", reason: "Same first name with a small surname spelling difference" };
  }
  if (leftName.surname === rightName.surname && leftName.givenInitial === rightName.givenInitial) {
    return { rank: 2, confidence: "check", reason: "Same surname and first initial" };
  }
  if (compatibleGivenNames(leftName.given, rightName.given) && surnameDistance <= 2) {
    return { rank: 2, confidence: "check", reason: "Similar first and surname spellings" };
  }
  return null;
}

function addToBucket(buckets, key, index) {
  if (!key) return;
  if (!buckets.has(key)) buckets.set(key, []);
  buckets.get(key).push(index);
}

export function buildDoctorIdentityCandidates(doctors = [], options = {}) {
  const identities = flattenDoctorIdentities(doctors);
  const bucketLimit = Math.max(2, Number(options.bucketLimit || DEFAULT_BUCKET_LIMIT));
  const candidateLimit = Math.max(1, Number(options.candidateLimit || DEFAULT_CANDIDATE_LIMIT));
  const buckets = new Map();
  identities.forEach((identity, index) => {
    const { exactKey, surname, given, givenInitial } = identity.name;
    addToBucket(buckets, `exact:${exactKey}`, index);
    addToBucket(buckets, `surname-initial:${surname}:${givenInitial}`, index);
    addToBucket(buckets, `given-surname-prefix:${given}:${surname.slice(0, 3)}`, index);
    addToBucket(buckets, `surname-given-prefix:${surname}:${given.slice(0, 3)}`, index);
  });

  const compared = new Set();
  const candidates = [];
  let comparisonCount = 0;
  let skippedLargeBuckets = 0;
  for (const bucket of buckets.values()) {
    if (bucket.length > bucketLimit) {
      skippedLargeBuckets += 1;
      continue;
    }
    for (let leftIndex = 0; leftIndex < bucket.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < bucket.length; rightIndex += 1) {
        const left = identities[bucket[leftIndex]];
        const right = identities[bucket[rightIndex]];
        const pairKey = doctorIdentityPairKey(left, right);
        if (!pairKey || compared.has(pairKey)) continue;
        compared.add(pairKey);
        comparisonCount += 1;
        if (left.personId && left.personId === right.personId) continue;
        const evidence = candidateEvidence(left, right);
        if (!evidence) continue;
        candidates.push({ pairKey, left, right, ...evidence });
      }
    }
  }

  candidates.sort((left, right) => (
    left.rank - right.rank
    || left.left.displayName.localeCompare(right.left.displayName)
    || left.right.displayName.localeCompare(right.right.displayName)
  ));
  return {
    candidates: candidates.slice(0, candidateLimit),
    identityCount: identities.length,
    comparisonCount,
    skippedLargeBuckets,
    truncated: candidates.length > candidateLimit,
  };
}

export function suggestDoctorPersonReference(displayName = "") {
  const { words } = normalizeDoctorIdentityName(displayName);
  if (words.length < 2) return words.join("").toLowerCase();
  const given = words[0].toLowerCase();
  const surname = words.slice(1).join("").toLowerCase();
  return `${surname}-${given}`.replace(/[^a-z0-9-]+/g, "").replace(/-+/g, "-").replace(/^-|-$/g, "");
}

