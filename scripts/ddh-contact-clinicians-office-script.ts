/*
 * Excel Office Script: Extract DDH clinician contacts
 *
 * Save this script against Daily Contact Sheet.xlsx, then select it in the
 * Power Automate "Run script from SharePoint library" action.  Power Automate
 * supplies the SharePoint modification date, version and Melbourne date.
 */
function main(
  workbook: ExcelScript.Workbook,
  sourceDate: string,
  providerModifiedAt: string,
  providerVersion: string,
) {
  const sheet = workbook.getWorksheet("ED Clinicians");
  if (!sheet) throw new Error("Worksheet ED Clinicians was not found.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(sourceDate || ""))) {
    throw new Error("A Melbourne source date in YYYY-MM-DD format is required.");
  }

  // DDH uses four data columns per period with a blank separator column:
  // A:D AM, F:I PM and K:N Night (labelled ND in the workbook).
  const values = sheet.getRange("A1:N120").getTexts();
  const blocks: Array<[string, number]> = [
    ["AM", 0],
    ["PM", 5],
    ["Night", 10],
  ];
  const contacts: Array<{
    area: string; shift: string; role: string; name: string; phone: string; isPopulated: boolean;
  }> = [];

  for (const [shift, firstColumn] of blocks) {
    let continuationRole = "";
    for (const cells of values) {
      const suppliedRole = String(cells[firstColumn] || "").trim();
      if (isHeading(suppliedRole) || isExcludedRole(suppliedRole)) {
        continuationRole = "";
        continue;
      }
      if (suppliedRole && shift !== "Night") {
        if (isContinuationStreamRole(suppliedRole)) continuationRole = suppliedRole;
        else if (isStructuredRole(suppliedRole)) continuationRole = "";
      }
      const rawName = String(cells[firstColumn + 1] || "");
      const name = clinicianName(rawName);
      const phone = clinicianPhone(String(cells[firstColumn + 3] || ""), rawName, String(cells[firstColumn + 2] || ""), shift);
      const role = suppliedRole || (name && phone ? continuationRole : "");
      if (!role) continue;
      contacts.push({
        area: "Dandenong Emergency",
        shift,
        role,
        name,
        phone,
        isPopulated: Boolean(name && /[a-z]/i.test(name)),
      });
    }
  }

  return {
    sourceId: "ddh-daily-contact-sheet",
    sourceDate,
    providerModifiedAt: String(providerModifiedAt || "").trim(),
    providerVersion: String(providerVersion || "").trim(),
    contacts,
  };
}

function clinicianName(value: string) {
  const raw = String(value || "").trim();
  const phone = telephoneMatch(raw);
  if (!phone) return raw;
  const trailing = raw.slice(phone.index + phone.text.length).replace(/[\s)\]}.,;:-]+/g, "");
  if (trailing) return raw;
  return raw.slice(0, phone.index)
    .replace(/\s*[([{-]\s*(?:ph(?:one)?\s*)?$/i, "")
    .replace(/\s+ph(?:one)?\s*$/i, "")
    .trim();
}

function clinicianPhone(standardPhone: string, rawName: string, emr: string, shift: string) {
  const candidates = shift === "Night"
    ? [rawName, emr, standardPhone]
    : [standardPhone, rawName, emr];
  return candidates.map(telephoneNumber).find(Boolean) || "";
}

function telephoneNumber(value: string) {
  const match = telephoneMatch(value);
  return match ? match.text : "";
}

function telephoneMatch(value: string) {
  const raw = String(value || "").trim();
  const pattern = /\(?\d(?:[\d ()-]*\d)?/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(raw))) {
    const candidate = String(match[0] || "").trim();
    const digitCount = candidate.replace(/\D/g, "").length;
    if (digitCount >= 5 && digitCount <= 10) {
      const text = candidate.startsWith("(") && !candidate.includes(")") ? candidate.slice(1) : candidate;
      return { text, index: match.index + (text === candidate ? 0 : 1) };
    }
  }
  return null;
}

function isExcludedRole(role: string) {
  return /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(role);
}

function isHeading(role: string) {
  return /^(?:ed clinician phones?|role|am|pm|nd|night)$/i.test(String(role || "").trim());
}

function isContinuationStreamRole(role: string) {
  return /^(?:orange\s+(?:dr|doctor)\b|silver\s+(?:dr|doctor)\b|(?:ft|fast\s+track)\b)/i.test(String(role || "").trim());
}

function isStructuredRole(role: string) {
  return /^(?:doctor\s+in\s+charge\b|avao\b|orange\b|silver\b|(?:ft|fast\s+track)\b|ssu\b|geriatrician\b|cart\b|miprep\b|resus\b|ed\s+care[\s-]*co\b|gap\b|clinical\s+support\b)/i.test(String(role || "").trim());
}
