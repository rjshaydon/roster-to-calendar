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
    let precedingRole = "";
    for (const cells of values) {
      const suppliedRole = String(cells[firstColumn] || "").trim();
      if (isHeading(suppliedRole) || isExcludedRole(suppliedRole)) {
        precedingRole = "";
        continue;
      }
      if (suppliedRole) precedingRole = suppliedRole;
      const rawName = String(cells[firstColumn + 1] || "");
      const name = clinicianName(rawName);
      const phone = clinicianPhone(String(cells[firstColumn + 3] || ""), rawName, String(cells[firstColumn + 2] || ""));
      const role = suppliedRole || (name && phone ? precedingRole : "");
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
  return String(value || "")
    .trim()
    .replace(/\s*[-–—]?\s*(?:\+?61\s*\d(?:[\s-]*\d){7,}|0\d(?:[\s-]*\d){7,}|\d{5})\s*$/i, "")
    .trim();
}

function clinicianPhone(standardPhone: string, rawName: string, emr: string) {
  return fiveDigitExtension(standardPhone) || fiveDigitExtension(rawName) || fiveDigitExtension(emr);
}

function fiveDigitExtension(value: string) {
  const match = String(value || "").trim().match(/(?:^|\D)(\d{5})(?!\d)/);
  return match ? match[1] : "";
}

function isExcludedRole(role: string) {
  return /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(role);
}

function isHeading(role: string) {
  return /^(?:ed clinician phones?|role|am|pm|nd|night)$/i.test(String(role || "").trim());
}
