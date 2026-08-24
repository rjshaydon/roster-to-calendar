/*
 * Excel Office Script: Extract MMC doctor shift contacts
 *
 * Save this in Excel Online against SHIFT ALLOCATIONS.xlsx, then select it in
 * the Power Automate "Run script from SharePoint library" action.
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("SHIFT ALLOCATIONS");
  if (!sheet) throw new Error("Worksheet SHIFT ALLOCATIONS was not found.");

  const values = sheet.getRange("A1:I42").getTexts();
  const sourceDate = sourceDateFromLabel(values[1][3]);
  if (!sourceDate) throw new Error("The date in SHIFT ALLOCATIONS!D2 could not be read.");

  const sections: Array<[string, number, number]> = [
    ["Adult Emergency", 6, 27],
    ["Paediatric Emergency", 31, 42],
  ];
  const shifts = ["AM", "PM", "Night"];
  const contacts: Array<{
    area: string; shift: string; role: string; name: string; phone: string; isPopulated: boolean;
  }> = [];

  for (const [area, firstRow, lastRow] of sections) {
    for (let shiftIndex = 0; shiftIndex < shifts.length; shiftIndex += 1) {
      for (let row = firstRow; row <= lastRow; row += 1) {
        const cells = values[row - 1];
        const role = cells[shiftIndex * 3].trim();
        const name = cells[shiftIndex * 3 + 1].trim();
        const phone = cells[shiftIndex * 3 + 2].trim();
        if (!role || isExcludedRole(role)) continue;
        contacts.push({
          area,
          shift: shifts[shiftIndex],
          role,
          name,
          phone,
          // Phone extensions remain after staff are removed. A name is the
          // signal that this is a live allocation.
          isPopulated: Boolean(name),
        });
      }
    }
  }
  return { sourceId: "mmc-shift-allocations", sourceDate, contacts };
}

function isExcludedRole(role: string) {
  return /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(role);
}

function sourceDateFromLabel(value: string) {
  const match = String(value || "").trim().match(/(\d{1,2})(?:st|nd|rd|th)?\s+([a-z]+)\s+(\d{4})/i);
  if (!match) return "";
  const months: { [key: string]: string } = {
    january: "01", february: "02", march: "03", april: "04", may: "05", june: "06",
    july: "07", august: "08", september: "09", october: "10", november: "11", december: "12",
  };
  const month = months[match[2].toLowerCase()];
  if (!month) return "";
  return `${match[3]}-${month}-${match[1].padStart(2, "0")}`;
}
