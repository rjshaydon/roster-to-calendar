/*
 * Excel Office Script: Extract VHH Zebra Allocations contacts
 *
 * Run from Power Automate's "Run script from SharePoint library" action.
 * It reads only the requested cells and never changes the workbook.
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("Zebra Allocations");
  if (!sheet) throw new Error("Worksheet Zebra Allocations was not found.");
  const values = sheet.getRange("C1:E80").getTexts();
  const cell = (row: number, column: number) => String(values[row - 1]?.[column - 3] || "").trim();

  if (cell(4, 3).toUpperCase() !== "CIC") throw new Error("Zebra Allocations!C4 no longer contains CIC.");
  const doctorsHeader = findRow(values, "DOCTORS");
  const clericalHeader = findRow(values, "CLERICAL");
  if (!doctorsHeader || !clericalHeader || clericalHeader <= doctorsHeader) {
    throw new Error("The DOCTORS and CLERICAL boundaries were not found in Zebra Allocations.");
  }

  const doctors: Array<{ role: string; phone: string; name: string }> = [];
  for (let row = doctorsHeader + 1; row < clericalHeader; row += 1) {
    const role = cell(row, 3);
    const phone = cell(row, 4);
    const name = cell(row, 5);
    if (role || phone || name) doctors.push({ role, phone, name });
  }
  return {
    schemaVersion: 1,
    sourceId: "vhh-shift-phone-allocations",
    cic: { phone: cell(4, 4), name: cell(4, 5) },
    doctors,
  };
}

function findRow(values: string[][], label: string) {
  const expected = label.toUpperCase();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index]?.[0] || "").trim().toUpperCase() === expected) return index + 1;
  }
  return 0;
}
