/* Excel Office Script: VHH Extract Shift Phone Allocations.
 * Existing read-only C1:E80 extraction with the sheet date added to its JSON.
 * The app normalizes C2's formatted date and matches current timed VHH shifts.
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
    const role = cell(row, 3), phone = cell(row, 4), name = cell(row, 5);
    if (role || phone || name) doctors.push({ role, phone, name });
  }
  return { schemaVersion: 1, sourceId: "vhh-shift-phone-allocations", sourceDate: cell(2, 3),
    cic: { phone: cell(4, 4), name: cell(4, 5) }, doctors };
}
function findRow(values: string[][], label: string) {
  const expected = label.toUpperCase();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index]?.[0] || "").trim().toUpperCase() === expected) return index + 1;
  }
  return 0;
}
