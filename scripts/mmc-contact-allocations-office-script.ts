/*
 * Excel Office Script: Extract MMC doctor shift contacts
 *
 * Save this in Excel Online against SHIFT ALLOCATIONS.xlsx, then select it in
 * the Power Automate "Run script from SharePoint library" action.
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("SHIFT ALLOCATIONS");
  if (!sheet) throw new Error("Worksheet SHIFT ALLOCATIONS was not found.");

  const values = sheet.getRange("A1:I64").getTexts();
  const sourceDate = values[1][3].trim();

  // Section headings move when clerks insert an extra allocation row.
  // Never read past the bounded doctors area into the nursing/contact tables.
  const paedsHeading = values.findIndex((row) => /^paediatric emergency$/i.test(row[0].trim()));
  const nursingHeading = values.findIndex((row, index) => index > paedsHeading && /^adults$/i.test(row[0].trim()));
  if (paedsHeading < 6 || nursingHeading <= paedsHeading) throw new Error("Doctor section boundaries could not be read.");
  const sections: Array<[string, number, number]> = [
    ["Adult Emergency", 6, paedsHeading],
    ["Paediatric Emergency", paedsHeading + 2, nursingHeading],
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
        // Keep diversion instructions and separate numbers as written.
        const phone = cells[shiftIndex * 3 + 2].trim();
        if (!role || /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(role) || /^(?:role|am|pm|nd|night)$/i.test(role)) continue;
        contacts.push({
          area,
          shift: shifts[shiftIndex],
          role,
          name,
          phone,
          // Phone extensions remain after staff are removed. A name is the
          // signal that this is a live allocation.
          isPopulated: /[a-z]/i.test(name),
        });
      }
    }
  }
  return { sourceDate, contacts };
}
