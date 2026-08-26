/*
 * Excel Office Script: Extract VHH Active Medical Roster
 *
 * Reads visible worksheets only. Each Shift Label block is retained through
 * the row before JMS Teaching Timetable. The script never writes to Excel.
 */
function main(workbook: ExcelScript.Workbook) {
  const blocks: Array<{
    sheetName: string; blockIndex: number; headerRow: number; teachingTimetableRow: number;
    dates: Array<{ sourceColumn: string; displayedDate: string; date: string }>;
    rows: Array<{ sourceRow: number; sourceShiftLabel: string; shiftLabel: string; assignments: Array<{ date: string; displayedDate: string; namesText: string; sourceCell: string }> }>;
  }> = [];

  for (const sheet of workbook.getWorksheets()) {
    if (sheet.getVisibility() !== ExcelScript.SheetVisibility.visible) continue;
    const used = sheet.getUsedRange(true);
    if (!used) continue;
    const rowCount = used.getRowIndex() + used.getRowCount();
    const values = sheet.getRangeByIndexes(0, 0, Math.max(1, rowCount), 15).getValues();
    const texts = sheet.getRangeByIndexes(0, 0, Math.max(1, rowCount), 15).getTexts();
    let blockIndex = 0;
    for (let header = 0; header < texts.length; header += 1) {
      if (text(texts[header]?.[0]) !== "SHIFT LABEL") continue;
      const dates: Array<{ sourceColumn: string; displayedDate: string; date: string }> = [];
      for (let column = 1; column <= 14; column += 1) {
        const date = excelDateToIso(values[header]?.[column]);
        if (date) dates.push({ sourceColumn: columnName(column), displayedDate: text(texts[header]?.[column]), date });
      }
      if (!dates.length) continue;
      let timetable = texts.length;
      for (let row = header + 1; row < texts.length; row += 1) {
        if (/^JMS\s+TEACHING\s+TIMETABLE$/i.test(text(texts[row]?.[0]))) { timetable = row; break; }
        if (text(texts[row]?.[0]) === "SHIFT LABEL") break;
      }
      const rows: Array<{ sourceRow: number; sourceShiftLabel: string; shiftLabel: string; assignments: Array<{ date: string; displayedDate: string; namesText: string; sourceCell: string }> }> = [];
      let inheritedShiftLabel = "";
      for (let row = header + 1; row < timetable; row += 1) {
        const sourceShiftLabel = text(texts[row]?.[0]);
        if (sourceShiftLabel) inheritedShiftLabel = sourceShiftLabel;
        if (!inheritedShiftLabel) continue;
        const assignments = dates.map((date) => {
          const column = columnIndex(date.sourceColumn);
          const namesText = text(texts[row]?.[column]);
          return namesText ? { date: date.date, displayedDate: date.displayedDate, namesText, sourceCell: `${date.sourceColumn}${row + 1}` } : null;
        }).filter((assignment): assignment is { date: string; displayedDate: string; namesText: string; sourceCell: string } => Boolean(assignment));
        if (assignments.length) rows.push({ sourceRow: row + 1, sourceShiftLabel, shiftLabel: inheritedShiftLabel, assignments });
      }
      if (rows.length) blocks.push({ sheetName: sheet.getName(), blockIndex: ++blockIndex, headerRow: header + 1, teachingTimetableRow: timetable + 1, dates, rows });
    }
  }
  if (!blocks.length) throw new Error("No populated VHH Shift Label blocks were found.");
  return { schemaVersion: 1, sourceId: "vhh-active-medical-roster", blocks };
}

function text(value: unknown) {
  return String(value ?? "").trim();
}

function excelDateToIso(value: unknown) {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + Math.floor(value) * 86400000).toISOString().slice(0, 10);
}

function columnName(column: number) {
  let value = column + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function columnIndex(label: string) {
  let value = 0;
  for (const character of label) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}
