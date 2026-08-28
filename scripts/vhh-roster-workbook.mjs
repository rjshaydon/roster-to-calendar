import * as XLSX from "xlsx";
import { VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";

const MAX_COLUMNS = 15;

export async function extractVhhRosterWorkbook(file, metadata = {}) {
  if (!(file instanceof File)) throw new Error("A VHH roster workbook is required.");
  const bytes = new Uint8Array(await file.arrayBuffer());
  let workbook;
  try {
    workbook = XLSX.read(bytes, { type: "array", cellDates: true, cellNF: false, cellHTML: false, cellStyles: false });
  } catch {
    throw new Error(`${file.name || "VHH roster"} is not a supported Excel workbook.`);
  }

  const visibility = new Map((workbook.Workbook?.Sheets || []).map((sheet) => [sheet.name, Number(sheet.Hidden || 0)]));
  const blocks = [];
  for (const sheetName of workbook.SheetNames || []) {
    if ((visibility.get(sheetName) || 0) !== 0) continue;
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let blockIndex = 0;
    for (let header = 0; header <= range.e.r; header += 1) {
      if (cellText(sheet, header, 0) !== "SHIFT LABEL") continue;
      const dates = [];
      for (let column = 1; column < MAX_COLUMNS; column += 1) {
        const date = cellIsoDate(sheet, header, column);
        if (date) dates.push({ sourceColumn: XLSX.utils.encode_col(column), displayedDate: cellText(sheet, header, column), date });
      }
      if (!dates.length) continue;

      let timetable = range.e.r + 1;
      for (let row = header + 1; row <= range.e.r; row += 1) {
        const label = cellText(sheet, row, 0);
        if (/^JMS\s+TEACHING\s+TIMETABLE$/i.test(label) || label === "SHIFT LABEL") {
          timetable = row;
          break;
        }
      }

      const rows = [];
      let inheritedShiftLabel = "";
      for (let row = header + 1; row < timetable; row += 1) {
        const sourceShiftLabel = cellText(sheet, row, 0);
        if (sourceShiftLabel) inheritedShiftLabel = sourceShiftLabel;
        if (!inheritedShiftLabel) continue;
        const assignments = dates.flatMap((date) => {
          const column = XLSX.utils.decode_col(date.sourceColumn);
          const namesText = cellText(sheet, row, column);
          return namesText ? [{
            date: date.date,
            displayedDate: date.displayedDate,
            namesText,
            sourceCell: XLSX.utils.encode_cell({ r: row, c: column }),
          }] : [];
        });
        if (assignments.length) rows.push({ sourceRow: row + 1, sourceShiftLabel, shiftLabel: inheritedShiftLabel, assignments });
      }
      if (rows.length) blocks.push({
        sheetName,
        blockIndex: ++blockIndex,
        headerRow: header + 1,
        teachingTimetableRow: timetable + 1,
        dates,
        rows,
      });
    }
  }

  if (!blocks.length) throw new Error("No populated VHH Shift Label blocks were found in the workbook.");
  return {
    schemaVersion: 1,
    sourceId: VHH_ROSTER_SOURCE_ID,
    fileName: file.name || "Active Medical Roster.xlsx",
    providerModifiedAt: String(metadata.providerModifiedAt || "").trim(),
    providerVersion: String(metadata.providerVersion || "").trim(),
    blocks,
  };
}

function cell(sheet, row, column) {
  return sheet[XLSX.utils.encode_cell({ r: row, c: column })];
}

function cellText(sheet, row, column) {
  const value = cell(sheet, row, column);
  return String(value?.w ?? value?.v ?? "").trim();
}

function cellIsoDate(sheet, row, column) {
  const value = cell(sheet, row, column)?.v;
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  return "";
}
