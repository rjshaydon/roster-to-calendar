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
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const visible = (visibility.get(sheetName) || 0) === 0;
    let blockIndex = 0;
    for (let header = 0; header <= range.e.r; header += 1) {
      if (!isShiftLabel(cellText(sheet, header, 0))) continue;
      const dates = [];
      for (let column = 1; column < MAX_COLUMNS; column += 1) {
        const date = cellIsoDate(sheet, header, column);
        if (date) dates.push({ sourceColumn: XLSX.utils.encode_col(column), displayedDate: cellText(sheet, header, column), date });
      }
      if (!dates.length) continue;

      let blockEnd = range.e.r + 1;
      for (let row = header + 1; row <= range.e.r; row += 1) {
        if (isShiftLabel(cellText(sheet, row, 0))) {
          blockEnd = row;
          break;
        }
      }
      const teachingStart = findRow(sheet, header + 1, blockEnd, /^JMS\s+TEACHING\s+TIMETABLE$/i);
      const teachingEnd = findRow(sheet, header + 1, blockEnd, /^ULTRASOUND\s+TEACHING\s+SESSIONS\s+ARE\s+AVAILABLE\s+FOR\s+BOOKING\s+VIA\b/i);
      if ((teachingStart >= 0) !== (teachingEnd >= 0) || (teachingStart >= 0 && teachingEnd < teachingStart)) {
        throw new Error(`VHH teaching timetable boundaries are incomplete on ${sheetName}, roster row ${header + 1}.`);
      }

      const rows = [];
      let inheritedShiftLabel = "";
      for (let row = header + 1; row < blockEnd; row += 1) {
        if (teachingStart >= 0 && row >= teachingStart && row <= teachingEnd) {
          inheritedShiftLabel = "";
          continue;
        }
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
        visible,
        headerRow: header + 1,
        teachingTimetableRow: teachingStart >= 0 ? teachingStart + 1 : 0,
        teachingTimetableEndRow: teachingEnd >= 0 ? teachingEnd + 1 : 0,
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
  const sourceCell = cell(sheet, row, column);
  const value = sourceCell?.v;
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const displayed = String(sourceCell?.w ?? value ?? "").trim();
  const match = displayed.match(/(?:^|\s)(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:$|\s)/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    const candidate = new Date(Date.UTC(year, month - 1, day));
    if (candidate.getUTCFullYear() === year && candidate.getUTCMonth() === month - 1 && candidate.getUTCDate() === day) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }
  return "";
}

function isShiftLabel(value) {
  return /^SHIFT\s+LABEL$/i.test(String(value || "").trim());
}

function findRow(sheet, startRow, endRow, expression) {
  for (let row = startRow; row < endRow; row += 1) {
    for (let column = 0; column < MAX_COLUMNS; column += 1) {
      if (expression.test(cellText(sheet, row, column))) return row;
    }
  }
  return -1;
}
