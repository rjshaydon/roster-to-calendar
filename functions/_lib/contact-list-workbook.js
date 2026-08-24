import * as XLSX from "xlsx";

const SOURCE_ID = "mmc-shift-allocations";
const TARGET_SHEET = "SHIFT ALLOCATIONS";
const SECTIONS = [
  ["Adult Emergency", 6, 27],
  ["Paediatric Emergency", 31, 42],
];
const SHIFTS = ["AM", "PM", "Night"];

export async function extractMmcDoctorContactsFromWorkbook(bytes, { providerModifiedAt = "" } = {}) {
  // sheetRows prevents Excel's accidentally formatted far-right/far-down cells
  // from expanding into the enormous used range present in this workbook.
  const workbook = XLSX.read(bytes, { type: "array", sheetRows: 42, cellText: true });
  const sheet = workbook.Sheets[TARGET_SHEET];
  if (!sheet) throw new Error(`Worksheet ${TARGET_SHEET} was not found.`);
  const values = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: "A1:I42",
    raw: false,
    defval: "",
    blankrows: true,
  });
  const sourceDate = sourceDateFromLabel(values?.[1]?.[3]);
  if (!sourceDate) throw new Error("The date in SHIFT ALLOCATIONS!D2 could not be read.");

  const contacts = [];
  for (const [area, firstRow, lastRow] of SECTIONS) {
    for (let shiftIndex = 0; shiftIndex < SHIFTS.length; shiftIndex += 1) {
      for (let row = firstRow; row <= lastRow; row += 1) {
        const cells = values[row - 1] || [];
        const role = text(cells[shiftIndex * 3]);
        const name = text(cells[shiftIndex * 3 + 1]);
        const phone = text(cells[shiftIndex * 3 + 2]);
        if (!role || isExcludedRole(role)) continue;
        contacts.push({
          area,
          shift: SHIFTS[shiftIndex],
          role,
          name,
          phone,
          isPopulated: Boolean(name),
        });
      }
    }
  }
  return {
    sourceId: SOURCE_ID,
    sourceDate,
    providerModifiedAt: String(providerModifiedAt || "").trim(),
    contacts,
  };
}

function text(value) {
  return String(value ?? "").trim();
}

function sourceDateFromLabel(value) {
  const match = text(value).match(/(\d{1,2})(?:st|nd|rd|th)?\s+([a-z]+)\s+(\d{4})/i);
  if (!match) return "";
  const months = {
    january: "01", february: "02", march: "03", april: "04", may: "05", june: "06",
    july: "07", august: "08", september: "09", october: "10", november: "11", december: "12",
  };
  const month = months[match[2].toLowerCase()];
  return month ? `${match[3]}-${month}-${match[1].padStart(2, "0")}` : "";
}

function isExcludedRole(role) {
  return /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(role);
}
