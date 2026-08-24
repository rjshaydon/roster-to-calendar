const SOURCE_ID = "mmc-shift-allocations";
const TARGET_SHEET = "SHIFT ALLOCATIONS";
const SECTIONS = [
  ["Adult Emergency", 6, 27],
  ["Paediatric Emergency", 31, 42],
];
const SHIFTS = ["AM", "PM", "Night"];

export async function extractMmcDoctorContactsFromWorkbook(bytes, { providerModifiedAt = "" } = {}) {
  const values = await readTargetCells(bytes);
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

async function readTargetCells(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const entries = readZipDirectory(bytes);
  const workbookXml = await readZipText(bytes, requiredEntry(entries, "xl/workbook.xml"));
  const relationshipId = workbookRelationshipId(workbookXml, TARGET_SHEET);
  const relationshipsXml = await readZipText(bytes, requiredEntry(entries, "xl/_rels/workbook.xml.rels"));
  const worksheetPath = worksheetPathForRelationship(relationshipsXml, relationshipId);
  const sharedStringsEntry = entries.get("xl/sharedStrings.xml");
  const sharedStrings = sharedStringsEntry
    ? parseSharedStrings(await readZipText(bytes, sharedStringsEntry))
    : [];
  const worksheetXml = await readZipText(bytes, requiredEntry(entries, worksheetPath), { stopBeforeRow: 43 });
  return parseWorksheetCells(worksheetXml, sharedStrings);
}

function readZipDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.byteLength - 22;
  const minimum = Math.max(0, bytes.byteLength - 65557);
  while (end >= minimum && view.getUint32(end, true) !== 0x06054b50) end -= 1;
  if (end < minimum) throw new Error("The workbook ZIP directory was not found.");
  const entryCount = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true);
  const decoder = new TextDecoder();
  const entries = new Map();
  for (let index = 0; index < entryCount; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("The workbook ZIP directory is invalid.");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + fileNameLength));
    entries.set(name, { method, compressedSize, localOffset });
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function requiredEntry(entries, path) {
  const entry = entries.get(path);
  if (!entry) throw new Error(`Workbook entry ${path} was not found.`);
  return entry;
}

async function readZipText(bytes, entry, { stopBeforeRow = 0 } = {}) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(entry.localOffset, true) !== 0x04034b50) throw new Error("The workbook ZIP entry is invalid.");
  const nameLength = view.getUint16(entry.localOffset + 26, true);
  const extraLength = view.getUint16(entry.localOffset + 28, true);
  const start = entry.localOffset + 30 + nameLength + extraLength;
  const compressed = bytes.subarray(start, start + entry.compressedSize);
  let stream;
  if (entry.method === 0) stream = new Blob([compressed]).stream();
  else if (entry.method === 8) stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  else throw new Error(`Unsupported workbook ZIP compression method ${entry.method}.`);

  const reader = stream.getReader();
  const decoder = new TextDecoder();
  let result = "";
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    result += decoder.decode(value, { stream: true });
    if (stopBeforeRow) {
      const marker = new RegExp(`<row\\b[^>]*\\br=["']${stopBeforeRow}["']`, "i").exec(result);
      if (marker) {
        result = result.slice(0, marker.index);
        await reader.cancel();
        break;
      }
    }
  }
  return result + decoder.decode();
}

function workbookRelationshipId(xml, sheetName) {
  for (const match of xml.matchAll(/<sheet\b([^>]*)\/?\s*>/gi)) {
    const attributes = match[1];
    const name = xmlAttribute(attributes, "name");
    if (decodeXml(name) === sheetName) return xmlAttribute(attributes, "r:id");
  }
  throw new Error(`Worksheet ${sheetName} was not found.`);
}

function worksheetPathForRelationship(xml, relationshipId) {
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/gi)) {
    const attributes = match[1];
    if (xmlAttribute(attributes, "Id") !== relationshipId) continue;
    const target = xmlAttribute(attributes, "Target").replace(/^\/?xl\//i, "").replace(/^\//, "");
    return `xl/${target}`;
  }
  throw new Error(`Worksheet relationship ${relationshipId} was not found.`);
}

function parseSharedStrings(xml) {
  const values = [];
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)) {
    const parts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)].map((part) => decodeXml(part[1]));
    values.push(parts.join(""));
  }
  return values;
}

function parseWorksheetCells(xml, sharedStrings) {
  const values = Array.from({ length: 42 }, () => Array(9).fill(""));
  for (const match of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/gi)) {
    const attributes = match[1];
    const reference = xmlAttribute(attributes, "r").match(/^([A-Z]+)(\d+)$/i);
    if (!reference) continue;
    const column = columnIndex(reference[1]);
    const row = Number(reference[2]);
    if (row < 1 || row > 42 || column < 0 || column > 8) continue;
    const type = xmlAttribute(attributes, "t");
    const content = match[2] || "";
    const valueMatch = content.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i);
    const inlineParts = [...content.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)].map((part) => decodeXml(part[1]));
    const raw = valueMatch ? decodeXml(valueMatch[1]) : inlineParts.join("");
    values[row - 1][column] = type === "s" ? String(sharedStrings[Number(raw)] ?? "") : raw;
  }
  return values;
}

function columnIndex(label) {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function xmlAttribute(attributes, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return attributes.match(new RegExp(`(?:^|\\s)${escaped}=["']([^"']*)["']`, "i"))?.[1] || "";
}

function decodeXml(value) {
  return String(value || "").replace(/&#x([0-9a-f]+);|&#(\d+);|&(amp|lt|gt|quot|apos);/gi, (_, hex, decimal, named) => {
    if (hex) return String.fromCodePoint(Number.parseInt(hex, 16));
    if (decimal) return String.fromCodePoint(Number(decimal));
    return { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }[named.toLowerCase()];
  });
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
