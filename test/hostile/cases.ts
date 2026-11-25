import type { ReadErrorCode, ReadLimits } from "../../src/read.js";
import type { Outcome, Plan, Refusal, Row, SheetOutcome } from "../helpers/outcome.ts";
import { type Layout, type Part, bytesOf, packageOf } from "../helpers/package.ts";
import {
  type Book,
  MAIN,
  RELATIONSHIPS,
  SHEET,
  STRICT_MAIN,
  STRICT_RELATIONSHIPS,
  STRINGS,
  bookOf,
  item,
  relationship,
  relationships,
  sheetEntry,
  styleSheet,
  table,
  workbookPart,
  worksheet,
} from "./book.ts";

/**
 * A file from a stranger and what reading it must end in: built at test time, read in a process
 * of its own as `plan` says, with `limits`, inside `milliseconds` and under a heap of
 * `heapMebibytes`. The outcome is a refusal with its code, or every sheet with its values.
 */
export interface HostileCase {
  readonly name: string;
  /** What the file is, and what it attacks. */
  readonly why: string;
  readonly file: () => Uint8Array | Blob;
  readonly plan?: Plan;
  readonly limits?: ReadLimits;
  /** The outcome, or a function that makes it where it is too large to make for every case. */
  readonly outcome: Outcome | (() => Outcome);
  readonly milliseconds?: number;
  readonly heapMebibytes?: number;
}

const MEBIBYTE = 1_048_576;

/**
 * The time a case may take unless it says otherwise, from its file built to its outcome: more than
 * fifteen times the slowest case, which read in 294 to 304 ms on the machine that set it. It fails a
 * reader that tests a format code with a pattern that backtracks.
 */
export const DEFAULT_MILLISECONDS = 5_000;

/**
 * The old space a case's process may hold unless the case says otherwise: twice the 64 MiB that the
 * largest case under it needs on Node 24 and 26. It fails a reader that keeps a copy of a shared
 * target for each of 100 000 sheets, joins text of four million pieces one piece at a time, or keeps
 * every entry of a directory it has not charged.
 */
export const DEFAULT_HEAP_MEBIBYTES = 128;

/**
 * The old space of the three sheets of one row of 119 MiB, read at once, which needed 64 to 96 MiB
 * in two measurements. A reader without `rowCharacters` exhausts it.
 */
const LARGE_HEAP_MEBIBYTES = 512;

const opening = (code: ReadErrorCode, more: Omit<Refusal, "code"> = {}): Outcome => ({
  refused: { code, ...more },
});

const read = (...sheets: SheetOutcome[]): Outcome => ({ sheets });

const rows = (name: string, ...listed: Row[]): SheetOutcome => ({ name, rows: listed });

/** A sheet whose read was refused, the refusal naming the sheet. */
const refused = (
  name: string,
  code: ReadErrorCode,
  more: Omit<Refusal, "code" | "sheet"> = {},
): SheetOutcome => ({ name, refused: { code, sheet: name, ...more } });

/** A sheet whose read was refused for the shared strings or the styles, which name no sheet. */
const tableRefused = (
  name: string,
  code: ReadErrorCode,
  more: Omit<Refusal, "code" | "sheet"> = {},
): SheetOutcome => ({ name, refused: { code, ...more } });

/** The one sheet of a default book refused with `code`. */
const sheetRefused = (code: ReadErrorCode, more: Omit<Refusal, "code" | "sheet"> = {}) =>
  read(refused("Sheet1", code, more));

/** What the default book reads: a shared string and a number. */
const DEFAULT_ROWS = read(rows("Sheet1", [1, "a", 1]));

const file =
  (book: Book = {}, layout: Layout = {}) =>
  () =>
    packageOf(bookOf(book), layout);

/** Text, or a function that makes it where it is too large to make for every case. */
type Text = string | (() => string);

const made = (text: Text) => (typeof text === "string" ? text : text());

/** A default book whose one sheet holds these rows. */
const sheet =
  (rowsXml: Text, book: Book = {}) =>
  () =>
    packageOf(bookOf({ ...book, sheets: [["Sheet1", worksheet(made(rowsXml))]] }));

/** A part's content as enough repetitions of a chunk of about a mebibyte to reach `bytes`, between two ends. */
function repeated(head: string, unit: string, bytes: number, tail: string): Uint8Array[] {
  const chunk = bytesOf(unit.repeat(Math.ceil(MEBIBYTE / unit.length)));
  const count = Math.ceil(bytes / chunk.length);
  return [bytesOf(head), ...Array.from({ length: count }, () => chunk), bytesOf(tail)];
}

const sheetData =
  (data: () => Part["data"], more: Partial<Part> = {}) =>
  () =>
    packageOf(bookOf({ parts: [{ path: SHEET, data: data(), ...more }] }));

const stringsData = (data: () => Part["data"]) => () =>
  packageOf(bookOf({ parts: [{ path: STRINGS, data: data() }] }));

const SHEET_HEAD = `<worksheet xmlns="${MAIN}"><sheetData>`;
const SHEET_TAIL = "</sheetData></worksheet>";

/**
 * A package of `entries` empty stored parts with seven-character names, each with a local header
 * of its own, and the zip64 end record that more than 65 535 entries need: nothing but a directory.
 */
function directoryFlood(entries: number): Uint8Array {
  const NAME = 7;
  const localBytes = 30 + NAME;
  const centralBytes = 46 + NAME;
  const directoryStart = entries * localBytes;
  const directorySize = entries * centralBytes;
  const bytes = new Uint8Array(directoryStart + directorySize + 56 + 20 + 22);
  const view = new DataView(bytes.buffer);
  for (let entry = 0; entry < entries; entry++) {
    const name = entry.toString(36).padStart(NAME, "0");
    const local = entry * localBytes;
    const central = directoryStart + entry * centralBytes;
    view.setUint32(local, 0x04034b50, true);
    view.setUint16(local + 4, 20, true);
    view.setUint16(local + 26, NAME, true);
    view.setUint32(central, 0x02014b50, true);
    view.setUint16(central + 4, 20, true);
    view.setUint16(central + 6, 20, true);
    view.setUint16(central + 28, NAME, true);
    view.setUint32(central + 42, local, true);
    for (let at = 0; at < NAME; at++) {
      bytes[local + 30 + at] = name.charCodeAt(at);
      bytes[central + 46 + at] = name.charCodeAt(at);
    }
  }
  const record = directoryStart + directorySize;
  view.setUint32(record, 0x06064b50, true);
  view.setBigUint64(record + 4, 44n, true);
  view.setUint16(record + 12, 45, true);
  view.setUint16(record + 14, 45, true);
  view.setBigUint64(record + 24, BigInt(entries), true);
  view.setBigUint64(record + 32, BigInt(entries), true);
  view.setBigUint64(record + 40, BigInt(directorySize), true);
  view.setBigUint64(record + 48, BigInt(directoryStart), true);
  const locator = record + 56;
  view.setUint32(locator, 0x07064b50, true);
  view.setBigUint64(locator + 8, BigInt(record), true);
  view.setUint32(locator + 16, 1, true);
  const end = locator + 20;
  view.setUint32(end, 0x06054b50, true);
  view.setUint16(end + 8, 0xffff, true);
  view.setUint16(end + 10, 0xffff, true);
  view.setUint32(end + 12, 0xffffffff, true);
  view.setUint32(end + 16, 0xffffffff, true);
  return bytes;
}

/** The default book with the bytes of its directory entry for `path` changed by `change`. */
function withDirectoryEntry(
  path: string,
  change: (view: DataView, entry: number) => void,
): () => Uint8Array {
  return () => {
    const bytes = packageOf(bookOf());
    const view = new DataView(bytes.buffer);
    const name = bytesOf(path);
    for (let at = bytes.length - 22 - 46; at >= 0; at--) {
      if (
        view.getUint32(at, true) === 0x02014b50 &&
        view.getUint16(at + 28, true) === name.length &&
        name.every((byte, index) => bytes[at + 46 + index] === byte)
      ) {
        change(view, at);
        return bytes;
      }
    }
    throw new Error(`no directory entry for ${path}`);
  };
}

const sheetWithLocalHeader = (change: (bytes: Uint8Array, local: number) => void) => () => {
  const bytes = packageOf(bookOf());
  const view = new DataView(bytes.buffer);
  const name = bytesOf(SHEET);
  for (let at = 0; at < bytes.length; at++) {
    if (
      view.getUint32(at, true) === 0x04034b50 &&
      name.every((byte, index) => bytes[at + 30 + index] === byte)
    ) {
      change(bytes, at);
      return bytes;
    }
  }
  throw new Error("no local header for the sheet");
};

function noise(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  let state = 0x2545f491;
  for (let at = 0; at < length; at++) {
    state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
    bytes[at] = state >>> 24;
  }
  return bytes;
}

const columnName = (column: number): string => {
  let name = "";
  for (let left = column; left > 0; left = Math.floor((left - 1) / 26)) {
    name = String.fromCharCode(65 + ((left - 1) % 26)) + name;
  }
  return name;
};

const ROW_CHARACTERS = 4_194_304;

/**
 * Three sheets that name one part, one row of 119 MiB: 16 384 cells of 7 600 characters, or one
 * inline string of 31 runs of 4 000 000 characters. One character of each value is outside
 * Latin-1, so an engine holds it at two bytes a character. Each read alone stays under every
 * default limit but `rowCharacters`; three at once held more than a 512 MiB heap before it.
 */
function threeSheetsOfOneRow(row: "cells" | "runs"): () => Uint8Array {
  return () => {
    const value = (length: number) => bytesOf(`Ā${"a".repeat(length - 1)}`);
    const chunks: Uint8Array[] = [bytesOf(`${SHEET_HEAD}<row r="1">`)];
    if (row === "cells") {
      const text = value(7_600);
      const close = bytesOf("</v></c>");
      for (let column = 1; column <= 16_384; column++) {
        chunks.push(bytesOf(`<c r="${columnName(column)}1" t="str"><v>`), text, close);
      }
    } else {
      const text = value(4_000_000);
      const [open, close] = [bytesOf("<r><t>"), bytesOf("</t></r>")];
      chunks.push(bytesOf('<c r="A1" t="inlineStr"><is>'));
      for (let run = 0; run < 31; run++) chunks.push(open, text, close);
      chunks.push(bytesOf("</is></c>"));
    }
    chunks.push(bytesOf(`</row>${SHEET_TAIL}`));
    return packageOf(
      bookOf({
        strings: null,
        workbook: workbookPart([1, 2, 3].map((at) => sheetEntry(`Sheet${at}`, "rId1")).join("")),
        relationships: relationships(relationship("rId1", "worksheet", "worksheets/sheet1.xml")),
        parts: [{ path: SHEET, data: chunks }],
      }).filter(({ path }) => !path.startsWith("xl/worksheets/") || path === SHEET),
    );
  };
}

const packageCases: HostileCase[] = [
  {
    name: "no bytes",
    why: "an empty file",
    file: () => new Uint8Array(0),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "three bytes",
    why: "the start of a local header and nothing else",
    file: () => bytesOf("PK\x03"),
    outcome: opening("NOT_XLSX"),
  },
  { name: "noise", why: "4 KiB of noise", file: () => noise(4096), outcome: opening("NOT_XLSX") },
  {
    name: "a compound file",
    why: "an old .xls or an encrypted workbook",
    file: () => {
      const bytes = new Uint8Array(4096);
      bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
      return bytes;
    },
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a CSV file",
    why: "text with the wrong extension",
    file: () => bytesOf("Article number,Quantity\n00123,4\n"),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a PNG file",
    why: "a picture with the wrong extension",
    file: () => {
      const bytes = new Uint8Array(80);
      bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      return bytes;
    },
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a zip that is no workbook",
    why: "a package without relationships of its own",
    file: () => packageOf([{ path: "hello.txt", data: "hello" }]),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a word-processing document",
    why: "a package whose office document is not a workbook",
    file: () =>
      packageOf([
        {
          path: "_rels/.rels",
          data: relationships(relationship("rId1", "officeDocument", "word/document.xml")),
        },
        {
          path: "word/document.xml",
          data: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        },
      ]),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "the first half of a workbook",
    why: "a file cut in the middle",
    file: () => {
      const bytes = packageOf(bookOf());
      return bytes.subarray(0, bytes.length >> 1);
    },
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "the last ten bytes cut",
    why: "an end record cut short",
    file: () => {
      const bytes = packageOf(bookOf());
      return bytes.subarray(0, bytes.length - 10);
    },
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "an end record alone",
    why: "the last 22 bytes of a workbook",
    file: () => packageOf(bookOf()).slice(-22),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "bytes after the end record",
    why: "refused although other tools read it (issue 82)",
    file: file({}, { after: bytesOf("tail") }),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "a comment that holds a second end record",
    why: "an end that reads two ways",
    file: file({}, { comment: secondEndRecord() }),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "bytes in front that the offsets do not count",
    why: "a directory elsewhere than the end record says (issue 82)",
    file: file({}, { before: new Uint8Array(4096) }),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "bytes in front that the offsets count",
    why: "a package written as a part of a larger file, which the reader accepts",
    file: file({}, { before: new Uint8Array(4096), offsetsCountBefore: true }),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a directory offset past the end",
    why: "an end record that points outside the file",
    file: file({}, { directoryOffset: 0x7fffffff }),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "an end record that claims 65 535 entries",
    why: "the walk of the directory stops at its size, not at a count the file states",
    file: file({}, { entries: 0xffff }),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "two workbook parts",
    why: "two parts with one name",
    file: () =>
      packageOf([
        ...bookOf(),
        { path: "xl/workbook.xml", data: workbookPart(sheetEntry("Other", "rId1")) },
      ]),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "two sheet parts",
    why: "two parts with the sheet's name, the second of another value",
    file: () =>
      packageOf([
        ...bookOf(),
        { path: SHEET, data: worksheet('<row r="1"><c r="A1"><v>666</v></c></row>') },
      ]),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "two part names that differ by the case of A to Z",
    why: "part names fold A to Z (issue 82)",
    file: () => packageOf([...bookOf(), { path: "XL/Worksheets/Sheet1.xml", data: worksheet("") }]),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "two part names that differ by the case of a letter outside A to Z",
    why: "part names fold A to Z only, so these are two parts (issue 82)",
    file: () =>
      packageOf(
        bookOf({
          strings: null,
          workbook: workbookPart(sheetEntry("Lower", "rId1") + sheetEntry("Upper", "rId2")),
          relationships: relationships(
            relationship("rId1", "worksheet", "worksheets/å.xml"),
            relationship("rId2", "worksheet", "worksheets/Å.xml"),
          ),
          sheets: [],
          parts: [
            {
              path: "xl/worksheets/å.xml",
              data: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>'),
            },
            {
              path: "xl/worksheets/Å.xml",
              data: worksheet('<row r="1"><c r="A1"><v>2</v></c></row>'),
            },
          ],
        }),
      ),
    outcome: read(rows("Lower", [1, 1]), rows("Upper", [1, 2])),
  },
  {
    name: "a local header that names another part",
    why: "a local header that disagrees with the directory",
    file: file({
      parts: [{ path: SHEET, data: bookOf()[4]!.data, localPath: "xl/worksheets/other.xml" }],
    }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "two directory entries over one part's data",
    why: "entries whose data overlap are read, each charged on its own (issue 82)",
    file: () =>
      packageOf([...bookOf(), { path: "xl/worksheets/sheet2.xml", data: "", sameDataAs: SHEET }]),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "60 000 entries more",
    why: "a directory of many parts nobody reads",
    file: () =>
      packageOf([
        ...bookOf(),
        ...Array.from({ length: 60_000 }, (_, at) => ({
          path: `junk/${at}.bin`,
          data: "",
          method: 0,
        })),
      ]),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a directory of 110 000 entries",
    why: "more parts than `parts` allows",
    file: () => directoryFlood(110_000),
    outcome: opening("LIMIT_EXCEEDED", { limit: "parts" }),
  },
  {
    name: "a directory of four million entries",
    why: "a directory larger than its share of `inflatedBytes`, refused before it is read (interrogation finding 4)",
    file: () => directoryFlood(4_000_000),
    outcome: opening("LIMIT_EXCEEDED", { limit: "inflatedBytes" }),
  },
  {
    name: "damaged compressed bytes",
    why: "a sheet whose deflate stream is damaged",
    file: sheetWithLocalHeader((bytes, local) => {
      const start = local + 30 + SHEET.length;
      for (let at = start + 8; at < start + 32; at++) bytes[at]! ^= 0xa5;
    }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "a compressed size of 2 GiB",
    why: "a part that lies outside the package",
    file: withDirectoryEntry(SHEET, (view, entry) => view.setUint32(entry + 20, 0x7ffffff0, true)),
    outcome: opening("INVALID_PACKAGE"),
  },
  {
    name: "a declared size of 1 000 bytes over 300 MiB",
    why: "a part that inflates past its declared size",
    file: sheetData(() => repeated(SHEET_HEAD, " ", 300 * MEBIBYTE, SHEET_TAIL), { size: 1_000 }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "a wrong CRC-32",
    why: "a part whose checksum is not its content's",
    file: sheetData(() => worksheet('<row r="1"><c r="A1"><v>1</v></c></row>'), { crc32: 1 }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "a compression method of 12",
    why: "a method other than stored or deflate",
    file: sheetData(() => worksheet('<row r="1"><c r="A1"><v>1</v></c></row>'), { method: 12 }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "an encrypted part",
    why: "a part whose flags say it is encrypted",
    file: sheetData(() => worksheet('<row r="1"><c r="A1"><v>1</v></c></row>'), { flags: 1 }),
    outcome: sheetRefused("INVALID_PACKAGE"),
  },
  {
    name: "stored parts",
    why: "a package without compression",
    file: () => packageOf(bookOf().map((part) => ({ ...part, method: 0 }))),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "data descriptors",
    why: "sizes and checksums after the data, as a streaming writer writes them",
    file: () => packageOf(bookOf().map((part) => ({ ...part, descriptor: true }))),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "zip64 records",
    why: "zip64 is accepted (ADR 0032)",
    file: () =>
      packageOf(
        bookOf().map((part) => ({ ...part, zip64: true })),
        { zip64: true },
      ),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a comment and extra fields",
    why: "fields a reader passes over",
    file: () =>
      packageOf(
        bookOf().map((part) => ({ ...part, extra: new Uint8Array([0x55, 0x54, 1, 0, 1]) })),
        { comment: bytesOf("a comment at the end of the package") },
      ),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a sheet of a gibibyte of white space",
    why: "a part whose declared size is more than `inflatedBytes`, refused before a byte is inflated",
    file: sheetData(() => repeated(SHEET_HEAD, " ", 1024 * MEBIBYTE, SHEET_TAIL)),
    outcome: sheetRefused("LIMIT_EXCEEDED", { limit: "inflatedBytes" }),
  },
  {
    name: "a sheet of a gibibyte of white space, from a Blob",
    why: "a Blob is read in slices and never whole",
    file: () =>
      new Blob([sheetData(() => repeated(SHEET_HEAD, " ", 1024 * MEBIBYTE, SHEET_TAIL))()]),
    outcome: sheetRefused("LIMIT_EXCEEDED", { limit: "inflatedBytes" }),
  },
  {
    name: "a workbook from a Blob",
    why: "the same values from a Blob as from bytes",
    file: () => new Blob([packageOf(bookOf())]),
    outcome: DEFAULT_ROWS,
  },
];

function secondEndRecord(): Uint8Array {
  const comment = new Uint8Array(27);
  comment.set([0x50, 0x4b, 0x05, 0x06]);
  comment.set(bytesOf(" tail"), 22);
  return comment;
}

const TARGETS: readonly (readonly [string, string])[] = [
  ["../../../../etc/passwd", ""],
  ["file:///etc/passwd", ' TargetMode="External"'],
  ["http://example.invalid/sheet.xml", ' TargetMode="External"'],
  ["worksheets/sheet1.xml", ' TargetMode="External"'],
  ["", ""],
  ["../../sheet1.xml", ""],
  ["worksheets//sheet1.xml", ""],
  ["worksheets/sheet1.xml/.", ""],
  ["worksheets/sheet1.xml/..", ""],
  ["worksheets/sheet1%2Exml", ""],
  ["worksheets/sheet1.xml?", ""],
  ["worksheets/sheet1.xml#a", ""],
  ["worksheets\\sheet1.xml", ""],
  ["c:/sheet1.xml", ""],
];

const DIALOG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const MACRO = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

const workbookCases: HostileCase[] = [
  {
    name: "no workbook part",
    why: "a package whose workbook is missing",
    file: () => packageOf(bookOf({ without: ["xl/workbook.xml"] })),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "no workbook relationships",
    why: "a workbook whose relationships are missing",
    file: () => packageOf(bookOf({ without: ["xl/_rels/workbook.xml.rels"] })),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "no sheet part",
    why: "a sheet whose part is missing, known when it is read",
    file: () => packageOf(bookOf({ without: [SHEET] })),
    outcome: sheetRefused("INVALID_WORKBOOK"),
  },
  {
    name: "no sheets",
    why: "a workbook that lists no sheets is refused, never read as an empty one",
    file: file({ workbook: workbookPart("") }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "no sheet list",
    why: "a workbook part without a sheet list",
    file: file({ workbook: `<workbook xmlns="${MAIN}"/>` }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  ...TARGETS.map(([target, more]): HostileCase => ({
    name: `a sheet's target of ${JSON.stringify(target)}${more}`,
    why: "a target outside the package, or one that cannot be followed inside it (issue 84)",
    file: file({
      relationships: relationships(
        relationship("rId1", "worksheet", target, more),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  })),
  {
    name: "a sheet's target without a Target",
    why: "a relationship that names no part",
    file: file({
      relationships: relationships(
        `<Relationship Id="rId1" Type="${RELATIONSHIPS}/worksheet"/>`,
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  },
  {
    name: "a sheet's target that leads back to the workbook part",
    why: "a part that is no worksheet, read as one",
    file: file({
      relationships: relationships(
        relationship("rId1", "worksheet", "worksheets/../../xl/workbook.xml"),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: sheetRefused("INVALID_WORKBOOK"),
  },
  {
    name: "an absolute target",
    why: "a target from the root of the package",
    file: file({
      relationships: relationships(
        relationship("rId1", "worksheet", "/xl/worksheets/sheet1.xml"),
        relationship("rIdS", "sharedStrings", "/xl/sharedStrings.xml"),
      ),
    }),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a damaged sheet relationship type",
    why: "one character of the type changed: refused, never read as a sheet without rows (interrogation finding 1)",
    file: file({
      relationships: relationships(
        `<Relationship Id="rId1" Type="${RELATIONSHIPS}/workshees" Target="worksheets/sheet1.xml"/>`,
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  },
  {
    name: "a sheet relationship without a type",
    why: "the type left out: refused, never read as a workbook without that sheet",
    file: file({
      relationships: relationships(
        '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>',
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  },
  {
    name: "a sheet without a relationship",
    why: "a sheet that names a relationship the workbook does not have",
    file: file({ workbook: workbookPart(sheetEntry("Sheet1", "rId9")) }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  },
  {
    name: "a chart sheet beside an over-limit string table, the chart sheet read alone",
    why: "the table is not read for a sheet without cells, and its refusal ends no process (interrogation finding 3)",
    file: chartAndTable(),
    limits: { sharedStringBytes: 1_024 },
    plan: ["Chart"],
    outcome: read(rows("Chart"), { name: "Sheet1" }),
  },
  {
    name: "a chart sheet beside an over-limit string table",
    why: "the chart sheet has no rows and the worksheet is refused for the table",
    file: chartAndTable(),
    limits: { sharedStringBytes: 1_024 },
    outcome: read(
      rows("Chart"),
      tableRefused("Sheet1", "LIMIT_EXCEEDED", { limit: "sharedStringBytes" }),
    ),
  },
  {
    name: "a dialog sheet and a macro sheet",
    why: "sheets without cells by their kind are listed and have no rows",
    file: file({
      workbook: workbookPart(
        sheetEntry("Dialog", "rIdD") + sheetEntry("Macro", "rIdM") + sheetEntry("Sheet1", "rId1"),
      ),
      relationships: relationships(
        `<Relationship Id="rIdD" Type="${DIALOG}" Target="dialogsheets/sheet1.xml"/>`,
        `<Relationship Id="rIdM" Type="${MACRO}" Target="macrosheets/sheet1.xml"/>`,
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: read(rows("Dialog"), rows("Macro"), rows("Sheet1", [1, "a", 1])),
  },
  {
    name: "sheets named like members of an object",
    why: "no name from the file becomes a property",
    file: () => {
      const names = ["__proto__", "constructor", "toString", "hasOwnProperty", "then", "0"];
      return packageOf(
        bookOf({
          strings: null,
          sheets: names.map((name, at) => [
            name,
            worksheet(`<row r="1"><c r="A1"><v>${at}</v></c></row>`),
          ]),
        }),
      );
    },
    outcome: read(
      ...["__proto__", "constructor", "toString", "hasOwnProperty", "then", "0"].map((name, at) =>
        rows(name, [1, at]),
      ),
    ),
  },
  {
    name: "two sheets with one name",
    why: "a caller chooses a sheet by its exact name",
    file: () =>
      packageOf(
        bookOf({
          sheets: [
            ["a", worksheet("")],
            ["a", worksheet("")],
          ],
        }),
      ),
    outcome: opening("INVALID_WORKBOOK", { sheet: "a" }),
  },
  {
    name: "sheets named with the Kelvin sign and with k",
    why: "names that differ by Unicode folding are two sheets, as LibreOffice writes them (interrogation finding 9)",
    file: () =>
      packageOf(
        bookOf({
          strings: null,
          sheets: [
            ["\u212a", worksheet('<row r="1"><c r="A1"><v>1</v></c></row>')],
            ["k", worksheet('<row r="1"><c r="A1"><v>2</v></c></row>')],
            ["K", worksheet('<row r="1"><c r="A1"><v>3</v></c></row>')],
          ],
        }),
      ),
    outcome: read(rows("\u212a", [1, 1]), rows("k", [1, 2]), rows("K", [1, 3])),
  },
  {
    name: "100 000 sheets that name one part, the last read by its name",
    why: "a workbook of 100 000 sheets opens, lists them all, and reads one by its name",
    file: () =>
      packageOf(
        bookOf({
          workbook: workbookPart(
            Array.from({ length: 100_000 }, (_, at) => sheetEntry(`S${at + 1}`, "rId1")).join(""),
          ),
        }),
      ),
    plan: ["S100000"],
    outcome: () =>
      read(
        ...Array.from({ length: 99_999 }, (_, at): SheetOutcome => ({ name: `S${at + 1}` })),
        rows("S100000", [1, "a", 1]),
      ),
  },
  {
    name: "100 000 sheets that name one long target",
    why: "a relationship is followed once however many sheets name it: a copy of its target each held 192 MiB from 7 kB (issue 84)",
    file: () => {
      const target = `worksheets/${"a".repeat(2_000)}.xml`;
      return packageOf(
        bookOf({
          workbook: workbookPart(
            Array.from({ length: 100_000 }, (_, at) => sheetEntry(`S${at + 1}`, "rId1")).join(""),
          ),
          relationships: relationships(
            relationship("rId1", "worksheet", target),
            relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
          ),
          parts: [
            {
              path: `xl/${target}`,
              data: worksheet(
                '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>',
              ),
            },
          ],
          without: [SHEET],
        }),
      );
    },
    plan: ["S1"],
    outcome: () =>
      read(
        rows("S1", [1, "a", 1]),
        ...Array.from({ length: 99_999 }, (_, at): SheetOutcome => ({ name: `S${at + 2}` })),
      ),
  },
  {
    name: "a workbook under a long folder with 2 000 relationships",
    why: "each part name a relationship resolves to is charged before it is made: the folder repeated in each held 115 MiB from 265 kB (issue 84)",
    file: () => {
      const folder = "f/".repeat(30_000);
      const sheets = Array.from({ length: 2_000 }, (_, at) => sheetEntry(`S${at}`, `rId${at}`));
      const related = Array.from({ length: 2_000 }, (_, at) =>
        relationship(`rId${at}`, "worksheet", `sheet${at}.xml`),
      );
      return packageOf([
        {
          path: "_rels/.rels",
          data: relationships(relationship("rId1", "officeDocument", `${folder}workbook.xml`)),
        },
        { path: `${folder}workbook.xml`, data: workbookPart(sheets.join("")) },
        { path: `${folder}_rels/workbook.xml.rels`, data: relationships(...related) },
      ]);
    },
    outcome: opening("LIMIT_EXCEEDED", { limit: "inflatedBytes" }),
  },
  {
    name: "a relationship type of 4 190 000 tabs",
    why: "decoding an attribute of many line ends once made a string per character: about 130 MiB of heap from 5 kB (issue 84)",
    file: () =>
      packageOf(
        bookOf({
          relationships: relationships(
            relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
            relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
            `<Relationship Id="rIdU" Type="${"\t".repeat(4_190_000)}" Target="unused.xml"/>`,
          ),
        }),
      ),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "a workbook target of 4 190 006 characters",
    why: "a part name is folded whole before it is looked up, which issue 84 left for a change of its own",
    file: () =>
      packageOf(
        bookOf({
          parts: [
            {
              path: "_rels/.rels",
              data: relationships(
                relationship("rId1", "officeDocument", `${"aA".repeat(2_095_000)}/w.xml`),
              ),
            },
          ],
        }),
      ),
    outcome: opening("INVALID_WORKBOOK"),
    // Folding the name whole needs more than 96 MiB of old space on Node 24 and 26.
    heapMebibytes: 256,
  },
  ...["TRUE", "yes", "off", " 1", ""].map((spelling): HostileCase => ({
    name: `date1904="${spelling}"`,
    why: "a date system that is neither true nor false (issue 86)",
    file: file({
      workbook: workbookPart(sheetEntry("Sheet1", "rId1"), `<workbookPr date1904="${spelling}"/>`),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  })),
  {
    name: "two workbook properties",
    why: "a date system stated twice (issue 86)",
    file: file({
      workbook: workbookPart(
        sheetEntry("Sheet1", "rId1"),
        '<workbookPr date1904="0"/><workbookPr date1904="1"/>',
      ),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a foreign-namespace date1904",
    why: "an attribute under a prefix is never the attribute of its local name (issue 105)",
    file: () =>
      packageOf(
        bookOf({
          workbook: workbookPart(
            sheetEntry("Sheet1", "rId1"),
            '<workbookPr xmlns:foo="urn:foo" foo:date1904="1"/>',
          ),
          styles: styleSheet([], [0, 14]),
          sheets: [["Sheet1", worksheet('<row r="1"><c r="A1" s="1"><v>1</v></c></row>')]],
        }),
      ),
    outcome: read(rows("Sheet1", [1, new Date(Date.UTC(1900, 0, 1))])),
  },
  ...(
    [
      ["a child of the sheet list that is no sheet", `${sheetEntry("Sheet1", "rId1")}<foo/>`],
      ["a child of a sheet", `<sheet name="Sheet1" sheetId="1" r:id="rId1"><foo/></sheet>`],
      ["a sheet without a name", '<sheet sheetId="1" r:id="rId1"/>'],
      [
        "a sheet below an AlternateContent",
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${sheetEntry("Sheet1", "rId1")}</mc:AlternateContent>`,
      ],
      [
        "a sheet of another namespace",
        `${sheetEntry("Sheet1", "rId1")}<x:sheet xmlns:x="urn:x" name="X" sheetId="2" r:id="rId1"/>`,
      ],
    ] as const
  ).map(([name, sheets]): HostileCase => ({
    name,
    why: "the sheet list holds sheets and nothing else: passing over one listed one sheet fewer than LibreOffice shows (issue 84)",
    file: file({ workbook: workbookPart(sheets) }),
    outcome: opening("INVALID_WORKBOOK"),
  })),
  {
    name: "a child of a relationship",
    why: "a relationship holds nothing: passing over what it holds could hide a sheet (issue 84)",
    file: file({
      relationships: relationships(
        `<Relationship Id="rId1" Type="${RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"><x/></Relationship>`,
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a relationships root of another namespace",
    why: "a root of another namespace is no list of relationships (issue 105)",
    file: file({
      relationships: `<Relationships xmlns="urn:other">${relationship("rId1", "worksheet", "worksheets/sheet1.xml")}</Relationships>`,
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a relationship id stated twice under two prefixes",
    why: "two attributes that are one after their prefixes resolve, where the reader asks for one by namespace (issue 105)",
    file: file({
      workbook: `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}" xmlns:q="${RELATIONSHIPS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" q:id="rId9"/></sheets></workbook>`,
    }),
    outcome: opening("INVALID_XML"),
  },
  {
    name: "a sheet list below an AlternateContent of another namespace",
    why: "an AlternateContent is known by its local name in any namespace (issue 105)",
    file: file({
      workbook: `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}"><foo:AlternateContent xmlns:foo="urn:foo"><sheets>${sheetEntry("Sheet1", "rId1")}</sheets></foo:AlternateContent></workbook>`,
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a styles relationship without a Target",
    why: "a styles part that cannot be found (issue 86)",
    file: file({
      styles: styleSheet([], [0, 14]),
      relationships: relationships(
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
        `<Relationship Id="rIdT" Type="${RELATIONSHIPS}/styles"/>`,
      ),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "two sheet lists",
    why: "a sheet list stated twice",
    file: file({
      workbook: `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}"><sheets>${sheetEntry("Sheet1", "rId1")}</sheets><sheets/></workbook>`,
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a sheet list below an AlternateContent",
    why: "markup compatibility could choose another list",
    file: file({
      workbook: `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheets>${sheetEntry("Sheet1", "rId1")}</sheets></mc:AlternateContent></workbook>`,
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a workbook root of another namespace",
    why: "a root of another namespace is no workbook (issue 105)",
    file: file({
      workbook: `<workbook xmlns="urn:other" xmlns:r="${RELATIONSHIPS}"><sheets>${sheetEntry("Sheet1", "rId1")}</sheets></workbook>`,
    }),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a workbook root of no namespace",
    why: "a root of no namespace is no workbook (issue 105)",
    file: file({
      workbook: `<workbook xmlns:r="${RELATIONSHIPS}"><sheets>${sheetEntry("Sheet1", "rId1")}</sheets></workbook>`,
    }),
    outcome: opening("NOT_XLSX"),
  },
  {
    name: "a relationship id of another namespace",
    why: "the one attribute read by namespace names no relationship under another (issue 105)",
    file: file({
      workbook: `<workbook xmlns="${MAIN}" xmlns:r="urn:other"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    }),
    outcome: opening("INVALID_WORKBOOK", { sheet: "Sheet1" }),
  },
  {
    name: "two workbooks",
    why: "a package with two office documents",
    file: () =>
      packageOf(
        bookOf({
          parts: [
            {
              path: "_rels/.rels",
              data: relationships(
                relationship("rId1", "officeDocument", "xl/workbook.xml"),
                relationship("rId2", "officeDocument", "xl/workbook.xml"),
              ),
            },
          ],
        }),
      ),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "two shared-string tables",
    why: "a workbook that names two tables",
    file: file({
      relationships: relationships(
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
        relationship("rIdU", "sharedStrings", "sharedStrings.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "two styles parts",
    why: "a workbook that names two styles parts (issue 86)",
    file: file({
      styles: styleSheet([], [0]),
      relationships: relationships(
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
        relationship("rIdT", "styles", "styles.xml"),
        relationship("rIdU", "styles", "styles.xml"),
      ),
    }),
    outcome: opening("INVALID_WORKBOOK"),
  },
  {
    name: "a styles relationship without a part",
    why: "a styles part that is missing (issue 86)",
    file: () => packageOf(bookOf({ styles: styleSheet([], [0]), without: ["xl/styles.xml"] })),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  },
  {
    name: "a shared-string relationship without a part",
    why: "a shared-string table that is missing",
    file: () => packageOf(bookOf({ without: [STRINGS] })),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  },
  {
    name: "Strict namespaces",
    why: "the other family of namespace names",
    file: () =>
      packageOf([
        {
          path: "_rels/.rels",
          data: relationships(
            `<Relationship Id="rId1" Type="${STRICT_RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>`,
          ),
        },
        {
          path: "xl/workbook.xml",
          data: `<workbook xmlns="${STRICT_MAIN}" xmlns:r="${STRICT_RELATIONSHIPS}"><sheets><sheet name="Strict" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        },
        {
          path: "xl/_rels/workbook.xml.rels",
          data: relationships(
            `<Relationship Id="rId1" Type="${STRICT_RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/>`,
          ),
        },
        {
          path: SHEET,
          data: `<worksheet xmlns="${STRICT_MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>strict</t></is></c><c r="B1" t="d"><v>2026-11-17T12:34:56Z</v></c><c r="C1"><v>7</v></c></row></sheetData></worksheet>`,
        },
      ]),
    outcome: read(rows("Strict", [1, "strict", new Date(Date.UTC(2026, 10, 17, 12, 34, 56)), 7])),
  },
  {
    name: "Strict elements in a Transitional part",
    why: "both families are read in one part, as LibreOffice reads them (issue 105)",
    file: sheet(
      `<row r="1"><c r="A1"><v>1</v></c></row><s:row xmlns:s="${STRICT_MAIN}" r="2"><s:c r="A2"><s:v>2</s:v></s:c></s:row>`,
    ),
    outcome: read(rows("Sheet1", [1, 1], [2, 2])),
  },
];

function chartAndTable(): () => Uint8Array {
  return file({
    workbook: workbookPart(sheetEntry("Chart", "rIdC") + sheetEntry("Sheet1", "rId1")),
    relationships: relationships(
      relationship("rIdC", "chartsheet", "chartsheets/sheet1.xml"),
      relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
      relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
    ),
    strings: table(...Array.from({ length: 200 }, (_, at) => item(`text ${at}`))),
    parts: [{ path: "xl/chartsheets/sheet1.xml", data: `<chartsheet xmlns="${MAIN}"/>` }],
  });
}

const BAD_REFERENCES = [
  "XFE1",
  "A1048577",
  "A0",
  "1A",
  "",
  "A".repeat(300),
  "a1",
  "A-1",
  "A1:B2",
  "$A$1",
  "A1.5",
  "ÅA1",
];

const BAD_ROW_NUMBERS = ["0", "1048577", "-1", "abc", "1.5", "99999999999999999999", ""];

const BAD_STRING_INDEXES = ["5", "-1", "abc", "99999999999", "1.5", "", "0x1"];

const BAD_NUMBERS = [
  "1e999",
  "NaN",
  "Infinity",
  "-Infinity",
  "0x10",
  "1,5",
  "",
  "--1",
  "1e",
  ".",
  "1 2",
  "٣",
  "INF",
];

const BROKEN_XML = [
  `${SHEET_HEAD}<row r="1"><c r="A1"`,
  `${SHEET_HEAD}<row r="1"><c r="A1"><v>1</c></row>`,
  `${SHEET_HEAD}<row r="1"><c r="A1"><v>1</v>`,
  worksheet("</row></row><c><v>1</v></c>"),
  worksheet('<row r="1" r="2"><c r="A1"><v>1</v></c></row>'),
  worksheet("<row r=1><c r=A1><v>1</v></c></row>"),
  worksheet("<<row>><c><v>1</v></c>"),
  `${worksheet('<row r="1"><c r="A1"><v>1</v></c></row>')}\0\0`,
];

/** A row of one cell of `text`, as a string a formula stored. */
const textCell = (text: string) => `<row r="1"><c r="A1" t="str"><v>${text}</v></c></row>`;

const worksheetCases: HostileCase[] = [
  ...BAD_REFERENCES.map((reference): HostileCase => ({
    name: `a cell reference of ${JSON.stringify(reference.length > 20 ? `${reference.slice(0, 20)}…` : reference)}`,
    why: "a reference outside the sheet or not a reference",
    file: sheet(`<row r="1"><c r="${reference}"><v>1</v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", {
      cell: reference.length > 10 ? `${reference.slice(0, 10)}…` : reference,
    }),
  })),
  ...BAD_ROW_NUMBERS.map((number): HostileCase => ({
    name: `a row number of ${JSON.stringify(number)}`,
    why: "a row outside the sheet or not a number",
    file: sheet(`<row r="${number}"><c><v>1</v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", number === "1048577" ? { cell: "A1048577" } : {}),
  })),
  {
    name: "rows 5, 2, 3 and 3",
    why: "rows out of order are refused, never sorted",
    file: sheet(
      '<row r="5"><c r="A5"><v>5</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row><row r="3"><c r="A3"><v>33</v></c></row>',
    ),
    outcome: sheetRefused("INVALID_CELL", { cell: "A2" }),
  },
  {
    name: "cells C1, A1, A1 and B2",
    why: "cells out of order, and one naming another row",
    file: sheet(
      '<row r="1"><c r="C1"><v>3</v></c><c r="A1"><v>1</v></c><c r="A1"><v>11</v></c><c r="B2"><v>2</v></c></row>',
    ),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  },
  {
    name: "a cell that names another row",
    why: "a reference names its own row (issue 85)",
    file: sheet('<row r="1"><c r="A2"><v>1</v></c></row>'),
    outcome: sheetRefused("INVALID_CELL", { cell: "A2" }),
  },
  {
    name: "a reference written with a character reference",
    why: "a raw attribute span is not decoded, so A&#49; is no reference (issue 83)",
    file: sheet('<row r="1"><c r="A&#49;"><v>1</v></c></row>'),
    outcome: sheetRefused("INVALID_CELL", { cell: "A&#49;" }),
  },
  {
    name: "one cell in the last row and column",
    why: "the row the ADR fixes: every empty cell before the value is null (interrogation finding 12)",
    file: sheet('<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>'),
    outcome: read(rows("Sheet1", [1_048_576, ...Array.from({ length: 16_383 }, () => null), 1])),
  },
  {
    name: "200 000 rows of one cell in the last column",
    why: "3.3 billion positions if rows are dense: more than `cellPositions` allows",
    file: sheet(() =>
      Array.from(
        { length: 200_000 },
        (_, at) => `<row r="${at + 1}"><c r="XFD${at + 1}"><v>1</v></c></row>`,
      ).join(""),
    ),
    outcome: sheetRefused("LIMIT_EXCEEDED", { limit: "cellPositions" }),
  },
  {
    name: "one row of 5 million cells without references",
    why: "a row past the last column",
    file: sheetData(() =>
      repeated(`${SHEET_HEAD}<row>`, "<c><v>1</v></c>", 75 * MEBIBYTE, `</row>${SHEET_TAIL}`),
    ),
    outcome: sheetRefused("INVALID_CELL", { cell: "XFE1" }),
  },
  {
    name: "23 million rows of one cell",
    why: "600 MiB of rows, more than `inflatedBytes` allows",
    file: sheetData(() =>
      repeated(SHEET_HEAD, "<row><c><v>1</v></c></row>", 600 * MEBIBYTE, SHEET_TAIL),
    ),
    outcome: sheetRefused("LIMIT_EXCEEDED", { limit: "inflatedBytes" }),
  },
  ...BAD_STRING_INDEXES.map((index): HostileCase => ({
    name: `a shared string of index ${JSON.stringify(index)} in a table of two`,
    why: "a shared string that does not exist",
    file: sheet(`<row r="1"><c r="A1" t="s"><v>${index}</v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  })),
  {
    name: "a shared string without a table",
    why: "a workbook that names no table",
    file: file({ strings: null }),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  },
  ...BAD_NUMBERS.map((number): HostileCase => ({
    name: `a number of ${JSON.stringify(number)}`,
    why: "a number that is not an XML Schema double, or no finite one (issue 85)",
    file: sheet(`<row r="1"><c r="A1"><v>${number}</v></c><c r="B1"><v>2</v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  })),
  {
    name: "numbers in every form XML Schema allows",
    why: "white space, a leading plus, an exponent and a negative zero",
    file: sheet(
      '<row r="1"><c r="A1"><v> +1.5 </v></c><c r="B1"><v>1E+3</v></c><c r="C1"><v>-0</v></c><c r="D1"><v>.5</v></c><c r="E1"><v>1e-400</v></c></row>',
    ),
    outcome: read(rows("Sheet1", [1, 1.5, 1000, -0, 0.5, 0])),
  },
  {
    name: "booleans that are not 0 or 1",
    why: "a boolean of 2",
    file: sheet('<row r="1"><c r="A1" t="b"><v>2</v></c><c r="B1" t="b"><v>yes</v></c></row>'),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  },
  {
    name: "a cell of an unknown type",
    why: "a type the format does not define",
    file: sheet('<row r="1"><c r="A1" t="zz"><v>1</v></c></row>'),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  },
  {
    name: "empty values of text and of an error",
    why: "an empty value is null for a formula's text and for an error (issue 85)",
    file: sheet(
      '<row r="1"><c r="A1" t="str"><v></v></c><c r="B1" t="e"><v></v></c><c r="C1"><v>1</v></c></row>',
    ),
    outcome: read(rows("Sheet1", [1, null, null, 1])),
  },
  ...(["n", "s", "b"] as const).map((type): HostileCase => ({
    name: `an empty value of type ${type}`,
    why: "an empty value is refused for a number, a shared string and a boolean (issue 85)",
    file: sheet(`<row r="1"><c r="A1" t="${type}"><v></v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  })),
  {
    name: "an inline string with a value",
    why: "an inline string refuses any v (issue 85)",
    file: sheet('<row r="1"><c r="A1" t="inlineStr"><v>1</v></c></row>'),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  },
  {
    name: "every type of cell",
    why: "a formula is its stored result, an error is its text, a formula without a result is null",
    file: sheet(
      '<row r="1"><c r="A1" t="str"><f>B1&amp;"x"</f><v>ax</v></c><c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="C1" t="b"><v>0</v></c><c r="D1" t="n"><v>-0.5</v></c><c r="E1"><f>SUM(A2:A3)</f></c><c r="F1" t="e"><v>#N/A</v></c><c r="G1"><f>1+1</f><v>2</v></c><c r="H1" t="str"><f>""</f><v></v></c><c r="I1" t="b"><f>TRUE()</f><v>1</v></c><c r="J1" t="inlineStr"><is><r><t>in</t></r><r><t>line</t></r></is></c></row>',
    ),
    outcome: read(
      rows("Sheet1", [1, "ax", "#DIV/0!", false, -0.5, null, "#N/A", 2, null, true, "inline"]),
    ),
  },
  ...(
    [
      ["two sheet data", worksheet("").replace("</worksheet>", "<sheetData/></worksheet>")],
      [
        "a row outside the sheet data",
        worksheet("").replace("</worksheet>", '<row r="1"/></worksheet>'),
      ],
      ["no sheet data", `<worksheet xmlns="${MAIN}"/>`],
      [
        "sheet data below an AlternateContent",
        `<worksheet xmlns="${MAIN}"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/></mc:AlternateContent><sheetData/></worksheet>`,
      ],
      ["an element in the sheet data that is no row", worksheet("<foo/>")],
      [
        "a worksheet root of another namespace",
        `<worksheet xmlns="urn:other"><sheetData/></worksheet>`,
      ],
      [
        "rows in sheet data of another namespace",
        `<worksheet xmlns="${MAIN}"><x:sheetData xmlns:x="urn:x"><row r="1"><c r="A1"><v>1</v></c></row></x:sheetData></worksheet>`,
      ],
      [
        "a row of another namespace",
        worksheet('<x:row xmlns:x="urn:x" r="1"><c r="A1"><v>1</v></c></x:row>'),
      ],
      ["a row of no namespace", worksheet('<row xmlns="" r="1"><c r="A1"><v>1</v></c></row>')],
    ] as const
  ).map(([name, xml]): HostileCase => ({
    name,
    why: "the sheet data is one, under the root, and holds rows only: anything else is a refusal, never a shorter sheet (issues 85 and 105)",
    file: file({ sheets: [["Sheet1", xml]] }),
    outcome: sheetRefused("INVALID_WORKBOOK"),
  })),
  ...(
    [
      ["an element in a row that is no cell", '<row r="1"><foo/></row>', "A1"],
      ["an element in a cell that is no value", '<row r="1"><c r="A1"><foo/></c></row>', "A1"],
      ["two values in a cell", '<row r="1"><c r="A1"><v>1</v><v>2</v></c></row>', "A1"],
      [
        "a cell of another namespace",
        '<row r="1"><x:c xmlns:x="urn:x" r="A1"><v>1</v></x:c></row>',
        "A1",
      ],
      [
        "a value of another namespace",
        '<row r="1"><c r="A1"><x:v xmlns:x="urn:x">1</x:v></c></row>',
        "A1",
      ],
      [
        "an inline string of another namespace",
        '<row r="1"><c r="A1" t="inlineStr"><x:is xmlns:x="urn:x"><t>a</t></x:is></c></row>',
        "A1",
      ],
    ] as const
  ).map(([name, rowsXml, cell]): HostileCase => ({
    name,
    why: "a row holds cells and extLst, a cell a formula, extLst and one value: anything else is refused (issues 85 and 105)",
    file: sheet(rowsXml),
    outcome: sheetRefused("INVALID_CELL", { cell }),
  })),
  {
    name: "extension lists in a row and a cell",
    why: "what a reader passes over holds no part of the answer",
    file: sheet(
      '<row r="1"><c r="A1"><extLst><ext><v>9</v></ext></extLst><v>1</v></c><extLst><ext><c r="B1"><v>9</v></c></ext></extLst></row>',
    ),
    outcome: read(rows("Sheet1", [1, 1])),
  },
  {
    name: "elements nested 300 deep in an extension list",
    why: "elements nest at most 256 deep (ADR 0032)",
    file: sheet(
      () => `<row r="1"><extLst>${"<a>".repeat(300)}${"</a>".repeat(300)}</extLst></row>`,
    ),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "2.8 million nested elements, never closed",
    why: "the file the prototypes were attacked with: refused for its shape before its depth",
    file: sheetData(() => repeated(SHEET_HEAD, "<a>", 8 * MEBIBYTE, "")),
    outcome: sheetRefused("INVALID_WORKBOOK"),
  },
  {
    name: "open elements whose names hold more than 4 194 304 characters together",
    why: "the third fixed cap: without it the depth cap does not bound memory (issue 83)",
    file: sheet(
      () =>
        `<row r="1"><extLst>${`<${"n".repeat(17_000)}>`.repeat(250)}${`</${"n".repeat(17_000)}>`.repeat(250)}</extLst></row>`,
    ),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "a tag of five million spaces",
    why: "one token is at most 4 194 304 characters (ADR 0032)",
    file: sheet(() => `<row r="1"><c r="A1"${" ".repeat(5_000_000)}><v>1</v></c></row>`),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "a value of 4 194 305 characters",
    why: "one token longer than any workbook holds",
    file: sheet(() => textCell("a".repeat(ROW_CHARACTERS + 1))),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "a row of 4 194 304 characters",
    why: "a row as long as `rowCharacters` allows is read",
    file: sheet(
      () =>
        `<row r="1"><c r="A1" t="str"><v>${"a".repeat(ROW_CHARACTERS - 1)}</v></c><c r="B1" t="str"><v>b</v></c></row>`,
    ),
    outcome: () => read(rows("Sheet1", [1, "a".repeat(ROW_CHARACTERS - 1), "b"])),
  },
  {
    name: "a row of 4 194 305 characters",
    why: "one character more than `rowCharacters` allows",
    file: sheet(
      () =>
        `<row r="1"><c r="A1" t="str"><v>${"a".repeat(ROW_CHARACTERS)}</v></c><c r="B1" t="str"><v>b</v></c></row>`,
    ),
    outcome: sheetRefused("LIMIT_EXCEEDED", { limit: "rowCharacters" }),
  },
  ...(["cells", "runs"] as const).map((row): HostileCase => ({
    name: `three sheets of one row of 119 MiB in ${row}, read at once`,
    why: "each read inside every default limit but `rowCharacters` (interrogation finding 5)",
    file: threeSheetsOfOneRow(row),
    plan: "together",
    outcome: read(
      ...["Sheet1", "Sheet2", "Sheet3"].map((name) =>
        refused(name, "LIMIT_EXCEEDED", { limit: "rowCharacters" }),
      ),
    ),
    heapMebibytes: LARGE_HEAP_MEBIBYTES,
  })),
  {
    name: "a value split by four million comments",
    why: "text of many pieces is joined in batches: one piece at a time, a 75 kB file cost 240 MiB (issue 85)",
    file: sheet(() => textCell("a<!---->".repeat(4_000_000))),
    outcome: () => read(rows("Sheet1", [1, "a".repeat(4_000_000)])),
  },
  {
    name: "a value of a million escapes",
    why: "escapes are decoded in batches",
    file: sheet(() => textCell("_x0041_".repeat(500_000))),
    outcome: () => read(rows("Sheet1", [1, "A".repeat(500_000)])),
  },
  {
    name: "six cells of the longest text of character references",
    why: "the most one read reached in trials, 106 MiB of heap (ADR 0032)",
    file: sheet(
      () =>
        `<row r="1">${Array.from(
          { length: 6 },
          (_, at) =>
            `<c r="${columnName(at + 1)}1" t="str"><v>${"&#256;".repeat(Math.floor(ROW_CHARACTERS / 6))}</v></c>`,
        ).join("")}</row>`,
    ),
    outcome: () =>
      read(
        rows("Sheet1", [
          1,
          ...Array.from({ length: 6 }, () => "Ā".repeat(Math.floor(ROW_CHARACTERS / 6))),
        ]),
      ),
  },
  {
    name: "a cell of a million duplicate attributes",
    why: "an attribute stated twice is not well-formed",
    file: sheet(() => `<row r="1"><c r="A1"${' a="b"'.repeat(1_000_000)}><v>1</v></c></row>`),
    outcome: sheetRefused("INVALID_XML"),
  },
  ...BROKEN_XML.map((xml, at): HostileCase => ({
    name: `malformed XML ${at}`,
    why: "a sheet part that is not well-formed",
    file: file({ sheets: [["Sheet1", xml]] }),
    outcome: sheetRefused("INVALID_XML"),
  })),
  {
    name: "a sheet in UTF-16",
    why: "a part in an encoding other than UTF-8",
    file: sheetData(() =>
      utf16(
        `<?xml version="1.0" encoding="UTF-16"?>${worksheet('<row r="1"><c r="A1"><v>1</v></c></row>')}`,
      ),
    ),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "a declaration of XML 1.1",
    why: "version 1.0 only: XML 1.1 reads NEL and U+2028 as line ends (issue 83)",
    file: file({
      sheets: [
        ["Sheet1", `<?xml version="1.1"?>${worksheet('<row r="1"><c r="A1"><v>1</v></c></row>')}`],
      ],
    }),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "entity expansion in a sheet",
    why: "a DOCTYPE is refused, whatever it declares",
    file: file({
      sheets: [
        [
          "Sheet1",
          `<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol0 "lol">${Array.from({ length: 9 }, (_, at) => `<!ENTITY lol${at + 1} "${`&lol${at};`.repeat(10)}">`).join("")}]>${worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>&lol9;</t></is></c></row>')}`,
        ],
      ],
    }),
    outcome: sheetRefused("INVALID_XML"),
  },
  {
    name: "a DOCTYPE in the workbook part",
    why: "an external entity in the part that lists the sheets",
    file: file({
      workbook: `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>${workbookPart(sheetEntry("Sheet1", "rId1"))}`,
    }),
    outcome: opening("INVALID_XML"),
  },
  {
    name: "a byte-order mark, comments, instructions and CDATA",
    why: "forms of XML any part may use",
    file: () =>
      packageOf(
        bookOf({
          strings: null,
          sheets: [
            [
              "Sheet1",
              `﻿<?xml version="1.0" encoding="UTF-8"?><!-- a comment --><worksheet xmlns="${MAIN}"><?pi x?><sheetData><row r="1"><c r="A1" t="str"><v><![CDATA[a<b && c]]></v></c><c r="B1" t="inlineStr"><is><t xml:space="preserve">  two  </t></is></c></row></sheetData></worksheet>`,
            ],
          ],
        }),
      ),
    outcome: read(rows("Sheet1", [1, "a<b && c", "  two  "])),
  },
  ...(
    [
      ["an element under a prefix nothing binds", '<row r="1"><x:c r="A1"><v>1</v></x:c></row>'],
      [
        "an attribute under a prefix nothing binds",
        '<row r="1"><c r="A1" x:s="1"><v>1</v></c></row>',
      ],
      ["a declaration of xmlns", '<row r="1" xmlns:xmlns="urn:x"><c r="A1"><v>1</v></c></row>'],
      [
        "xml bound to another namespace",
        '<row r="1" xmlns:xml="urn:x"><c r="A1"><v>1</v></c></row>',
      ],
      ["a prefix bound to no namespace", '<row r="1" xmlns:p=""><c r="A1"><v>1</v></c></row>'],
      [
        "1 024 namespace declarations in scope",
        `<row r="1" ${Array.from({ length: 1_024 }, (_, at) => `xmlns:p${at}="urn:${at}"`).join(" ")}><c r="A1"><v>1</v></c></row>`,
      ],
    ] as const
  ).map(([name, rowsXml]): HostileCase => ({
    name,
    why: "the corners of the namespace rules, decided as Python's namespace-aware parser decides them (issue 105)",
    file: sheet(rowsXml),
    outcome: sheetRefused("INVALID_XML"),
  })),
  {
    name: "1 000 namespace declarations in scope",
    why: "fewer than 1 024 declarations fit beside the names of the open elements (issue 105)",
    file: sheet(
      `<row r="1" ${Array.from({ length: 1_000 }, (_, at) => `xmlns:p${at}="urn:${at}"`).join(" ")}><c r="A1"><v>1</v></c></row>`,
    ),
    outcome: read(rows("Sheet1", [1, 1])),
  },
  {
    name: "a declaration on every row",
    why: "a declaration is resolved once, where it stands",
    file: sheet(() =>
      Array.from(
        { length: 10_000 },
        (_, at) =>
          `<x:row xmlns:x="${MAIN}" r="${at + 1}"><x:c r="A${at + 1}"><x:v>${at}</x:v></x:c></x:row>`,
      ).join(""),
    ),
    outcome: () =>
      read(rows("Sheet1", ...Array.from({ length: 10_000 }, (_, at): Row => [at + 1, at]))),
  },
  {
    name: "the bytes of a date that is no date",
    why: "the file the prototypes were attacked with: serials outside the calendar and a date of text that is none",
    file: sheet(
      '<row r="1"><c r="A1" s="1"><v>-5</v></c><c r="B1" s="1"><v>1e300</v></c><c r="C1" s="1"><v>2958466</v></c><c r="D1" t="d"><v>not a date</v></c></row>',
      { styles: styleSheet([], [0, 14]) },
    ),
    outcome: sheetRefused("INVALID_CELL", { cell: "D1" }),
  },
];

function utf16(text: string): Uint8Array {
  const bytes = new Uint8Array(2 + text.length * 2);
  bytes.set([0xff, 0xfe]);
  for (let at = 0; at < text.length; at++) {
    bytes[2 + at * 2] = text.charCodeAt(at) & 0xff;
    bytes[3 + at * 2] = text.charCodeAt(at) >> 8;
  }
  return bytes;
}

const SST_HEAD = `<sst xmlns="${MAIN}">`;

/** A book whose table is `items` and whose sheet names its first entry. */
const strings = (items: Text) => () =>
  packageOf(
    bookOf({
      strings: table(made(items)),
      sheets: [["Sheet1", worksheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>')]],
    }),
  );

const stringCases: HostileCase[] = [
  {
    name: "a table of two shared strings that states 4 294 967 295",
    why: "`count` and `uniqueCount` are the file's word: nothing is sized from them",
    file: file({
      strings: `<sst xmlns="${MAIN}" count="4294967295" uniqueCount="4294967295">${item("a")}${item("b")}</sst>`,
    }),
    outcome: DEFAULT_ROWS,
  },
  {
    name: "20 million shared strings",
    why: "340 MiB of table, more than `sharedStringBytes` allows",
    file: stringsData(() => repeated(SST_HEAD, "<si><t>a</t></si>", 340 * MEBIBYTE, "</sst>")),
    outcome: read(tableRefused("Sheet1", "LIMIT_EXCEEDED", { limit: "sharedStringBytes" })),
  },
  {
    name: "2 million shared strings in 34 MiB",
    why: "under `inflatedBytes`, over `sharedStringBytes`",
    file: stringsData(() => repeated(SST_HEAD, "<si><t>a</t></si>", 34 * MEBIBYTE, "</sst>")),
    outcome: read(tableRefused("Sheet1", "LIMIT_EXCEEDED", { limit: "sharedStringBytes" })),
  },
  {
    name: "1 000 001 shared strings in 13 MB",
    why: "one entry more than `sharedStrings` allows, inside `sharedStringBytes`",
    file: stringsData(() => [
      bytesOf(SST_HEAD),
      bytesOf("<si><t/></si>".repeat(1_000_001)),
      bytesOf("</sst>"),
    ]),
    outcome: read(tableRefused("Sheet1", "LIMIT_EXCEEDED", { limit: "sharedStrings" })),
  },
  {
    name: "one shared string of 200 MiB",
    why: "a table larger than `sharedStringBytes`",
    file: stringsData(() => repeated(`${SST_HEAD}<si><t>`, "a", 200 * MEBIBYTE, "</t></si></sst>")),
    outcome: read(tableRefused("Sheet1", "LIMIT_EXCEEDED", { limit: "sharedStringBytes" })),
  },
  {
    name: "2 million nested runs in a shared string, never closed",
    why: "the file the prototypes were attacked with: refused for its shape",
    file: stringsData(() => repeated(`${SST_HEAD}<si>`, "<r>", 6 * MEBIBYTE, "")),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  },
  ...(
    [
      ["an entry of two texts", "<si><t>a</t><t>b</t></si>"],
      ["a run of two texts", "<si><r><t>a</t><t>b</t></r></si>"],
      ["an entry with an unknown child", "<si><t>a</t><x/></si>"],
      ["an entry of another namespace", '<x:si xmlns:x="urn:x"><t>a</t></x:si><si><t>b</t></si>'],
      ["an entry in an extension list", "<extLst><si><t>a</t></si></extLst><si><t>b</t></si>"],
    ] as const
  ).map(([name, items]): HostileCase => ({
    name,
    why: "an entry holds one text or runs of one text each: anything else is refused, never read as another index (issues 85 and 105)",
    file: strings(items),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  })),
  {
    name: "a table root of another namespace",
    why: "a root of another namespace is no table (issue 105)",
    file: file({ strings: `<sst xmlns="urn:other">${item("a")}</sst>` }),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  },
  {
    name: "rich text and phonetic runs",
    why: "runs are joined, phonetic text is passed over",
    file: strings(
      '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> and plain</t></r><rPh sb="0" eb="1"><t>PHONETIC</t></rPh><phoneticPr fontId="1"/></si>',
    ),
    outcome: read(rows("Sheet1", [1, "Bold and plain"])),
  },
  {
    name: "escapes in shared strings, inline strings and a formula's text",
    why: "_xHHHH_ is decoded in each, within one run, and half a surrogate pair reads as written (issue 85)",
    file: file({
      strings: table(item("line_x000D__x000A_next"), item("_x005F_x000D_"), item("a_xD800_")),
      sheets: [
        [
          "Sheet1",
          worksheet(
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="inlineStr"><is><r><t>_x00</t></r><r><t>41_</t></r></is></c><c r="E1" t="str"><f>A1</f><v>_x263A_</v></c><c r="F1" t="e"><v>_x0041_</v></c></row>',
          ),
        ],
      ],
    }),
    outcome: read(
      rows("Sheet1", [1, "line\r\nnext", "_x000D_", "a_xD800_", "_x0041_", "☺", "_x0041_"]),
    ),
  },
  {
    name: "a DOCTYPE in the shared strings",
    why: "an external entity in the table",
    file: file({
      strings: `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>${table(item("&e;"), item("b"))}`,
    }),
    outcome: read(tableRefused("Sheet1", "INVALID_XML")),
  },
  {
    name: "entities XML does not define",
    why: "an entity no DOCTYPE may declare",
    file: strings(item("&nbsp;&bogus;")),
    outcome: read(tableRefused("Sheet1", "INVALID_XML")),
  },
  {
    name: "bytes that are not UTF-8 in a shared string",
    why: "a part that is not UTF-8",
    file: stringsData(() => [
      bytesOf(`${SST_HEAD}<si><t>`),
      new Uint8Array([0xff, 0xfe, 0xc3, 0x28]),
      bytesOf("</t></si></sst>"),
    ]),
    outcome: read(tableRefused("Sheet1", "INVALID_XML")),
  },
  {
    name: "a million unterminated entities in a shared string",
    why: "a reference that never ends",
    file: strings(() => item("&amp".repeat(1_000_000))),
    outcome: read(tableRefused("Sheet1", "INVALID_XML")),
  },
  {
    name: "1.5 million half-written escapes",
    why: "text that starts an escape and never finishes one",
    file: strings(() => item("_x".repeat(1_500_000))),
    outcome: () => read(rows("Sheet1", [1, "_x".repeat(1_500_000)])),
  },
];

/** A book of one sheet whose cells are `cellsXml` in row 1, with the styles `styles`. */
const styled =
  (cellsXml: string, styles: Text, more: Book = {}) =>
  () =>
    packageOf(
      bookOf({
        ...more,
        strings: null,
        styles: made(styles),
        sheets: [["Sheet1", worksheet(`<row r="1">${cellsXml}</row>`)]],
      }),
    );

const DATE = styleSheet([], [0, 14]);

const SLOW_FORMATS: readonly (readonly [string, () => string, Row])[] = [
  ["unbalanced quotes", () => "&quot;a".repeat(100_000), [1, 5]],
  ["open brackets", () => "[$-".repeat(70_000), [1, 5]],
  ["backslashes", () => "\\".repeat(200_001), [1, 5]],
  [
    "letters then a quote",
    () => `${"y".repeat(200_000)}&quot;`,
    [1, new Date(Date.UTC(1900, 0, 5))],
  ],
  ["nested elapsed", () => `${"[h".repeat(100_000)}]`, [1, 5]],
  ["sections", () => `0${";".repeat(200_000)}`, [1, 5]],
];

const styleCases: HostileCase[] = [
  {
    name: "General written as a custom code",
    why: "General is not a date: its letters read as era tokens turned 28 of 84 workbooks' numbers into dates (interrogation finding 2)",
    file: styled('<c r="A1" s="1"><v>1234.5</v></c>', styleSheet([[164, "General"]], [0, 164])),
    outcome: read(rows("Sheet1", [1, 1234.5])),
  },
  ...SLOW_FORMATS.map(([name, code, row]): HostileCase => ({
    name: `a format code of 200 000 characters: ${name}`,
    why: "a format code is tokenized once, with no regular expression",
    file: styled('<c r="A1" s="1"><v>5</v></c>', () => styleSheet([[200, code()]], [0, 200])),
    outcome: read(rows("Sheet1", row)),
  })),
  {
    name: "style indexes and format ids far out of range",
    why: "the file the prototypes were attacked with: a negative format id",
    file: styled(
      '<c r="A1" s="4000000000"><v>5</v></c><c r="B1" s="-1"><v>5</v></c><c r="C1" s="x"><v>5</v></c>',
      styleSheet(
        [
          [4_000_000_000, "yyyy"],
          [-5, "yyyy"],
        ],
        [0, 4_000_000_000, -5],
      ),
    ),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  },
  ...(
    [
      [
        "formats after the cell styles",
        `<styleSheet xmlns="${MAIN}"><cellXfs><xf numFmtId="164"/></cellXfs><numFmts><numFmt numFmtId="164" formatCode="yyyy"/></numFmts></styleSheet>`,
      ],
      [
        "a second list of formats in an extension list",
        `<styleSheet xmlns="${MAIN}"><numFmts/><cellXfs><xf numFmtId="164"/></cellXfs><extLst><numFmts><numFmt numFmtId="164" formatCode="yyyy"/></numFmts></extLst></styleSheet>`,
      ],
      [
        "a second list of cell styles",
        `<styleSheet xmlns="${MAIN}"><cellXfs><xf numFmtId="0"/></cellXfs><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>`,
      ],
      [
        "an extension list in the cell styles",
        `<styleSheet xmlns="${MAIN}"><cellXfs><xf numFmtId="14"/><extLst/></cellXfs></styleSheet>`,
      ],
      [
        "a format with a child",
        `<styleSheet xmlns="${MAIN}"><numFmts><numFmt numFmtId="164" formatCode="yyyy"><x/></numFmt></numFmts><cellXfs><xf numFmtId="164"/></cellXfs></styleSheet>`,
      ],
      [
        "a format id with a sign",
        styleSheet([], [0, 14]).replace('numFmtId="14"', 'numFmtId="+14"'),
      ],
      [
        "a format id with white space",
        styleSheet([], [0, 14]).replace('numFmtId="14"', 'numFmtId=" 14"'),
      ],
      [
        "a format id of sixteen digits",
        styleSheet([], [0, 14]).replace('numFmtId="14"', 'numFmtId="0000000000000014"'),
      ],
      [
        "cell styles below an AlternateContent",
        `<styleSheet xmlns="${MAIN}"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><cellXfs><xf numFmtId="14"/></cellXfs></mc:AlternateContent></styleSheet>`,
      ],
      [
        "a styles root of another namespace",
        `<styleSheet xmlns="urn:other"><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>`,
      ],
    ] as const
  ).map(([name, styles]): HostileCase => ({
    name,
    why: "a styles part LibreOffice and openpyxl read in different ways is refused (issues 86 and 105)",
    file: styled('<c r="A1" s="1"><v>5</v></c>', styles),
    outcome: read(tableRefused("Sheet1", "INVALID_WORKBOOK")),
  })),
  ...(
    [
      ["a style that does not exist", "9"],
      ["a style that is not digits", "x"],
      ["a style of minus one", "-1"],
    ] as const
  ).map(([name, style]): HostileCase => ({
    name: `${name}, on a number`,
    why: "where a style can change the value, one that names no style is refused (issue 86)",
    file: styled(`<c r="A1" s="${style}"><v>5</v></c>`, DATE),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  })),
  {
    name: "a style that does not exist, on text",
    why: "where a style cannot change the value it is not read, so not refused (issue 86)",
    file: styled(
      '<c r="A1" s="9" t="inlineStr"><is><t>a</t></is></c><c r="B1" s="9" t="b"><v>1</v></c>',
      DATE,
    ),
    outcome: read(rows("Sheet1", [1, "a", true])),
  },
  {
    name: "text, a boolean and an error under a date style",
    why: "only a number becomes a date (issue 86)",
    file: styled(
      '<c r="A1" s="1" t="inlineStr"><is><t>5</t></is></c><c r="B1" s="1" t="b"><v>1</v></c><c r="C1" s="1" t="e"><v>#N/A</v></c><c r="D1" s="1" t="str"><f>A1</f><v>5</v></c><c r="E1" s="1"><f>1+4</f><v>5</v></c>',
      DATE,
    ),
    outcome: read(rows("Sheet1", [1, "5", true, "#N/A", "5", new Date(Date.UTC(1900, 0, 5))])),
  },
  {
    name: "a format id the part does not define",
    why: "an id past the built-in ones without a code shows no date (issue 86)",
    file: styled('<c r="A1" s="1"><v>5</v></c>', styleSheet([], [0, 99])),
    outcome: read(rows("Sheet1", [1, 5])),
  },
  {
    name: "a format id under a prefix",
    why: "foo:numFmtId names no format (issue 105)",
    file: styled(
      '<c r="A1" s="1"><v>5</v></c>',
      `<styleSheet xmlns="${MAIN}" xmlns:foo="urn:foo"><cellXfs><xf numFmtId="0"/><xf foo:numFmtId="14"/></cellXfs></styleSheet>`,
    ),
    outcome: read(rows("Sheet1", [1, 5])),
  },
  {
    name: "serials at the edges of the calendar",
    why: "a Date only for a serial the writer can write, checked after rounding to the millisecond (interrogation finding 7)",
    file: styled(
      [-1, 0.5, 1, 59, 59.99999999999, 60, 61, 2_958_465, 2_958_465.99999, 2_958_465.9999999995]
        .map((serial, at) => `<c r="${columnName(at + 1)}1" s="1"><v>${serial}</v></c>`)
        .join(""),
      DATE,
    ),
    outcome: read(
      rows("Sheet1", [
        1,
        -1,
        0.5,
        new Date(Date.UTC(1900, 0, 1)),
        new Date(Date.UTC(1900, 1, 28)),
        59.99999999999,
        60,
        new Date(Date.UTC(1900, 2, 1)),
        new Date(Date.UTC(9999, 11, 31)),
        new Date(Date.UTC(9999, 11, 31, 23, 59, 59, 136)),
        2_958_465.9999999995,
      ]),
    ),
  },
  {
    name: "serials in the 1904 date system",
    why: "serial 0 is 1904-01-01 and a negative serial stays a number",
    file: styled('<c r="A1" s="1"><v>0</v></c><c r="B1" s="1"><v>-1</v></c>', DATE, {
      workbook: workbookPart(sheetEntry("Sheet1", "rId1"), '<workbookPr date1904="true"/>'),
    }),
    outcome: read(rows("Sheet1", [1, new Date(Date.UTC(1904, 0, 1)), -1])),
  },
  {
    name: "dates stored as text in every form read",
    why: "a date, a date and time with decimals rounded half up, Z, and white space around (issue 86)",
    file: sheet(
      '<row r="1"><c r="A1" t="d"><v>2026-11-17</v></c><c r="B1" t="d"><v> 2026-11-17T12:34:56.5Z </v></c><c r="C1" t="d"><v>2026-11-17T12:34:56.12350</v></c></row>',
    ),
    outcome: read(
      rows("Sheet1", [
        1,
        new Date(Date.UTC(2026, 10, 17)),
        new Date(Date.UTC(2026, 10, 17, 12, 34, 56, 500)),
        new Date(Date.UTC(2026, 10, 17, 12, 34, 56, 124)),
      ]),
    ),
  },
  ...[
    "2026-11-17T12:34:56+01:00",
    "2026-02-30",
    "1899-12-31",
    "9999-12-31T23:59:59.9995",
    "17/11/2026",
  ].map((text): HostileCase => ({
    name: `a date stored as ${JSON.stringify(text)}`,
    why: "an offset, a day that does not exist, a date the writer cannot write, another form (issue 86)",
    file: sheet(`<row r="1"><c r="A1" t="d"><v>${text}</v></c></row>`),
    outcome: sheetRefused("INVALID_CELL", { cell: "A1" }),
  })),
  {
    name: "800 000 cell styles",
    why: "the styles are held as one byte per cell style",
    file: styled(
      '<c r="A1" s="799999"><v>5</v></c>',
      () =>
        `<styleSheet xmlns="${MAIN}"><cellXfs>${'<xf numFmtId="0"/>'.repeat(799_999)}<xf numFmtId="14"/></cellXfs></styleSheet>`,
    ),
    outcome: read(rows("Sheet1", [1, new Date(Date.UTC(1900, 0, 5))])),
  },
];

export const hostileCases: readonly HostileCase[] = [
  ...packageCases,
  ...workbookCases,
  ...worksheetCases,
  ...stringCases,
  ...styleCases,
];
