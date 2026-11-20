/**
 * The read side of the workbook: from an opened package to the list of its sheets. Owns the three
 * parts that answer it (the package's relationships, the workbook part, the workbook's
 * relationships) and every rule about them, so the caller asks one question and never sees a
 * relationship.
 *
 * Elements are matched by where they stand, since the tokenizer drops prefixes and looks at no
 * namespace: the root's local name, then the path down from it. On that path an element the schema
 * does not allow is refused, never skipped: skipped, it could hide a sheet another reader lists.
 */
import { XlsxReadError } from "../errors/read.js";
import { type PartName, partName } from "../opc/read.js";
import { type ByteChunks, type XmlVisitor, readXml } from "../xml/read.js";
import {
  refuse,
  relationshipsOf,
  relationshipsVisitor,
  resolveTarget,
} from "./read/relationships.js";

export interface ListedSheet {
  readonly name: string;
  /**
   * The part that holds the sheet's cells, or `undefined` for a sheet that has none by its kind: a
   * chart sheet, a dialog sheet, a macro sheet. Whether the package holds the part is known only
   * when it is opened.
   */
  readonly worksheet: PartName<"worksheet"> | undefined;
}

/**
 * How the workbook reader gets a part's bytes: `undefined` when the package has no such part. The
 * caller charges each part to the meter of its use, so this module holds no meter and no package.
 */
export type OpenPart = (name: string) => ByteChunks | undefined;

const TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const MACRO_SHEET = "http://schemas.microsoft.com/office/2006/relationships/";

/**
 * Whether a sheet of a relationship type has cells. A type that is not here, even by one character,
 * is no kind of sheet: read as a sheet without cells, it would make a damaged workbook an empty one.
 */
const SHEET_TYPES = new Map<string | undefined, boolean>([
  [`${TRANSITIONAL}worksheet`, true],
  [`${STRICT}worksheet`, true],
  [`${TRANSITIONAL}chartsheet`, false],
  [`${STRICT}chartsheet`, false],
  [`${TRANSITIONAL}dialogsheet`, false],
  [`${STRICT}dialogsheet`, false],
  [`${MACRO_SHEET}xlMacrosheet`, false],
  [`${MACRO_SHEET}xlIntlMacrosheet`, false],
]);

const WORKBOOK_TYPES = new Set<string | undefined>([
  `${TRANSITIONAL}officeDocument`,
  `${STRICT}officeDocument`,
]);

const UNRELATED = "a sheet's relationship is missing or not to a sheet";

const UNFOLLOWED = "a relationship does not name a part of the package";

const PACKAGE_ROOT = "";

function notWorkbook(): never {
  throw new XlsxReadError("NOT_XLSX", "not an .xlsx file");
}

/**
 * Reads the sheet list from three parts, opened through `open` one after the other, and nothing
 * else.
 *
 * Refuses, never repairs. `NOT_XLSX` when the package has no relationships of its own, none to a
 * workbook, or a workbook part whose root is not `workbook`. `INVALID_WORKBOOK` when there are two
 * workbooks, a part is missing or damaged, the sheet list is missing, empty, stated twice or holds
 * anything but sheets, a sheet list or a sheet stands below an `AlternateContent`, a sheet has no
 * name, no relationship, or one of a type that is not a kind of sheet, a target cannot be followed
 * inside the package, or two sheets have one name, character for character.
 */
export async function readWorkbook(open: OpenPart): Promise<readonly ListedSheet[]> {
  const root = await readXml(
    open(relationshipsOf(PACKAGE_ROOT)) ?? notWorkbook(),
    relationshipsVisitor(),
  );
  const [document, another] = [...root.values()].filter(({ type }) => WORKBOOK_TYPES.has(type));
  if (document === undefined) notWorkbook();
  if (another !== undefined) refuse("the package holds two workbooks");
  const workbook = resolveTarget(PACKAGE_ROOT, document.target) ?? refuse(UNFOLLOWED);
  const listed = await readXml(
    open(workbook) ?? refuse("the workbook part is missing"),
    workbookVisitor(),
  );
  const related = await readXml(
    open(relationshipsOf(workbook)) ?? refuse("the workbook's relationships are missing"),
    relationshipsVisitor(),
  );
  return listed.map(({ name, relationship }) => {
    const { type, target } = related.get(relationship) ?? refuse(UNRELATED, { sheet: name });
    const cells = SHEET_TYPES.get(type);
    if (cells === undefined) refuse(UNRELATED, { sheet: name });
    const part = resolveTarget(workbook, target) ?? refuse(UNFOLLOWED, { sheet: name });
    return { name, worksheet: cells ? partName(part, "worksheet") : undefined };
  });
}

interface SheetEntry {
  readonly name: string;
  readonly relationship: string | undefined;
}

/**
 * Inside the sheet list the schema allows `sheet` and nothing in it. Markup compatibility lets an
 * application take a `sheets` or a `sheet` from below an `AlternateContent` anywhere, so one there is
 * refused too.
 */
function workbookVisitor(): XmlVisitor<never, readonly SheetEntry[]> {
  const names = new Set<string>();
  let sheets: SheetEntry[] | undefined;
  let list: SheetEntry[] | undefined;
  let alternate = 0;
  return {
    open(xml) {
      if (xml.depth === 1 && xml.name !== "workbook") notWorkbook();
      if (alternate === 0 && xml.name === "AlternateContent") alternate = xml.depth;
      if (alternate !== 0 && (xml.name === "sheets" || xml.name === "sheet")) {
        refuse("a sheet stands below an AlternateContent");
      }
      if (xml.depth === 2) {
        list = undefined;
        if (xml.name !== "sheets") return;
        if (sheets !== undefined) refuse("the workbook lists its sheets twice");
        list = sheets = [];
        return;
      }
      if (list === undefined) return;
      if (xml.depth !== 3 || xml.name !== "sheet") refuse("the sheet list holds more than sheets");
      const name = xml.attribute("name");
      if (!name) refuse("a sheet has no name");
      if (names.has(name)) refuse("two sheets have one name", { sheet: name });
      names.add(name);
      list.push({ name, relationship: xml.attribute("id") });
    },
    close(xml) {
      if (xml.depth === alternate) alternate = 0;
      return undefined;
    },
    end() {
      if (!sheets?.length) refuse("the workbook lists no sheets");
      return sheets;
    },
  };
}
