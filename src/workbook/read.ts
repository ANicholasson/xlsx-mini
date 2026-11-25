/**
 * The read side of the workbook: from an opened package to the list of its sheets. Owns the three
 * parts that answer it (the package's relationships, the workbook part, the workbook's
 * relationships) and every rule about them, so the caller asks one question and never sees a
 * relationship. Owns the namespace of SpreadsheetML too, which every part reader of the workbook
 * asks for.
 */
import { type Meter, XlsxReadError } from "../errors/read.js";
import { type PartName, partName } from "../opc/read.js";
import { type ByteChunks, type XmlVisitor, readXml } from "../xml/read.js";
import {
  refuse,
  relationshipsOf,
  relationshipsVisitor,
  resolveTarget,
} from "./read/relationships.js";

/**
 * What opening a workbook reads: its sheets, its date system, and where its shared-string table and
 * its styles are.
 */
export interface ListedWorkbook {
  readonly sheets: readonly ListedSheet[];
  /**
   * The shared-string table's part, or `undefined` for a workbook without one. Whether the package
   * holds the part is known only when it is opened.
   */
  readonly sharedStrings: PartName<"sharedStrings"> | undefined;
  /** The styles part, or `undefined` for a workbook without one; known to be there when opened. */
  readonly styles: PartName<"structure"> | undefined;
  /** Whether serials count from 1904-01-01 (`date1904` of `workbookPr`) rather than 1899-12-30. */
  readonly date1904: boolean;
}

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

/** The number a part reader of SpreadsheetML gives its namespace. */
export const SPREADSHEET = 1;
const RELATIONSHIP = 2;

/**
 * SpreadsheetML's namespace in both families, Transitional and Strict (ECMA-376 part 1, annex A),
 * for a part reader's `namespaces`.
 */
export const SPREADSHEET_NAMESPACES: ReadonlyMap<string, number> = new Map([
  ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", SPREADSHEET],
  ["http://purl.oclc.org/ooxml/spreadsheetml/main", SPREADSHEET],
]);

const TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MACRO_SHEET = "http://schemas.microsoft.com/office/2006/relationships/";

const WORKBOOK_NAMESPACES = new Map([
  ...SPREADSHEET_NAMESPACES,
  [TRANSITIONAL, RELATIONSHIP],
  [STRICT, RELATIONSHIP],
]);

/**
 * Whether a sheet of a relationship type has cells. A type that is not here, even by one character,
 * is no kind of sheet: read as a sheet without cells, it would make a damaged workbook an empty one.
 */
const SHEET_TYPES = new Map<string | undefined, boolean>([
  [`${TRANSITIONAL}/worksheet`, true],
  [`${STRICT}/worksheet`, true],
  [`${TRANSITIONAL}/chartsheet`, false],
  [`${STRICT}/chartsheet`, false],
  [`${TRANSITIONAL}/dialogsheet`, false],
  [`${STRICT}/dialogsheet`, false],
  [`${MACRO_SHEET}xlMacrosheet`, false],
  [`${MACRO_SHEET}xlIntlMacrosheet`, false],
]);

const WORKBOOK_TYPES = new Set<string | undefined>([
  `${TRANSITIONAL}/officeDocument`,
  `${STRICT}/officeDocument`,
]);

const STRING_TYPES = new Set<string | undefined>([
  `${TRANSITIONAL}/sharedStrings`,
  `${STRICT}/sharedStrings`,
]);

const STYLE_TYPES = new Set<string | undefined>([`${TRANSITIONAL}/styles`, `${STRICT}/styles`]);

const UNRELATED = "a sheet's relationship is missing or not to a sheet";

const UNFOLLOWED = "a relationship does not name a part of the package";

const DATE_SYSTEM = "the date system is stated twice or is neither true nor false";

const PACKAGE_ROOT = "";

function notWorkbook(): never {
  throw new XlsxReadError("NOT_XLSX", "not an .xlsx file");
}

/**
 * Reads the sheet list from three parts, opened through `open` one after the other, and nothing
 * else; the workbook part also states its date system, and its relationships name its shared-string
 * table and its styles.
 *
 * A relationship is followed once, however many sheets name it: a workbook may point every sheet
 * at one long target, and resolving it per sheet would hold a copy of the path for each. The part
 * name a relationship resolves to is charged to `names` before it is made, by the length of the
 * workbook's own name and of the target: a workbook under a long folder would otherwise repeat
 * that folder for every relationship.
 *
 * Refuses, never repairs. `NOT_XLSX` when the package has no relationships of its own, none to a
 * workbook, or a workbook part whose root is not SpreadsheetML's `workbook`. A sheet names its
 * relationship by `id` in the namespace of relationships, Transitional or Strict; an `id` without a
 * prefix, or under a prefix bound elsewhere, names none. `INVALID_WORKBOOK` when there are two
 * workbooks, two shared-string tables or two styles parts, a part is missing or damaged, the sheet
 * list is missing, empty, stated twice or holds anything but sheets, a sheet list, a sheet or the
 * workbook's properties stand below an `AlternateContent`, a sheet has no name, no relationship, or
 * one of a type that is not a kind of sheet, a target cannot be followed inside the package, two
 * sheets have one name, character for character, or the properties are stated twice or say
 * `date1904` in any spelling but `1`, `true`, `0` and `false`.
 */
export async function readWorkbook(
  open: OpenPart,
  names: Meter<"inflatedBytes">,
): Promise<ListedWorkbook> {
  const root = await readXml(
    open(relationshipsOf(PACKAGE_ROOT)) ?? notWorkbook(),
    relationshipsVisitor(),
  );
  const [document, another] = [...root.values()].filter(({ type }) => WORKBOOK_TYPES.has(type));
  if (document === undefined) notWorkbook();
  if (another !== undefined) refuse("the package holds two workbooks");
  const workbook = resolveTarget(PACKAGE_ROOT, document.target) ?? refuse(UNFOLLOWED);
  const { entries, date1904 } = await readXml(
    open(workbook) ?? refuse("the workbook part is missing"),
    workbookVisitor(),
  );
  const related = await readXml(
    open(relationshipsOf(workbook)) ?? refuse("the workbook's relationships are missing"),
    relationshipsVisitor(),
  );
  const follow = (target: string | undefined, sheet?: string): string => {
    names.charge(workbook.length + (target?.length ?? 0));
    return resolveTarget(workbook, target) ?? refuse(UNFOLLOWED, { sheet });
  };
  const only = (types: ReadonlySet<string | undefined>, what: string): string | undefined => {
    const [first, second] = [...related.values()].filter(({ type }) => types.has(type));
    if (second !== undefined) refuse(`the workbook holds two ${what}`);
    return first && follow(first.target);
  };
  const strings = only(STRING_TYPES, "shared-string tables");
  const styles = only(STYLE_TYPES, "styles parts");
  const parts = new Map<string | undefined, PartName<"worksheet"> | undefined>();
  return {
    sheets: entries.map(({ name, relationship }) => {
      if (!parts.has(relationship)) {
        const { type, target } = related.get(relationship) ?? refuse(UNRELATED, { sheet: name });
        const cells = SHEET_TYPES.get(type);
        if (cells === undefined) refuse(UNRELATED, { sheet: name });
        const part = follow(target, name);
        parts.set(relationship, cells ? partName(part, "worksheet") : undefined);
      }
      return { name, worksheet: parts.get(relationship) };
    }),
    sharedStrings: strings === undefined ? undefined : partName(strings, "sharedStrings"),
    styles: styles === undefined ? undefined : partName(styles, "structure"),
    date1904,
  };
}

interface SheetEntry {
  readonly name: string;
  readonly relationship: string | undefined;
}

interface WorkbookEntries {
  readonly entries: readonly SheetEntry[];
  readonly date1904: boolean;
}

const DATE_SYSTEMS = new Map([
  ["1", true],
  ["true", true],
  ["0", false],
  ["false", false],
  [undefined, false],
]);

/**
 * Markup compatibility lets an application take a `sheets`, a `sheet` or a `workbookPr` from below
 * an `AlternateContent` anywhere, so one there is refused. An `AlternateContent` is known by its
 * local name in any namespace: refusing more never gives a value the application does not show.
 * A `workbookPr` deeper than the root's own, as in an extension, states nothing about
 * this workbook. An element in another namespace than SpreadsheetML's is no element of these.
 */
function workbookVisitor(): XmlVisitor<never, WorkbookEntries> {
  const names = new Set<string>();
  let sheets: SheetEntry[] | undefined;
  let list: SheetEntry[] | undefined;
  let date1904: boolean | undefined;
  let alternate = 0;
  return {
    namespaces: WORKBOOK_NAMESPACES,
    open(xml) {
      const element = xml.namespace === SPREADSHEET ? xml.name : "";
      if (xml.depth === 1 && element !== "workbook") notWorkbook();
      if (alternate === 0 && xml.name === "AlternateContent") alternate = xml.depth;
      if (
        alternate !== 0 &&
        (element === "sheets" || element === "sheet" || element === "workbookPr")
      ) {
        refuse("a sheet or the date system stands below an AlternateContent");
      }
      if (xml.depth === 2) {
        list = undefined;
        if (element === "workbookPr") {
          const first = date1904 === undefined;
          date1904 = DATE_SYSTEMS.get(xml.attribute("date1904"));
          if (!first || date1904 === undefined) refuse(DATE_SYSTEM);
        }
        if (element !== "sheets") return;
        if (sheets !== undefined) refuse("the workbook lists its sheets twice");
        list = sheets = [];
        return;
      }
      if (list === undefined) return;
      if (xml.depth !== 3 || element !== "sheet") refuse("the sheet list holds more than sheets");
      const name = xml.attribute("name");
      if (!name) refuse("a sheet has no name");
      if (names.has(name)) refuse("two sheets have one name", { sheet: name });
      names.add(name);
      list.push({ name, relationship: xml.attribute("id", RELATIONSHIP) });
    },
    close(xml) {
      if (xml.depth === alternate) alternate = 0;
    },
    end() {
      if (!sheets?.length) refuse("the workbook lists no sheets");
      return { entries: sheets, date1904: date1904 ?? false };
    },
  };
}
