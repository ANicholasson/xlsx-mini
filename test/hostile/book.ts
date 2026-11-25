import type { Part } from "../helpers/package.ts";

export const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships";
export const STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main";
export const STRICT_RELATIONSHIPS = "http://purl.oclc.org/ooxml/officeDocument/relationships";

export const worksheet = (rows: string, around: { before?: string; after?: string } = {}) =>
  `<worksheet xmlns="${MAIN}">${around.before ?? ""}<sheetData>${rows}</sheetData>${around.after ?? ""}</worksheet>`;

/** A shared-string table of these items, each written whole, as `<si>…</si>`. */
export const table = (...items: string[]) => `<sst xmlns="${MAIN}">${items.join("")}</sst>`;

export const item = (text: string) => `<si><t>${text}</t></si>`;

/** A styles part with these custom formats and one cell style per format id. */
export const styleSheet = (formats: readonly (readonly [number, string])[], styles: number[]) =>
  `<styleSheet xmlns="${MAIN}">${
    formats.length === 0
      ? ""
      : `<numFmts>${formats.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${code}"/>`).join("")}</numFmts>`
  }<cellXfs>${styles.map((id) => `<xf numFmtId="${id}"/>`).join("")}</cellXfs></styleSheet>`;

export const relationship = (id: string, type: string, target: string, more = "") =>
  `<Relationship Id="${id}" Type="${RELATIONSHIPS}/${type}" Target="${target}"${more}/>`;

export const relationships = (...listed: string[]) =>
  `<Relationships xmlns="${PACKAGE}">${listed.join("")}</Relationships>`;

export const workbookPart = (sheets: string, before = "") =>
  `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}">${before}<sheets>${sheets}</sheets></workbook>`;

export const sheetEntry = (name: string, id: string) =>
  `<sheet name="${name}" sheetId="1" r:id="${id}"/>`;

/** The parts a workbook needs, each of which a case may write differently or leave out. */
export interface Book {
  /** Each sheet's name and its part's text; one sheet of a shared string and a number by default. */
  readonly sheets?: readonly (readonly [name: string, xml: string])[];
  /** The shared-string table's text, `a` and `b` by default; `null` for a workbook without one. */
  readonly strings?: string | null;
  /** The styles part's text, one cell style of format 0 by default. */
  readonly styles?: string;
  /** The workbook part's text, in place of the one that lists `sheets`. */
  readonly workbook?: string;
  /** The workbook's relationships, in place of the ones to `sheets`, the table and the styles. */
  readonly relationships?: string;
  /** Parts to leave out, by name. */
  readonly without?: readonly string[];
  /** Parts written in place of the part of their name, or after the others. */
  readonly parts?: readonly Part[];
}

const TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml";

/** The content types of a workbook's parts, as an application states them. */
const contentTypes = (strings: boolean) =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${TYPE}.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="${TYPE}.styles+xml"/>${strings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${TYPE}.sharedStrings+xml"/>` : ""}</Types>`;

export const SHEET = "xl/worksheets/sheet1.xml";
export const STRINGS = "xl/sharedStrings.xml";

const ONE_ROW = worksheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>');

/** The parts of a workbook as an application writes them, with what `book` changes. */
export function bookOf(book: Book = {}): Part[] {
  const sheets = book.sheets ?? [["Sheet1", ONE_ROW]];
  const strings = book.strings === undefined ? table(item("a"), item("b")) : book.strings;
  const styles = book.styles ?? styleSheet([], [0]);
  const listed = sheets.map(([name], at) => sheetEntry(name, `rId${at + 1}`)).join("");
  const parts: Part[] = [
    { path: "[Content_Types].xml", data: contentTypes(strings !== null) },
    {
      path: "_rels/.rels",
      data: relationships(relationship("rId1", "officeDocument", "xl/workbook.xml")),
    },
    { path: "xl/workbook.xml", data: book.workbook ?? workbookPart(listed) },
    {
      path: "xl/_rels/workbook.xml.rels",
      data:
        book.relationships ??
        relationships(
          ...sheets.map((_, at) =>
            relationship(`rId${at + 1}`, "worksheet", `worksheets/sheet${at + 1}.xml`),
          ),
          strings === null ? "" : relationship("rIdS", "sharedStrings", "sharedStrings.xml"),
          relationship("rIdT", "styles", "styles.xml"),
        ),
    },
    ...sheets.map(([, xml], at) => ({ path: `xl/worksheets/sheet${at + 1}.xml`, data: xml })),
    ...(strings === null ? [] : [{ path: STRINGS, data: strings }]),
    { path: "xl/styles.xml", data: styles },
  ];
  const replaced = new Map((book.parts ?? []).map((part) => [part.path, part]));
  const kept = parts
    .filter(({ path }) => !book.without?.includes(path))
    .map((part) => replaced.get(part.path) ?? part);
  const added = (book.parts ?? []).filter(({ path }) => !parts.some((part) => part.path === path));
  return [...kept, ...added];
}
