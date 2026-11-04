import { XML_DECLARATION, escapeXml } from "../xml/index.js";

export type { Workbook } from "./lib/model.js";
export { validateWorkbook } from "./lib/validate.js";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const SPREADSHEET_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Renders `[Content_Types].xml`: one override per part, plus the defaults readers expect. */
export function renderContentTypes(sheetCount: number): string {
  let overrides = "";
  for (let i = 1; i <= sheetCount; i++) {
    overrides += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }
  return (
    `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NAMESPACE}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    overrides +
    "</Types>"
  );
}

/** Renders `_rels/.rels`: the package points at the workbook part. */
export function renderRootRelationships(): string {
  return (
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument" Target="xl/workbook.xml"/>` +
    "</Relationships>"
  );
}

/** What the workbook part needs per sheet: its name and, when filtered, its autofilter range. */
export interface SheetEntry {
  readonly name: string;
  readonly filterRange: string | undefined;
}

/**
 * Renders `xl/workbook.xml`: the sheet list in tab order, each pointing at its relationship, and
 * the hidden `_xlnm._FilterDatabase` name Excel keeps for every autofilter.
 */
export function renderWorkbook(entries: readonly SheetEntry[]): string {
  let sheets = "";
  let definedNames = "";
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]!;
    sheets += `<sheet name="${escapeXml(entry.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
    if (entry.filterRange !== undefined) {
      definedNames += `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${escapeXml(quoteSheetName(entry.name))}!${absolute(entry.filterRange)}</definedName>`;
    }
  }
  return (
    `${XML_DECLARATION}<workbook xmlns="${SPREADSHEET_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}">` +
    `<sheets>${sheets}</sheets>` +
    (definedNames.length === 0 ? "" : `<definedNames>${definedNames}</definedNames>`) +
    "</workbook>"
  );
}

/** A sheet name as a formula reference: quoted, with apostrophes doubled. */
function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

/** `A1:C4` as `$A$1:$C$4`. */
function absolute(range: string): string {
  return range.replace(/([A-Z]+)(\d+)/g, "$$$1$$$2");
}

/** Renders `xl/_rels/workbook.xml.rels`: sheets first (rId1..N), then styles and shared strings. */
export function renderWorkbookRelationships(sheetCount: number): string {
  let relationships = "";
  for (let i = 1; i <= sheetCount; i++) {
    relationships += `<Relationship Id="rId${i}" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/worksheet" Target="worksheets/sheet${i}.xml"/>`;
  }
  relationships += `<Relationship Id="rId${sheetCount + 1}" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/styles" Target="styles.xml"/>`;
  relationships += `<Relationship Id="rId${sheetCount + 2}" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/sharedStrings" Target="sharedStrings.xml"/>`;
  return (
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    relationships +
    "</Relationships>"
  );
}
