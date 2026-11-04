import { createPackage, type PackagePart } from "./opc/index.js";
import { createStringTable } from "./shared-strings/index.js";
import { renderStyles } from "./styles/index.js";
import {
  type Workbook,
  renderContentTypes,
  renderRootRelationships,
  renderWorkbook,
  renderWorkbookRelationships,
} from "./workbook/index.js";
import { renderWorksheet } from "./worksheet/index.js";

const encoder = new TextEncoder();

/**
 * Writes a workbook to the bytes of an .xlsx file.
 *
 * The result is a plain `Uint8Array`: hand it to `fs.writeFile`, a `Response`, or `new Blob([bytes])`.
 * Identical input gives identical bytes.
 *
 * @example
 * const bytes = await writeXlsx({
 *   sheets: [{ name: "Orders", rows: [["Order", "Total"], ["A-1001", 1234.5]] }],
 * });
 */
export async function writeXlsx(workbook: Workbook): Promise<Uint8Array> {
  const strings = createStringTable();
  const sheets = workbook.sheets;

  // Sheets render first: they fill the shared-strings table the sharedStrings part is built from.
  const sheetParts: PackagePart[] = [];
  for (let i = 0; i < sheets.length; i++) {
    sheetParts.push(part(`xl/worksheets/sheet${i + 1}.xml`, renderWorksheet(sheets[i]!, strings)));
  }

  return createPackage([
    part("[Content_Types].xml", renderContentTypes(sheets.length)),
    part("_rels/.rels", renderRootRelationships()),
    part("xl/workbook.xml", renderWorkbook(sheets.map((sheet) => sheet.name))),
    part("xl/_rels/workbook.xml.rels", renderWorkbookRelationships(sheets.length)),
    part("xl/styles.xml", renderStyles()),
    part("xl/sharedStrings.xml", strings.render()),
    ...sheetParts,
  ]);
}

function part(path: string, xml: string): PackagePart {
  return { path, data: encoder.encode(xml) };
}
