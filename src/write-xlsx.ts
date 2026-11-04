import { createPackage, type PackagePart } from "./opc/index.js";
import { createStringTable } from "./shared-strings/index.js";
import { createStyleTable } from "./styles/index.js";
import {
  type Workbook,
  renderContentTypes,
  renderRootRelationships,
  renderWorkbook,
  renderWorkbookRelationships,
  validateWorkbook,
} from "./workbook/index.js";
import { type Coerce, renderWorksheet } from "./worksheet/index.js";

/** How to write, as opposed to what to write. */
export interface WriteOptions {
  /** Converts values the writer does not handle natively; see `Coerce`. */
  readonly coerce?: Coerce;
}

const encoder = new TextEncoder();

/**
 * Writes a workbook to the bytes of an .xlsx file.
 *
 * The result is a plain `Uint8Array`: hand it to `fs.writeFile`, a `Response`, or `new Blob([bytes])`.
 * Identical input gives identical bytes. Throws `XlsxMiniError` on input Excel would refuse.
 *
 * @example
 * const bytes = await writeXlsx({
 *   sheets: [{ name: "Orders", rows: [["Order", "Total"], ["A-1001", 1234.5]] }],
 * });
 */
export async function writeXlsx(
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  validateWorkbook(workbook);
  const strings = createStringTable();
  const styles = createStyleTable();
  const sheets = workbook.sheets;

  // Sheets render first: they fill the shared-strings and styles tables those parts are built from.
  const sheetParts: PackagePart[] = [];
  for (let i = 0; i < sheets.length; i++) {
    const context = { strings, styles, coerce: options.coerce, sheetIndex: i };
    sheetParts.push(part(`xl/worksheets/sheet${i + 1}.xml`, renderWorksheet(sheets[i]!, context)));
  }

  return createPackage([
    part("[Content_Types].xml", renderContentTypes(sheets.length)),
    part("_rels/.rels", renderRootRelationships()),
    part("xl/workbook.xml", renderWorkbook(sheets.map((sheet) => sheet.name))),
    part("xl/_rels/workbook.xml.rels", renderWorkbookRelationships(sheets.length)),
    part("xl/styles.xml", styles.render()),
    part("xl/sharedStrings.xml", strings.render()),
    ...sheetParts,
  ]);
}

function part(path: string, xml: string): PackagePart {
  return { path, data: encoder.encode(xml) };
}
