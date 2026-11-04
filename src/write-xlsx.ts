import {
  type CompressionMode,
  createPackage,
  type PackagePart,
  selectCompression,
} from "./opc/index.js";
import { createStringTable } from "./shared-strings/index.js";
import { createStyleTable } from "./styles/index.js";
import {
  type SheetEntry,
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
  /**
   * `"deflate"` (the default) compresses every part through the runtime's `CompressionStream`.
   * `"none"` writes STORE entries: larger, but byte-identical on every platform, which is what
   * golden tests want. Runtimes that cannot deflate get STORE either way.
   */
  readonly compression?: CompressionMode;
}

const encoder = new TextEncoder();

/**
 * Writes a workbook to the bytes of an .xlsx file.
 *
 * The result is a plain `Uint8Array`: hand it to `fs.writeFile`, a `Response`, or `new Blob([bytes])`.
 * Identical input gives identical bytes for `compression: "none"`. Throws `XlsxMiniError` on input Excel
 * would refuse.
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
  const entries: SheetEntry[] = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i]!;
    const context = { strings, styles, coerce: options.coerce, sheetIndex: i };
    const rendered = renderWorksheet(sheet, context);
    sheetParts.push(part(`xl/worksheets/sheet${i + 1}.xml`, rendered.xml));
    entries.push({ name: sheet.name, filterRange: rendered.filterRange });
  }

  const compression = selectCompression(options.compression ?? "deflate");
  return createPackage(
    [
      part("[Content_Types].xml", renderContentTypes(sheets.length)),
      part("_rels/.rels", renderRootRelationships()),
      part("xl/workbook.xml", renderWorkbook(entries)),
      part("xl/_rels/workbook.xml.rels", renderWorkbookRelationships(sheets.length)),
      part("xl/styles.xml", styles.render()),
      part("xl/sharedStrings.xml", strings.render()),
      ...sheetParts,
    ],
    compression,
  );
}

function part(path: string, xml: string): PackagePart {
  return { path, data: encoder.encode(xml) };
}
