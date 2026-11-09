import {
  type CompressionMode,
  createPackage,
  type PackagePart,
  type StreamedPart,
  selectCompression,
  selectStreamCompression,
  streamPackage,
} from "./opc/index.js";
import { createStringTable } from "./shared-strings/index.js";
import { createStyleTable } from "./styles/index.js";
import {
  type SheetEntry,
  type StreamedWorkbook,
  type Workbook,
  renderContentTypes,
  renderRootRelationships,
  renderWorkbook,
  renderWorkbookRelationships,
  validateWorkbook,
} from "./workbook/index.js";
import { type Coerce, renderWorksheet, streamWorksheet } from "./worksheet/index.js";

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
 * The result is an ArrayBuffer-backed `Uint8Array`: hand it to `fs.writeFile`, a `Response`, or
 * `new Blob([bytes])` as it is.
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
): Promise<Uint8Array<ArrayBuffer>> {
  validateWorkbook(workbook);
  const strings = createStringTable();
  const styles = createStyleTable();
  const sheets = workbook.sheets;

  // Sheets render first: they fill the shared-strings and styles tables those parts are built from.
  const sheetParts: PackagePart[] = [];
  const entries: SheetEntry[] = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i]!;
    // The same shape as the streamed context, so the per-cell check reads a present property.
    const context = {
      strings,
      styles,
      coerce: options.coerce,
      sheetIndex: i,
      inlineStrings: false,
    };
    const rendered = renderWorksheet(sheet, context);
    sheetParts.push({ path: `xl/worksheets/sheet${i + 1}.xml`, data: rendered.data });
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

/** How to stream, on top of how to write. */
export interface StreamOptions extends WriteOptions {
  /**
   * `"inline"` (the default) writes text into each cell, so memory stays flat however many
   * distinct strings the rows hold. `"shared"` keeps a shared-strings table as `writeXlsx` does:
   * a smaller file, at the cost of holding every distinct string until the end.
   */
  readonly strings?: "inline" | "shared";
}

/**
 * Writes a workbook as a stream of the bytes of an .xlsx file, produced as the consumer pulls
 * them: a `Response` body, `pipeTo` a file, or `for await` over the chunks. Rows may be an async
 * iterable, so a database cursor can feed a sheet. Peak memory does not grow with the row count.
 *
 * Throws `XlsxMiniError` before returning when the workbook itself is invalid; a bad cell met
 * while rendering errors the stream instead. Cancelling the stream returns the row iterators.
 *
 * @example
 * export function GET() {
 *   const workbook = { sheets: [{ name: "Orders", rows: orders() }] };
 *   return new Response(streamXlsx(workbook), { headers: { "Content-Type": XLSX_MIME_TYPE } });
 * }
 */
export function streamXlsx(
  workbook: StreamedWorkbook,
  options: StreamOptions = {},
): ReadableStream<Uint8Array> {
  validateWorkbook(workbook);
  const strings = createStringTable();
  const styles = createStyleTable();
  const sheets = workbook.sheets;
  const inlineStrings = (options.strings ?? "inline") === "inline";
  const compression = selectStreamCompression(options.compression ?? "deflate");

  // The packager takes each part only once the previous one has gone out, so the parts after the
  // sheets see every string and style the sheets registered.
  async function* parts(): AsyncGenerator<StreamedPart> {
    yield bytesPart("[Content_Types].xml", renderContentTypes(sheets.length));
    yield bytesPart("_rels/.rels", renderRootRelationships());
    yield bytesPart("xl/_rels/workbook.xml.rels", renderWorkbookRelationships(sheets.length));
    const entries: SheetEntry[] = [];
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i]!;
      const context = { strings, styles, coerce: options.coerce, sheetIndex: i, inlineStrings };
      const streamed = streamWorksheet(sheet, context);
      yield { path: `xl/worksheets/sheet${i + 1}.xml`, body: streamed.body };
      entries.push({ name: sheet.name, filterRange: streamed.filterRange() });
    }
    yield bytesPart("xl/workbook.xml", renderWorkbook(entries));
    yield bytesPart("xl/styles.xml", styles.render());
    yield bytesPart("xl/sharedStrings.xml", strings.render());
  }

  return streamPackage(parts(), compression);
}

function bytesPart(path: string, xml: string): StreamedPart {
  const data = encoder.encode(xml);
  return {
    path,
    body: new ReadableStream({
      start(controller) {
        controller.enqueue(data);
        controller.close();
      },
    }),
  };
}
