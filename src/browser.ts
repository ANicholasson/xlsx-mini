import { XlsxMiniError } from "./errors/index.js";
import { XLSX_MIME_TYPE } from "./workbook/index.js";

export { XLSX_MIME_TYPE };

/**
 * How long the object URL outlives the click. Revoking it at once cancels the download in some
 * browsers; forty seconds is the delay FileSaver.js settled on after years of bug reports.
 */
const REVOKE_DELAY_MS = 40_000;

/**
 * Saves `bytes` as a file through the browser's download mechanism: a `Blob` with the .xlsx media
 * type behind an object URL, clicked through an anchor with a `download` attribute. The `.xlsx`
 * extension is added when `filename` lacks it. Throws `NO_DOM` outside a browser; on a server,
 * return the bytes in a `Response` with `XLSX_MIME_TYPE` instead.
 *
 * @example
 * import { writeXlsx } from "xlsx-mini";
 * import { downloadXlsx } from "xlsx-mini/browser";
 *
 * downloadXlsx(await writeXlsx(workbook), "orders.xlsx");
 */
export function downloadXlsx(bytes: Uint8Array, filename: string): void {
  if (typeof document === "undefined") {
    throw new XlsxMiniError(
      "NO_DOM",
      "downloadXlsx",
      "downloadXlsx needs a browser document; on a server send the bytes in a Response",
    );
  }
  // The writer's output is ArrayBuffer-backed; the Blob constructor insists on that in its types.
  const blob = new Blob([bytes as Uint8Array<ArrayBuffer>], { type: XLSX_MIME_TYPE });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename.toLowerCase().endsWith(".xlsx") ? filename : `${filename}.xlsx`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), REVOKE_DELAY_MS);
}
