import type { Workbook, WriteOptions } from "../../src/index.js";

/** What every host reports back after running the scenario. */
export interface HostReport {
  /** The STORE package, base64: byte-identical on every host. */
  readonly stored: string;
  /** The default package, base64: DEFLATE entries where the host can deflate, STORE otherwise. */
  readonly deflated: string;
  /** The media type exported by the browser entry. */
  readonly mime: string;
  /** The error code `downloadXlsx` threw on a host without a document. */
  readonly noDom: string | undefined;
}

/** The shape of the core entry the scenario needs, as any host loads it. */
export interface CoreLike {
  writeXlsx(workbook: Workbook, options?: WriteOptions): Promise<Uint8Array>;
  XlsxMiniError: new (...args: never[]) => Error;
}

/** The shape of the browser entry the scenario needs. */
export interface BrowserLike {
  downloadXlsx(bytes: Uint8Array, filename: string): void;
  XLSX_MIME_TYPE: string;
}

export declare const workbook: Workbook;
export declare function run(lib: CoreLike, browser: BrowserLike): Promise<HostReport>;
