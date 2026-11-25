import type { StreamOptions, StreamedWorkbook, Workbook, WriteOptions } from "../../src/index.js";
import type { OpenedWorkbook } from "../../src/read.js";

/** What every host reports back after running the scenario. */
export interface HostReport {
  /** The STORE package, base64: byte-identical on every host. */
  readonly stored: string;
  /** The default package, base64: DEFLATE entries where the host can deflate, STORE otherwise. */
  readonly deflated: string;
  /** The workbook streamed with shared strings and stored blocks from async rows, base64: byte-identical on every host. */
  readonly streamed: string;
  /** The workbook streamed with the defaults, base64: raw deflate where the host has it, stored blocks otherwise. */
  readonly streamedDeflated: string;
  /** The media type exported by the browser entry. */
  readonly mime: string;
  /** The error code `downloadXlsx` threw on a host without a document. */
  readonly noDom: string | undefined;
  /** Each package above read back by the read entry, and the default package once more from a `Blob`. */
  readonly read: {
    readonly stored: ReadBack;
    readonly deflated: ReadBack;
    readonly streamed: ReadBack;
    readonly streamedDeflated: ReadBack;
    readonly blob: ReadBack;
  };
}

/** A value as JSON carries it: a `Date` as `{ date }` with its ISO instant. */
export type PlainValue = string | number | boolean | null | { readonly date: string };

/** The sheets and every sheet's rows, each as its number and its cells; or the refusal's code. */
export type ReadBack =
  | {
      readonly sheets: readonly string[];
      readonly rows: readonly (readonly [number, ...PlainValue[]])[];
    }
  | { readonly refused: string; readonly coded: boolean };

/** The shape of the core entry the scenario needs, as any host loads it. */
export interface CoreLike {
  writeXlsx(workbook: Workbook, options?: WriteOptions): Promise<Uint8Array>;
  streamXlsx(workbook: StreamedWorkbook, options?: StreamOptions): ReadableStream<Uint8Array>;
  XlsxMiniError: new (...args: never[]) => Error;
}

/** The shape of the browser entry the scenario needs. */
export interface BrowserLike {
  downloadXlsx(bytes: Uint8Array, filename: string): void;
  XLSX_MIME_TYPE: string;
}

/** The shape of the read entry the scenario needs. */
export interface ReadLike {
  openXlsx(file: Uint8Array | Blob): Promise<OpenedWorkbook>;
  XlsxReadError: new (...args: never[]) => Error;
}

export declare const workbook: Workbook;
export declare function run(
  lib: CoreLike,
  browser: BrowserLike,
  read: ReadLike,
): Promise<HostReport>;
