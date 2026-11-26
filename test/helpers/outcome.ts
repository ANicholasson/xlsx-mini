import { crc32 } from "node:zlib";
import {
  type OpenedWorkbook,
  type ReadErrorCode,
  type ReadLimits,
  type ReadValue,
  XlsxReadError,
  openXlsx,
} from "../../src/read.js";

/** An `XlsxReadError` as a value, with only the fields it has. */
export interface Refusal {
  readonly code: ReadErrorCode;
  readonly limit?: keyof ReadLimits;
  readonly sheet?: string;
  readonly cell?: string;
}

export type Row = readonly [number, ...ReadValue[]];

/** One sheet of an opened workbook: its rows, its refusal, or neither when it was not read. */
export interface SheetOutcome {
  readonly name: string;
  readonly rows?: readonly Row[];
  readonly refused?: Refusal;
}

/** What reading a file ends in: a refusal to open it, or every sheet it lists, in order. */
export type Outcome = { readonly refused: Refusal } | { readonly sheets: readonly SheetOutcome[] };

/**
 * Which sheets a read reads: every sheet one after another, every sheet at once, or the sheets
 * named, one after another.
 */
export type Plan = "each" | "together" | readonly string[];

function refusalOf(error: unknown): Refusal {
  if (!(error instanceof XlsxReadError)) throw error;
  const { code, limit, sheet, cell } = error;
  return {
    code,
    ...(limit === undefined ? {} : { limit }),
    ...(sheet === undefined ? {} : { sheet }),
    ...(cell === undefined ? {} : { cell }),
  };
}

/** A turn of the event loop, after which a rejection that nothing awaits is reported. */
const turn = () => new Promise<void>((resolve) => setImmediate(resolve));

async function sheetOf(book: OpenedWorkbook, name: string): Promise<SheetOutcome> {
  const rows: Row[] = [];
  try {
    for await (const { number, cells } of book.rows(name)) rows.push([number, ...cells]);
  } catch (error) {
    return { name, refused: refusalOf(error) };
  }
  return { name, rows };
}

/**
 * Opens `file` and reads its sheets as `plan` says: one after another with a turn of the event loop
 * after each, as a caller that does other work between them would, or all at once with a turn after
 * them. Throws whatever is thrown that is not an `XlsxReadError`.
 */
export async function outcomeOf(
  file: Uint8Array | Blob,
  plan: Plan = "each",
  limits?: ReadLimits,
): Promise<Outcome> {
  let book: OpenedWorkbook;
  try {
    book = await openXlsx(file, limits === undefined ? undefined : { limits });
  } catch (error) {
    return { refused: refusalOf(error) };
  }
  const read = new Map<string, SheetOutcome>();
  if (plan === "together") {
    for (const sheet of await Promise.all(book.sheets.map((name) => sheetOf(book, name)))) {
      read.set(sheet.name, sheet);
    }
  } else {
    for (const name of plan === "each" ? book.sheets : plan) {
      read.set(name, await sheetOf(book, name));
      await turn();
    }
  }
  await turn();
  return { sheets: book.sheets.map((name) => read.get(name) ?? { name }) };
}

/** A value written short enough for a log: long text by its ends, its length and its CRC-32. */
export function render(value: unknown): string {
  if (value instanceof Date) return `Date(${value.toISOString()})`;
  if (typeof value === "number") return Object.is(value, -0) ? "-0" : String(value);
  if (typeof value === "string") {
    if (value.length <= 80) return JSON.stringify(value);
    const digest = crc32(value).toString(16).padStart(8, "0");
    return `${JSON.stringify(value.slice(0, 40))}…${JSON.stringify(value.slice(-20))} (${value.length} characters, CRC-32 ${digest})`;
  }
  if (Array.isArray(value)) {
    const shown = value.length > 12 ? [...value.slice(0, 6), "…", ...value.slice(-5)] : value;
    const items = shown.map((item: unknown) => (item === "…" ? "…" : render(item)));
    return `[${items.join(", ")}]${value.length > 12 ? ` (${value.length} items)` : ""}`;
  }
  if (value !== null && typeof value === "object") {
    const fields = Object.entries(value).map(([key, field]) => `${key}: ${render(field)}`);
    return `{ ${fields.join(", ")} }`;
  }
  return String(value);
}

/** Sends one message to the process that forked this one, and settles once it has left. */
export function send(message: unknown): Promise<void> {
  return new Promise((resolve, reject) => {
    if (process.send === undefined) reject(new Error("not forked with a channel"));
    else process.send(message, undefined, {}, (error) => (error ? reject(error) : resolve()));
  });
}
