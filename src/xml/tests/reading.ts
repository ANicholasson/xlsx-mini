import { type ByteChunks, type XmlCursor, type XmlVisitor, readXml } from "../read.js";

export type Token =
  | readonly ["open", name: string, depth: number, attributes: Readonly<Record<string, string>>]
  | readonly ["close", name: string, depth: number, text: string];

const encoder = new TextEncoder();

export function utf8(text: string): Uint8Array {
  return encoder.encode(text);
}

export interface Chunked extends ByteChunks {
  readonly reads: number;
  readonly closes: number;
}

export function chunked(part: string | Uint8Array, cuts: readonly number[] = []): Chunked {
  const bytes = typeof part === "string" ? utf8(part) : part;
  const ends = [...cuts.filter((cut) => cut <= bytes.length), bytes.length];
  const reused = new Uint8Array(bytes.length);
  let last = 0;
  let from = 0;
  let reads = 0;
  let closes = 0;
  return {
    get reads() {
      return reads;
    },
    get closes() {
      return closes;
    },
    read() {
      const end = ends[reads++];
      if (end === undefined) return Promise.resolve(undefined);
      reused.fill(0xff, 0, last);
      const chunk = reused.subarray(0, end - from);
      chunk.set(bytes.subarray(from, end));
      last = chunk.length;
      from = end;
      return Promise.resolve(chunk);
    },
    close() {
      closes++;
    },
  };
}

/** `urn:one` has a second spelling, as SpreadsheetML has Transitional and Strict. */
export const NAMESPACES: ReadonlyMap<string, number> = new Map([
  ["urn:one", 1],
  ["urn:one:strict", 1],
  ["urn:two", 2],
]);

export function collector(asked: readonly string[] = []): XmlVisitor<never, Token[]> {
  const tokens: Token[] = [];
  return {
    namespaces: NAMESPACES,
    open(xml: XmlCursor) {
      const attributes: Record<string, string> = {};
      for (const name of asked) {
        const value = xml.attribute(name);
        if (value !== undefined) attributes[name] = value;
      }
      tokens.push(["open", xml.name, xml.depth, attributes]);
    },
    close(xml: XmlCursor) {
      tokens.push(["close", xml.name, xml.depth, xml.text()]);
      return undefined;
    },
    end: () => tokens,
  };
}

interface Reading {
  readonly cuts?: readonly number[];
  readonly asked?: readonly string[];
}

export function tokensOf(part: string | Uint8Array, reading: Reading = {}): Promise<Token[]> {
  return readXml(chunked(part, reading.cuts), collector(reading.asked));
}

export async function outcomeOf(
  part: string | Uint8Array,
  reading: Reading = {},
): Promise<Token[] | string> {
  try {
    return await tokensOf(part, reading);
  } catch (error) {
    return (error as { code: string }).code;
  }
}
