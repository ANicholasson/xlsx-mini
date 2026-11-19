// What the tests of the read side of XML share: a part as chunks, and its tokens as values to compare.
import { type ByteChunks, type XmlCursor, type XmlVisitor, readXml } from "../read.js";

/** An open tag with the attributes that were asked for and found, or a close tag with its text. */
export type Token =
  | readonly ["open", name: string, depth: number, attributes: Readonly<Record<string, string>>]
  | readonly ["close", name: string, depth: number, text: string];

const encoder = new TextEncoder();

export function utf8(text: string): Uint8Array {
  return encoder.encode(text);
}

/** A part given in chunks that end at `cuts`, which counts what is asked of it. */
export interface Chunked extends ByteChunks {
  readonly reads: number;
  readonly closes: number;
}

export function chunked(part: string | Uint8Array, cuts: readonly number[] = []): Chunked {
  const bytes = typeof part === "string" ? utf8(part) : part;
  const ends = [...cuts.filter((cut) => cut <= bytes.length), bytes.length];
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
      const chunk = bytes.subarray(from, end);
      from = end;
      return Promise.resolve(chunk);
    },
    close() {
      closes++;
    },
  };
}

/** Collects every token, asking each open tag for the attributes in `asked`. */
export function collector(asked: readonly string[] = []): XmlVisitor<never, Token[]> {
  const tokens: Token[] = [];
  return {
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

/** The tokens of a part, or the code it is refused with. */
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
