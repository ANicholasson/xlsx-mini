// Run by read-memory.test.ts in a process of its own, with a small heap. Prints one line of JSON, a
// `DecodeReport`.
import { type ByteChunks, readXml } from "../read.js";

export interface Decoded {
  readonly name: string;
  readonly outcome: string;
  readonly characters: number;
  /** Characters of the result other than the one every character should be. */
  readonly wrong: number;
}

export interface DecodeReport {
  readonly tokenCharacters: number;
  readonly decoded: readonly Decoded[];
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);
const encoder = new TextEncoder();

const TOKEN_CHARACTERS = 4_194_304;
const CHUNK_BYTES = 16_384;

/** `head`, then `unit` as often as one token may hold it, then `tail`, in chunks of 16 KiB. */
function part(head: string, unit: string, tail: string): ByteChunks {
  const perChunk = Math.floor(CHUNK_BYTES / unit.length);
  const body = encoder.encode(unit.repeat(perChunk));
  let left = Math.floor((TOKEN_CHARACTERS - 64) / unit.length);
  const ends = [encoder.encode(head), encoder.encode(tail)];
  return {
    read() {
      if (ends.length === 2) return Promise.resolve(ends.shift());
      if (left === 0) return Promise.resolve(ends.shift());
      const units = Math.min(left, perChunk);
      left -= units;
      return Promise.resolve(body.subarray(0, units * unit.length));
    },
    close() {},
  };
}

const cases: [name: string, head: string, unit: string, tail: string, read: string][] = [
  ["letters in a text", "<a>", "a", "</a>", "a"],
  ["carriage returns in a text", "<a>", "\r", "</a>", "\n"],
  ["CR LF pairs in a text", "<a>", "\r\n", "</a>", "\n"],
  ["carriage returns in a CDATA section", "<a><![CDATA[", "\r", "]]></a>", "\n"],
  ["character references in a text", "<a>", "&#13;", "</a>", "\r"],
  ["entity references in a text", "<a>", "&lt;", "</a>", "<"],
  ["tabs in an attribute value", '<a b="', "\t", '"/>', " "],
  ["line feeds in an attribute value", '<a b="', "\n", '"/>', " "],
  ["CR LF pairs in an attribute value", '<a b="', "\r\n", '"/>', " "],
  ["character references in an attribute value", '<a b="', "&#9;", '"/>', "\t"],
];

const decoded: Decoded[] = [];
for (const [name, head, unit, tail, read] of cases) {
  collect();
  let text = "";
  const outcome = await readXml(part(head, unit, tail), {
    open(xml) {
      text = xml.attribute("b") ?? "";
    },
    close(xml) {
      text ||= xml.text();
      return undefined;
    },
    end: () => "read",
  }).catch((error: unknown) => (error as { code: string }).code);
  let wrong = 0;
  for (let at = 0; at < text.length; at++) if (text[at] !== read) wrong++;
  decoded.push({ name, outcome, characters: text.length, wrong });
}

const report: DecodeReport = { tokenCharacters: TOKEN_CHARACTERS, decoded };
console.log(JSON.stringify(report));
