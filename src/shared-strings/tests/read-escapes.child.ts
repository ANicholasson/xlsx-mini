// Run by read-memory.test.ts in a process of its own, with a small heap. Prints one line of JSON, an
// `EscapesReport`.
import { decodeEscapes } from "../read.js";

interface Decoded {
  readonly name: string;
  readonly characters: number;
  /** Characters of the result other than the ones the escapes name, at their place. */
  readonly wrong: number;
}

export interface EscapesReport {
  readonly tokenCharacters: number;
  readonly decoded: readonly Decoded[];
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);

const TOKEN_CHARACTERS = 4_194_304;

const cases: [name: string, unit: string, read: string][] = [
  ["letters", "A", "A"],
  ["escapes of a letter", "_x0041_", "A"],
  ["escapes of a letter outside Latin-1", "_x0100_", "Ā"],
  ["escapes of a surrogate pair", "_xD83D__xDE00_", "😀"],
];

const decoded: Decoded[] = [];
for (const [name, unit, read] of cases) {
  collect();
  const text = decodeEscapes(unit.repeat(Math.floor((TOKEN_CHARACTERS - 64) / unit.length)));
  let wrong = 0;
  for (let at = 0; at < text.length; at++) {
    if (text.charCodeAt(at) !== read.charCodeAt(at % read.length)) wrong++;
  }
  decoded.push({ name, characters: text.length, wrong });
}

const report: EscapesReport = { tokenCharacters: TOKEN_CHARACTERS, decoded };
console.log(JSON.stringify(report));
