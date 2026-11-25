// node test/corpus/anonymize.ts FILE...
//
// Blanks the two fields of a workbook Excel saved that name a person and a home folder: the last
// author in docProps/core.xml and the folder in xl/workbook.xml that Excel recorded the file was
// saved in.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { crc32, deflateRawSync, inflateRawSync } from "node:zlib";

const LOCAL_HEADER = 0x04034b50;
const CENTRAL_HEADER = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const ZIP64_LOCATOR = 0x07064b50;

export const NEUTRAL_PATH = "/corpus/excel/";

interface Edit {
  readonly part: string;
  readonly pattern: RegExp;
  readonly blank: string;
}

const EDITS: readonly Edit[] = [
  {
    part: "docProps/core.xml",
    pattern: /(<cp:lastModifiedBy>)([^<]+)(<\/cp:lastModifiedBy>)/g,
    blank: "",
  },
  { part: "xl/workbook.xml", pattern: /(<\w+:absPath url=")([^"]*)(")/g, blank: NEUTRAL_PATH },
];

interface Entry {
  readonly name: string;
  readonly central: number;
  readonly centralLength: number;
  readonly local: number;
  readonly method: number;
  readonly flags: number;
  readonly compressedSize: number;
}

function fail(message: string): never {
  throw new Error(message);
}

function directory(file: Uint8Array): { entries: Entry[]; start: number; end: number } {
  const view = new DataView(file.buffer, file.byteOffset, file.byteLength);
  let end = file.byteLength - 22;
  while (end >= 0 && view.getUint32(end, true) !== END_OF_CENTRAL_DIRECTORY) end--;
  if (end < 0) fail("no end of central directory");
  if (end >= 20 && view.getUint32(end - 20, true) === ZIP64_LOCATOR) fail("a zip64 package");
  const count = view.getUint16(end + 10, true);
  const start = view.getUint32(end + 16, true);
  if (count === 0xffff || start === 0xffffffff) fail("a zip64 package");
  const entries: Entry[] = [];
  let at = start;
  for (let index = 0; index < count; index++) {
    if (view.getUint32(at, true) !== CENTRAL_HEADER) fail(`no central header at ${at}`);
    const nameLength = view.getUint16(at + 28, true);
    const centralLength =
      46 + nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    entries.push({
      name: new TextDecoder("utf-8", { fatal: true }).decode(
        file.subarray(at + 46, at + 46 + nameLength),
      ),
      central: at,
      centralLength,
      local: view.getUint32(at + 42, true),
      method: view.getUint16(at + 10, true),
      flags: view.getUint16(at + 8, true),
      compressedSize: view.getUint32(at + 20, true),
    });
    at += centralLength;
  }
  return { entries, start, end };
}

function dataStart(file: Uint8Array, entry: Entry): number {
  const view = new DataView(file.buffer, file.byteOffset, file.byteLength);
  if (view.getUint32(entry.local, true) !== LOCAL_HEADER) fail(`no local header for ${entry.name}`);
  const nameLength = view.getUint16(entry.local + 26, true);
  return entry.local + 30 + nameLength + view.getUint16(entry.local + 28, true);
}

function contentOf(file: Uint8Array, entry: Entry): Uint8Array {
  const start = dataStart(file, entry);
  const stored = file.subarray(start, start + entry.compressedSize);
  if (entry.method === 0) return stored;
  if (entry.method === 8) return new Uint8Array(inflateRawSync(stored));
  return fail(`method ${entry.method} for ${entry.name}`);
}

export function replaceEntries(
  file: Uint8Array,
  contents: ReadonlyMap<string, Uint8Array>,
): Uint8Array {
  const { entries, start, end } = directory(file);
  const packed = new Map<Entry, { data: Uint8Array; crc: number; size: number }>();
  for (const entry of entries) {
    const content = contents.get(entry.name);
    if (content === undefined) continue;
    if ((entry.flags & 0b1001) !== 0) fail(`${entry.name} is encrypted or has a data descriptor`);
    if (entry.method !== 0 && entry.method !== 8) fail(`method ${entry.method} for ${entry.name}`);
    const data = entry.method === 8 ? new Uint8Array(deflateRawSync(content)) : content;
    packed.set(entry, { data, crc: crc32(content), size: content.byteLength });
  }
  const byOffset = [...entries].sort((a, b) => a.local - b.local);
  const pieces: Uint8Array[] = [file.subarray(0, byOffset[0]?.local ?? start)];
  const moved = new Map<Entry, number>();
  let offset = pieces[0]!.byteLength;
  for (const [index, entry] of byOffset.entries()) {
    const next = byOffset[index + 1]?.local ?? start;
    const replacement = packed.get(entry);
    moved.set(entry, offset);
    if (replacement === undefined) {
      pieces.push(file.subarray(entry.local, next));
      offset += next - entry.local;
      continue;
    }
    const begin = dataStart(file, entry);
    const header = file.slice(entry.local, begin);
    const view = new DataView(header.buffer);
    view.setUint32(14, replacement.crc, true);
    view.setUint32(18, replacement.data.byteLength, true);
    view.setUint32(22, replacement.size, true);
    const gap = file.subarray(begin + entry.compressedSize, next);
    pieces.push(header, replacement.data, gap);
    offset += header.byteLength + replacement.data.byteLength + gap.byteLength;
  }
  const centralStart = offset;
  for (const entry of entries) {
    const record = file.slice(entry.central, entry.central + entry.centralLength);
    const view = new DataView(record.buffer);
    view.setUint32(42, moved.get(entry)!, true);
    const replacement = packed.get(entry);
    if (replacement !== undefined) {
      view.setUint32(16, replacement.crc, true);
      view.setUint32(20, replacement.data.byteLength, true);
      view.setUint32(24, replacement.size, true);
    }
    pieces.push(record);
  }
  const tail = file.slice(end);
  new DataView(tail.buffer).setUint32(16, centralStart, true);
  pieces.push(tail);
  const out = new Uint8Array(pieces.reduce((sum, piece) => sum + piece.byteLength, 0));
  let at = 0;
  for (const piece of pieces) {
    out.set(piece, at);
    at += piece.byteLength;
  }
  return out;
}

export function anonymize(file: Uint8Array): Uint8Array | undefined {
  const { entries } = directory(file);
  const contents = new Map<string, Uint8Array>();
  for (const { part, pattern, blank } of EDITS) {
    const entry = entries.find(({ name }) => name === part);
    if (entry === undefined) continue;
    const xml = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(
      contentOf(file, entry),
    );
    const found = [...xml.matchAll(pattern)];
    if (found.length > 1) fail(`${part} states the field ${found.length} times`);
    if (found.length === 0 || found[0]![2] === blank) continue;
    contents.set(part, new TextEncoder().encode(xml.replace(pattern, `$1${blank}$3`)));
  }
  return contents.size === 0 ? undefined : replaceEntries(file, contents);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  for (const path of process.argv.slice(2)) {
    const changed = anonymize(new Uint8Array(readFileSync(path)));
    if (changed !== undefined) writeFileSync(path, changed);
    console.log(`${path}: ${changed === undefined ? "nothing to blank" : "blanked"}`);
  }
}
