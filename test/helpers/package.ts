import { constants, crc32, deflateRawSync } from "node:zlib";

/**
 * One part as `packageOf` writes it. Without the optional fields its records are what any zip
 * writer would write; each optional field writes one of them another way, true or not.
 */
export interface Part {
  /** The part's name, in the directory and, unless `localPath` says otherwise, in its local header. */
  readonly path: string;
  /**
   * The content, as text, bytes, or chunks of bytes. The deflate of a chunk is made once however
   * often the chunk repeats, so a part of a gibibyte of one repeated chunk is quick to build.
   */
  readonly data: string | Uint8Array | readonly Uint8Array[];
  /** The compression method the records state: 8 deflates the content, any other stores it. */
  readonly method?: number;
  /** The general purpose flags of both records, in place of the ones a writer sets. */
  readonly flags?: number;
  /** The inflated size both records state, in place of the true one. */
  readonly size?: number;
  /** The compressed size the directory states, in place of the true one. */
  readonly compressedSize?: number;
  /** The CRC-32 both records state, in place of the true one. */
  readonly crc32?: number;
  /** The name the local header gives, in place of `path`. */
  readonly localPath?: string;
  /** The checksum and sizes after the data, in a data descriptor, and zeros in the local header. */
  readonly descriptor?: boolean;
  /** The sizes in a zip64 extra field of both records, and all ones in their 32-bit fields. */
  readonly zip64?: boolean;
  /** An extra field for both records. */
  readonly extra?: Uint8Array;
  /** A directory entry over the data of the part named so, written before it: no data of its own. */
  readonly sameDataAs?: string;
}

/** Where the records of the package stand, and what they state about one another. */
export interface Layout {
  /** Bytes in front of the first local header. */
  readonly before?: Uint8Array;
  /** Whether the offsets count the bytes in front, as they do when they were written as a part. */
  readonly offsetsCountBefore?: boolean;
  /** Bytes after the end record. */
  readonly after?: Uint8Array;
  readonly comment?: Uint8Array;
  /** The number of entries the end record states, in place of the true one. */
  readonly entries?: number;
  /** Where the end record says the directory starts, in place of the true place. */
  readonly directoryOffset?: number;
  /** A zip64 end record and its locator, with all ones in the end record's fields. */
  readonly zip64?: boolean;
}

const LOCAL = 0x04034b50;
const DESCRIPTOR = 0x08074b50;
const CENTRAL = 0x02014b50;
const END = 0x06054b50;
const ZIP64_END = 0x06064b50;
const ZIP64_LOCATOR = 0x07064b50;
const ALL_ONES = 0xffffffff;
const FINAL_EMPTY_BLOCK = deflateRawSync(new Uint8Array(0));

const encoder = new TextEncoder();
export const bytesOf = (text: string): Uint8Array<ArrayBuffer> => encoder.encode(text);

/** What the directory says about a part once its data is written. */
interface Written {
  readonly part: Part;
  readonly offset: number;
  readonly crc32: number;
  readonly size: number;
  readonly compressedSize: number;
}

interface Output {
  bytes: Uint8Array<ArrayBuffer>;
  view: DataView;
  length: number;
}

function reserve(output: Output, length: number): number {
  const at = output.length;
  if (at + length > output.bytes.length) {
    const grown = new Uint8Array(Math.max(output.bytes.length * 2, at + length));
    grown.set(output.bytes.subarray(0, at));
    output.bytes = grown;
    output.view = new DataView(grown.buffer);
  }
  output.length += length;
  return at;
}

function write(output: Output, bytes: Uint8Array): void {
  const at = reserve(output, bytes.length);
  output.bytes.set(bytes, at);
}

const NO_BYTES = new Uint8Array(0);

/** The zip64 extra field of a part's records, when it has one: its sizes, in this order. */
function writeZip64Extra(output: Output, at: number, size: number, compressedSize: number): void {
  output.view.setUint16(at, 1, true);
  output.view.setUint16(at + 2, 16, true);
  output.view.setBigUint64(at + 4, BigInt(size), true);
  output.view.setBigUint64(at + 12, BigInt(compressedSize), true);
}

const extraLength = (part: Part) => (part.zip64 ? 20 : 0) + (part.extra?.length ?? 0);

function writeExtra(output: Output, at: number, entry: Written): void {
  if (entry.part.zip64) writeZip64Extra(output, at, entry.size, entry.compressedSize);
  output.bytes.set(entry.part.extra ?? NO_BYTES, at + (entry.part.zip64 ? 20 : 0));
}

/**
 * The bytes of a zip package of `parts`, in order, written as `layout` says. Only what a field
 * names is changed; every other record states the truth about what was written, so a lie is the
 * only thing wrong with the package.
 */
export function packageOf(parts: readonly Part[], layout: Layout = {}): Uint8Array<ArrayBuffer> {
  const first = new Uint8Array(65_536);
  const output: Output = { bytes: first, view: new DataView(first.buffer), length: 0 };
  const front = layout.before ?? NO_BYTES;
  write(output, front);
  const base = layout.offsetsCountBefore ? 0 : front.length;
  const deflated = new Map<Uint8Array, Uint8Array>();
  const written: Written[] = [];
  for (const part of parts) {
    if (part.sameDataAs !== undefined) {
      const same = written.find((entry) => entry.part.path === part.sameDataAs);
      if (same === undefined) throw new Error(`no part ${part.sameDataAs} to share the data of`);
      written.push({ ...same, part });
      continue;
    }
    const method = part.method ?? 8;
    const local = bytesOf(part.localPath ?? part.path);
    const header = reserve(output, 30 + local.length + extraLength(part));
    const chunks =
      typeof part.data === "string"
        ? [bytesOf(part.data)]
        : part.data instanceof Uint8Array
          ? [part.data]
          : part.data;
    let checksum = 0;
    let size = 0;
    for (const chunk of chunks) {
      checksum = crc32(chunk, checksum);
      size += chunk.length;
      if (method !== 8) {
        write(output, chunk);
        continue;
      }
      let compressed = deflated.get(chunk);
      if (compressed === undefined) {
        compressed = deflateRawSync(chunk, { finishFlush: constants.Z_FULL_FLUSH });
        deflated.set(chunk, compressed);
      }
      write(output, compressed);
    }
    if (method === 8) write(output, FINAL_EMPTY_BLOCK);
    const entry: Written = {
      part,
      offset: header - base,
      crc32: part.crc32 ?? checksum,
      size: part.size ?? size,
      compressedSize: output.length - header - 30 - local.length - extraLength(part),
    };
    written.push(entry);
    const view = output.view;
    view.setUint32(header, LOCAL, true);
    view.setUint16(header + 4, part.zip64 ? 45 : 20, true);
    view.setUint16(header + 6, flagsOf(part), true);
    view.setUint16(header + 8, method, true);
    view.setUint16(header + 12, 0x21, true);
    if (!part.descriptor) {
      view.setUint32(header + 14, entry.crc32, true);
      view.setUint32(header + 18, part.zip64 ? ALL_ONES : entry.compressedSize, true);
      view.setUint32(header + 22, part.zip64 ? ALL_ONES : entry.size, true);
    }
    view.setUint16(header + 26, local.length, true);
    view.setUint16(header + 28, extraLength(part), true);
    output.bytes.set(local, header + 30);
    writeExtra(output, header + 30 + local.length, entry);
    if (part.descriptor) {
      const at = reserve(output, 16);
      output.view.setUint32(at, DESCRIPTOR, true);
      output.view.setUint32(at + 4, entry.crc32, true);
      output.view.setUint32(at + 8, entry.compressedSize, true);
      output.view.setUint32(at + 12, entry.size, true);
    }
  }

  const directoryStart = output.length - base;
  for (const entry of written) {
    const name = bytesOf(entry.part.path);
    const at = reserve(output, 46 + name.length + extraLength(entry.part));
    const view = output.view;
    const zip64 = entry.part.zip64 === true;
    view.setUint32(at, CENTRAL, true);
    view.setUint16(at + 4, zip64 ? 45 : 20, true);
    view.setUint16(at + 6, zip64 ? 45 : 20, true);
    view.setUint16(at + 8, flagsOf(entry.part), true);
    view.setUint16(at + 10, entry.part.method ?? 8, true);
    view.setUint16(at + 14, 0x21, true);
    view.setUint32(at + 16, entry.crc32, true);
    view.setUint32(
      at + 20,
      zip64 ? ALL_ONES : (entry.part.compressedSize ?? entry.compressedSize),
      true,
    );
    view.setUint32(at + 24, zip64 ? ALL_ONES : entry.size, true);
    view.setUint16(at + 28, name.length, true);
    view.setUint16(at + 30, extraLength(entry.part), true);
    view.setUint32(at + 42, entry.offset, true);
    output.bytes.set(name, at + 46);
    writeExtra(output, at + 46 + name.length, entry);
  }
  const directorySize = output.length - base - directoryStart;
  const count = layout.entries ?? written.length;
  const offset = layout.directoryOffset ?? directoryStart;
  if (layout.zip64) {
    const record = reserve(output, 56);
    const view = output.view;
    view.setUint32(record, ZIP64_END, true);
    view.setBigUint64(record + 4, 44n, true);
    view.setUint16(record + 12, 45, true);
    view.setUint16(record + 14, 45, true);
    view.setBigUint64(record + 24, BigInt(count), true);
    view.setBigUint64(record + 32, BigInt(count), true);
    view.setBigUint64(record + 40, BigInt(directorySize), true);
    view.setBigUint64(record + 48, BigInt(offset), true);
    const locator = reserve(output, 20);
    output.view.setUint32(locator, ZIP64_LOCATOR, true);
    output.view.setBigUint64(locator + 8, BigInt(record - base), true);
    output.view.setUint32(locator + 16, 1, true);
  }
  const comment = layout.comment ?? NO_BYTES;
  const end = reserve(output, 22);
  const view = output.view;
  view.setUint32(end, END, true);
  view.setUint16(end + 8, layout.zip64 ? 0xffff : count & 0xffff, true);
  view.setUint16(end + 10, layout.zip64 ? 0xffff : count & 0xffff, true);
  view.setUint32(end + 12, layout.zip64 ? ALL_ONES : directorySize, true);
  view.setUint32(end + 16, layout.zip64 ? ALL_ONES : offset, true);
  view.setUint16(end + 20, comment.length, true);
  write(output, comment);
  write(output, layout.after ?? NO_BYTES);
  return output.bytes.slice(0, output.length);
}

/** The flags a writer sets: a name outside ASCII in UTF-8, and sizes after the data in a descriptor. */
function flagsOf(part: Part): number {
  const name = part.localPath ?? part.path;
  const utf8 = bytesOf(name).length === name.length ? 0 : 0x800;
  return part.flags ?? utf8 | (part.descriptor ? 0x08 : 0);
}
