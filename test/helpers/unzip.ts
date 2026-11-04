import { inflateRawSync, crc32 as nodeCrc32 } from "node:zlib";

/** One entry read back from a zip archive, with its payload already decompressed. */
export interface ZipEntry {
  readonly name: string;
  readonly method: 0 | 8;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly localHeaderOffset: number;
  readonly data: Uint8Array;
}

const LOCAL_HEADER = 0x04034b50;
const CENTRAL_HEADER = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY = 0x06054b50;

/**
 * Minimal zip reader for tests. Independent of the library's writer: it walks the
 * central directory the way a real reader does, checks every signature and CRC with
 * Node's own crc32, and inflates DEFLATE entries with node:zlib.
 */
export function unzip(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEndOfCentralDirectory(view);
  const entryCount = view.getUint16(eocd + 10, true);
  const centralDirectorySize = view.getUint32(eocd + 12, true);
  const centralDirectoryOffset = view.getUint32(eocd + 16, true);
  if (centralDirectoryOffset + centralDirectorySize !== eocd) {
    throw new Error("central directory does not end at the end-of-central-directory record");
  }

  const entries: ZipEntry[] = [];
  let cursor = centralDirectoryOffset;
  for (let i = 0; i < entryCount; i++) {
    if (view.getUint32(cursor, true) !== CENTRAL_HEADER) {
      throw new Error(`bad central directory signature at ${cursor}`);
    }
    const method = view.getUint16(cursor + 10, true);
    const crc = view.getUint32(cursor + 16, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const uncompressedSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localHeaderOffset = view.getUint32(cursor + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
    cursor += 46 + nameLength + extraLength + commentLength;

    if (view.getUint32(localHeaderOffset, true) !== LOCAL_HEADER) {
      throw new Error(`bad local header signature for ${name}`);
    }
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
    if (method !== 0 && method !== 8) throw new Error(`unsupported method ${method} for ${name}`);
    const data = method === 8 ? new Uint8Array(inflateRawSync(compressed)) : compressed;
    if (data.byteLength !== uncompressedSize) {
      throw new Error(`size mismatch for ${name}: ${data.byteLength} != ${uncompressedSize}`);
    }
    if (nodeCrc32(data) !== crc) throw new Error(`crc mismatch for ${name}`);
    entries.push({
      name,
      method,
      crc32: crc,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
      data,
    });
  }
  return entries;
}

export function readPart(bytes: Uint8Array, name: string): string {
  const entry = unzip(bytes).find((e) => e.name === name);
  if (!entry) throw new Error(`part ${name} not found`);
  return new TextDecoder().decode(entry.data);
}

function findEndOfCentralDirectory(view: DataView): number {
  for (let i = view.byteLength - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === END_OF_CENTRAL_DIRECTORY) return i;
  }
  throw new Error("end-of-central-directory record not found");
}
