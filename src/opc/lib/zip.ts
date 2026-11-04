import { crc32 } from "./crc32.js";

/** A file inside the package: its zip entry name (forward slashes, no leading slash) and bytes. */
export interface PackagePart {
  readonly path: string;
  readonly data: Uint8Array;
}

const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const LOCAL_HEADER_SIZE = 30;
const CENTRAL_HEADER_SIZE = 46;
const END_OF_CENTRAL_DIRECTORY_SIZE = 22;
const VERSION_2_0 = 20;
const METHOD_STORE = 0;
/** 1980-01-01 00:00:00, the DOS epoch. Fixed so identical input gives identical bytes. */
const DOS_TIME = 0x0000;
const DOS_DATE = 0x0021;
const MAX_32_BIT = 0xffffffff;

const encoder = new TextEncoder();

/**
 * Packs parts into a zip archive using STORE entries, in the order given, with a central
 * directory and no data descriptors. Every size and offset is known up front, so the
 * archive is written into one preallocated buffer.
 */
export function createPackage(parts: readonly PackagePart[]): Uint8Array {
  const count = parts.length;
  const names: Uint8Array[] = [];
  let localBytes = 0;
  let centralBytes = 0;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = encoder.encode(part.path);
    names.push(name);
    localBytes += LOCAL_HEADER_SIZE + name.length + part.data.length;
    centralBytes += CENTRAL_HEADER_SIZE + name.length;
  }
  const total = localBytes + centralBytes + END_OF_CENTRAL_DIRECTORY_SIZE;
  if (total > MAX_32_BIT) {
    throw new RangeError("package exceeds the 4 GiB limit of a plain zip archive");
  }

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  const checksums = new Uint32Array(count);
  const offsets = new Uint32Array(count);

  let cursor = 0;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = names[i]!;
    const checksum = crc32(part.data);
    checksums[i] = checksum;
    offsets[i] = cursor;
    view.setUint32(cursor, LOCAL_HEADER_SIGNATURE, true);
    view.setUint16(cursor + 4, VERSION_2_0, true);
    view.setUint16(cursor + 6, 0, true); // general purpose flags
    view.setUint16(cursor + 8, METHOD_STORE, true);
    view.setUint16(cursor + 10, DOS_TIME, true);
    view.setUint16(cursor + 12, DOS_DATE, true);
    view.setUint32(cursor + 14, checksum, true);
    view.setUint32(cursor + 18, part.data.length, true);
    view.setUint32(cursor + 22, part.data.length, true);
    view.setUint16(cursor + 26, name.length, true);
    view.setUint16(cursor + 28, 0, true); // extra field length
    cursor += LOCAL_HEADER_SIZE;
    out.set(name, cursor);
    cursor += name.length;
    out.set(part.data, cursor);
    cursor += part.data.length;
  }

  const centralDirectoryOffset = cursor;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = names[i]!;
    view.setUint32(cursor, CENTRAL_HEADER_SIGNATURE, true);
    view.setUint16(cursor + 4, VERSION_2_0, true); // version made by
    view.setUint16(cursor + 6, VERSION_2_0, true); // version needed
    view.setUint16(cursor + 8, 0, true); // general purpose flags
    view.setUint16(cursor + 10, METHOD_STORE, true);
    view.setUint16(cursor + 12, DOS_TIME, true);
    view.setUint16(cursor + 14, DOS_DATE, true);
    view.setUint32(cursor + 16, checksums[i]!, true);
    view.setUint32(cursor + 20, part.data.length, true);
    view.setUint32(cursor + 24, part.data.length, true);
    view.setUint16(cursor + 28, name.length, true);
    view.setUint16(cursor + 30, 0, true); // extra field length
    view.setUint16(cursor + 32, 0, true); // comment length
    view.setUint16(cursor + 34, 0, true); // disk number start
    view.setUint16(cursor + 36, 0, true); // internal attributes
    view.setUint32(cursor + 38, 0, true); // external attributes
    view.setUint32(cursor + 42, offsets[i]!, true);
    cursor += CENTRAL_HEADER_SIZE;
    out.set(name, cursor);
    cursor += name.length;
  }

  view.setUint32(cursor, END_OF_CENTRAL_DIRECTORY_SIGNATURE, true);
  view.setUint16(cursor + 4, 0, true); // this disk
  view.setUint16(cursor + 6, 0, true); // disk with central directory
  view.setUint16(cursor + 8, count, true);
  view.setUint16(cursor + 10, count, true);
  view.setUint32(cursor + 12, centralBytes, true);
  view.setUint32(cursor + 16, centralDirectoryOffset, true);
  view.setUint16(cursor + 20, 0, true); // comment length
  return out;
}
