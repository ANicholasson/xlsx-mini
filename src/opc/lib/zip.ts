import { XlsxMiniError } from "../../errors/index.js";
import type { Compression } from "./compression.js";
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
/** 1980-01-01 00:00:00, the DOS epoch. Fixed so identical input gives identical bytes. */
const DOS_TIME = 0x0000;
const DOS_DATE = 0x0021;
const MAX_32_BIT = 0xffffffff;

const encoder = new TextEncoder();

/**
 * Packs parts into a zip archive in the order given, with a central directory and no data
 * descriptors. Parts are compressed up front so every size and offset is known, and the archive
 * is then written into one preallocated buffer.
 */
export async function createPackage(
  parts: readonly PackagePart[],
  compression: Compression,
): Promise<Uint8Array<ArrayBuffer>> {
  const count = parts.length;
  const names: Uint8Array[] = [];
  const payloads: Uint8Array[] = [];
  const checksums = new Uint32Array(count);
  let localBytes = 0;
  let centralBytes = 0;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = encoder.encode(part.path);
    // Sequential on purpose: one part in flight keeps peak memory at one compressed copy.
    const payload = await compression.compress(part.data);
    names.push(name);
    payloads.push(payload);
    checksums[i] = crc32(part.data);
    localBytes += LOCAL_HEADER_SIZE + name.length + payload.length;
    centralBytes += CENTRAL_HEADER_SIZE + name.length;
  }
  const total = localBytes + centralBytes + END_OF_CENTRAL_DIRECTORY_SIZE;
  if (total > MAX_32_BIT) {
    throw new XlsxMiniError(
      "PACKAGE_TOO_LARGE",
      "workbook",
      "The package exceeds the 4 GiB limit of a plain zip archive",
    );
  }

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  const offsets = new Uint32Array(count);
  const method = compression.method;

  let cursor = 0;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = names[i]!;
    const payload = payloads[i]!;
    offsets[i] = cursor;
    view.setUint32(cursor, LOCAL_HEADER_SIGNATURE, true);
    view.setUint16(cursor + 4, VERSION_2_0, true);
    view.setUint16(cursor + 6, 0, true); // general purpose flags
    view.setUint16(cursor + 8, method, true);
    view.setUint16(cursor + 10, DOS_TIME, true);
    view.setUint16(cursor + 12, DOS_DATE, true);
    view.setUint32(cursor + 14, checksums[i]!, true);
    view.setUint32(cursor + 18, payload.length, true);
    view.setUint32(cursor + 22, part.data.length, true);
    view.setUint16(cursor + 26, name.length, true);
    view.setUint16(cursor + 28, 0, true); // extra field length
    cursor += LOCAL_HEADER_SIZE;
    out.set(name, cursor);
    cursor += name.length;
    out.set(payload, cursor);
    cursor += payload.length;
  }

  const centralDirectoryOffset = cursor;
  for (let i = 0; i < count; i++) {
    const part = parts[i]!;
    const name = names[i]!;
    view.setUint32(cursor, CENTRAL_HEADER_SIGNATURE, true);
    view.setUint16(cursor + 4, VERSION_2_0, true); // version made by
    view.setUint16(cursor + 6, VERSION_2_0, true); // version needed
    view.setUint16(cursor + 8, 0, true); // general purpose flags
    view.setUint16(cursor + 10, method, true);
    view.setUint16(cursor + 12, DOS_TIME, true);
    view.setUint16(cursor + 14, DOS_DATE, true);
    view.setUint32(cursor + 16, checksums[i]!, true);
    view.setUint32(cursor + 20, payloads[i]!.length, true);
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
