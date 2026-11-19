import { type Meter, type PartMeter, XlsxReadError } from "../../errors/read.js";
import {
  CENTRAL_HEADER_SIGNATURE,
  CENTRAL_HEADER_SIZE,
  END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  END_OF_CENTRAL_DIRECTORY_SIZE,
  LOCAL_HEADER_SIGNATURE,
  LOCAL_HEADER_SIZE,
  MAX_32_BIT,
} from "../lib/zip-records.js";
import {
  COMPOUND_FILE_SIGNATURE,
  FLAG_ENCRYPTED,
  MAX_COMMENT,
  ZIP64_END_RECORD_SIGNATURE,
  ZIP64_END_RECORD_SIZE,
  ZIP64_EXTRA_FIELD,
  ZIP64_LOCATOR_SIGNATURE,
  ZIP64_LOCATOR_SIZE,
  foldName,
  names,
  refuse,
  u16,
  u32,
  u64,
} from "./layout.js";
import type { ByteSource } from "./source.js";

/** What the directory says about one part. The only truth about it: a local header is checked against this. */
export interface DirectoryEntry {
  /** The name as written, which the local header must repeat. */
  readonly name: string;
  /** The length of the name in bytes. */
  readonly nameLength: number;
  /** Where the local header starts in the file. */
  readonly offset: number;
  readonly compressedSize: number;
  /** The declared inflated size: what is charged, and what the content must amount to exactly. */
  readonly size: number;
  readonly crc32: number;
  readonly method: number;
  readonly encrypted: boolean;
}

export interface Directory {
  /** Every part by its folded name. A `Map`, so a name from the file never reaches a prototype. */
  readonly entries: ReadonlyMap<string, DirectoryEntry>;
  /** Where the directory starts: no part's data may reach past it. */
  readonly start: number;
}

/** The end of a file as far back as an end record can start, with a zip64 one in front of it. */
const TAIL_BYTES =
  ZIP64_END_RECORD_SIZE + ZIP64_LOCATOR_SIZE + END_OF_CENTRAL_DIRECTORY_SIZE + MAX_COMMENT;

/** Where the end of the file says the directory is. */
interface Extent {
  /** Where the directory ends: at the end record, or at the zip64 one in front of it. */
  readonly end: number;
  readonly size: number;
  /** Where the directory starts, as stated. It must be `end - size`. */
  readonly offset: number;
}

/**
 * Finds the end record, reads the directory and range-checks every entry, the parts nobody will
 * read included.
 *
 * The search for the end record is bounded by the longest comment. The directory's size is
 * charged before the directory is read, and every entry before it is kept. The walk ignores the
 * number of entries the end record states, stops at the directory's own size and moves forward on
 * every step or refuses. Nothing is allocated from a count or a size the file states.
 */
export async function readDirectory(
  source: ByteSource,
  parts: Meter<"parts">,
  structure: PartMeter<"structure">,
): Promise<Directory> {
  const extent = await locate(source);
  const start = extent.end - extent.size;
  if (start !== extent.offset) refuse("the directory is not where the package says");
  structure.charge(extent.size);
  const directory = await source.read(start, extent.end);

  const entries = new Map<string, DirectoryEntry>();
  for (let at = 0; at < extent.size;) {
    const nameStart = at + CENTRAL_HEADER_SIZE;
    const nameEnd = nameStart + u16(directory, at + 28);
    const extraEnd = nameEnd + u16(directory, at + 30);
    const next = extraEnd + u16(directory, at + 32);
    if (u32(directory, at) !== CENTRAL_HEADER_SIGNATURE || next > extent.size) {
      refuse("the directory of the package is damaged");
    }
    parts.charge(1);

    let compressedSize = u32(directory, at + 20);
    let size = u32(directory, at + 24);
    let offset = u32(directory, at + 42);
    if (size === MAX_32_BIT || compressedSize === MAX_32_BIT || offset === MAX_32_BIT) {
      let field = nameEnd;
      while (field < extraEnd && u16(directory, field) !== ZIP64_EXTRA_FIELD) {
        field += 4 + u16(directory, field + 2);
      }
      if (field < extraEnd) {
        // The field holds a value for each 32-bit field that is all ones, in this order.
        let value = field + 4;
        if (size === MAX_32_BIT) {
          size = u64(directory, value);
          value += 8;
        }
        if (compressedSize === MAX_32_BIT) {
          compressedSize = u64(directory, value);
          value += 8;
        }
        if (offset === MAX_32_BIT) {
          offset = u64(directory, value);
          value += 8;
        }
        if (value > Math.min(field + 4 + u16(directory, field + 2), extraEnd)) {
          refuse("the directory of the package is damaged");
        }
      }
    }

    const name = names.decode(directory.subarray(nameStart, nameEnd));
    const key = foldName(name);
    if (entries.has(key)) refuse("two parts have one name");
    const nameLength = nameEnd - nameStart;
    if (offset + LOCAL_HEADER_SIZE + nameLength + compressedSize > start) {
      refuse("a part lies outside the package");
    }
    entries.set(key, {
      name,
      nameLength,
      offset,
      compressedSize,
      size,
      crc32: u32(directory, at + 16),
      method: u16(directory, at + 10),
      encrypted: (u16(directory, at + 8) & FLAG_ENCRYPTED) !== 0,
    });
    at = next;
  }
  return { entries, start };
}

async function locate(source: ByteSource): Promise<Extent> {
  const tailStart = Math.max(0, source.size - TAIL_BYTES);
  const tail = await source.read(tailStart, source.size);
  const at = endRecordIn(tail);
  if (at === undefined) return refuseEnd(source);

  const size = u32(tail, at + 12);
  const offset = u32(tail, at + 16);
  const locator = at - ZIP64_LOCATOR_SIZE;
  if (u32(tail, locator) !== ZIP64_LOCATOR_SIGNATURE) return { end: tailStart + at, size, offset };

  const record = locator - ZIP64_END_RECORD_SIZE;
  const extent = {
    end: tailStart + record,
    size: u64(tail, record + 40),
    offset: u64(tail, record + 48),
  };
  // The end record states the same size and place, or leaves each to the zip64 one with all ones.
  if (
    u32(tail, record) !== ZIP64_END_RECORD_SIGNATURE ||
    u64(tail, locator + 8) !== extent.end ||
    (size !== MAX_32_BIT && size !== extent.size) ||
    (offset !== MAX_32_BIT && offset !== extent.offset)
  ) {
    refuse("the zip64 end of the package disagrees with itself");
  }
  return extent;
}

/**
 * Where the end record starts in the end of a file, or `undefined` when the file does not end in
 * one end record. The last end signature that a record fits behind must be the record, with a
 * comment that ends the file, and no record before it may end the file too: an end that reads two
 * ways is refused, not chosen from.
 */
function endRecordIn(tail: Uint8Array): number | undefined {
  let found: number | undefined;
  for (let at = tail.length - END_OF_CENTRAL_DIRECTORY_SIZE; at >= 0; at--) {
    if (u32(tail, at) !== END_OF_CENTRAL_DIRECTORY_SIGNATURE) continue;
    const endsFile = at + END_OF_CENTRAL_DIRECTORY_SIZE + u16(tail, at + 20) === tail.length;
    if (found === undefined) {
      if (!endsFile) return undefined;
      found = at;
    } else if (endsFile) {
      return undefined;
    }
  }
  return found;
}

async function refuseEnd(source: ByteSource): Promise<never> {
  const signature = u32(await source.read(0, 4), 0);
  if (signature === LOCAL_HEADER_SIGNATURE) {
    refuse("the package does not end in one end record");
  }
  throw new XlsxReadError(
    "NOT_XLSX",
    signature === COMPOUND_FILE_SIGNATURE
      ? "a compound file such as an old .xls or an encrypted workbook"
      : "not an .xlsx file",
  );
}
