/**
 * What only reading needs to know of a package's layout: zip64, the compound-file signature,
 * the methods and the one flag. The signatures and record sizes that writing needs too are in
 * `lib/zip-records.ts`.
 */
import { XlsxReadError } from "../../errors/read.js";

export const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
export const ZIP64_END_RECORD_SIGNATURE = 0x06064b50;
/** The first four bytes of a compound file, which is what an old `.xls` and an encrypted workbook are. */
export const COMPOUND_FILE_SIGNATURE = 0xe011cfd0;

export const ZIP64_LOCATOR_SIZE = 20;
/** A zip64 end record without the extensible data no writer of workbooks adds. */
export const ZIP64_END_RECORD_SIZE = 56;
/** The longest comment an end record can carry, which bounds the search for the record. */
export const MAX_COMMENT = 0xffff;
/** The extra field that holds the real value of a 32-bit field set to `MAX_32_BIT`. */
export const ZIP64_EXTRA_FIELD = 1;

export const METHOD_STORED = 0;
export const METHOD_DEFLATED = 8;
export const FLAG_ENCRYPTED = 1;
/** Set in a header that was written before its data: the CRC-32 and the sizes follow the data. */
export const FLAG_DATA_DESCRIPTOR = 8;

/** The most a 64-bit field may hold in its upper half and still be an exact number. */
const MAX_EXACT_UPPER_HALF = 0x1fffff;

export function refuse(message: string): never {
  throw new XlsxReadError("INVALID_PACKAGE", message);
}

/**
 * Little-endian fields. A read outside `bytes` gives 0 and never throws, so a record that is cut
 * short fails the check of its signature or of its bounds like any other damaged one.
 */
export function u16(bytes: Uint8Array, at: number): number {
  return (bytes[at] ?? 0) | ((bytes[at + 1] ?? 0) << 8);
}

export function u32(bytes: Uint8Array, at: number): number {
  return (u16(bytes, at) | (u16(bytes, at + 2) << 16)) >>> 0;
}

export function u64(bytes: Uint8Array, at: number): number {
  const upper = u32(bytes, at + 4);
  if (upper > MAX_EXACT_UPPER_HALF) refuse("the directory of the package is damaged");
  return upper * 2 ** 32 + u32(bytes, at);
}

/**
 * Decodes a part's name. Bytes that are not UTF-8 read as U+FFFD and never throw. A byte order
 * mark is kept: a name that starts with one is not the name without it.
 */
export const names = new TextDecoder("utf-8", { ignoreBOM: true });

/**
 * The form in which two part names are compared: the package rules make them equal without regard
 * to the case of the letters A to Z, and of no other. The Kelvin sign, whose lower case is `k`,
 * stays what it is.
 */
export function foldName(name: string): string {
  return name.replace(/[A-Z]+/g, (letters) => letters.toLowerCase());
}
