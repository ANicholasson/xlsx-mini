const POLYNOMIAL = 0xedb88320;
const SLICES = 8;

let tables: Uint32Array | undefined;

/**
 * Eight 256-entry tables in one array: table k holds the CRC after the byte and k zero bytes, so
 * eight bytes advance the checksum with eight independent lookups instead of eight dependent ones.
 */
function buildTables(): Uint32Array {
  const built = new Uint32Array(SLICES * 256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? POLYNOMIAL ^ (c >>> 1) : c >>> 1;
    }
    built[n] = c >>> 0;
  }
  for (let n = 0; n < 256; n++) {
    let c = built[n]!;
    for (let k = 1; k < SLICES; k++) {
      c = (built[c & 0xff]! ^ (c >>> 8)) >>> 0;
      built[k * 256 + n] = c;
    }
  }
  return built;
}

/** The state a running CRC-32 starts from; feed chunks through `crc32Update`, then `crc32Finish`. */
export const CRC32_INITIAL = 0xffffffff;

/**
 * Advances a running CRC-32 over one chunk. Slice-by-8: the tables are built on first use and
 * shared afterwards.
 */
export function crc32Update(state: number, bytes: Uint8Array): number {
  const t = (tables ??= buildTables());
  const length = bytes.length;
  // Read as an unsigned 32-bit value: a plain parameter leaves the loop's type open, and the byte
  // path measured a quarter slower on text-heavy sheets until this was added.
  let crc = state >>> 0;
  let i = 0;
  // Indexes are in range by construction: every lookup masks to a byte within a 256-entry slice.
  for (const end = length - (length % 8); i < end; i += 8) {
    const low =
      (crc ^ (bytes[i]! | (bytes[i + 1]! << 8) | (bytes[i + 2]! << 16) | (bytes[i + 3]! << 24))) >>>
      0;
    const high =
      (bytes[i + 4]! | (bytes[i + 5]! << 8) | (bytes[i + 6]! << 16) | (bytes[i + 7]! << 24)) >>> 0;
    crc =
      (t[1792 + (low & 0xff)]! ^
        t[1536 + ((low >>> 8) & 0xff)]! ^
        t[1280 + ((low >>> 16) & 0xff)]! ^
        t[1024 + (low >>> 24)]! ^
        t[768 + (high & 0xff)]! ^
        t[512 + ((high >>> 8) & 0xff)]! ^
        t[256 + ((high >>> 16) & 0xff)]! ^
        t[high >>> 24]!) >>>
      0;
  }
  for (; i < length; i++) {
    crc = (t[(crc ^ bytes[i]!) & 0xff]! ^ (crc >>> 8)) >>> 0;
  }
  return crc;
}

/** The checksum of everything fed so far, as an unsigned 32-bit integer. */
export function crc32Finish(state: number): number {
  return (state ^ 0xffffffff) >>> 0;
}

/**
 * CRC-32 (IEEE 802.3, the polynomial zip uses) of a byte array, as an unsigned 32-bit integer.
 *
 * @example crc32(new TextEncoder().encode("123456789")) // 0xcbf43926
 */
export function crc32(bytes: Uint8Array): number {
  return crc32Finish(crc32Update(CRC32_INITIAL, bytes));
}
