let table: Uint32Array | undefined;

function buildTable(): Uint32Array {
  const built = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    built[n] = c >>> 0;
  }
  return built;
}

/**
 * CRC-32 (IEEE 802.3, the polynomial zip uses) of a byte array, as an unsigned 32-bit integer.
 * The lookup table is built on first use and shared afterwards.
 *
 * @example crc32(new TextEncoder().encode("123456789")) // 0xcbf43926
 */
export function crc32(bytes: Uint8Array): number {
  const lookup = (table ??= buildTable());
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    // Index is in range by construction of the loop and the 256-entry table.
    crc = (lookup[(crc ^ bytes[i]!) & 0xff]! ^ (crc >>> 8)) >>> 0;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
