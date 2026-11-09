/**
 * DEFLATE without compression: the input cut into stored blocks, each a five-byte header and up
 * to 65 535 bytes copied as they are. Every inflater reads it, it is deterministic, and it needs
 * no `CompressionStream`, so a streamed entry can carry a data descriptor on every runtime.
 */

/** The most data one stored block can carry: its length field is 16 bits. */
export const MAX_STORED_BLOCK = 65_535;
const HEADER_SIZE = 5;

/** One stored block: `BFINAL` in bit 0, `BTYPE` 00, then LEN and its one's complement, little-endian. */
export function storedBlock(data: Uint8Array, final: boolean): Uint8Array {
  const length = data.length;
  const block = new Uint8Array(HEADER_SIZE + length);
  block[0] = final ? 1 : 0;
  block[1] = length & 0xff;
  block[2] = length >>> 8;
  block[3] = ~length & 0xff;
  block[4] = (~length >>> 8) & 0xff;
  block.set(data, HEADER_SIZE);
  return block;
}

/**
 * A transform that writes every chunk as stored blocks and closes the stream with an empty final
 * block, so the last data block never has to know it is last.
 */
export function createStoredBlocksStream(): TransformStream<Uint8Array, Uint8Array> {
  return new TransformStream<Uint8Array, Uint8Array>({
    transform(chunk, controller) {
      for (let offset = 0; offset < chunk.length; offset += MAX_STORED_BLOCK) {
        controller.enqueue(storedBlock(chunk.subarray(offset, offset + MAX_STORED_BLOCK), false));
      }
    },
    flush(controller) {
      controller.enqueue(storedBlock(new Uint8Array(0), true));
    },
  });
}
