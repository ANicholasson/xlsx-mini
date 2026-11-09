import { createStoredBlocksStream } from "./stored-blocks.js";

/** How entries are stored in the package: the zip method id and the transform that produces them. */
export interface Compression {
  /** Zip method: 0 for STORE, 8 for DEFLATE. */
  readonly method: 0 | 8;
  compress(data: Uint8Array): Promise<Uint8Array>;
}

/** What a caller may ask for. The runtime may still force STORE when it cannot deflate. */
export type CompressionMode = "deflate" | "none";

/** STORE: bytes go in as they are. Byte-deterministic on every platform. */
export const store: Compression = {
  method: 0,
  compress: (data) => Promise.resolve(data),
};

/** DEFLATE through the web-standard `CompressionStream`, raw (no zlib or gzip wrapper), as zip requires. */
const deflateRaw: Compression = {
  method: 8,
  compress: (data) => drain(new CompressionStream("deflate-raw"), data),
};

/**
 * Picks the strategy for a requested mode. Deflate is confirmed by constructing the stream, not by
 * checking that the class exists: some runtimes list `deflate-raw` without implementing it. Where
 * it is unavailable the package is written with STORE entries instead.
 */
export function selectCompression(mode: CompressionMode): Compression {
  if (mode === "none") return store;
  try {
    void new CompressionStream("deflate-raw");
    return deflateRaw;
  } catch {
    return store;
  }
}

/**
 * How a streamed entry is produced: always method 8, through the runtime's raw deflate where it
 * exists and through stored blocks where it does not, because a STORE entry whose size follows its
 * data cannot be read by a sequential reader.
 */
export interface StreamCompression {
  readonly method: 8;
  transform(): ReadableWritablePair<Uint8Array, Uint8Array>;
}

const streamDeflateRaw: StreamCompression = {
  method: 8,
  transform: () => {
    const stream = new CompressionStream("deflate-raw");
    // The compressor takes any BufferSource; every chunk here is a Uint8Array.
    return { readable: stream.readable, writable: stream.writable as WritableStream<Uint8Array> };
  },
};

const streamStoredBlocks: StreamCompression = {
  method: 8,
  transform: createStoredBlocksStream,
};

/** The streamed counterpart of `selectCompression`: `"none"` and a runtime without raw deflate both give stored blocks. */
export function selectStreamCompression(mode: CompressionMode): StreamCompression {
  if (mode === "none") return streamStoredBlocks;
  try {
    void new CompressionStream("deflate-raw");
    return streamDeflateRaw;
  } catch {
    return streamStoredBlocks;
  }
}

/** Writes one chunk through a transform stream and gathers everything it emits. */
async function drain(stream: CompressionStream, data: Uint8Array): Promise<Uint8Array> {
  const writer = stream.writable.getWriter();
  // Start the write without awaiting it: the readable side must be drained concurrently.
  // Parts come from TextEncoder and are ArrayBuffer-backed; the stream API insists on that in its types.
  const written = writer.write(data as Uint8Array<ArrayBuffer>).then(() => writer.close());
  const reader = stream.readable.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.byteLength;
  }
  await written;
  if (chunks.length === 1) return chunks[0]!;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}
