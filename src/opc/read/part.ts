import { XlsxReadError } from "../../errors/read.js";
import { CRC32_INITIAL, crc32Finish, crc32Update } from "../lib/crc32.js";
import { LOCAL_HEADER_SIGNATURE, LOCAL_HEADER_SIZE, MAX_32_BIT } from "../lib/zip-records.js";
import type { DirectoryEntry } from "./directory.js";
import {
  FLAG_DATA_DESCRIPTOR,
  FLAG_ENCRYPTED,
  METHOD_DEFLATED,
  METHOD_STORED,
  names,
  refuse,
  u16,
  u32,
} from "./layout.js";
import type { ByteSource } from "./source.js";

/** One part being read: its content a chunk at a time, held to what the directory says of it. */
export interface PartBytes {
  /**
   * The next chunk of the part's content, or `undefined` after the last. The first call checks
   * the local header against the directory. Every call holds the content to the declared size, and
   * the call after the last chunk to the CRC-32. A refusal is final: every later call repeats it.
   * A chunk is a view to read from, valid until the next call. One call at a time: a second waits
   * for the first.
   */
  read(): Promise<Uint8Array | undefined>;
  /**
   * Stops reading and lets go of the inflate. A `read` after it gives `undefined`, which then says
   * nothing about the part, or repeats the refusal the part ended in. Safe to call twice, before
   * the first read and after the last.
   */
  close(): void;
}

/**
 * The most bytes of a part given to the inflate at a time. Deflate inflates at most 1 032 times,
 * so one slice is at most 16 908 288 bytes of content: what an inflate that does not wait for its
 * consumer holds ahead of it, and the largest chunk an inflate gives that makes one of each slice.
 */
const SLICE_BYTES = 16_384;

/**
 * The most bytes of a part read from the file at a time. A read of a `Blob` over a file costs
 * about 30 microseconds before its first byte, measured on Node 20, 22 and 24: read slice by
 * slice, the file reads of a part took seven times as long as in these blocks.
 */
const BLOCK_BYTES = 262_144;

type Chunks<Chunk extends Uint8Array> = AsyncGenerator<Chunk, undefined, undefined>;

/**
 * One part's content as chunks, with every promise the directory made about it enforced. The
 * caller has charged `entry.size` already: nothing here sees a meter, so nothing here can skip one.
 */
export function openPart(
  source: ByteSource,
  directoryStart: number,
  entry: DirectoryEntry,
): PartBytes {
  const chunks = verified(source, directoryStart, entry);
  let failure: { readonly reason: unknown } | undefined;
  return {
    async read() {
      try {
        const step = await chunks.next();
        if (failure === undefined) return step.value;
      } catch (reason) {
        failure ??= { reason };
      }
      throw failure.reason;
    },
    close() {
      void chunks.return(undefined);
    },
  };
}

async function* verified(
  source: ByteSource,
  directoryStart: number,
  entry: DirectoryEntry,
): Chunks<Uint8Array> {
  if (entry.encrypted) refuse("a part is encrypted");
  if (entry.method !== METHOD_STORED && entry.method !== METHOD_DEFLATED) {
    refuse("a part is neither stored nor deflated");
  }
  const headerEnd = entry.offset + LOCAL_HEADER_SIZE + entry.nameLength;
  const header = await source.read(entry.offset, headerEnd);
  const flags = u16(header, 6);
  // A header written before its data leaves its CRC-32 and its sizes zero and says so in a flag.
  const leftToDescriptor = (flags & FLAG_DATA_DESCRIPTOR) !== 0;
  const agrees = (at: number, stated: number): boolean => {
    const local = u32(header, at);
    return local === stated || (leftToDescriptor && local === 0);
  };
  const dataStart = headerEnd + u16(header, 28);
  const dataEnd = dataStart + entry.compressedSize;
  if (
    u32(header, 0) !== LOCAL_HEADER_SIGNATURE ||
    (flags & FLAG_ENCRYPTED) !== 0 ||
    u16(header, 8) !== entry.method ||
    !agrees(14, entry.crc32) ||
    // All ones in a size: zip64 moved it out of the field.
    (!agrees(18, entry.compressedSize) && u32(header, 18) !== MAX_32_BIT) ||
    (!agrees(22, entry.size) && u32(header, 22) !== MAX_32_BIT) ||
    u16(header, 26) !== entry.nameLength ||
    names.decode(header.subarray(LOCAL_HEADER_SIZE)) !== entry.name ||
    dataEnd > directoryStart
  ) {
    refuse("a local header disagrees with the directory");
  }

  const slices = sliced(source, dataStart, dataEnd);
  let read = 0;
  let crc = CRC32_INITIAL;
  for await (const chunk of entry.method === METHOD_STORED
    ? slices
    : inflated(member(slices, entry))) {
    read += chunk.length;
    if (read > entry.size) refuse("a part is larger than the directory states");
    crc = crc32Update(crc, chunk);
    yield chunk;
  }
  if (read !== entry.size) refuse("a part is smaller than the directory states");
  if (crc32Finish(crc) !== entry.crc32) refuse("a part has another CRC-32 than stated");
}

/** How a gzip member starts: the two bytes that name one, deflate, and nothing more said of it. */
const GZIP_HEADER = new Uint8Array([0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 0xff]);

/**
 * A part's deflate stream as a gzip member: a header in front of it, and behind it what a member
 * ends in, which is the CRC-32 and the size that the directory states. So the platform holds the
 * stream to its end by itself. Given the stream alone, Node 20 and 22 report no bytes after its
 * end and Deno 2.9 no stream that is cut short. Node 20 still lets one kind of bytes after the end
 * pass: a copy of that CRC-32 and size, where the lowest byte of the CRC-32 is zero.
 */
async function* member(
  slices: Chunks<Uint8Array<ArrayBuffer>>,
  entry: DirectoryEntry,
): Chunks<Uint8Array<ArrayBuffer>> {
  yield GZIP_HEADER;
  yield* slices;
  const trailer = new Uint8Array(8);
  const view = new DataView(trailer.buffer);
  view.setUint32(0, entry.crc32, true);
  // A member states its size in 32 bits, which is what this keeps of a larger one.
  view.setUint32(4, entry.size, true);
  yield trailer;
}

async function* sliced(
  source: ByteSource,
  start: number,
  end: number,
): Chunks<Uint8Array<ArrayBuffer>> {
  for (let at = start; at < end; at += BLOCK_BYTES) {
    const block = await source.read(at, Math.min(end, at + BLOCK_BYTES));
    for (let from = 0; from < block.length; from += SLICE_BYTES) {
      yield block.subarray(from, from + SLICE_BYTES);
    }
  }
}

function refuseInflate(): never {
  return refuse("a part does not inflate to what the directory states");
}

/**
 * The platform's inflate, driven by demand: a slice goes in only while a chunk is waited for and
 * the slice before it has been taken. Piping the slices through would leave that to the platform,
 * and Node 20 and 22 then read a part's every compressed byte ahead of a consumer that has asked
 * for one chunk.
 */
async function* inflated(slices: Chunks<Uint8Array<ArrayBuffer>>): Chunks<Uint8Array> {
  const { readable, writable } = platformInflate();
  const reader = readable.getReader();
  const writer = writable.getWriter();
  let feeding: Promise<undefined> | undefined;
  let fed = false;
  const feed = async (): Promise<undefined> => {
    const slice = await slices.next();
    if (slice.done) {
      fed = true;
      // What the end of the stream brings, a failure too, comes out of the readable side.
      void Promise.allSettled([writer.close()]);
    } else {
      await writer.write(slice.value).catch(refuseInflate);
    }
  };
  try {
    for (;;) {
      const reading = reader.read().then((step) => ({ step }), refuseInflate);
      let read = fed ? await reading : undefined;
      while (read === undefined) {
        feeding ??= feed();
        read = await Promise.race([reading, feeding]);
        if (read === undefined) {
          feeding = undefined;
          if (fed) read = await reading;
        }
      }
      if (read.step.done) return undefined;
      yield read.step.value;
    }
  } finally {
    void Promise.allSettled([reader.cancel(), writer.abort()]);
  }
}

function platformInflate(): DecompressionStream {
  try {
    return new DecompressionStream("gzip");
  } catch {
    throw new XlsxReadError("NO_INFLATE", "a part is deflated and the platform has no inflate");
  }
}
