import { XlsxMiniError } from "../../errors/index.js";
import type { StreamCompression } from "./compression.js";
import { CRC32_INITIAL, crc32Finish, crc32Update } from "./crc32.js";

/** A file inside a streamed package: its zip entry name and its bytes as they are produced. */
export interface StreamedPart {
  readonly path: string;
  readonly body: ReadableStream<Uint8Array>;
}

const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const LOCAL_HEADER_SIZE = 30;
const DATA_DESCRIPTOR_SIZE = 16;
const CENTRAL_HEADER_SIZE = 46;
const END_OF_CENTRAL_DIRECTORY_SIZE = 22;
const VERSION_2_0 = 20;
/** Bit 3: CRC and sizes follow the data in a descriptor. */
const DESCRIPTOR_FLAG = 0x0008;
/** 1980-01-01 00:00:00, the DOS epoch. Fixed so identical input gives identical bytes. */
const DOS_TIME = 0x0000;
const DOS_DATE = 0x0021;
const MAX_32_BIT = 0xffffffff;

const encoder = new TextEncoder();

/** What the central directory needs about an entry once its data has gone out. */
interface Entry {
  readonly name: Uint8Array;
  readonly offset: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
}

/** The entry whose data is going out now. */
interface Open {
  readonly entry: Entry;
  readonly reader: ReadableStreamDefaultReader<Uint8Array>;
  /** Aborting it cancels the part's body; `piped` settles once that has happened. */
  readonly abort: AbortController;
  readonly piped: Promise<void>;
  /** Errors the tap on both sides, so a write blocked by backpressure lets the body's pipe shut down. */
  readonly errorTap: (reason: unknown) => void;
}

/**
 * Packs parts into a zip archive as they are pulled: each entry's local header goes out before
 * its data, its CRC and sizes follow in a data descriptor, and the central directory closes the
 * archive. Parts are taken from the iterator one at a time, the next only once the previous is
 * complete, so a part rendered late may depend on everything before it. Cancelling the stream
 * cancels the open part and returns the iterator. The caller keeps the entry count within the
 * 16-bit field; the sizes and offsets are checked here, since they are only known as data goes out.
 */
export function streamPackage(
  parts: AsyncIterator<StreamedPart>,
  compression: StreamCompression,
): ReadableStream<Uint8Array> {
  const entries: Entry[] = [];
  let offset = 0;
  let open: Open | undefined;
  /** Set by cancel; a part that finishes opening afterwards is cancelled rather than started. */
  let cancelled: { readonly reason: unknown } | undefined;
  /** The part-opening in flight, so cancel can wait for it. */
  let opening: Promise<Uint8Array | undefined> | undefined;

  const tooLarge = (): XlsxMiniError =>
    new XlsxMiniError(
      "PACKAGE_TOO_LARGE",
      "workbook",
      "The package exceeds the 4 GiB limit of a plain zip archive",
    );

  const openNext = async (): Promise<Uint8Array | undefined> => {
    const next = await parts.next();
    if (next.done) return undefined;
    const part = next.value;
    if (cancelled !== undefined) {
      await part.body.cancel(cancelled.reason);
      return undefined;
    }
    const name = encoder.encode(part.path);
    const entry: Entry = {
      name,
      offset,
      crc: CRC32_INITIAL,
      compressedSize: 0,
      uncompressedSize: 0,
    };
    entries.push(entry);
    let tapController!: TransformStreamDefaultController<Uint8Array>;
    const tap = new TransformStream<Uint8Array, Uint8Array>({
      start(controller) {
        tapController = controller;
      },
      transform(chunk, controller) {
        entry.uncompressedSize += chunk.length;
        if (entry.uncompressedSize > MAX_32_BIT) throw tooLarge();
        entry.crc = crc32Update(entry.crc, chunk);
        controller.enqueue(chunk);
      },
    });
    const abort = new AbortController();
    // Piped by hand rather than with pipeThrough so that cancelling can wait for the body's own
    // cancel to finish; a failure surfaces through the reader, so the pipe's rejection is spent.
    const piped = part.body.pipeTo(tap.writable, { signal: abort.signal }).catch(() => undefined);
    open = {
      entry,
      reader: tap.readable.pipeThrough(compression.transform()).getReader(),
      abort,
      piped,
      errorTap: (reason) => tapController.error(reason),
    };
    const header = new Uint8Array(LOCAL_HEADER_SIZE + name.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, LOCAL_HEADER_SIGNATURE, true);
    view.setUint16(4, VERSION_2_0, true);
    view.setUint16(6, DESCRIPTOR_FLAG, true);
    view.setUint16(8, compression.method, true);
    view.setUint16(10, DOS_TIME, true);
    view.setUint16(12, DOS_DATE, true);
    // CRC and both sizes stay zero: they are only known once the data has gone out.
    view.setUint16(26, name.length, true);
    header.set(name, LOCAL_HEADER_SIZE);
    offset += header.length;
    return header;
  };

  const closeOpen = (entry: Entry): Uint8Array => {
    entry.crc = crc32Finish(entry.crc);
    const descriptor = new Uint8Array(DATA_DESCRIPTOR_SIZE);
    const view = new DataView(descriptor.buffer);
    view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
    view.setUint32(4, entry.crc, true);
    view.setUint32(8, entry.compressedSize, true);
    view.setUint32(12, entry.uncompressedSize, true);
    offset += descriptor.length;
    open = undefined;
    return descriptor;
  };

  const centralDirectory = (): Uint8Array => {
    const count = entries.length;
    let centralBytes = 0;
    for (const entry of entries) centralBytes += CENTRAL_HEADER_SIZE + entry.name.length;
    const centralDirectoryOffset = offset;
    if (centralDirectoryOffset + centralBytes + END_OF_CENTRAL_DIRECTORY_SIZE > MAX_32_BIT) {
      throw tooLarge();
    }
    const out = new Uint8Array(centralBytes + END_OF_CENTRAL_DIRECTORY_SIZE);
    const view = new DataView(out.buffer);
    let cursor = 0;
    for (const entry of entries) {
      view.setUint32(cursor, CENTRAL_HEADER_SIGNATURE, true);
      view.setUint16(cursor + 4, VERSION_2_0, true); // version made by
      view.setUint16(cursor + 6, VERSION_2_0, true); // version needed
      view.setUint16(cursor + 8, DESCRIPTOR_FLAG, true);
      view.setUint16(cursor + 10, compression.method, true);
      view.setUint16(cursor + 12, DOS_TIME, true);
      view.setUint16(cursor + 14, DOS_DATE, true);
      view.setUint32(cursor + 16, entry.crc, true);
      view.setUint32(cursor + 20, entry.compressedSize, true);
      view.setUint32(cursor + 24, entry.uncompressedSize, true);
      view.setUint16(cursor + 28, entry.name.length, true);
      view.setUint16(cursor + 30, 0, true); // extra field length
      view.setUint16(cursor + 32, 0, true); // comment length
      view.setUint16(cursor + 34, 0, true); // disk number start
      view.setUint16(cursor + 36, 0, true); // internal attributes
      view.setUint32(cursor + 38, 0, true); // external attributes
      view.setUint32(cursor + 42, entry.offset, true);
      cursor += CENTRAL_HEADER_SIZE;
      out.set(entry.name, cursor);
      cursor += entry.name.length;
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
  };

  /** Cancels the open part, lets a part still arriving be cancelled on arrival, and returns the iterator. */
  const release = async (reason: unknown): Promise<void> => {
    cancelled = { reason };
    // The open part is captured first: a pending read may resolve as done while this waits, and
    // closeOpen then clears `open`. The iterator is returned whatever happens on the way.
    const current = open;
    try {
      // A part still being taken from the iterator is cancelled by openNext once it arrives.
      await opening?.catch(() => undefined);
      if (current !== undefined) {
        // Older runtimes do not error a transform's writable side when its readable is cancelled,
        // so a write held back by the compressor would never settle and the body's pipe never
        // shut down.
        current.errorTap(reason);
        current.abort.abort(reason);
        await current.piped;
        await current.reader.cancel(reason).catch(() => undefined);
      }
    } finally {
      await parts.return?.(undefined);
    }
  };

  /** One pull's work: a pull that enqueues nothing is never called again, so this loops until it has. */
  const produce = async (
    controller: ReadableStreamDefaultController<Uint8Array>,
  ): Promise<void> => {
    for (;;) {
      if (open === undefined) {
        opening = openNext();
        const header = await opening;
        opening = undefined;
        if (cancelled !== undefined) return;
        if (header === undefined) {
          controller.enqueue(centralDirectory());
          controller.close();
          return;
        }
        controller.enqueue(header);
        return;
      }
      const { done, value } = await open.reader.read();
      if (done) {
        controller.enqueue(closeOpen(open.entry));
        return;
      }
      if (value.length === 0) continue;
      open.entry.compressedSize += value.length;
      offset += value.length;
      if (offset > MAX_32_BIT) throw tooLarge();
      controller.enqueue(value);
      return;
    }
  };

  return new ReadableStream<Uint8Array>(
    {
      async pull(controller) {
        try {
          await produce(controller);
        } catch (error) {
          // An errored stream never has its cancel called: let go of the part and the iterator here.
          await release(error);
          throw error;
        }
      },
      cancel: release,
    },
    { highWaterMark: 0 },
  );
}
