import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateRawSync } from "node:zlib";
import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { unzip } from "../../../test/helpers/unzip.js";
import {
  MAX_STORED_BLOCK,
  type StreamCompression,
  type StreamedPart,
  createPackage,
  selectStreamCompression,
  store,
  storedBlock,
  streamPackage,
} from "../index.js";

const encode = (text: string) => new TextEncoder().encode(text);
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

/** A part whose body arrives in the chunks given. */
function part(path: string, ...chunks: readonly Uint8Array[]): StreamedPart {
  return {
    path,
    body: new ReadableStream({
      start(controller) {
        for (const chunk of chunks) controller.enqueue(chunk);
        controller.close();
      },
    }),
  };
}

async function* parts(...list: readonly StreamedPart[]): AsyncGenerator<StreamedPart> {
  for (const item of list) yield item;
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of stream) {
    chunks.push(chunk);
    total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

const stored = selectStreamCompression("none");
const deflated = selectStreamCompression("deflate");

const sample = () =>
  parts(
    part("[Content_Types].xml", encode("<Types/>")),
    part(
      "xl/worksheets/sheet1.xml",
      encode("<worksheet>"),
      encode("æøå ".repeat(500)),
      encode("</worksheet>"),
    ),
    part("empty.txt"),
  );

function passesSystemUnzip(bytes: Uint8Array): boolean {
  const dir = mkdtempSync(join(tmpdir(), "xlsx-mini-"));
  const file = join(dir, "package.zip");
  writeFileSync(file, bytes);
  return execFileSync("unzip", ["-t", file], { encoding: "utf8" }).includes("No errors detected");
}

describe("stored blocks", () => {
  it("frames data so an inflater returns it unchanged, with the final bit only where asked", () => {
    const data = encode("hello");
    const block = storedBlock(data, false);
    expect([...block.subarray(0, 5)]).toEqual([0, 5, 0, 0xfa, 0xff]);
    expect(storedBlock(data, true)[0]).toBe(1);
    const stream = new Uint8Array([...block, ...storedBlock(new Uint8Array(0), true)]);
    expect(decode(inflateRawSync(stream))).toBe("hello");
  });

  it("splits a chunk longer than one block and closes with an empty final block", async () => {
    const big = new Uint8Array(MAX_STORED_BLOCK + 10);
    for (let i = 0; i < big.length; i++) big[i] = i & 0xff;
    const out = await collect(
      part("x", big, new Uint8Array(0)).body.pipeThrough(stored.transform()),
    );
    // Two data blocks, 5 bytes of header each, then the 5-byte empty final block.
    expect(out.length).toBe(big.length + 15);
    expect(out[0]).toBe(0);
    expect(out[5 + MAX_STORED_BLOCK]).toBe(0);
    expect(out[out.length - 5]).toBe(1);
    expect(new Uint8Array(inflateRawSync(out))).toEqual(big);
  });
});

describe("selectStreamCompression", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("is method 8 either way, through CompressionStream where raw deflate exists", async () => {
    expect(stored.method).toBe(8);
    expect(deflated.method).toBe(8);
    const out = await collect(
      part("x", encode("a".repeat(1000))).body.pipeThrough(deflated.transform()),
    );
    expect(out.length).toBeLessThan(100);
    expect(decode(inflateRawSync(out))).toBe("a".repeat(1000));
  });

  it("falls back to stored blocks where the runtime cannot deflate", async () => {
    vi.stubGlobal(
      "CompressionStream",
      class {
        constructor() {
          throw new TypeError("deflate-raw is not supported");
        }
      },
    );
    const fallback = selectStreamCompression("deflate");
    const out = await collect(part("x", encode("abc")).body.pipeThrough(fallback.transform()));
    expect([...out]).toEqual([0, 3, 0, 0xfc, 0xff, 97, 98, 99, 1, 0, 0, 0xff, 0xff]);
  });
});

describe("streamPackage", () => {
  it("writes entries with data descriptors that an independent reader walks back in order", async () => {
    const bytes = await collect(streamPackage(sample(), stored));
    const entries = unzip(bytes);
    expect(entries.map((e) => e.name)).toEqual([
      "[Content_Types].xml",
      "xl/worksheets/sheet1.xml",
      "empty.txt",
    ]);
    expect(entries.every((e) => e.method === 8)).toBe(true);
    expect(decode(entries[1]!.data)).toBe(`<worksheet>${"æøå ".repeat(500)}</worksheet>`);
    expect(entries[2]!.data.length).toBe(0);
    const view = new DataView(bytes.buffer);
    // Local header: flag bit 3 set, CRC and sizes zero; the descriptor follows the data.
    expect(view.getUint16(6, true)).toBe(8);
    expect(view.getUint32(14, true)).toBe(0);
    expect(view.getUint32(18, true)).toBe(0);
    expect(view.getUint32(22, true)).toBe(0);
    const first = entries[0]!;
    const descriptor = 30 + first.name.length + first.compressedSize;
    expect(view.getUint32(descriptor, true)).toBe(0x08074b50);
    expect(view.getUint32(descriptor + 4, true)).toBe(first.crc32);
    expect(view.getUint32(descriptor + 8, true)).toBe(first.compressedSize);
    expect(view.getUint32(descriptor + 12, true)).toBe(first.uncompressedSize);
    expect(entries[1]!.localHeaderOffset).toBe(descriptor + 16);
    expect(passesSystemUnzip(bytes)).toBe(true);
  });

  it("holds the same parts as the buffered package, and the deflated form inflates to them", async () => {
    const buffered = unzip(
      await createPackage(
        [
          { path: "[Content_Types].xml", data: encode("<Types/>") },
          {
            path: "xl/worksheets/sheet1.xml",
            data: encode(`<worksheet>${"æøå ".repeat(500)}</worksheet>`),
          },
          { path: "empty.txt", data: new Uint8Array(0) },
        ],
        store,
      ),
    );
    for (const compression of [stored, deflated]) {
      const streamed = unzip(await collect(streamPackage(sample(), compression)));
      expect(streamed.map((e) => [e.name, decode(e.data), e.crc32])).toEqual(
        buffered.map((e) => [e.name, decode(e.data), e.crc32]),
      );
    }
    expect(passesSystemUnzip(await collect(streamPackage(sample(), deflated)))).toBe(true);
  });

  it("is byte-identical across calls with stored blocks", async () => {
    expect(await collect(streamPackage(sample(), stored))).toEqual(
      await collect(streamPackage(sample(), stored)),
    );
  });

  it("takes the next part only once the previous one has gone out, and renders nothing before the first read", async () => {
    const order: string[] = [];
    async function* lazy(): AsyncGenerator<StreamedPart> {
      order.push("first asked");
      yield part("a", encode("A"));
      order.push("second asked");
      yield part("b", encode("B"));
    }
    const stream = streamPackage(lazy(), stored);
    await Promise.resolve();
    expect(order).toEqual([]);
    const reader = stream.getReader();
    const isHeaderOf = (chunk: Uint8Array, name: string) =>
      chunk.length === 30 + name.length &&
      new DataView(chunk.buffer, chunk.byteOffset).getUint32(0, true) === 0x04034b50 &&
      decode(chunk.subarray(30)) === name;
    const chunks: Uint8Array[] = [];
    const read = async () => {
      const chunk = (await reader.read()).value!;
      chunks.push(chunk);
      return chunk;
    };
    expect(isHeaderOf(await read(), "a")).toBe(true);
    expect(order).toEqual(["first asked"]);
    let chunk = await read();
    // Data blocks and the descriptor of a go out before b is even asked for.
    while (!isHeaderOf(chunk, "b")) {
      expect(order).toEqual(["first asked"]);
      chunk = await read();
    }
    expect(order).toEqual(["first asked", "second asked"]);
    for (let next = await reader.read(); !next.done; next = await reader.read()) {
      chunks.push(next.value);
    }
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let offset = 0;
    for (const c of chunks) {
      bytes.set(c, offset);
      offset += c.length;
    }
    expect(unzip(bytes).map((e) => e.name)).toEqual(["a", "b"]);
  });

  it("cancels the open part and returns the part iterator when the consumer cancels", async () => {
    let bodyCancelled: unknown;
    let returned = false;
    const endless: StreamedPart = {
      path: "endless",
      body: new ReadableStream({
        pull(controller) {
          controller.enqueue(encode("more "));
        },
        cancel(reason) {
          bodyCancelled = reason;
        },
      }),
    };
    async function* source(): AsyncGenerator<StreamedPart> {
      try {
        yield endless;
        yield part("never", encode("x"));
      } finally {
        returned = true;
      }
    }
    const stream = streamPackage(source(), stored);
    const reader = stream.getReader();
    await reader.read();
    await reader.read();
    await reader.cancel("gone");
    expect(bodyCancelled).toBe("gone");
    expect(returned).toBe(true);
  });

  it("skips empty chunks a compressor may emit", async () => {
    const chatty: StreamCompression = {
      method: 8,
      transform: () =>
        new TransformStream<Uint8Array, Uint8Array>({
          transform(chunk, controller) {
            controller.enqueue(new Uint8Array(0));
            controller.enqueue(storedBlock(chunk, false));
          },
          flush(controller) {
            controller.enqueue(storedBlock(new Uint8Array(0), true));
          },
        }),
    };
    const bytes = await collect(streamPackage(parts(part("a", encode("abc"))), chatty));
    expect(decode(unzip(bytes)[0]!.data)).toBe("abc");
  });

  it("returns the part iterator when cancelled before anything was pulled", async () => {
    let started = false;
    async function* source(): AsyncGenerator<StreamedPart> {
      started = true;
      yield part("a", encode("x"));
    }
    const iterator = source();
    await streamPackage(iterator, stored).cancel("never read");
    // A generator returned before its first `next` closes without running its body.
    expect(started).toBe(false);
    expect(await iterator.next()).toEqual({ done: true, value: undefined });
  });

  it("cancels a part that was still being taken from the iterator when the consumer cancelled", async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    let bodyCancelled: unknown;
    let returned = false;
    const late: StreamedPart = {
      path: "late",
      body: new ReadableStream({
        pull(controller) {
          controller.enqueue(encode("row "));
        },
        cancel(reason) {
          bodyCancelled = reason;
        },
      }),
    };
    async function* source(): AsyncGenerator<StreamedPart> {
      try {
        yield part("a", encode("x"));
        await gate;
        yield late;
      } finally {
        returned = true;
      }
    }
    const stream = streamPackage(source(), stored);
    const reader = stream.getReader();
    // Read a to its descriptor, then start the pull that waits on the gate for the next part.
    for (let i = 0; i < 4; i++) await reader.read();
    const pending = reader.read();
    // The pull behind that read starts on a later microtask; let it reach the iterator first.
    await new Promise((resolve) => setTimeout(resolve, 0));
    const cancelling = reader.cancel("gone");
    release();
    await cancelling;
    await pending;
    expect(bodyCancelled).toBe("gone");
    expect(returned).toBe(true);
  });

  it("still settles a cancel when the part iterator fails while it is being awaited", async () => {
    let fail!: (error: Error) => void;
    const gate = new Promise<never>((_resolve, reject) => {
      fail = reject;
    });
    async function* source(): AsyncGenerator<StreamedPart> {
      yield part("a", encode("x"));
      await gate;
    }
    const stream = streamPackage(source(), stored);
    const reader = stream.getReader();
    for (let i = 0; i < 4; i++) await reader.read();
    const pending = reader.read();
    await new Promise((resolve) => setTimeout(resolve, 0));
    const cancelling = reader.cancel("gone");
    fail(new Error("cursor broke"));
    await expect(cancelling).resolves.toBeUndefined();
    await expect(pending).resolves.toMatchObject({ done: true });
  });

  it("lets go of the open part and the iterator when the package itself fails", async () => {
    let bodyCancelled = false;
    let returned = false;
    const endless: StreamedPart = {
      path: "endless",
      body: new ReadableStream({
        pull(controller) {
          controller.enqueue(encode("more "));
        },
        cancel() {
          bodyCancelled = true;
        },
      }),
    };
    async function* source(): AsyncGenerator<StreamedPart> {
      try {
        yield endless;
      } finally {
        returned = true;
      }
    }
    const huge: StreamCompression = {
      method: 8,
      transform: () =>
        new TransformStream<Uint8Array, Uint8Array>({
          transform(_chunk, controller) {
            controller.enqueue({ length: 0x1_0000_0000 } as Uint8Array);
          },
        }),
    };
    await expect(collect(streamPackage(source(), huge))).rejects.toMatchObject({
      code: "PACKAGE_TOO_LARGE",
    });
    expect(bodyCancelled).toBe(true);
    expect(returned).toBe(true);
  });

  it("errors the stream when a part fails", async () => {
    const failing: StreamedPart = {
      path: "bad",
      body: new ReadableStream({
        pull() {
          throw new Error("row 3 is bad");
        },
      }),
    };
    await expect(collect(streamPackage(parts(failing), stored))).rejects.toThrow("row 3 is bad");
  });

  it("errors with PACKAGE_TOO_LARGE when an entry or the archive exceeds the zip limits", async () => {
    // A transform that reports huge chunks without allocating them.
    const huge = (length: number): StreamCompression => ({
      method: 8,
      transform: () =>
        new TransformStream<Uint8Array, Uint8Array>({
          transform(_chunk, controller) {
            controller.enqueue({ length } as Uint8Array);
          },
        }),
    });
    await expect(
      collect(streamPackage(parts(part("a", encode("x"))), huge(0x1_0000_0000))),
    ).rejects.toMatchObject({ code: "PACKAGE_TOO_LARGE" });
    await expect(
      collect(
        streamPackage(parts(part("a", encode("x")), part("b", encode("x"))), huge(0xffff_ff00)),
      ),
    ).rejects.toMatchObject({ code: "PACKAGE_TOO_LARGE" });
    // Data that fits the 32-bit offsets exactly, so only the central directory pushes past them.
    await expect(
      collect(streamPackage(parts(part("a", encode("x"))), huge(0xffff_ffff - 31 - 16))),
    ).rejects.toMatchObject({ code: "PACKAGE_TOO_LARGE" });
    const uncompressedTooLarge: StreamedPart = {
      path: "u",
      body: new ReadableStream({
        start(controller) {
          controller.enqueue({ length: 0x1_0000_0000 } as Uint8Array);
          controller.close();
        },
      }),
    };
    await expect(
      collect(streamPackage(parts(uncompressedTooLarge), huge(1))),
    ).rejects.toMatchObject({ code: "PACKAGE_TOO_LARGE" });
  });
});
