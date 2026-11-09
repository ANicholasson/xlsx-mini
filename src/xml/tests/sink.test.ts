import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { MAX_CHUNK_SIZE, createChunkSink, createXmlSink } from "../index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

describe("createXmlSink", () => {
  it("returns the prefix and suffix alone when nothing was written", () => {
    expect(decoder.decode(createXmlSink().finish("<a>", "</a>"))).toBe("<a></a>");
  });

  it("encodes a body far larger than its first block, multi-byte characters crossing block edges", () => {
    const sink = createXmlSink();
    const piece = "<c>Ærlig «ekte» 🎉</c>";
    // 20 000 pieces of 30 bytes cross the block edges at 16, 32, 64, 128, 256 and 512 KiB.
    let expected = "";
    for (let i = 0; i < 20_000; i++) {
      sink.write(piece);
      expected += piece;
    }
    const actual = sink.finish("<row>", "</row>");
    const wanted = encoder.encode(`<row>${expected}</row>`);
    // A byte comparison: element-wise deep equality on a large typed array takes seconds.
    expect(actual.byteLength).toBe(wanted.byteLength);
    expect(Buffer.compare(actual, wanted)).toBe(0);
  });

  it("agrees with encoding the concatenated text for arbitrary pieces", () => {
    fc.assert(
      fc.property(
        fc.array(fc.string({ unit: "grapheme" })),
        fc.string(),
        fc.string(),
        (pieces, prefix, suffix) => {
          const sink = createXmlSink();
          for (const piece of pieces) sink.write(piece);
          expect(sink.finish(prefix, suffix)).toEqual(
            encoder.encode(prefix + pieces.join("") + suffix),
          );
        },
      ),
    );
  });
});

describe("createChunkSink", () => {
  const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

  it("hands out nothing until a chunk could be full, then everything at end", () => {
    const empty = createChunkSink();
    empty.end();
    expect(empty.take()).toBeUndefined();
    const sink = createChunkSink();
    sink.write("<row/>");
    expect(sink.take()).toBeUndefined();
    sink.end();
    expect(decode(sink.take()!)).toBe("<row/>");
    expect(sink.take()).toBeUndefined();
  });

  it("cuts chunks at the stored-block limit without splitting a character", () => {
    const sink = createChunkSink();
    // Three-byte characters: a full chunk of 65 535 bytes holds 21 845 of them exactly, so one
    // more pushes the boundary onto a character that must move whole into the next chunk.
    const text = "æ".repeat(40_000) + "€".repeat(30_000);
    sink.write(text);
    const chunks: Uint8Array[] = [];
    for (let chunk = sink.take(); chunk !== undefined; chunk = sink.take()) chunks.push(chunk);
    sink.end();
    for (let chunk = sink.take(); chunk !== undefined; chunk = sink.take()) chunks.push(chunk);
    expect(chunks.length).toBeGreaterThan(1);
    expect(chunks.every((c) => c.length <= MAX_CHUNK_SIZE)).toBe(true);
    expect(chunks.slice(0, -1).every((c) => c.length > MAX_CHUNK_SIZE - 4)).toBe(true);
    expect(chunks.map(decode).join("")).toBe(text);
  });

  it("encodes pending text in pieces as it is written, so a part never exists as one string", () => {
    const sink = createChunkSink();
    for (let i = 0; i < 20_000; i++) sink.write(`<c r="A${i}"><v>${i}</v></c>`);
    const early = sink.take();
    expect(early).toBeDefined();
    expect(early!.length).toBe(MAX_CHUNK_SIZE);
    sink.end();
    let total = early!.length;
    for (let chunk = sink.take(); chunk !== undefined; chunk = sink.take()) total += chunk.length;
    let expected = 0;
    for (let i = 0; i < 20_000; i++) expected += `<c r="A${i}"><v>${i}</v></c>`.length;
    expect(total).toBe(expected);
  });
});
