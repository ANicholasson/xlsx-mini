import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { createXmlSink } from "../index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

describe("createXmlSink", () => {
  it("returns the prefix and suffix alone when nothing was written", () => {
    expect(decoder.decode(createXmlSink().finish("<a>", "</a>"))).toBe("<a></a>");
  });

  it("encodes a body far larger than its first block, multi-byte characters crossing block edges", () => {
    const sink = createXmlSink();
    const piece = "<c>Ærlig «ekte» 🎉</c>";
    let expected = "";
    for (let i = 0; i < 100_000; i++) {
      sink.write(piece);
      expected += piece;
    }
    expect(sink.finish("<row>", "</row>")).toEqual(encoder.encode(`<row>${expected}</row>`));
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
