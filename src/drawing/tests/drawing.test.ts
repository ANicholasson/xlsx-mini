import { readFileSync } from "node:fs";
import { crc32 } from "node:zlib";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import type { XlsxMiniError } from "../../errors/index.js";
import { createMediaTable, createSheetDrawing } from "../index.js";

const images = join(dirname(fileURLToPath(import.meta.url)), "../../../test/fixtures/images");
const stripes = new Uint8Array(readFileSync(join(images, "stripes.png")));
const badge = new Uint8Array(readFileSync(join(images, "badge.png")));

const failure = (run: () => void): XlsxMiniError => {
  try {
    run();
  } catch (error) {
    return error as XlsxMiniError;
  }
  throw new Error("expected a throw");
};

describe("media table", () => {
  it("reads the pixel size from the PNG header and numbers distinct images across calls", () => {
    const media = createMediaTable();
    expect(media.add(stripes, "p")).toEqual({ id: 1, width: 120, height: 40 });
    expect(media.add(badge, "p")).toEqual({ id: 2, width: 48, height: 48 });
    expect(media.takeParts().map((part) => [part.path, part.data.length])).toEqual([
      ["xl/media/image1.png", stripes.length],
      ["xl/media/image2.png", badge.length],
    ]);
    expect(media.takeParts()).toEqual([]);
  });

  it("writes identical bytes once, by value rather than by reference", () => {
    const media = createMediaTable();
    media.add(stripes, "p");
    expect(media.add(new Uint8Array(stripes), "p").id).toBe(1);
    const twin = new Uint8Array(stripes);
    twin[twin.length - 1] = twin[twin.length - 1]! ^ 0x01;
    expect(media.add(twin, "p").id).toBe(2);
    expect(media.takeParts()).toHaveLength(2);
  });

  it("tells apart two different images of the same length and the same CRC-32", () => {
    // Two PNG-shaped buffers that collide: the header is checked, the trailing bytes are not, so a
    // birthday search over mixed tails finds a pair with one CRC-32 in a few hundred thousand tries.
    const seen = new Map<number, Uint8Array>();
    let pair: [Uint8Array, Uint8Array] | undefined;
    for (let i = 0; pair === undefined; i++) {
      const buffer = new Uint8Array(32);
      buffer.set(stripes.subarray(0, 24));
      new DataView(buffer.buffer).setUint32(24, i);
      new DataView(buffer.buffer).setUint32(28, Math.imul(i, 2654435761) >>> 0);
      const crc = crc32(buffer);
      const other = seen.get(crc);
      if (other !== undefined) pair = [other, buffer];
      else seen.set(crc, buffer);
    }
    const [first, second] = pair;
    expect(crc32(first)).toBe(crc32(second));
    expect(first).not.toEqual(second);
    const media = createMediaTable();
    expect(media.add(first, "p").id).toBe(1);
    expect(media.add(second, "p").id).toBe(2);
    expect(media.add(first, "p").id).toBe(1);
  });

  it("refuses what is not a PNG, at the image property", () => {
    const media = createMediaTable();
    expect(failure(() => media.add("png", "sheets[0].rows[1][2]"))).toMatchObject({
      code: "INVALID_IMAGE",
      path: "sheets[0].rows[1][2].image",
    });
    expect(failure(() => media.add(new Uint8Array(10), "p")).message).toContain("too short");
    expect(failure(() => media.add(new Uint8Array(30), "p")).message).toContain("signature");
    const noHeader = new Uint8Array(stripes);
    noHeader[12] = 0x58;
    expect(failure(() => media.add(noHeader, "p")).message).toContain("IHDR");
    const zero = new Uint8Array(stripes);
    zero[16] = zero[17] = zero[18] = zero[19] = 0;
    expect(failure(() => media.add(zero, "p")).message).toContain("impossible size");
  });
});

describe("sheet drawing", () => {
  it("anchors each picture at its cell, scaling one given dimension by the other", () => {
    const media = createMediaTable();
    const drawing = createSheetDrawing();
    expect(drawing.hasImages).toBe(false);
    const stripesMedia = media.add(stripes, "p");
    expect(drawing.add(2, 1, stripesMedia, undefined, undefined, "p")).toEqual({
      width: 120,
      height: 40,
    });
    expect(drawing.add(2, 2, stripesMedia, 60, undefined, "p")).toEqual({ width: 60, height: 20 });
    expect(drawing.add(3, 1, media.add(badge, "p"), undefined, 24, "p")).toEqual({
      width: 24,
      height: 24,
    });
    expect(drawing.add(3, 2, stripesMedia, 10, 10, "p")).toEqual({ width: 10, height: 10 });
    expect(drawing.hasImages).toBe(true);
    const xml = drawing.render();
    expect(xml).toContain(
      '<xdr:from><xdr:col>2</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from><xdr:ext cx="1143000" cy="381000"/>',
    );
    expect(xml).toContain('<xdr:ext cx="571500" cy="190500"/>');
    expect(xml).toContain('<xdr:cNvPr id="2" name="Picture 1"/>');
    expect(xml).toContain('<xdr:cNvPr id="5" name="Picture 4"/>');
    expect(xml.match(/r:embed="rId1"/g)).toHaveLength(3);
    expect(xml.match(/r:embed="rId2"/g)).toHaveLength(1);
    expect(drawing.relationships()).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/></Relationships>',
    );
  });

  it("rejects a size whose other dimension, scaled by the aspect ratio, passes the limit", () => {
    const drawing = createSheetDrawing();
    const media = createMediaTable().add(stripes, "p");
    expect(failure(() => drawing.add(0, 0, media, undefined, 50_000, "c"))).toMatchObject({
      code: "INVALID_IMAGE",
      path: "c.height",
    });
    const widest = drawing.add(0, 0, media, 100_000, undefined, "c");
    expect(widest.width).toBe(100_000);
    expect(widest.height).toBeCloseTo(33_333.33, 1);
    expect(drawing.add(0, 0, media, undefined, 30_000, "c")).toEqual({
      width: 90_000,
      height: 30_000,
    });
    // A tall, thin image: the header is read, not the pixels, so one can be declared.
    const tall = new Uint8Array(stripes);
    new DataView(tall.buffer).setUint32(16, 1);
    new DataView(tall.buffer).setUint32(20, 100_000);
    const tallMedia = createMediaTable().add(tall, "p");
    expect(failure(() => drawing.add(0, 0, tallMedia, 2, undefined, "c"))).toMatchObject({
      code: "INVALID_IMAGE",
      path: "c.width",
    });
  });

  it("rejects a width or height that is not a positive number of pixels, at the property", () => {
    const drawing = createSheetDrawing();
    const media = createMediaTable().add(badge, "p");
    for (const bad of [0, -1, Number.NaN, Number.POSITIVE_INFINITY, "40", 100_001]) {
      expect(failure(() => drawing.add(0, 0, media, bad as never, undefined, "c"))).toMatchObject({
        code: "INVALID_IMAGE",
        path: "c.width",
      });
      expect(failure(() => drawing.add(0, 0, media, undefined, bad as never, "c"))).toMatchObject({
        code: "INVALID_IMAGE",
        path: "c.height",
      });
    }
  });
});
