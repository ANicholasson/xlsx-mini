import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { unzip } from "../../../test/helpers/unzip.js";
import { resolveLimits } from "../../errors/read.js";
import { createPackage, selectCompression, selectStreamCompression, store } from "../index.js";
import { type PartName, partName } from "../read.js";
import {
  WatchedBlob,
  commented,
  encode,
  gather,
  join as joined,
  noise,
  open,
  read,
  streamed,
} from "./reading.js";

const here = dirname(fileURLToPath(import.meta.url));

const decode = (bytes: Uint8Array | undefined) =>
  bytes === undefined ? undefined : new TextDecoder().decode(bytes);

function same(read: Uint8Array | undefined, written: Uint8Array): boolean {
  return read !== undefined && Buffer.compare(read, written) === 0;
}

const sheet = "xl/worksheets/sheet1.xml";

const parts = [
  { path: "[Content_Types].xml", data: encode("<Types/>") },
  { path: sheet, data: encode(`<worksheet>${"æøå ".repeat(500)}</worksheet>`) },
  { path: "empty.txt", data: new Uint8Array(0) },
  {
    path: "xl/worksheets/sheet2.xml",
    data: encode(Array.from({ length: 200_000 }, (_, row) => `<row r="${row}"/>`).join("\n")),
  },
  { path: "xl/media/image1.png", data: noise(100_000) },
];

const deflate = selectCompression("deflate");

describe("openPackage", () => {
  it("reads back every part of what createPackage and streamPackage write, stored and deflated, from bytes, an ArrayBuffer and a Blob", async () => {
    const packages = {
      "createPackage, stored": await createPackage(parts, store),
      "createPackage, deflated": await createPackage(parts, deflate),
      "streamPackage, stored blocks": await streamed(parts, selectStreamCompression("none")),
      "streamPackage, deflated": await streamed(parts, selectStreamCompression("deflate")),
    };
    for (const [written, bytes] of Object.entries(packages)) {
      const files = { bytes, "an ArrayBuffer": bytes.buffer, "a Blob": new Blob([bytes]) };
      for (const [kind, file] of Object.entries(files)) {
        const pkg = await open(file);
        for (const { path, data } of parts) {
          expect(same(await read(pkg, path), data), `${written}, from ${kind}: ${path}`).toBe(true);
        }
      }
    }
  });

  it("reads every part of every golden fixture as an independent reader does", async () => {
    const fixtures = join(here, "..", "..", "..", "test", "fixtures");
    const workbooks = readdirSync(fixtures).filter((name) => name.endsWith(".xlsx"));
    expect(workbooks.length).toBeGreaterThan(20);
    for (const name of workbooks) {
      const bytes = new Uint8Array(readFileSync(join(fixtures, name)));
      const pkg = await open(bytes);
      for (const entry of unzip(bytes)) {
        expect(same(await read(pkg, entry.name), entry.data), `${name}: ${entry.name}`).toBe(true);
      }
    }
  });

  // Written by tools that are not this library: Zip 3.0 by Info-ZIP as `zip -X -fz`, which
  // gives every entry a zip64 field and the package a zip64 end record, and zipfile of Python
  // 3.9.6 as `open(info, "w", force_zip64=True)`, which puts a zip64 field in the local headers
  // only.
  it.each(["zip64-info-zip.zip", "zip64-python.zip"])(
    "reads the zip64 package %s that another tool wrote",
    async (fixture) => {
      const pkg = await open(new Uint8Array(readFileSync(join(here, "fixtures", fixture))));
      expect(decode(await read(pkg, "[Content_Types].xml"))).toBe("<Types/>");
      expect(decode(await read(pkg, sheet))).toBe(`<worksheet>${"æøå ".repeat(500)}</worksheet>`);
      expect(await read(pkg, "xl/empty.xml")).toEqual(new Uint8Array(0));
    },
  );

  it("finds a part whatever the case of the letters A to Z in its name, and has no part it was not given", async () => {
    const named = [
      ...parts,
      { path: "xl/media/Bilde-Æ.png", data: encode("Æ") },
      { path: "xl/wor\u212Abook.xml", data: encode("named with the Kelvin sign") },
    ];
    const pkg = await open(await createPackage(named, store));
    expect(decode(await read(pkg, "[CONTENT_TYPES].XML"))).toBe("<Types/>");
    expect(decode(await read(pkg, "[content_types].xml"))).toBe("<Types/>");
    expect(decode(await read(pkg, "XL/MEDIA/bilde-Æ.PNG"))).toBe("Æ");
    expect(await read(pkg, "xl/media/bilde-æ.png")).toBeUndefined();
    expect(await read(pkg, "xl/workbook.xml")).toBeUndefined();
    expect(decode(await read(pkg, "XL/WOR\u212ABOOK.XML"))).toBe("named with the Kelvin sign");
    expect(await read(pkg, "constructor")).toBeUndefined();
    expect(await read(pkg, "__proto__")).toBeUndefined();
  });

  it("keeps a byte order mark at the start of a name: it is another name than the one without", async () => {
    const marked = [...parts, { path: "\uFEFFxl/workbook.xml", data: encode("<marked/>") }];
    const pkg = await open(await createPackage(marked, store));
    expect(await read(pkg, "xl/workbook.xml")).toBeUndefined();
    expect(decode(await read(pkg, "\uFEFFxl/workbook.xml"))).toBe("<marked/>");
    const both = [...marked, { path: "xl/workbook.xml", data: encode("<workbook/>") }];
    expect(
      decode(await read(await open(await createPackage(both, store)), "xl/workbook.xml")),
    ).toBe("<workbook/>");
  });

  it("reads bytes that are a view into a larger buffer, and bytes over shared memory", async () => {
    const bytes = await createPackage(parts, deflate);
    const larger = new Uint8Array(bytes.length + 64).fill(0x50);
    larger.set(bytes, 32);
    const shared = new Uint8Array(new SharedArrayBuffer(bytes.length));
    shared.set(bytes);
    for (const file of [larger.subarray(32, 32 + bytes.length), shared]) {
      const pkg = await open(file);
      for (const { path, data } of parts)
        expect(same(await read(pkg, path), data), path).toBe(true);
    }
  });

  it("refuses what is neither bytes nor a Blob with NOT_XLSX, the bytes of a package in another form included", async () => {
    const bytes = await createPackage(parts, store);
    const others = ["orders.xlsx", undefined, 4096, new DataView(bytes.buffer), Array.from(bytes)];
    for (const file of others) {
      await expect(open(file as never), typeof file).rejects.toMatchObject({ code: "NOT_XLSX" });
    }
  });

  it("views the bytes it is given: a stored part comes as chunks over the caller's buffer", async () => {
    const bytes = await createPackage(parts, store);
    const pkg = await open(bytes);
    const part = pkg.open(partName(sheet, "worksheet"), resolveLimits(undefined).rows().inflate)!;
    expect((await part.read())!.buffer).toBe(bytes.buffer);
  });

  it("reads a package with a comment after its end record", async () => {
    const bytes = await createPackage(parts, deflate);
    const pkg = await open(commented(bytes, encode("exported 2026-11-19")));
    for (const { path, data } of parts) expect(same(await read(pkg, path), data), path).toBe(true);
  });
});

describe("a part", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("is charged to its meter before anything of it is read, and refused when it does not fit", async () => {
    const data = parts[1]!.data;
    const file = new WatchedBlob([await createPackage(parts, deflate)]);
    const exact = resolveLimits({ inflatedBytes: data.length, sharedStringBytes: data.length });
    const pkg = await open(file, resolveLimits(undefined));
    const opened = file.reads.length;

    expect(same(await read(pkg, sheet, exact), data)).toBe(true);
    const strings = pkg.open(partName(sheet, "sharedStrings"), exact.book.stringBytes);
    expect(same(await gather(strings!), data)).toBe(true);

    const small = resolveLimits({ inflatedBytes: data.length - 1, sharedStringBytes: 100 });
    const inflate = vi.fn();
    vi.stubGlobal("DecompressionStream", inflate);
    const readSoFar = file.reads.length;
    expect(readSoFar).toBeGreaterThan(opened);
    expect(() => pkg.open(partName(sheet, "worksheet"), small.rows().inflate)).toThrow(
      expect.objectContaining({ code: "LIMIT_EXCEEDED", limit: "inflatedBytes" }),
    );
    expect(() => pkg.open(partName(sheet, "sharedStrings"), small.book.stringBytes)).toThrow(
      expect.objectContaining({ code: "LIMIT_EXCEEDED", limit: "sharedStringBytes" }),
    );
    expect(file.reads.length).toBe(readSoFar);
    expect(inflate).not.toHaveBeenCalled();
  });

  it("is read from its start by every open, each on its own", async () => {
    const pkg = await open(await createPackage(parts, deflate));
    const limits = resolveLimits(undefined);
    const name = partName("xl/worksheets/sheet2.xml", "worksheet");
    const first = pkg.open(name, limits.rows().inflate)!;
    const second = pkg.open(name, limits.rows().inflate)!;
    const head = (await first.read())!.slice();
    const whole = await gather(second);
    expect(same(whole, parts[3]!.data)).toBe(true);
    expect(same(joined([head, await gather(first)]), parts[3]!.data)).toBe(true);
  });

  it("answers reads that overlap one after the other, in the order they were asked", async () => {
    for (const compression of [store, deflate]) {
      const pkg = await open(new Blob([await createPackage(parts, compression)]));
      const part = pkg.open(
        partName("xl/media/image1.png", "worksheet"),
        resolveLimits(undefined).rows().inflate,
      )!;
      const chunks: Uint8Array[] = [];
      for (;;) {
        const asked = [part.read(), part.read(), part.read()];
        const answered = await Promise.all(asked.map(async (chunk) => (await chunk)?.slice()));
        chunks.push(...answered.filter((chunk) => chunk !== undefined));
        if (answered.includes(undefined)) break;
      }
      expect(same(joined(chunks), parts[4]!.data)).toBe(true);
    }
  });

  it("gives the platform's inflate the part as a gzip member that ends in the directory's CRC-32 and size, no more than 16 384 bytes at a time", async () => {
    const Inflate = DecompressionStream;
    const formats: string[] = [];
    const given: Uint8Array[] = [];
    vi.stubGlobal(
      "DecompressionStream",
      class {
        readonly readable: ReadableStream<Uint8Array>;
        readonly writable: WritableStream<Uint8Array<ArrayBuffer>>;

        constructor(format: CompressionFormat) {
          formats.push(format);
          const inflate = new Inflate(format);
          const writer = inflate.writable.getWriter();
          this.readable = inflate.readable;
          this.writable = new WritableStream({
            write: (chunk) => {
              given.push(chunk.slice());
              return writer.write(chunk);
            },
            close: () => writer.close(),
            abort: (reason) => writer.abort(reason),
          });
        }
      },
    );
    const bytes = await createPackage(parts, deflate);
    const image = unzip(bytes).find(({ name }) => name === "xl/media/image1.png")!;
    expect(image.compressedSize).toBeGreaterThan(6 * 16_384);
    const dataStart = image.localHeaderOffset + 30 + image.name.length;
    const trailer = new DataView(new ArrayBuffer(8));
    trailer.setUint32(0, image.crc32, true);
    trailer.setUint32(4, image.uncompressedSize, true);

    expect(same(await read(await open(bytes), image.name), image.data)).toBe(true);

    expect(formats).toEqual(["gzip"]);
    expect(Math.max(...given.map((chunk) => chunk.length))).toBeLessThanOrEqual(16_384);
    expect(
      same(
        joined(given),
        joined([
          new Uint8Array([0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 0xff]),
          bytes.subarray(dataStart, dataStart + image.compressedSize),
          new Uint8Array(trailer.buffer),
        ]),
      ),
    ).toBe(true);
  });

  it("lets go of the platform's inflate when it is closed, and when it refuses", async () => {
    const Inflate = DecompressionStream;
    const letGo: string[] = [];
    vi.stubGlobal(
      "DecompressionStream",
      class {
        readonly readable: ReadableStream<Uint8Array>;
        readonly writable: WritableStream<Uint8Array<ArrayBuffer>>;

        constructor(format: CompressionFormat) {
          const inflate = new Inflate(format);
          const reader = inflate.readable.getReader();
          const writer = inflate.writable.getWriter();
          this.readable = new ReadableStream(
            {
              pull: async (controller) => {
                const { done, value } = await reader.read();
                if (done) controller.close();
                else controller.enqueue(value);
              },
              cancel: (reason) => {
                letGo.push("its readable side cancelled");
                return reader.cancel(reason);
              },
            },
            { highWaterMark: 0 },
          );
          this.writable = new WritableStream({
            write: (chunk) => writer.write(chunk),
            close: () => writer.close(),
            abort: (reason) => {
              letGo.push("its writable side aborted");
              return writer.abort(reason);
            },
          });
        }
      },
    );
    const both = ["its readable side cancelled", "its writable side aborted"];
    const image = partName("xl/media/image1.png", "worksheet");
    const meter = () => resolveLimits(undefined).rows().inflate;

    const pkg = await open(await createPackage(parts, deflate));
    const closed = pkg.open(image, meter())!;
    expect(await closed.read()).toBeInstanceOf(Uint8Array);
    expect(letGo).toEqual([]);
    closed.close();
    await vi.waitFor(() => expect([...letGo].sort()).toEqual(both));

    // Six times what the part states: refused while the writable side still takes its bytes.
    letGo.length = 0;
    const lying = await createPackage(parts, {
      method: 8,
      compress: (data) =>
        deflate.compress(data === parts[4]!.data ? joined([data, noise(500_000)]) : data),
    });
    const refused = (await open(lying)).open(image, meter())!;
    await expect(gather(refused)).rejects.toMatchObject({ code: "INVALID_PACKAGE" });
    await vi.waitFor(() => expect([...letGo].sort()).toEqual(both));
  });

  it("refuses with a code when the platform's inflate fails a write and gives nothing to read", async () => {
    vi.stubGlobal(
      "DecompressionStream",
      class {
        readonly readable = new ReadableStream<Uint8Array>();
        readonly writable = new WritableStream<Uint8Array>({
          write() {
            throw new TypeError("the platform takes no such chunk");
          },
        });
      },
    );
    const pkg = await open(await createPackage(parts, deflate));
    const part = pkg.open(partName(sheet, "worksheet"), resolveLimits(undefined).rows().inflate)!;
    await expect(part.read()).rejects.toMatchObject({
      name: "XlsxReadError",
      code: "INVALID_PACKAGE",
    });
  });

  it("gives undefined again after its last chunk", async () => {
    const pkg = await open(await createPackage(parts, deflate));
    for (const { path } of parts) {
      const part = pkg.open(partName(path, "worksheet"), resolveLimits(undefined).rows().inflate)!;
      await gather(part);
      expect([await part.read(), await part.read()], path).toEqual([undefined, undefined]);
    }
  });

  it("stops reading the file when closed, before the first read, between two and after the last", async () => {
    for (const compression of [store, deflate]) {
      const file = new WatchedBlob([await createPackage(parts, compression)]);
      const pkg = await open(file);
      const meter = () => resolveLimits(undefined).rows().inflate;
      const name = partName("xl/media/image1.png", "worksheet");

      const unread = pkg.open(name, meter())!;
      const opened = file.reads.length;
      unread.close();
      expect(await unread.read()).toBeUndefined();
      expect(file.reads.length).toBe(opened);

      const begun = pkg.open(name, meter())!;
      expect(await begun.read()).toBeInstanceOf(Uint8Array);
      begun.close();
      begun.close();
      const closed = file.reads.length;
      expect([await begun.read(), await begun.read()]).toEqual([undefined, undefined]);
      expect(file.reads.length).toBe(closed);

      const ended = pkg.open(name, meter())!;
      await gather(ended);
      ended.close();
      expect(await ended.read()).toBeUndefined();
    }
  });

  it("takes a name made for one use and only a meter of that use", async () => {
    const pkg = await open(await createPackage(parts, store));
    const limits = resolveLimits(undefined);
    const rows = limits.rows();
    const styles: PartName<"structure"> = partName("xl/styles.xml", "structure");
    const strings = partName("xl/sharedStrings.xml", "sharedStrings");
    const worksheet = partName(sheet, "worksheet");
    const either = <First, Second>(first: First, second: Second): First | Second => first ?? second;
    const stylesOrWorksheet = either(styles, worksheet);
    const widened: PartName<"structure" | "worksheet"> = styles;
    const eitherMeter = either(limits.book.structure, rows.inflate);
    const misuses = [
      // @ts-expect-error the styles are a structure part: a worksheet's budget does not open them
      () => pkg.open(styles, rows.inflate),
      // @ts-expect-error the shared strings are not opened outside their own cap
      () => pkg.open(strings, limits.book.structure),
      // @ts-expect-error nor under a worksheet's budget
      () => pkg.open(strings, rows.inflate),
      // @ts-expect-error a worksheet is not read on the workbook's structure share
      () => pkg.open(worksheet, limits.book.structure),
      // @ts-expect-error a meter that counts something other than bytes opens nothing
      () => pkg.open(worksheet, limits.book.parts),
      // @ts-expect-error a name that is one of two opens under neither's meter
      () => pkg.open(stylesOrWorksheet, rows.inflate),
      // @ts-expect-error nor does a name made for either of two uses
      () => pkg.open(partName(sheet, either("structure", "worksheet")), rows.inflate),
      // @ts-expect-error nor a name widened to either of two, even under the meter of the use it has
      () => pkg.open(widened, limits.book.structure),
      // @ts-expect-error a meter that is one of two opens no part
      () => pkg.open(worksheet, eitherMeter),
      // @ts-expect-error a use spelled out as either of two takes a name for neither
      () => pkg.open<"structure" | "worksheet">(worksheet, rows.inflate),
      // @ts-expect-error a plain string is not a part name
      () => pkg.open(sheet, rows.inflate),
      // @ts-expect-error nor is an object written out by hand: nothing outside the module names the use
      () => pkg.open({ name: sheet }, rows.inflate),
      // @ts-expect-error a use that is none of the three is not one
      () => partName(sheet, "drawing"),
    ];
    // @ts-expect-error a plain string is not assignable to a part name either
    const plain: PartName<"worksheet"> = sheet;
    expect([plain, ...misuses]).toHaveLength(14);

    expect(pkg.open(styles, limits.book.structure)).toBeUndefined();
    expect(pkg.open(strings, limits.book.stringBytes)).toBeUndefined();
    expect(await gather(pkg.open(worksheet, rows.inflate)!)).toEqual(parts[1]!.data);
  });
});

describe("a platform without an inflate", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("reads a stored part and refuses a deflated one with NO_INFLATE", async () => {
    const stored = await createPackage(parts, store);
    const deflated = await createPackage(parts, deflate);
    // Stubbed first, so that the stub's undoing brings the global back.
    vi.stubGlobal("DecompressionStream", undefined);
    Reflect.deleteProperty(globalThis, "DecompressionStream");
    expect(typeof DecompressionStream).toBe("undefined");

    expect(same(await read(await open(stored), sheet), parts[1]!.data)).toBe(true);
    const pkg = await open(deflated);
    const part = pkg.open(partName(sheet, "worksheet"), resolveLimits(undefined).rows().inflate)!;
    await expect(part.read()).rejects.toMatchObject({ code: "NO_INFLATE" });
    await expect(part.read()).rejects.toMatchObject({ code: "NO_INFLATE" });
  });

  it("refuses a deflated part with NO_INFLATE where the platform names the inflate and has none for the format asked of it", async () => {
    const deflated = await createPackage(parts, deflate);
    vi.stubGlobal(
      "DecompressionStream",
      class {
        constructor(format: string) {
          throw new TypeError(`unsupported format ${format}`);
        }
      },
    );
    await expect(read(await open(deflated), sheet)).rejects.toMatchObject({ code: "NO_INFLATE" });
  });
});
