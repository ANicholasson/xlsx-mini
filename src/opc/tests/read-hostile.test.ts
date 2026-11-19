import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { constants, crc32, deflateRawSync, gzipSync, inflateRawSync } from "node:zlib";
import { describe, expect, it } from "vite-plus/test";
import { resolveLimits } from "../../errors/read.js";
import { createPackage, selectCompression, store } from "../index.js";
import { partName } from "../read.js";
import {
  CENTRAL,
  END,
  LOCAL,
  WatchedBlob,
  changed,
  commented,
  encode,
  gather,
  join as joined,
  noise,
  open,
  read,
  records,
} from "./reading.js";

const here = dirname(fileURLToPath(import.meta.url));

/** Milliseconds within which every package below is read or refused: none of them may hang. */
const BOUND = 2_000;

const sheet = "xl/worksheets/sheet1.xml";
const content = encode(`<worksheet>${"æøå ".repeat(500)}</worksheet>`);
const parts = [
  { path: "[Content_Types].xml", data: encode("<Types/>") },
  { path: sheet, data: content },
  { path: "xl/styles.xml", data: encode("<styleSheet/>") },
];

const deflate = selectCompression("deflate");

const refusal = (code: string) => expect.objectContaining({ name: "XlsxReadError", code });
const damaged = refusal("INVALID_PACKAGE");

/** The package the writer would write if its deflate gave `payload` for the sheet. */
function withPayload(payload: Uint8Array): Promise<Uint8Array<ArrayBuffer>> {
  return createPackage(parts, {
    method: 8,
    compress: (data) => Promise.resolve(data === content ? payload : deflateRawSync(data)),
  });
}

/** Reads the sheet of a package to its end. */
async function readSheet(file: Uint8Array | Blob): Promise<Uint8Array | undefined> {
  return read(await open(file), sheet);
}

/** What the platform's own inflate makes of a raw deflate stream: its bytes, or "refused". */
async function platformInflates(stream: Uint8Array<ArrayBuffer>): Promise<number | "refused"> {
  const inflated = new Blob([stream]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Response(inflated).arrayBuffer().then(
    (bytes) => bytes.byteLength,
    () => "refused" as const,
  );
}

describe("a package that lies about a part's content", () => {
  it("is read when it does not lie", async () => {
    expect(await readSheet(await withPayload(deflateRawSync(content)))).toEqual(content);
  });

  it(
    "is refused when the part inflates past its declared size, before a byte past it is given out",
    async () => {
      for (const more of ["<row/>", "\n"]) {
        const longer = joined([content, encode(more)]);
        const pkg = await open(await withPayload(deflateRawSync(longer)));
        const part = pkg.open(
          partName(sheet, "worksheet"),
          resolveLimits(undefined).rows().inflate,
        )!;
        let given = 0;
        const reading = async () => {
          for (let chunk = await part.read(); chunk !== undefined; chunk = await part.read()) {
            given += chunk.length;
          }
        };
        await expect(reading(), more).rejects.toEqual(damaged);
        expect(given, more).toBeLessThanOrEqual(content.length);
      }
    },
    BOUND,
  );

  it(
    "is refused at the first chunk past the declared size, with the rest of a bomb left in the file",
    async () => {
      // Four gibibytes of zeros in four megabytes: 4 096 runs of a mebibyte, then the end.
      const run = deflateRawSync(new Uint8Array(1_048_576), {
        finishFlush: constants.Z_SYNC_FLUSH,
      });
      const bomb = joined([...Array.from({ length: 4096 }, () => run), new Uint8Array([3, 0])]);
      expect(bomb.length).toBeGreaterThan(4_000_000);
      expect(inflateRawSync(joined([run, run, new Uint8Array([3, 0])])).length).toBe(2 * 1_048_576);
      const file = new WatchedBlob([await withPayload(bomb)]);
      const pkg = await open(file);
      const opened = file.reads.length;

      await expect(read(pkg, sheet)).rejects.toEqual(damaged);

      const bytesRead = file.reads
        .slice(opened)
        .reduce((total, [start, end]) => total + end - start, 0);
      expect(bytesRead, "the local header and 262 144 bytes of the bomb").toBeLessThanOrEqual(
        30 + sheet.length + 262_144,
      );
    },
    BOUND,
  );

  it(
    "is refused when the part inflates to fewer bytes than declared",
    async () => {
      const shorter = content.subarray(0, content.length - 1);
      await expect(readSheet(await withPayload(deflateRawSync(shorter)))).rejects.toEqual(damaged);
      await expect(readSheet(await withPayload(deflateRawSync(new Uint8Array(0))))).rejects.toEqual(
        damaged,
      );
    },
    BOUND,
  );

  it(
    "is refused when the part inflates to fewer bytes than declared and those have its CRC-32",
    async () => {
      // Four bytes with the CRC-32 that no bytes at all have.
      const data = new Uint8Array([0x9d, 0x0a, 0xd9, 0x6d]);
      expect([crc32(data), crc32(new Uint8Array(0))]).toEqual([0, 0]);
      const lying = (method: 0 | 8, payload: Uint8Array) =>
        createPackage([{ path: sheet, data }], {
          method,
          compress: () => Promise.resolve(payload),
        });
      expect(await readSheet(await lying(0, data))).toEqual(data);
      await expect(readSheet(await lying(0, new Uint8Array(0)))).rejects.toEqual(damaged);
      await expect(readSheet(await lying(8, deflateRawSync(new Uint8Array(0))))).rejects.toEqual(
        damaged,
      );
    },
    BOUND,
  );

  it(
    "is refused when the part has another CRC-32 than its directory entry states",
    async () => {
      const other = content.slice();
      other[100] = other[100]! ^ 1;
      await expect(readSheet(await withPayload(deflateRawSync(other)))).rejects.toEqual(damaged);

      const stored = await createPackage(parts, store);
      const flipped = changed(stored, (_, bytes) => {
        const at = records(stored).local[1]! + LOCAL.name + sheet.length + 100;
        bytes[at] = bytes[at]! ^ 1;
      });
      expect(await readSheet(stored)).toEqual(content);
      await expect(readSheet(flipped)).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused with a code, never with the platform's own error, when the deflate stream is damaged",
    async () => {
      const stream = new Uint8Array(deflateRawSync(content));
      for (let at = 4; at < 24; at++) stream[at] = stream[at]! ^ 0xff;
      expect(await platformInflates(stream)).toBe("refused");
      await expect(readSheet(await withPayload(stream))).rejects.toEqual(damaged);
      await expect(readSheet(await withPayload(noise(300)))).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused when the deflate stream ends before the part's last byte",
    async () => {
      const data = noise(60_000);
      const stream = new Uint8Array(deflateRawSync(data));
      const cut = (bytes: number) =>
        createPackage([{ path: sheet, data }], {
          method: 8,
          compress: () => Promise.resolve(stream.subarray(0, stream.length - bytes)),
        });
      expect(await readSheet(await cut(0))).toEqual(data);
      for (const bytes of [1, 5, 30_000, stream.length]) {
        await expect(readSheet(await cut(bytes)), `${bytes} bytes short`).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is refused when the deflate stream lacks only its end, with all of the part's content in what there is",
    async () => {
      const whole = new Uint8Array(deflateRawSync(content));
      const stream = whole.slice(0, whole.length - 1);
      expect(inflateRawSync(stream, { finishFlush: constants.Z_SYNC_FLUSH })).toEqual(
        Buffer.from(content),
      );
      await expect(readSheet(await withPayload(stream))).rejects.toEqual(damaged);

      const empty = (payload: Uint8Array) =>
        createPackage([{ path: sheet, data: new Uint8Array(0) }], {
          method: 8,
          compress: () => Promise.resolve(payload),
        });
      expect(await readSheet(await empty(deflateRawSync(new Uint8Array(0))))).toEqual(
        new Uint8Array(0),
      );
      await expect(readSheet(await empty(new Uint8Array(0)))).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused when bytes follow the end of the deflate stream where the platform's inflate reports them, and is the stream's content where it stops at the end, as Node 20 and 22 do",
    async () => {
      const whole = deflateRawSync(content);
      const after = {
        text: encode("more"),
        "one zero byte": new Uint8Array(1),
        "the stream once more": whole,
        "a block of noise": noise(40_000),
        "the stream's gzip trailer and a gzip member": joined([
          gzipSync(content).subarray(-8),
          gzipSync(encode("<row/>")),
        ]),
      };
      for (const [what, bytes] of Object.entries(after)) {
        const stream = joined([whole, bytes]);
        const file = await withPayload(stream);
        if ((await platformInflates(stream)) === "refused") {
          await expect(readSheet(file), what).rejects.toEqual(damaged);
        } else {
          expect(await readSheet(file), what).toEqual(content);
        }
      }
    },
    BOUND,
  );

  // Written by a verifier of this reader. The sheet's deflate stream ends early and a gzip member
  // follows it: after an empty stream and eight zero bytes in the first, after the stream and the
  // gzip trailer of its 95 bytes in the second. Python's zipfile and Info-ZIP unzip read the first
  // stream, 0 and 95 bytes, where the directory states 90 and 140.
  it.each(["empty-stream-then-gzip-member.zip", "stream-trailer-then-gzip-member.zip"])(
    "is refused when a gzip member follows the end of the deflate stream: %s",
    async (fixture) => {
      const bytes = new Uint8Array(readFileSync(join(here, "fixtures", fixture)));
      await expect(readSheet(bytes)).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "repeats its refusal on every later read, and never ends the part after one",
    async () => {
      const longer = joined([content, noise(40_000)]);
      const cases = [
        deflateRawSync(longer),
        deflateRawSync(content.subarray(1)),
        noise(300),
        deflateRawSync(content).subarray(0, -1),
      ];
      for (const [index, payload] of cases.entries()) {
        const pkg = await open(await withPayload(payload));
        const part = pkg.open(
          partName(sheet, "worksheet"),
          resolveLimits(undefined).rows().inflate,
        )!;
        const first = await gather(part).then(
          () => undefined,
          (error: unknown) => error,
        );
        expect(first, `case ${index}`).toEqual(damaged);
        await expect(part.read(), `case ${index}`).rejects.toBe(first);
        part.close();
        await expect(part.read(), `case ${index}`).rejects.toBe(first);
      }
    },
    BOUND,
  );
});

describe("a package whose headers lie", () => {
  it(
    "is refused when a part read is encrypted, or compressed by another method than deflate, whether the directory alone says so or the local header with it",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { central, local } = records(bytes);
      const stating = (field: "flags" | "method", value: number, inBoth: boolean) =>
        changed(bytes, (view) => {
          view.setUint16(central[1]! + CENTRAL[field], value, true);
          if (inBoth) view.setUint16(local[1]! + LOCAL[field], value, true);
        });
      for (const inBoth of [false, true]) {
        const encrypted = stating("flags", 1, inBoth);
        await expect(readSheet(encrypted), `encrypted, ${inBoth}`).rejects.toEqual(damaged);
        expect(await read(await open(encrypted), "xl/styles.xml")).toEqual(parts[2]!.data);
        for (const method of [1, 9, 12, 14, 93, 99]) {
          const other = stating("method", method, inBoth);
          await expect(readSheet(other), `method ${method}, ${inBoth}`).rejects.toEqual(damaged);
          expect(await read(await open(other), "xl/styles.xml")).toEqual(parts[2]!.data);
        }
      }
    },
    BOUND,
  );

  it(
    "is refused when a part read has a local header that disagrees with the directory",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { local } = records(bytes);
      const at = local[1]!;
      const disagreeing = {
        "names another part": changed(bytes, (_, copy) => {
          copy[at + LOCAL.name + 3] = 0x58;
        }),
        "has a shorter name": changed(bytes, (view) =>
          view.setUint16(at + LOCAL.nameLength, sheet.length - 1, true),
        ),
        "has a longer name": changed(bytes, (view) =>
          view.setUint16(at + LOCAL.nameLength, sheet.length + 1, true),
        ),
        "is no local header": changed(bytes, (view) =>
          view.setUint32(at + LOCAL.signature, 0x02014b50, true),
        ),
        "puts the data into the directory": changed(bytes, (view) =>
          view.setUint16(at + LOCAL.extraLength, 0xffff, true),
        ),
        "states another method": changed(bytes, (view) =>
          view.setUint16(at + LOCAL.method, 0, true),
        ),
        "says the part is encrypted": changed(bytes, (view) =>
          view.setUint16(at + LOCAL.flags, 1, true),
        ),
        "states another CRC-32": changed(bytes, (view) =>
          view.setUint32(at + LOCAL.crc32, view.getUint32(at + LOCAL.crc32, true) ^ 1, true),
        ),
        "states another compressed size": changed(bytes, (view) =>
          view.setUint32(
            at + LOCAL.compressedSize,
            view.getUint32(at + LOCAL.compressedSize, true) - 1,
            true,
          ),
        ),
        "states another size": changed(bytes, (view) =>
          view.setUint32(at + LOCAL.size, content.length + 1, true),
        ),
        "states no CRC-32 and no data descriptor to hold one": changed(bytes, (view) =>
          view.setUint32(at + LOCAL.crc32, 0, true),
        ),
        "states no size and no data descriptor to hold one": changed(bytes, (view) =>
          view.setUint32(at + LOCAL.size, 0, true),
        ),
      };
      for (const [how, file] of Object.entries(disagreeing)) {
        await expect(readSheet(file), how).rejects.toEqual(damaged);
        expect(await read(await open(file), "xl/styles.xml"), how).toEqual(parts[2]!.data);
      }
    },
    BOUND,
  );

  it(
    "is refused when only the local header of a stored part states another method, or says the part is encrypted",
    async () => {
      const bytes = await createPackage(parts, store);
      const at = records(bytes).local[1]!;
      const other = changed(bytes, (view) => view.setUint16(at + LOCAL.method, 99, true));
      await expect(readSheet(other)).rejects.toEqual(damaged);
      const encrypted = changed(bytes, (view) => view.setUint16(at + LOCAL.flags, 1, true));
      await expect(readSheet(encrypted)).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is read when a local header leaves its CRC-32 and its sizes to the directory: zero where a data descriptor follows the data, all ones where zip64 moved a size",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const at = records(bytes).local[1]!;
      const withDescriptor = changed(bytes, (view) => {
        view.setUint16(at + LOCAL.flags, 8, true);
        view.setUint32(at + LOCAL.crc32, 0, true);
        view.setUint32(at + LOCAL.compressedSize, 0, true);
        view.setUint32(at + LOCAL.size, 0, true);
      });
      expect(await readSheet(withDescriptor)).toEqual(content);
      const moved = changed(bytes, (view) => {
        view.setUint32(at + LOCAL.compressedSize, 0xffffffff, true);
        view.setUint32(at + LOCAL.size, 0xffffffff, true);
      });
      expect(await readSheet(moved)).toEqual(content);
    },
    BOUND,
  );

  it(
    "is refused when a local header places a part's data behind the directory, although the part is there",
    async () => {
      const bytes = await createPackage(parts, store);
      const local = records(bytes).local[1]!;
      const dataStart = local + LOCAL.name + sheet.length;
      const withTheSheetAsComment = commented(bytes, content);
      expect(await readSheet(withTheSheetAsComment)).toEqual(content);

      const moved = changed(withTheSheetAsComment, (view) =>
        view.setUint16(local + LOCAL.extraLength, bytes.length - dataStart, true),
      );
      await expect(readSheet(moved)).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused when two parts have one name, whatever the case of the letters A to Z",
    async () => {
      const twice = [...parts, { path: sheet, data: encode("<worksheet/>") }];
      await expect(open(await createPackage(twice, store))).rejects.toEqual(damaged);
      const cased = [...parts, { path: "XL/Worksheets/Sheet1.xml", data: encode("<worksheet/>") }];
      await expect(open(await createPackage(cased, deflate))).rejects.toEqual(damaged);

      const kelvin = [
        ...parts,
        { path: "xl/wor\u212Asheets/sheet1.xml", data: encode("<worksheet/>") },
      ];
      expect(await readSheet(await createPackage(kelvin, store))).toEqual(content);
    },
    BOUND,
  );

  it(
    "is refused when its directory is out of bounds",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { end } = records(bytes);
      const view = new DataView(bytes.buffer);
      const offset = view.getUint32(end + END.offset, true);
      const size = view.getUint32(end + END.size, true);
      const lies = {
        "starts past the end of the file": [0x7fffffff, size],
        "starts one byte late": [offset + 1, size],
        "starts one byte early": [offset - 1, size],
        "is one byte smaller": [offset, size - 1],
        "is larger than the file": [offset, bytes.length + 1],
        "is as large as a 32-bit field can say": [offset, 0xffffffff],
      };
      for (const [how, [statedOffset, statedSize]] of Object.entries(lies)) {
        const file = changed(bytes, (copy) => {
          copy.setUint32(end + END.offset, statedOffset!, true);
          copy.setUint32(end + END.size, statedSize!, true);
        });
        await expect(open(file), how).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is walked by the size of its directory, whatever number of entries it states, and refused at what is no entry",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { end, central } = records(bytes);
      const view = new DataView(bytes.buffer);
      const size = view.getUint32(end + END.size, true);

      for (const entries of [0, 2, 65_535]) {
        const miscounted = changed(bytes, (copy) =>
          copy.setUint16(end + END.entries, entries, true),
        );
        expect(await readSheet(miscounted), `${entries} entries stated`).toEqual(content);
      }
      const unnamed = [{ path: "", data: encode("<unnamed/>") }, ...parts];
      const pkg = await open(await createPackage(unnamed, deflate));
      expect(await read(pkg, ""), "an entry without a name").toEqual(unnamed[0]!.data);
      expect(await read(pkg, sheet), "the entry after one without a name").toEqual(content);
      const lies = {
        "a directory that starts inside an entry": changed(bytes, (copy) => {
          copy.setUint32(end + END.size, size - 10, true);
          copy.setUint32(end + END.offset, copy.getUint32(end + END.offset, true) + 10, true);
        }),
        "an entry whose name runs past the directory": changed(bytes, (copy) =>
          copy.setUint16(central[2]! + CENTRAL.nameLength, 0xffff, true),
        ),
        "an entry whose extra field runs past the directory": changed(bytes, (copy) =>
          copy.setUint16(central[2]! + CENTRAL.extraLength, 1, true),
        ),
        "an entry that is none": changed(bytes, (copy) => copy.setUint32(central[1]!, 0, true)),
      };
      for (const [how, file] of Object.entries(lies)) {
        await expect(open(file), how).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is refused when a part, read or not, lies outside it",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { end, central } = records(bytes);
      const directoryStart = new DataView(bytes.buffer).getUint32(end + END.offset, true);
      const lies = {
        "a part that starts in the directory": changed(bytes, (view) =>
          view.setUint32(central[2]! + CENTRAL.offset, directoryStart, true),
        ),
        "a part of two gibibytes compressed": changed(bytes, (view) =>
          view.setUint32(central[2]! + CENTRAL.compressedSize, 0x80000000, true),
        ),
        "a part one byte too long": changed(bytes, (view) =>
          view.setUint32(
            central[2]! + CENTRAL.compressedSize,
            view.getUint32(central[2]! + CENTRAL.compressedSize, true) + 1,
            true,
          ),
        ),
      };
      for (const [how, file] of Object.entries(lies)) {
        await expect(open(file), how).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );
});

describe("a file that is cut short, or no package", () => {
  it(
    "is refused as a damaged package when it starts as one and its end is missing",
    async () => {
      const bytes = await createPackage(parts, deflate);
      for (const length of [bytes.length - 1, bytes.length - 10, bytes.length >> 1, 30, 4]) {
        await expect(open(bytes.subarray(0, length)), `${length} bytes`).rejects.toEqual(damaged);
      }
      await expect(open(joined([bytes, new Uint8Array(1)]))).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused as a damaged package when bytes are missing from it or stand in front of it",
    async () => {
      const bytes = await createPackage(parts, store);
      const { local } = records(bytes);
      const others = {
        "its first byte is missing": bytes.subarray(1),
        "a byte of its middle is missing": joined([
          bytes.subarray(0, local[1]!),
          bytes.subarray(local[1]! + 1),
        ]),
        "one byte stands in front of it": joined([new Uint8Array(1), bytes]),
        "four kibibytes stand in front of it": joined([noise(4096), bytes]),
        "another package stands in front of it": joined([bytes, bytes]),
      };
      for (const [how, file] of Object.entries(others)) {
        await expect(open(file), how).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is refused with NOT_XLSX when it is a compound file, as an old .xls and an encrypted workbook are",
    async () => {
      const compound = new Uint8Array(4096);
      compound.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
      await expect(open(compound)).rejects.toEqual(refusal("NOT_XLSX"));
      await expect(open(new Blob([compound]))).rejects.toEqual(refusal("NOT_XLSX"));
    },
    BOUND,
  );

  it(
    "is refused with NOT_XLSX when it is anything else",
    async () => {
      const others = {
        "no bytes": new Uint8Array(0),
        "three bytes": new Uint8Array([0x50, 0x4b, 0x03]),
        "a CSV file": encode("Order,Total\n1001,250.00\n"),
        "a PNG file": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]),
        noise: noise(100_000),
        "an end record cut by one byte": new Uint8Array([0x50, 0x4b, 5, 6, ...new Uint8Array(17)]),
      };
      for (const [what, file] of Object.entries(others)) {
        await expect(open(file), what).rejects.toEqual(refusal("NOT_XLSX"));
        await expect(open(new Blob([file])), what).rejects.toEqual(refusal("NOT_XLSX"));
      }
    },
    BOUND,
  );

  it(
    "is opened when it is a package without parts, and holds none",
    async () => {
      const pkg = await open(await createPackage([], store));
      expect(await read(pkg, sheet)).toBeUndefined();
    },
    BOUND,
  );

  it(
    "is refused when a second end record fits in its comment, whether or not that record ends the file, and read when the comment only ends in an end signature",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const signature = new Uint8Array([0x50, 0x4b, 5, 6]);
      const record = (size: number, offset: number) =>
        changed(joined([signature, new Uint8Array(18)]), (view) => {
          view.setUint32(END.size, size, true);
          view.setUint32(END.offset, offset, true);
        });
      const refused = {
        "a record of nothing but zeros": record(0, 0),
        "a record that states an empty directory where it stands itself": record(0, bytes.length),
        "text, then a record": joined([encode("exported "), record(0, 0)]),
        "a record that bytes follow": joined([record(0, 0), new Uint8Array(22)]),
      };
      for (const [what, comment] of Object.entries(refused)) {
        await expect(open(commented(bytes, comment)), what).rejects.toEqual(damaged);
      }
      const read = {
        "no end signature": new Uint8Array(44),
        "an end signature with no room for a record behind it": joined([
          new Uint8Array(40),
          signature,
        ]),
      };
      for (const [what, comment] of Object.entries(read)) {
        expect(await readSheet(commented(bytes, comment)), what).toEqual(content);
      }
    },
    BOUND,
  );
});

describe("a zip64 package that lies", () => {
  const fixture = new Uint8Array(readFileSync(join(here, "fixtures", "zip64-info-zip.zip")));

  /** Where every record with this signature starts. The fixture holds none by accident. */
  function find(signature: number): number[] {
    const view = new DataView(fixture.buffer);
    const found: number[] = [];
    for (let at = 0; at + 4 <= fixture.length; at++) {
      if (view.getUint32(at, true) === signature) found.push(at);
    }
    return found;
  }
  const central = find(0x02014b50);
  const [record] = find(0x06064b50);
  const [locator] = find(0x07064b50);
  /** The zip64 field of an entry: its header, then the entry's size as eight bytes. */
  const fieldOf = (entry: number) =>
    entry + CENTRAL.name + new DataView(fixture.buffer).getUint16(entry + CENTRAL.nameLength, true);

  it("is the package another tool wrote, with three entries and one zip64 end record", () => {
    expect([central.length, record, locator]).toEqual([3, 492, 548]);
  });

  it(
    "is read when the value its entry keeps in the zip64 field is the compressed size, or where the part starts",
    async () => {
      const view = new DataView(fixture.buffer);
      const entry = central[1]!;
      const compressedSize = view.getUint32(entry + CENTRAL.compressedSize, true);
      const size = view.getUint32(fieldOf(entry) + 4, true);
      const offset = view.getUint32(entry + CENTRAL.offset, true);
      const moved = {
        "the compressed size": changed(fixture, (copy) => {
          copy.setUint32(entry + CENTRAL.size, size, true);
          copy.setUint32(entry + CENTRAL.compressedSize, 0xffffffff, true);
          copy.setUint32(fieldOf(entry) + 4, compressedSize, true);
        }),
        "where the part starts": changed(fixture, (copy) => {
          copy.setUint32(entry + CENTRAL.size, size, true);
          copy.setUint32(entry + CENTRAL.offset, 0xffffffff, true);
          copy.setUint32(fieldOf(entry) + 4, offset, true);
        }),
      };
      for (const [what, file] of Object.entries(moved)) {
        expect(await readSheet(file), what).toEqual(content);
      }
    },
    BOUND,
  );

  it(
    "is read when other fields stand in front of an entry's zip64 field, a thousand of no length among them: the search moves on from each",
    async () => {
      const at = fieldOf(central[1]!);
      const others = {
        "one field of two bytes": new Uint8Array([0xfe, 0xca, 2, 0, 7, 7]),
        "a thousand fields of no length": joined(
          Array.from({ length: 1000 }, () => new Uint8Array([0xfe, 0xca, 0, 0])),
        ),
      };
      for (const [what, other] of Object.entries(others)) {
        const file = joined([fixture.subarray(0, at), other, fixture.subarray(at)]);
        const view = new DataView(file.buffer);
        const grown = (field: number) =>
          view.setUint32(field, view.getUint32(field, true) + other.length, true);
        view.setUint16(central[1]! + CENTRAL.extraLength, 12 + other.length, true);
        // The directory is larger, in both end records, and the zip64 one starts later.
        grown(record! + other.length + 40);
        grown(locator! + other.length + 8);
        grown(file.length - END.bytes + END.size);
        expect(await readSheet(file), what).toEqual(content);
      }
    },
    BOUND,
  );

  it(
    "is refused when its locator leads to no zip64 end record",
    async () => {
      const file = changed(fixture, (view) => view.setUint32(record!, 0, true));
      await expect(open(file)).rejects.toEqual(damaged);
      const short = fixture.subarray(locator! - 10);
      await expect(open(short)).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused when its locator points away from the zip64 end record in front of it",
    async () => {
      for (const place of [record! - 1, record! + 1, 0]) {
        const file = changed(fixture, (view) => view.setUint32(locator! + 8, place, true));
        await expect(open(file), `at ${place}`).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is refused when its end record states another size or place for the directory than the zip64 one, and read when it states the same or leaves it with all ones",
    async () => {
      const end = fixture.length - END.bytes;
      const view = new DataView(fixture.buffer);
      const size = Number(view.getBigUint64(record! + 40, true));
      const offset = Number(view.getBigUint64(record! + 48, true));
      expect([
        view.getUint32(end + END.size, true),
        view.getUint32(end + END.offset, true),
      ]).toEqual([size, 0xffffffff]);
      const stating = (field: number, value: number) =>
        changed(fixture, (copy) => copy.setUint32(end + field, value, true));

      expect(await readSheet(stating(END.size, 0xffffffff))).toEqual(content);
      expect(await readSheet(stating(END.offset, offset))).toEqual(content);
      await expect(open(stating(END.size, size - 1))).rejects.toEqual(damaged);
      await expect(open(stating(END.size, 0))).rejects.toEqual(damaged);
      await expect(open(stating(END.offset, offset + 1))).rejects.toEqual(damaged);
      await expect(open(stating(END.offset, 0))).rejects.toEqual(damaged);
    },
    BOUND,
  );

  it(
    "is refused when it states a size or a place that no number holds exactly",
    async () => {
      const upperHalves = {
        "the directory's size": record! + 44,
        "the directory's place": record! + 52,
        "a part's size": fieldOf(central[1]!) + 8,
      };
      for (const [what, at] of Object.entries(upperHalves)) {
        const exact = changed(fixture, (view) => view.setUint32(at, 0x1fffff, true));
        await expect(
          open(exact).then((pkg) => read(pkg, sheet)),
          what,
        ).rejects.toEqual(what === "a part's size" ? refusal("LIMIT_EXCEEDED") : damaged);
        const inexact = changed(fixture, (view) => view.setUint32(at, 0x200000, true));
        await expect(open(inexact), what).rejects.toEqual(damaged);
      }
    },
    BOUND,
  );

  it(
    "is refused when an entry's zip64 field is too short for what it must hold, and read as stated when the entry has none",
    async () => {
      const entry = central[1]!;
      const twoValues = changed(fixture, (view) =>
        view.setUint32(entry + CENTRAL.compressedSize, 0xffffffff, true),
      );
      await expect(open(twoValues)).rejects.toEqual(damaged);
      const shortField = changed(fixture, (view) => view.setUint16(fieldOf(entry) + 2, 4, true));
      await expect(open(shortField)).rejects.toEqual(damaged);
      // The last entry: what follows its field is the end of the directory, which reads as zeros.
      const last = central[2]!;
      const pastTheEntry = changed(fixture, (view) => {
        view.setUint32(last + CENTRAL.compressedSize, 0xffffffff, true);
        view.setUint16(fieldOf(last) + 2, 16, true);
      });
      await expect(open(pastTheEntry)).rejects.toEqual(damaged);

      const noField = changed(fixture, (view) => view.setUint16(fieldOf(entry), 0xcafe, true));
      await expect(readSheet(noField)).rejects.toEqual(refusal("LIMIT_EXCEEDED"));
    },
    BOUND,
  );
});

describe("the directory of a package", () => {
  it(
    "is charged by its size before it is read, and every entry before it is kept",
    async () => {
      const bytes = await createPackage(parts, deflate);
      const { end } = records(bytes);
      const size = new DataView(bytes.buffer).getUint32(end + END.size, true);

      const fitting = new WatchedBlob([bytes]);
      const pkg = await open(
        fitting,
        resolveLimits({ inflatedBytes: size * 8, parts: parts.length }),
      );
      expect(fitting.reads).toEqual([
        [0, bytes.length],
        [end - size, end],
      ]);
      expect(
        pkg.open(partName(sheet, "worksheet"), resolveLimits(undefined).rows().inflate),
      ).toBeDefined();

      const tooLarge = new WatchedBlob([bytes]);
      await expect(open(tooLarge, resolveLimits({ inflatedBytes: size * 8 - 8 }))).rejects.toEqual(
        expect.objectContaining({ code: "LIMIT_EXCEEDED", limit: "inflatedBytes" }),
      );
      expect(tooLarge.reads).toEqual([[0, bytes.length]]);

      await expect(open(bytes, resolveLimits({ parts: parts.length - 1 }))).rejects.toEqual(
        expect.objectContaining({ code: "LIMIT_EXCEEDED", limit: "parts" }),
      );
    },
    BOUND,
  );

  it(
    "is found by a read of the end of the file that a comment cannot lengthen",
    async () => {
      const large = [{ path: "xl/media/image1.png", data: noise(200_000) }, ...parts];
      const bytes = await createPackage(large, deflate);
      const file = commented(bytes, new Uint8Array(0xffff));
      const watched = new WatchedBlob([file]);
      expect(await readSheet(watched)).toEqual(content);
      expect(watched.reads[0]).toEqual([file.length - 22 - 0xffff - 76, file.length]);
    },
    BOUND,
  );
});

describe("a Blob that cannot be read", () => {
  it("fails with its own error, which is none of the reader's", async () => {
    class Unreadable extends WatchedBlob {
      constructor(
        parts: BlobPart[],
        private readonly after: number,
      ) {
        super(parts);
      }

      override slice(start?: number, end?: number): Blob {
        if (this.reads.length >= this.after) throw new DOMException("gone", "NotReadableError");
        return super.slice(start, end);
      }
    }
    const bytes = await createPackage(parts, deflate);
    const gone = expect.objectContaining({ name: "NotReadableError" });
    await expect(open(new Unreadable([bytes], 0))).rejects.toEqual(gone);
    await expect(open(new Unreadable([bytes], 1))).rejects.toEqual(gone);
    await expect(readSheet(new Unreadable([bytes], 2))).rejects.toEqual(gone);
    await expect(readSheet(new Unreadable([bytes], 3))).rejects.toEqual(gone);
    expect(await readSheet(new Unreadable([bytes], 4))).toEqual(content);
  });
});
