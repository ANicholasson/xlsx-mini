import { describe, expect, it } from "vite-plus/test";
import { XlsxMiniError } from "../index.js";
import { XlsxReadError, resolveLimits } from "../read.js";

function exceeded(limit: string) {
  return expect.objectContaining({ code: "LIMIT_EXCEEDED", limit });
}

describe("XlsxReadError", () => {
  it("carries a code and where in the file, and reads like an Error", () => {
    const error = new XlsxReadError("INVALID_CELL", "a cell out of order", {
      sheet: "Orders",
      cell: "C41",
    });
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe("XlsxReadError");
    expect(error.code).toBe("INVALID_CELL");
    expect([error.sheet, error.cell, error.limit, error.header]).toEqual([
      "Orders",
      "C41",
      undefined,
      undefined,
    ]);
    expect(error.message).toBe("a cell out of order");
  });

  it("is not the writer's error: one class per direction", () => {
    expect(new XlsxReadError("NOT_XLSX", "not a workbook")).not.toBeInstanceOf(XlsxMiniError);
  });
});

describe("resolveLimits", () => {
  it("refuses a limit that is not a positive finite number with INVALID_LIMIT naming it", () => {
    const names = [
      "inflatedBytes",
      "cellPositions",
      "rowCharacters",
      "sharedStrings",
      "sharedStringBytes",
      "parts",
    ] as const;
    const notPositiveFinite: unknown[] = [0, -1, -0.5, Number.NaN, Infinity, -Infinity, "5", null];
    for (const limit of names) {
      for (const value of notPositiveFinite) {
        expect(
          () => resolveLimits({ [limit]: value as number }),
          `${limit}: ${String(value)}`,
        ).toThrow(expect.objectContaining({ code: "INVALID_LIMIT", limit }));
      }
    }
    expect(() => resolveLimits({ parts: Number.NaN })).toThrow(XlsxReadError);
  });

  it("holds each limit at the default ADR 0032 gives it", () => {
    const limits = resolveLimits(undefined);
    const rows = limits.rows();
    const defaults: [string, { charge(amount: number): void }, number][] = [
      ["inflatedBytes", rows.inflate, 134_217_728],
      ["cellPositions", rows.positions, 5_000_000],
      ["rowCharacters", rows.rowText, 4_194_304],
      ["sharedStrings", limits.book.strings, 1_000_000],
      ["sharedStringBytes", limits.book.stringBytes, 33_554_432],
      ["parts", limits.book.parts, 100_000],
      ["inflatedBytes", limits.book.structure, 134_217_728 / 8],
    ];
    for (const [limit, meter, allowance] of defaults) {
      expect(() => meter.charge(allowance), limit).not.toThrow();
      expect(() => meter.charge(1), limit).toThrow(exceeded(limit));
    }
  });

  it("takes a caller's value in place of a default, a fraction included", () => {
    const limits = resolveLimits({ parts: 2.5, sharedStrings: 10 });
    limits.book.parts.charge(2.5);
    expect(() => limits.book.parts.charge(0.5)).toThrow(exceeded("parts"));
    limits.book.strings.charge(10);
    expect(() => limits.book.strings.charge(1)).toThrow(exceeded("sharedStrings"));
  });
});

describe("a meter", () => {
  it("refuses a charge that does not fit and takes nothing for it", () => {
    const { parts } = resolveLimits({ parts: 10 }).book;
    parts.charge(6);
    expect(() => parts.charge(5)).toThrow(exceeded("parts"));
    parts.charge(4);
    expect(() => parts.charge(1)).toThrow(exceeded("parts"));
  });

  it("never fits NaN or a negative amount, and takes or gives back nothing for one", () => {
    const limits = resolveLimits({
      parts: 10,
      inflatedBytes: 80,
      sharedStringBytes: 75,
      rowCharacters: 10,
    });
    const { parts, structure, stringBytes } = limits.book;
    const { rowText } = limits.rows();
    const meters = [
      ["parts", parts],
      ["inflatedBytes", structure],
      ["rowCharacters", rowText],
    ] as const;
    for (const [limit, meter] of meters) {
      for (const amount of [Number.NaN, -5, -Infinity]) {
        expect(() => meter.charge(amount), `${limit}: ${amount}`).toThrow(exceeded(limit));
      }
      meter.charge(10);
      expect(() => meter.charge(1), limit).toThrow(exceeded(limit));
    }
    stringBytes.charge(70);
    expect(() => stringBytes.charge(1), "opening the workbook got nothing back").toThrow(
      exceeded("inflatedBytes"),
    );
  });

  it("charges opening the workbook once, and a charge it refuses takes nothing from a share", () => {
    const { book } = resolveLimits({ inflatedBytes: 80, sharedStringBytes: 75 });
    expect(() => book.stringBytes.charge(76)).toThrow(exceeded("sharedStringBytes"));
    book.stringBytes.charge(72);
    expect(() => book.structure.charge(9)).toThrow(exceeded("inflatedBytes"));
    book.structure.charge(8);
    expect(() => book.structure.charge(1)).toThrow(exceeded("inflatedBytes"));
    expect(() => book.stringBytes.charge(1)).toThrow(exceeded("inflatedBytes"));
  });

  it("gives every rows iteration meters of its own, apart from opening the workbook", () => {
    const limits = resolveLimits({ inflatedBytes: 80, cellPositions: 3 });
    limits.book.stringBytes.charge(80);
    for (const rows of [limits.rows(), limits.rows()]) {
      rows.inflate.charge(80);
      rows.positions.charge(3);
      expect(() => rows.positions.charge(1)).toThrow(exceeded("cellPositions"));
    }
  });

  it("starts the text of a row again for every row", () => {
    const { rowText } = resolveLimits({ rowCharacters: 5 }).rows();
    rowText.charge(5);
    expect(() => rowText.charge(1)).toThrow(exceeded("rowCharacters"));
    rowText.next();
    rowText.charge(5);
    expect(() => rowText.charge(1)).toThrow(exceeded("rowCharacters"));
  });

  it("cannot be written out by hand, copied by spread or used for another kind of part", () => {
    const limits = resolveLimits(undefined);
    const { parts, structure } = limits.book;
    const { inflate } = limits.rows();
    // @ts-expect-error a meter written out by hand lacks the brand
    const handmade: typeof parts = { limit: "parts", charge() {} };
    // @ts-expect-error naming the brand does not make one: it is private to its class
    const named: typeof parts = { limit: "parts", charge() {}, metered: true };
    // @ts-expect-error a copy by spread drops the brand
    const copied: typeof parts = { ...parts, charge() {} };
    // @ts-expect-error the workbook's structure share does not pay for a worksheet
    const worksheet: typeof inflate = structure;
    // @ts-expect-error nor a worksheet's budget for the structure
    const lists: typeof structure = inflate;
    expect([handmade, named, copied, worksheet, lists]).toHaveLength(5);
  });
});
