import { describe, expect, expectTypeOf, it } from "vite-plus/test";
import { type Cell, type CellValue, type Coerce, writeXlsx } from "../src/index.js";
import { readPart } from "./helpers/unzip.js";

/** A consumer's own type, made writable by augmenting Register (see the README). */
class Money {
  constructor(readonly cents: number) {}
}

declare module "../src/index.js" {
  interface Register {
    cellValue: Money | bigint;
  }
}

describe("Register augmentation", () => {
  it("widens CellValue and types coerce to exactly the registered types", () => {
    expectTypeOf<Money>().toMatchTypeOf<CellValue>();
    expectTypeOf<bigint>().toMatchTypeOf<CellValue>();
    expectTypeOf<Map<string, string>>().not.toMatchTypeOf<CellValue>();
    expectTypeOf<Parameters<Coerce>[0]>().toEqualTypeOf<
      Date | import("../src/index.js").TemporalLike | Money | bigint
    >();
    const cell: Cell = new Money(1);
    expect(cell).toBeInstanceOf(Money);
  });

  it("round-trips a registered type through coerce at runtime", async () => {
    const bytes = await writeXlsx(
      { sheets: [{ name: "S", rows: [[new Money(1234), 5n]] }] },
      {
        coerce: (value) =>
          value instanceof Money
            ? value.cents / 100
            : typeof value === "bigint"
              ? Number(value)
              : value,
      },
    );
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="A1"><v>12.34</v></c><c r="B1"><v>5</v></c>');
  });
});
