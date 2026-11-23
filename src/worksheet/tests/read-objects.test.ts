import { describe, expect, expectTypeOf, it } from "vite-plus/test";
import { XlsxReadError } from "../../errors/read.js";
import { type ObjectColumn, sheetFromObjects } from "../index.js";
import {
  type ReadColumn,
  type ReadObject,
  type ReadRow,
  type ReadValue,
  objectsFromRows,
} from "../read.js";

type Rows = readonly (readonly [number, ...ReadValue[]])[];

/** Rows as `rows` yields them, iterable again, with what each iteration did written to `log`. */
function sheet(rows: Rows) {
  const log: string[] = [];
  return {
    log,
    async *[Symbol.asyncIterator](): AsyncGenerator<ReadRow> {
      log.push("start");
      try {
        for (const [number, ...cells] of rows) {
          log.push(`row ${number}`);
          yield { number, cells };
        }
      } finally {
        log.push("end");
      }
    },
  };
}

type Read = [number, Readonly<Record<string, ReadValue>>][];

async function readObjects<Key extends string>(
  rows: AsyncIterable<ReadRow>,
  columns: readonly ReadColumn<Key>[],
): Promise<Read> {
  const read: Read = [];
  for await (const { number, values } of objectsFromRows(rows, columns)) {
    read.push([number, values]);
  }
  return read;
}

async function outcomeOf<Key extends string>(
  rows: AsyncIterable<ReadRow>,
  columns: readonly ReadColumn<Key>[],
): Promise<Read | object> {
  try {
    return await readObjects(rows, columns);
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    const { code, header, sheet: named, cell, limit } = error;
    return { code, header, sheet: named, cell, limit };
  }
}

const missing = (header: string | undefined) => ({
  code: "MISSING_HEADER",
  header,
  sheet: undefined,
  cell: undefined,
  limit: undefined,
});

describe("objectsFromRows", () => {
  it("takes the first row as the header row whatever its number, and gives every later row every declared key in the order of the columns, null where the row ends before a column", async () => {
    const placed = new Date(Date.UTC(2026, 0, 2, 9, 30));
    const rows = sheet([
      [2, "Order", "Total", "Placed", "Note"],
      [3, "A-1", 10.5],
      [7, "A-2", null, placed, "late"],
    ]);
    const columns = [
      { key: "total", header: "Total" },
      { key: "id", header: "Order" },
      { key: "Note" },
      { key: "placed", header: "Placed" },
    ] as const;
    const read = await readObjects(rows, columns);
    expect(read).toEqual([
      [3, { total: 10.5, id: "A-1", Note: null, placed: null }],
      [7, { total: null, id: "A-2", Note: "late", placed }],
    ]);
    expect(read.map(([, values]) => Object.keys(values))).toEqual([
      ["total", "id", "Note", "placed"],
      ["total", "id", "Note", "placed"],
    ]);
    expect(read[1]![1]["placed"]).toBe(placed);
  });

  it("gives nothing for a header row without rows below it", async () => {
    expect(await readObjects(sheet([[1, "Order"]]), [{ key: "id", header: "Order" }])).toEqual([]);
  });

  it("matches a header cell only when it holds the declared text exactly: not by case or white space, and never a number, a boolean, a date or an empty cell", async () => {
    const date = new Date(Date.UTC(2024, 0, 1));
    const unmatched: readonly (readonly [ReadValue, string])[] = [
      ["total", "Total"],
      [" Total", "Total"],
      [2024, "2024"],
      [1.5, "1.5"],
      [true, "true"],
      [true, "TRUE"],
      [date, "2024-01-01"],
      [date, date.toISOString()],
      [null, ""],
    ];
    for (const [cell, header] of unmatched) {
      expect(await outcomeOf(sheet([[1, cell, "Other"]]), [{ key: "a", header }])).toEqual(
        missing(header),
      );
    }
    const header = sheet([
      [1, 2024, "2024", true, "true", null, ""],
      [2, 1, 2, 3, 4, 5, 6],
    ]);
    expect(
      await readObjects(header, [
        { key: "year", header: "2024" },
        { key: "flag", header: "true" },
        { key: "empty", header: "" },
      ]),
    ).toEqual([[2, { year: 2, flag: 4, empty: 6 }]]);
  });

  it("gives a column the first of two cells with its header, and a second column with that header the second cell; a header the caller did not declare is ignored", async () => {
    const rows = sheet([
      [1, "Total", "Extra", "Total", "__proto__", "Total"],
      [2, 1, 2, 3, 4, 5],
    ]);
    expect(await readObjects(rows, [{ key: "total", header: "Total" }])).toEqual([
      [2, { total: 1 }],
    ]);
    expect(
      await readObjects(rows, [
        { key: "net", header: "Total" },
        { key: "gross", header: "Total" },
        { key: "tax", header: "Total" },
      ]),
    ).toEqual([[2, { net: 1, gross: 3, tax: 5 }]]);
    const four = Array.from({ length: 4 }, (_, at) => ({ key: `total${at}`, header: "Total" }));
    expect(await outcomeOf(rows, four)).toEqual(missing("Total"));
  });

  it("gives a key declared twice the value of its first column, and still needs the header of the second", async () => {
    const rows = sheet([
      [1, "Gross", "Net"],
      [2, 125, 100],
    ]);
    expect(
      await readObjects(rows, [
        { key: "total", header: "Net" },
        { key: "total", header: "Gross" },
      ]),
    ).toEqual([[2, { total: 100 }]]);
    expect(
      await outcomeOf(rows, [
        { key: "total", header: "Net" },
        { key: "total", header: "Tax" },
      ]),
    ).toEqual(missing("Tax"));
  });

  it("throws MISSING_HEADER with the header text of the first declared column the header row lacks, and closes the rows", async () => {
    const rows = sheet([
      [4, "Order", "Total"],
      [5, "A-1", 10.5],
    ]);
    const columns = [
      { key: "id", header: "Order" },
      { key: "placed", header: "Placed" },
      { key: "note" },
    ] as const;
    expect(await outcomeOf(rows, columns)).toEqual(missing("Placed"));
    expect(rows.log).toEqual(["start", "row 4", "end"]);
    expect(await outcomeOf(rows, [{ key: "id", header: "Order" }, { key: "note" }])).toEqual(
      missing("note"),
    );
  });

  it("throws MISSING_HEADER with the first column's header when the rows have no header row, or no header when no column is declared", async () => {
    expect(await outcomeOf(sheet([]), [{ key: "id", header: "Order" }, { key: "total" }])).toEqual(
      missing("Order"),
    );
    expect(await outcomeOf(sheet([]), [{ key: "id" }])).toEqual(missing("id"));
    expect(await outcomeOf(sheet([]), [])).toEqual(missing(undefined));
    expect(await readObjects(sheet([[1, "Order"], [2]]), [])).toEqual([[2, {}]]);
  });

  it("throws at the first step of an iteration, not at the call, and reads nothing before it is iterated", async () => {
    const rows = sheet([]);
    const objects = objectsFromRows(rows, [{ key: "id", header: "Order" }]);
    const iterator = objects[Symbol.asyncIterator]();
    expect(rows.log).toEqual([]);
    await expect(iterator.next()).rejects.toThrow(
      new XlsxReadError("MISSING_HEADER", "the header row lacks a declared column", {
        header: "Order",
      }),
    );
    expect(rows.log).toEqual(["start", "end"]);
  });

  it("passes an error of the rows through as it is, with every field", async () => {
    const refusal = new XlsxReadError("INVALID_CELL", "a cell is of an unknown type", {
      sheet: "Orders",
      cell: "B7",
    });
    for (const error of [refusal, new Error("the disk is gone")]) {
      const rows: AsyncIterable<ReadRow> = {
        async *[Symbol.asyncIterator]() {
          yield { number: 1, cells: ["Order"] };
          yield { number: 2, cells: ["A-1"] };
          throw error;
        },
      };
      const read: number[] = [];
      await expect(
        (async () => {
          for await (const { number } of objectsFromRows(rows, [{ key: "id", header: "Order" }])) {
            read.push(number);
          }
        })(),
      ).rejects.toBe(error);
      expect(read).toEqual([2]);
    }
  });

  it("ends the iteration of the rows when the loop is left early, by a break or a throw", async () => {
    const rows = sheet([
      [1, "Order"],
      [2, "A-1"],
      [3, "A-2"],
      [4, "A-3"],
    ]);
    for await (const { number } of objectsFromRows(rows, [{ key: "id", header: "Order" }])) {
      if (number === 2) break;
    }
    expect(rows.log).toEqual(["start", "row 1", "row 2", "end"]);
    rows.log.length = 0;
    await expect(
      (async () => {
        for await (const { number } of objectsFromRows(rows, [{ key: "id", header: "Order" }])) {
          throw new Error(`left at row ${number}`);
        }
      })(),
    ).rejects.toThrow("left at row 2");
    expect(rows.log).toEqual(["start", "row 1", "row 2", "end"]);
  });

  it("iterates the rows again from their start on every iteration, with new objects", async () => {
    const rows = sheet([
      [1, "Order"],
      [2, "A-1"],
    ]);
    const objects = objectsFromRows(rows, [{ key: "id", header: "Order" }]);
    const first: object[] = [];
    for await (const { values } of objects) first.push(values);
    const second: object[] = [];
    for await (const { values } of objects) second.push(values);
    expect(second).toEqual(first);
    expect(second[0]).not.toBe(first[0]);
    expect(rows.log).toEqual(["start", "row 1", "row 2", "end", "start", "row 1", "row 2", "end"]);
  });

  it("gives values an object without a prototype, so a header or a key named like a property of Object changes nothing outside it", async () => {
    const snapshot = () =>
      Object.getOwnPropertyNames(Object.prototype).map((name) => [
        name,
        Object.getOwnPropertyDescriptor(Object.prototype, name),
      ]);
    const before = snapshot();
    const placed = new Date(Date.UTC(2026, 0, 2));
    const names = ["__proto__", "constructor", "toString", "hasOwnProperty", "valueOf"];
    const rows = sheet([
      [1, ...names, "id"],
      [2, placed, "a", "b", "c", "d", "A-1"],
    ]);
    const columns = [
      ...names.slice(0, 4).map((key) => ({ key })),
      { key: "valueOf", header: "id" },
    ] as const;
    const values = (await readObjects(rows, columns))[0]![1];
    expect(snapshot()).toEqual(before);
    expect(Object.getPrototypeOf(values)).toBeNull();
    expect(Object.keys(values)).toEqual([
      "__proto__",
      "constructor",
      "toString",
      "hasOwnProperty",
      "valueOf",
    ]);
    expect(Object.entries(values)).toEqual([
      ["__proto__", placed],
      ["constructor", "a"],
      ["toString", "b"],
      ["hasOwnProperty", "c"],
      ["valueOf", "A-1"],
    ]);
    expect(({} as Record<string, unknown>)["polluted"]).toBeUndefined();
  });

  it("gives an object that JSON.stringify, structuredClone and a spread copy, that Object.assign does not, and that String cannot convert", async () => {
    const placed = new Date(Date.UTC(2026, 0, 2));
    const rows = sheet([
      [1, "__proto__", "Total"],
      [2, placed, 10.5],
    ]);
    const values = (
      await readObjects(rows, [{ key: "__proto__" }, { key: "total", header: "Total" }])
    )[0]![1];
    expect(JSON.stringify(values)).toBe('{"__proto__":"2026-01-02T00:00:00.000Z","total":10.5}');
    for (const copy of [{ ...values }, structuredClone(values)]) {
      expect(Object.getPrototypeOf(copy)).toBe(Object.prototype);
      expect(Object.entries(copy)).toEqual([
        ["__proto__", placed],
        ["total", 10.5],
      ]);
    }
    const assigned: object = Object.assign({}, values);
    expect(Object.getPrototypeOf(assigned)).toBe(placed);
    expect(Object.keys(assigned)).toEqual(["total"]);
    const unknown: unknown = values;
    expect(() => String(unknown)).toThrow(TypeError);
  });

  it("reads one cell of a row per declared column, however wide the row", async () => {
    const read: string[] = [];
    const watched = (cells: ReadValue[]): ReadValue[] =>
      new Proxy(cells, {
        get(target, property, receiver) {
          read.push(String(property));
          return Reflect.get(target, property, receiver);
        },
      });
    const header = Array.from({ length: 16_384 }, (_, at) => `Column ${at}`);
    const wide = Array.from({ length: 16_384 }, (_, at) => at);
    const rows: AsyncIterable<ReadRow> = {
      async *[Symbol.asyncIterator]() {
        yield { number: 1, cells: header };
        yield { number: 2, cells: watched(wide) };
        yield { number: 3, cells: watched([1]) };
      },
    };
    const columns = [
      { key: "last", header: "Column 16383" },
      { key: "first", header: "Column 0" },
    ];
    expect(await readObjects(rows, columns)).toEqual([
      [2, { last: 16_383, first: 0 }],
      [3, { last: null, first: 1 }],
    ]);
    expect(read).toEqual(["16383", "0", "16383", "0"]);
  });

  it("finds 16 384 columns in a header row of 16 384 long texts in time linear in the text, also when every header is the same", async () => {
    const width = 16_384;
    const prefix = "Quantity per week and store, ".repeat(8);
    const distinct = Array.from({ length: width }, (_, at) => `${prefix}${at}`);
    const same = distinct.map(() => prefix);
    const values = Array.from({ length: width }, (_, at) => at);
    const declared = (header: readonly string[]) =>
      header.map((text, at) => ({ key: `k${at}`, header: text }));
    for (const [header, columns] of [
      [distinct, declared(distinct).reverse()],
      [same, declared(same)],
    ] as const) {
      const started = performance.now();
      const read = await readObjects(
        sheet([
          [1, ...header],
          [2, ...values],
        ]),
        columns,
      );
      const took = performance.now() - started;
      const found = Object.entries(read[0]![1]);
      expect(found).toHaveLength(width);
      expect(found.filter(([key, value]) => key !== `k${String(value)}`)).toEqual([]);
      // About thirty times what this takes on a 2026 laptop, under coverage or not; a search of the
      // header row for each column took 900 ms there.
      expect(took).toBeLessThan(250);
    }
  });
});

describe("the types of objectsFromRows", () => {
  interface Order {
    readonly id: string;
    readonly placed: Date;
    readonly total: number;
    readonly tags: readonly string[];
  }
  const rows = sheet([
    [1, "Order", "Placed", "total", "Tags", "id"],
    [2, "A-1", null, 10.5, "a, b", "A-1"],
  ]);

  it("takes the column list of sheetFromObjects as it is, and gives every key of T for a list typed ObjectColumn<T>[], undefined where the list does not declare it", async () => {
    const columns: ObjectColumn<Order>[] = [
      { key: "id", header: "Order", width: 12, hidden: false },
      { key: "total", style: { numberFormat: "0.00" } },
      { key: "tags", header: "Tags", value: (order) => order.tags.join(", ") },
    ];
    sheetFromObjects<Order>([], { name: "Orders", columns });
    const objects = objectsFromRows(rows, columns);
    expectTypeOf(objects).toEqualTypeOf<AsyncIterable<ReadObject<keyof Order>>>();
    expectTypeOf<ReadObject<keyof Order>["values"]>().toEqualTypeOf<{
      readonly id: ReadValue;
      readonly placed: ReadValue;
      readonly total: ReadValue;
      readonly tags: ReadValue;
    }>();
    const read: ReadObject<keyof Order>["values"][] = [];
    for await (const { values } of objects) read.push(values);
    expect(read).toEqual([{ id: "A-1", total: 10.5, tags: "a, b" }]);
    expect(read[0]!.placed).toBeUndefined();
  });

  it("gives only the declared keys for a list written as const, or inline", () => {
    const columns = [
      { key: "id", header: "Order" },
      { key: "total", header: "Total" },
    ] as const satisfies readonly ObjectColumn<Order>[];
    sheetFromObjects<Order>([], { name: "Orders", columns });
    expectTypeOf(objectsFromRows(rows, columns)).toEqualTypeOf<
      AsyncIterable<ReadObject<"id" | "total">>
    >();
    expectTypeOf(objectsFromRows(rows, [{ key: "id" }])).toEqualTypeOf<
      AsyncIterable<ReadObject<"id">>
    >();
  });

  it("refuses a key that was not declared, a value used as a type it was not checked to be, and a write", async () => {
    const columns = [{ key: "id" }, { key: "total" }] as const;
    for await (const { number, values } of objectsFromRows(rows, columns)) {
      const total: ReadValue = values.total;
      // @ts-expect-error a key that was not declared is not on the object
      void values.placed;
      // @ts-expect-error a value is not a number until the caller has checked it
      const unchecked: number = values.total;
      // @ts-expect-error the object is read only
      values.id = "A-1";
      void [number, total, unchecked];
    }
  });
});
