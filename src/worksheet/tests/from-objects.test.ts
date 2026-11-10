import { describe, expect, expectTypeOf, it } from "vite-plus/test";
import { type Row, type Sheet, sheetFromObjects } from "../index.js";

interface Order {
  id: string;
  customer: string;
  total: number;
  placedAt: Date;
  tags: string[];
  note?: string;
}

const orders: Order[] = [
  {
    id: "A-1001",
    customer: "Kari",
    total: 1234.5,
    placedAt: new Date(Date.UTC(2026, 10, 1)),
    tags: ["vip"],
  },
  {
    id: "A-1002",
    customer: "Ola",
    total: 99,
    placedAt: new Date(Date.UTC(2026, 10, 2)),
    tags: [],
    note: "gift",
  },
];

const rowsOf = (sheet: Sheet): Row[] => [...sheet.rows];

describe("sheetFromObjects", () => {
  it("builds a sheet with a bold header row, a frozen header and an autofilter by default", () => {
    const sheet = sheetFromObjects(orders, {
      name: "Orders",
      columns: [
        { key: "id", header: "Order" },
        { key: "customer", header: "Customer", width: 32 },
        { key: "total", header: "Total", style: { numberFormat: '#,##0.00 "kr"' } },
      ],
    });
    expect(sheet.name).toBe("Orders");
    expect(sheet.freeze).toEqual({ rows: 1 });
    expect(sheet.autoFilter).toBe(true);
    expect(sheet.columns).toEqual([
      {},
      { width: 32 },
      { style: { numberFormat: '#,##0.00 "kr"' } },
    ]);
    expect(rowsOf(sheet)).toEqual([
      [
        { value: "Order", style: { bold: true } },
        { value: "Customer", style: { bold: true } },
        { value: "Total", style: { bold: true } },
      ],
      ["A-1001", "Kari", 1234.5],
      ["A-1002", "Ola", 99],
    ]);
  });

  it("uses the key as the header when none is given and reuses one header style object for every header cell", () => {
    const sheet = sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id" }, { key: "total" }],
    });
    const [header] = rowsOf(sheet) as [{ value: string; style: object }[]];
    expect(header.map((cell) => cell.value)).toEqual(["id", "total"]);
    expect(header[0]!.style).toBe(header[1]!.style);
  });

  it("derives a cell through a value function, which receives the whole object", () => {
    const sheet = sheetFromObjects(orders, {
      name: "S",
      columns: [
        { key: "tags", header: "Tags", value: (order) => order.tags.join(", ") },
        { key: "note", value: (order) => order.note ?? "-" },
      ],
    });
    expect(rowsOf(sheet).slice(1)).toEqual([
      ["vip", "-"],
      ["", "gift"],
    ]);
  });

  it("switches each default off independently", () => {
    const plain = sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id" }],
      header: { bold: false },
      freeze: false,
      autoFilter: false,
    });
    expect(plain.freeze).toBeUndefined();
    expect(plain.autoFilter).toBe(false);
    expect(rowsOf(plain)[0]).toEqual([{ value: "id", style: { bold: false } }]);
    const custom = sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id" }],
      header: { bold: true, numberFormat: "@" },
    });
    expect(rowsOf(custom)[0]).toEqual([{ value: "id", style: { bold: true, numberFormat: "@" } }]);
  });

  it("passes auto widths through: the sheet flag and a column's own", () => {
    const whole = sheetFromObjects(orders, { name: "S", autoWidth: true });
    expect(whole.autoWidth).toBe(true);
    expect(whole.columns).toBeUndefined();
    const some = sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id", width: "auto" }, { key: "customer", width: 20 }, { key: "total" }],
    });
    expect(some.autoWidth).toBeUndefined();
    expect(some.columns).toEqual([{ width: "auto" }, { width: 20 }, {}]);
  });

  it("drops the header row, and with it the freeze and filter defaults, for header: false", () => {
    const sheet = sheetFromObjects(orders, { name: "S", columns: [{ key: "id" }], header: false });
    expect(rowsOf(sheet)).toEqual([["A-1001"], ["A-1002"]]);
    expect(sheet.freeze).toBeUndefined();
    expect(sheet.autoFilter).toBe(false);
    const filtered = sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id" }],
      header: false,
      freeze: true,
      autoFilter: true,
    });
    expect(filtered.freeze).toEqual({ rows: 1 });
    expect(filtered.autoFilter).toBe(true);
  });

  it("takes the columns from the first object's own keys when none are given", () => {
    const sheet = sheetFromObjects(
      [
        { b: 2, a: 1 },
        { b: 4, a: 3, extra: 5 },
      ],
      { name: "S" },
    );
    expect(sheet.columns).toBeUndefined();
    expect(rowsOf(sheet)).toEqual([
      [
        { value: "b", style: { bold: true } },
        { value: "a", style: { bold: true } },
      ],
      [2, 1],
      [4, 3],
    ]);
  });

  it("infers columns without a header row too", () => {
    expect(rowsOf(sheetFromObjects([{ a: 1, b: "x" }], { name: "S", header: false }))).toEqual([
      [1, "x"],
    ]);
  });

  it("consumes a one-shot iterable exactly once, with and without columns", () => {
    function* stream() {
      yield { id: "x", n: 1 };
      yield { id: "y", n: 2 };
    }
    expect(rowsOf(sheetFromObjects(stream(), { name: "S" }))).toEqual([
      [
        { value: "id", style: { bold: true } },
        { value: "n", style: { bold: true } },
      ],
      ["x", 1],
      ["y", 2],
    ]);
    expect(rowsOf(sheetFromObjects(stream(), { name: "S", columns: [{ key: "n" }] }))).toEqual([
      [{ value: "n", style: { bold: true } }],
      [1],
      [2],
    ]);
  });

  it("produces only the header for empty data with columns, and nothing at all without them", () => {
    expect(
      rowsOf(sheetFromObjects([] as Order[], { name: "S", columns: [{ key: "id" }] })),
    ).toEqual([[{ value: "id", style: { bold: true } }]]);
    expect(rowsOf(sheetFromObjects([] as Order[], { name: "S" }))).toEqual([]);
  });

  it("types keys against the object: only writable properties without a value function", () => {
    sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "id" }, { key: "total" }, { key: "placedAt" }, { key: "note" }],
    });
    // @ts-expect-error a misspelled key is rejected
    sheetFromObjects(orders, { name: "S", columns: [{ key: "ids" }] });
    // @ts-expect-error an array property cannot be written without a value function
    sheetFromObjects(orders, { name: "S", columns: [{ key: "tags" }] });
    sheetFromObjects(orders, {
      name: "S",
      columns: [{ key: "tags", value: (order) => order.tags.length }],
    });
    expectTypeOf(sheetFromObjects(orders, { name: "S" })).toEqualTypeOf<Sheet>();
  });
});
