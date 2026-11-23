import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { type ObjectColumn, type Workbook, sheetFromObjects, writeXlsx } from "../src/index.js";
import { createPackage, store } from "../src/opc/index.js";
import {
  type OpenedWorkbook,
  type ReadColumn,
  type ReadRow,
  type ReadValue,
  XlsxReadError,
  objectsFromRows,
  openXlsx,
} from "../src/read.js";
import {
  appearanceWorkbook,
  autoWidthWorkbook,
  cellTypesOptions,
  cellTypesWorkbook,
  escapesWorkbook,
  formulasWorkbook,
  futureFunctionsWorkbook,
  hiddenWorkbook,
  hyperlinksWorkbook,
  imagePaddingWorkbook,
  imagesWorkbook,
  layoutWorkbook,
  minimalWorkbook,
  multiSheetWorkbook,
  ordersWorkbook,
  printWorkbook,
  rowStylesWorkbook,
  sheetViewWorkbook,
  stylesWorkbook,
  validationWorkbook,
} from "./fixtures/workbooks.js";
import { expectedRows } from "./helpers/expected-rows.js";
import { unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const fixture = (name: string) => new Uint8Array(readFileSync(join(fixturesDir, name)));

const GOLDEN: Readonly<Record<string, Workbook>> = {
  "appearance.xlsx": appearanceWorkbook,
  "auto-width.xlsx": autoWidthWorkbook,
  "cell-types.xlsx": cellTypesWorkbook,
  "escapes.xlsx": escapesWorkbook,
  "formulas.xlsx": formulasWorkbook,
  "future-functions.xlsx": futureFunctionsWorkbook,
  "hidden.xlsx": hiddenWorkbook,
  "hyperlinks.xlsx": hyperlinksWorkbook,
  "image-padding.xlsx": imagePaddingWorkbook,
  "images.xlsx": imagesWorkbook,
  "layout.xlsx": layoutWorkbook,
  "minimal-deflate.xlsx": minimalWorkbook,
  "minimal.xlsx": minimalWorkbook,
  "multi-sheet.xlsx": multiSheetWorkbook,
  "orders.xlsx": ordersWorkbook,
  "print.xlsx": printWorkbook,
  "row-styles.xlsx": rowStylesWorkbook,
  "sheet-view.xlsx": sheetViewWorkbook,
  "streamed.xlsx": appearanceWorkbook,
  "styles.xlsx": stylesWorkbook,
  "validation.xlsx": validationWorkbook,
};

type Rows = (readonly [number, ...ReadValue[]])[];

async function readRows(book: OpenedWorkbook, sheet: string | number): Promise<Rows> {
  const rows: Rows = [];
  for await (const { number, cells } of book.rows(sheet)) rows.push([number, ...cells]);
  return rows;
}

async function outcomeOf(book: OpenedWorkbook, sheet: string | number): Promise<Rows | object> {
  try {
    return await readRows(book, sheet);
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    const { code, sheet: named, cell, limit } = error;
    return { code, sheet: named, cell, limit };
  }
}

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const relationship = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${RELATIONSHIPS}/${type}" Target="${target}"/>`;
const sheetPart = (rows: string) =>
  `<worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`;
const tablePart = (...texts: string[]) =>
  `<sst xmlns="${MAIN}">${texts.map((text) => `<si><t>${text}</t></si>`).join("")}</sst>`;

function workbookOf(
  sheets: string,
  relationships: string,
  more: Readonly<Record<string, string>> = {},
): Promise<Uint8Array> {
  const parts = {
    "_rels/.rels": `<Relationships>${relationship("rId1", "officeDocument", "xl/workbook.xml")}</Relationships>`,
    "xl/workbook.xml": `<workbook xmlns:r="${RELATIONSHIPS}"><sheets>${sheets}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships>${relationships}</Relationships>`,
    ...more,
  };
  const encoder = new TextEncoder();
  return createPackage(
    Object.entries(parts).map(([path, text]) => ({ path, data: encoder.encode(text) })),
    store,
  );
}

const stylesPart = (...formats: number[]) =>
  `<styleSheet xmlns="${MAIN}"><cellXfs>${formats.map((id) => `<xf numFmtId="${id}"/>`).join("")}</cellXfs></styleSheet>`;

/** Sheets One and Two, each a row of two shared strings and a date, and a chart sheet between them. */
function sharingWorkbook(table: string, styles = stylesPart(0, 14)): Promise<Uint8Array> {
  const row = (number: number) =>
    `<row r="${number}"><c r="A${number}" t="s"><v>0</v></c><c r="B${number}" t="s"><v>1</v></c><c r="C${number}" s="1"><v>45000</v></c></row>`;
  return workbookOf(
    '<sheet name="One" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/><sheet name="Two" sheetId="3" r:id="rId3"/>',
    relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
      relationship("rId2", "chartsheet", "chartsheets/sheet1.xml") +
      relationship("rId3", "worksheet", "worksheets/sheet2.xml") +
      relationship("rId4", "sharedStrings", "sharedStrings.xml") +
      relationship("rId5", "styles", "styles.xml"),
    {
      "xl/worksheets/sheet1.xml": sheetPart(row(1)),
      "xl/worksheets/sheet2.xml": sheetPart(row(2)),
      "xl/sharedStrings.xml": table,
      "xl/styles.xml": styles,
    },
  );
}

const MARCH_15 = new Date(Date.UTC(2023, 2, 15));

describe("openXlsx", () => {
  it("lists the sheets of every golden fixture in order, hidden ones too, from the writer's input", async () => {
    expect(
      readdirSync(fixturesDir)
        .filter((name) => name.endsWith(".xlsx"))
        .sort(),
    ).toEqual(Object.keys(GOLDEN).sort());
    for (const [name, input] of Object.entries(GOLDEN)) {
      const book = await openXlsx(fixture(name));
      expect(book.sheets, name).toEqual(input.sheets.map((sheet) => sheet.name));
    }
    expect((await openXlsx(fixture("hidden.xlsx"))).sheets).toEqual(["Lists", "Orders", "Notes"]);
  });

  it("gives the sheet names as a list that no holder can change", async () => {
    const book = await openXlsx(fixture("multi-sheet.xlsx"));
    expect(Object.isFrozen(book.sheets)).toBe(true);
    expect(book.sheets).toEqual(["Orders & Returns", "Ærlig <oversikt>", "Empty"]);
  });

  it("reads null options as no options, as it reads a value that is not an object", async () => {
    const names = ["Orders & Returns", "Ærlig <oversikt>", "Empty"];
    for (const options of [null, 5, "limits"]) {
      // @ts-expect-error: the types refuse these values, which a caller without types can pass.
      const book = await openXlsx(fixture("multi-sheet.xlsx"), options);
      expect(book.sheets, String(options)).toEqual(names);
    }
  });

  it("opens from an ArrayBuffer and from a Blob as from bytes", async () => {
    const bytes = fixture("multi-sheet.xlsx");
    const names = ["Orders & Returns", "Ærlig <oversikt>", "Empty"];
    expect((await openXlsx(bytes.slice().buffer)).sheets).toEqual(names);
    expect((await openXlsx(new Blob([bytes]))).sheets).toEqual(names);
  });

  it("reads the end of the file, the directory and the three parts of the sheet list, and no picture", async () => {
    for (const bytes of [fixture("images.xlsx"), await writeXlsx(imagesWorkbook)]) {
      const entries = unzip(bytes).sort((a, b) => a.localHeaderOffset - b.localHeaderOffset);
      expect(entries.filter(({ name }) => name.startsWith("xl/media/")).length).toBeGreaterThan(0);
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      const directory = view.getUint32(bytes.length - 22 + 16, true);
      const blob = new WatchedBlob([bytes]);
      await openXlsx(blob);
      const read = blob.reads.map(([start, end]) => {
        if (end === bytes.length) return "end of file";
        if (start === directory) return "directory";
        const at = entries.filter((entry) => entry.localHeaderOffset <= start).length - 1;
        const next = entries[at + 1]?.localHeaderOffset ?? directory;
        return end <= next ? entries[at]?.name : `across parts at ${start}`;
      });
      expect(new Set(read)).toEqual(
        new Set([
          "end of file",
          "directory",
          "_rels/.rels",
          "xl/workbook.xml",
          "xl/_rels/workbook.xml.rels",
        ]),
      );
    }
  });

  it("charges the directory, the three parts of the sheet list and the part names of each sheet, the shared strings and the styles to an eighth of inflatedBytes, and nothing more", async () => {
    const bytes = await writeXlsx(imagesWorkbook);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directory = view.getUint32(bytes.length - 22 + 12, true);
    const listParts = ["_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"];
    const charged = unzip(bytes)
      .filter(({ name }) => listParts.includes(name))
      .reduce(
        (total, entry) => total + entry.uncompressedSize,
        directory +
          3 * "xl/workbook.xml".length +
          "worksheets/sheet1.xml".length +
          "sharedStrings.xml".length +
          "styles.xml".length,
      );
    const opened = await openXlsx(bytes, { limits: { inflatedBytes: 8 * charged } });
    expect(opened.sheets).toEqual(["Catalogue"]);
    const styles = unzip(bytes).find(({ name }) => name === "xl/styles.xml")!.uncompressedSize;
    const styled = await openXlsx(bytes, { limits: { inflatedBytes: 8 * (charged + styles) } });
    expect(await readRows(styled, 0)).toEqual(await readRows(await openXlsx(bytes), 0));
    const unstyled = await openXlsx(bytes, {
      limits: { inflatedBytes: 8 * (charged + styles) - 8 },
    });
    expect(await outcomeOf(unstyled, 0)).toEqual({
      code: "LIMIT_EXCEEDED",
      sheet: undefined,
      cell: undefined,
      limit: "inflatedBytes",
    });
    await expect(
      openXlsx(bytes, { limits: { inflatedBytes: 8 * charged - 8 } }),
    ).rejects.toMatchObject({ code: "LIMIT_EXCEEDED", limit: "inflatedBytes" });
    await expect(openXlsx(bytes, { limits: { parts: 2 } })).rejects.toMatchObject({
      code: "LIMIT_EXCEEDED",
      limit: "parts",
    });
    await expect(openXlsx(bytes, { limits: { inflatedBytes: 0 } })).rejects.toMatchObject({
      code: "INVALID_LIMIT",
    });
  });

  it("refuses a file that is no workbook package with NOT_XLSX", async () => {
    const notes = await createPackage(
      [{ path: "notes.txt", data: new TextEncoder().encode("not a workbook") }],
      store,
    );
    for (const file of [notes, new TextEncoder().encode("Name,City\n")]) {
      await expect(openXlsx(file)).rejects.toMatchObject({ code: "NOT_XLSX" });
    }
  });
});

describe("rows", () => {
  it("throws UNKNOWN_SHEET at the call for a name or an index the workbook does not have", async () => {
    const book = await openXlsx(fixture("multi-sheet.xlsx"));
    const unknown = ["orders & returns", "Empty ", " Empty", "0", "", 3, -1, 1.5, Number.NaN];
    for (const sheet of [...unknown, Number.POSITIVE_INFINITY]) {
      let thrown: unknown;
      try {
        book.rows(sheet);
      } catch (error) {
        thrown = error;
      }
      expect(thrown, String(sheet)).toBeInstanceOf(XlsxReadError);
      expect(thrown).toMatchObject({
        code: "UNKNOWN_SHEET",
        sheet: typeof sheet === "string" ? sheet : undefined,
      });
    }
    for (const sheet of ["Empty", "Ærlig <oversikt>", 0, 2, -0]) {
      expect(() => book.rows(sheet), String(sheet)).not.toThrow();
    }
  });

  it("finds each of 100 000 sheets by its name within a second", async () => {
    const names = Array.from({ length: 100_000 }, (_, at) => `Sheet ${at}`);
    const book = await openXlsx(
      await workbookOf(
        names.map((name) => `<sheet name="${name}" sheetId="1" r:id="rId1"/>`).join(""),
        `<Relationship Id="rId1" Type="${RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/>`,
      ),
    );
    const started = performance.now();
    for (const name of names) book.rows(name);
    expect(performance.now() - started).toBeLessThan(1_000);
  });

  it("reads nothing until the rows are iterated", async () => {
    const blob = new WatchedBlob([fixture("multi-sheet.xlsx")]);
    const book = await openXlsx(blob);
    const reads = blob.reads.length;
    book.rows(0);
    book.rows("Empty");
    expect(blob.reads.length).toBe(reads);
  });

  it("reads back every golden fixture as the rows the writer was given, a formula as its cached value or nothing", async () => {
    for (const [name, input] of Object.entries(GOLDEN)) {
      const book = await openXlsx(fixture(name));
      const coerce = name === "cell-types.xlsx" ? cellTypesOptions.coerce : undefined;
      for (const [index, sheet] of input.sheets.entries()) {
        const expected = expectedRows(sheet, coerce);
        expect(await readRows(book, index), `${name}, ${sheet.name}`).toEqual(expected);
        expect(await readRows(book, sheet.name), `${name}, ${sheet.name}`).toEqual(expected);
      }
    }
    const formulas = await readRows(await openXlsx(fixture("formulas.xlsx")), "Budget");
    expect(formulas.slice(1)).toEqual([
      [2, "Desk", 2, 4999],
      [3, "Chair", 4, 1299.5],
      [4, "Lamp", 1, 349],
      [5, "Sum", 7],
      [7, "Items", 3],
      [8, "Over budget", true],
      [9, "Label", "Items: 3"],
      [10, "Not a formula", "=SUM(B2:B4)"],
    ]);
  });

  it("reads a number whose style shows a date as a Date, in the date system of the workbook", async () => {
    const rows =
      '<row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" s="1"><v>44000.5</v></c><c r="C1"><v>44000.5</v></c><c r="D1" t="d"><v>2024-06-19</v></c></row>';
    const bytes = (properties: string) =>
      workbookOf(
        "",
        relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
          relationship("rId2", "styles", "styles.xml"),
        {
          "xl/workbook.xml": `<workbook xmlns:r="${RELATIONSHIPS}">${properties}<sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets></workbook>`,
          "xl/worksheets/sheet1.xml": sheetPart(rows),
          "xl/styles.xml": stylesPart(0, 14),
        },
      );
    const june = (day: number, hour: number) =>
      new Date(Date.UTC(2024 - (day === 19 ? 0 : 4), 5, day, hour));
    expect(await readRows(await openXlsx(await bytes('<workbookPr date1904="1"/>')), 0)).toEqual([
      [1, new Date(Date.UTC(1904, 0, 1)), june(19, 12), 44000.5, june(19, 0)],
    ]);
    expect(await readRows(await openXlsx(await bytes("<workbookPr/>")), 0)).toEqual([
      [1, 0, june(18, 12), 44000.5, june(19, 0)],
    ]);
  });

  it("reads a cell's style only where it can change the value, so a style that does not exist is refused only on a number in a workbook with a date style", async () => {
    const rows = '<row r="1"><c r="A1" s="9" t="b"><v>1</v></c><c r="B1" s="9"><v>2</v></c></row>';
    const book = async (styles?: string) =>
      openXlsx(
        await workbookOf(
          '<sheet name="One" sheetId="1" r:id="rId1"/>',
          relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
            (styles === undefined ? "" : relationship("rId2", "styles", "styles.xml")),
          {
            "xl/worksheets/sheet1.xml": sheetPart(rows),
            ...(styles === undefined ? {} : { "xl/styles.xml": styles }),
          },
        ),
      );
    for (const styles of [undefined, stylesPart(0, 2)]) {
      expect(await readRows(await book(styles), 0), styles).toEqual([[1, true, 2]]);
    }
    expect(await outcomeOf(await book(stylesPart(0, 14)), 0)).toEqual({
      code: "INVALID_CELL",
      sheet: "One",
      cell: "B1",
      limit: undefined,
    });
  });

  it("reads the styles part once, by the first read of a worksheet, whichever sheet asks first", async () => {
    const bytes = await sharingWorkbook(tablePart("first", "second"));
    const styles = unzip(bytes).find(({ name }) => name === "xl/styles.xml")!;
    for (const order of [
      ["One", "Two", "One"],
      ["Two", "One", "Two"],
    ]) {
      const blob = new WatchedBlob([bytes.slice()]);
      const book = await openXlsx(blob);
      const ofStyles = () =>
        blob.reads.filter(
          ([start]) =>
            start >= styles.localHeaderOffset &&
            start < styles.localHeaderOffset + styles.compressedSize + 30 + styles.name.length,
        ).length;
      expect(ofStyles()).toBe(0);
      for (const sheet of order) {
        expect(await readRows(book, sheet)).toEqual([
          [sheet === "One" ? 1 : 2, "first", "second", MARCH_15],
        ]);
      }
      expect(ofStyles()).toBeGreaterThan(0);
      const once = ofStyles();
      await readRows(book, order[0]!);
      expect(ofStyles()).toBe(once);
    }
  });

  it("reads back every date-like value the writer takes as a Date, under each built-in id the writer gives one", async () => {
    const civil = {
      hour: 23,
      minute: 59,
      second: 59,
      millisecond: 999,
      microsecond: 0,
      nanosecond: 0,
    };
    const values = [
      new Date(Date.UTC(1900, 0, 1)),
      new Date(Date.UTC(1900, 1, 28, 23, 59, 59, 999)),
      new Date(Date.UTC(1900, 2, 1)),
      new Date(Date.UTC(9999, 11, 31, 23, 59, 59, 999)),
      { [Symbol.toStringTag]: "Temporal.PlainDate", year: 1900, month: 3, day: 1 },
      { [Symbol.toStringTag]: "Temporal.PlainDateTime", year: 9999, month: 12, day: 31, ...civil },
      { [Symbol.toStringTag]: "Temporal.ZonedDateTime", year: 1900, month: 2, day: 28, ...civil },
      { [Symbol.toStringTag]: "Temporal.Instant", epochMilliseconds: Date.UTC(2026, 10, 17) },
    ] as const;
    const sheet = { name: "Dates", rows: values.map((value) => [value]) };
    const bytes = await writeXlsx({ sheets: [sheet] });
    const text = (name: string) =>
      new TextDecoder().decode(unzip(bytes).find((entry) => entry.name === name)!.data);
    const formats = [
      ...text("xl/styles.xml")
        .split("<cellXfs")[1]!
        .matchAll(/<xf numFmtId="(\d+)"/g),
    ];
    const used = [...text("xl/worksheets/sheet1.xml").matchAll(/ s="(\d+)"/g)].map(([, style]) =>
      Number(formats[Number(style)]![1]),
    );
    expect(new Set(used)).toEqual(new Set([14, 22]));
    expect(await readRows(await openXlsx(bytes), 0)).toEqual(expectedRows(sheet));
  });

  it("writes back every Date it reads from a serial at the edges of the writer's range, and reads it again as itself", async () => {
    const serials = [
      "1",
      "59.9999999",
      "61",
      "2958465.9999999884",
      "45000.9999999999",
      "60",
      "0.5",
      "2958466",
    ];
    const rows = serials.map(
      (serial, at) => `<row r="${at + 1}"><c r="A${at + 1}" s="1"><v>${serial}</v></c></row>`,
    );
    const read = await readRows(
      await openXlsx(
        await workbookOf(
          '<sheet name="Edges" sheetId="1" r:id="rId1"/>',
          relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
            relationship("rId2", "styles", "styles.xml"),
          {
            "xl/worksheets/sheet1.xml": sheetPart(rows.join("")),
            "xl/styles.xml": stylesPart(0, 22),
          },
        ),
      ),
      0,
    );
    expect(read.map(([, value]) => value instanceof Date)).toEqual([
      true,
      true,
      true,
      true,
      true,
      false,
      false,
      false,
    ]);
    const dates = read.filter(([, value]) => value instanceof Date);
    const again = await openXlsx(
      await writeXlsx({ sheets: [{ name: "Edges", rows: dates.map(([, date]) => [date]) }] }),
    );
    expect(await readRows(again, 0)).toEqual(dates.map(([, date], at) => [at + 1, date]));
  });

  it("refuses a worksheet whose part is missing, or that holds no sheet data, naming it, and never reads it as empty", async () => {
    const book = await openXlsx(
      await workbookOf(
        '<sheet name="Gone" sheetId="1" r:id="rId1"/><sheet name="Bare" sheetId="2" r:id="rId2"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
          relationship("rId2", "worksheet", "worksheets/sheet2.xml"),
        {
          "xl/worksheets/sheet2.xml": `<worksheet xmlns="${MAIN}"><dimension ref="A1"/></worksheet>`,
        },
      ),
    );
    for (const sheet of ["Gone", "Bare"]) {
      expect(await outcomeOf(book, sheet)).toEqual({
        code: "INVALID_WORKBOOK",
        sheet,
        cell: undefined,
        limit: undefined,
      });
    }
  });

  it("names the sheet and the cell of a refused cell, and the sheet and the key of a limit a sheet passes", async () => {
    const book = await openXlsx(
      await workbookOf(
        '<sheet name="Cells" sheetId="1" r:id="rId1"/><sheet name="Wide" sheetId="2" r:id="rId2"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
          relationship("rId2", "worksheet", "worksheets/sheet2.xml"),
        {
          "xl/worksheets/sheet1.xml": sheetPart('<row r="41"><c r="C41"><v>0x10</v></c></row>'),
          "xl/worksheets/sheet2.xml": sheetPart('<row r="1"><c r="XFD1"><v>1</v></c></row>'),
        },
      ),
      { limits: { cellPositions: 16_383 } },
    );
    expect(await outcomeOf(book, "Cells")).toEqual({
      code: "INVALID_CELL",
      sheet: "Cells",
      cell: "C41",
      limit: undefined,
    });
    expect(await outcomeOf(book, 1)).toEqual({
      code: "LIMIT_EXCEEDED",
      sheet: "Wide",
      cell: undefined,
      limit: "cellPositions",
    });
  });

  it("charges each iteration's inflatedBytes the worksheet's declared size, and nothing of another iteration", async () => {
    const rows = Array.from({ length: 5_000 }, (_, at) => [at, at / 4]);
    const bytes = await writeXlsx({ sheets: [{ name: "Numbers", rows }] });
    const sheet = unzip(bytes).find(({ name }) => name === "xl/worksheets/sheet1.xml")!;
    const opened = (inflatedBytes: number) => openXlsx(bytes, { limits: { inflatedBytes } });
    const fits = await opened(sheet.uncompressedSize);
    for (let read = 0; read < 3; read++) expect(await readRows(fits, 0)).toHaveLength(5_000);
    expect(await outcomeOf(await opened(sheet.uncompressedSize - 1), 0)).toEqual({
      code: "LIMIT_EXCEEDED",
      sheet: "Numbers",
      cell: undefined,
      limit: "inflatedBytes",
    });
  });

  it("gives a chart sheet no rows and starts no shared strings or styles, even ones it would refuse, and leaves no rejection unhandled", async () => {
    const bytes = await sharingWorkbook(tablePart("first", "second"), "<styles/>");
    const unhandled: unknown[] = [];
    const listen = (reason: unknown) => unhandled.push(reason);
    process.on("unhandledRejection", listen);
    try {
      const blob = new WatchedBlob([bytes.slice()]);
      const book = await openXlsx(blob, { limits: { sharedStrings: 1 } });
      const reads = blob.reads.length;
      expect(await readRows(book, "Chart")).toEqual([]);
      expect(await readRows(book, 1)).toEqual([]);
      await new Promise((resolve) => setTimeout(resolve, 10));
      expect(blob.reads.length).toBe(reads);
      expect(unhandled).toEqual([]);
      expect(await outcomeOf(book, "One")).toEqual({
        code: "LIMIT_EXCEEDED",
        sheet: undefined,
        cell: undefined,
        limit: "sharedStrings",
      });
    } finally {
      process.off("unhandledRejection", listen);
    }
  });

  it("reads the shared strings once, charged to the opening's inflatedBytes, whichever sheet asks first", async () => {
    const texts = Array.from({ length: 4_000 }, (_, at) => `text number ${at}`);
    const bytes = await sharingWorkbook(tablePart(...texts));
    const table = unzip(bytes).find(({ name }) => name === "xl/sharedStrings.xml")!;
    const once = Math.floor(table.uncompressedSize * 1.5);
    for (const order of [
      ["One", "Two", "One"],
      ["Two", "One", "Two"],
    ]) {
      const book = await openXlsx(bytes, { limits: { inflatedBytes: once } });
      for (const sheet of order) {
        expect(await readRows(book, sheet)).toEqual([
          [sheet === "One" ? 1 : 2, ...texts.slice(0, 2), MARCH_15],
        ]);
      }
    }
    const book = await openXlsx(bytes, { limits: { inflatedBytes: table.uncompressedSize } });
    expect(await outcomeOf(book, "Two")).toMatchObject({
      code: "LIMIT_EXCEEDED",
      sheet: undefined,
      limit: "inflatedBytes",
    });
  });

  it("gives a refusal of the shared strings to every iteration of a worksheet, each time, and names no sheet", async () => {
    const refusals = [
      [tablePart("a", "b"), { sharedStringBytes: 20 }, "LIMIT_EXCEEDED", "sharedStringBytes"],
      [tablePart("a", "b"), { sharedStrings: 1 }, "LIMIT_EXCEEDED", "sharedStrings"],
      [`<sst xmlns="${MAIN}"><si><t>a</t></si><item/></sst>`, {}, "INVALID_WORKBOOK", undefined],
      [tablePart("a", "b"), {}, "INVALID_WORKBOOK", undefined, `<styles xmlns="${MAIN}"/>`],
      [
        tablePart("a", "b"),
        { inflatedBytes: 64_000 },
        "LIMIT_EXCEEDED",
        "inflatedBytes",
        stylesPart(...Array.from({ length: 2_000 }, () => 0)),
      ],
    ] as const;
    for (const [table, limits, code, limit, styles] of refusals) {
      const book = await openXlsx(await sharingWorkbook(table, styles), { limits });
      for (const sheet of ["One", "Two", "One", 0]) {
        expect(await outcomeOf(book, sheet), `${code} ${sheet}`).toEqual({
          code,
          sheet: undefined,
          cell: undefined,
          limit,
        });
      }
    }
    const missing = await openXlsx(
      await workbookOf(
        '<sheet name="One" sheetId="1" r:id="rId1"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
          relationship("rId2", "sharedStrings", "sharedStrings.xml"),
        { "xl/worksheets/sheet1.xml": sheetPart("") },
      ),
    );
    expect(await outcomeOf(missing, "One")).toMatchObject({
      code: "INVALID_WORKBOOK",
      sheet: undefined,
    });
    const unstyled = await openXlsx(
      await workbookOf(
        '<sheet name="One" sheetId="1" r:id="rId1"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml") +
          relationship("rId2", "styles", "styles.xml"),
        { "xl/worksheets/sheet1.xml": sheetPart("") },
      ),
    );
    expect(await outcomeOf(unstyled, "One")).toEqual({
      code: "INVALID_WORKBOOK",
      sheet: undefined,
      cell: undefined,
      limit: undefined,
    });
  });

  it("reads a cell that names a shared string of a workbook without a table as one that does not exist", async () => {
    const book = await openXlsx(
      await workbookOf(
        '<sheet name="One" sheetId="1" r:id="rId1"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        {
          "xl/worksheets/sheet1.xml": sheetPart(
            '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>',
          ),
        },
      ),
    );
    expect(await outcomeOf(book, 0)).toEqual({
      code: "INVALID_CELL",
      sheet: "One",
      cell: "B1",
      limit: undefined,
    });
  });

  it("passes on an error of the Blob itself as it is, and names the sheet of a refusal with the refusal's own message", async () => {
    const blob = new FailingBlob([fixture("minimal.xlsx")]);
    const book = await openXlsx(blob);
    expect(await readRows(book, 0)).toHaveLength(5);
    blob.failing = true;
    await expect(readRows(book, 0)).rejects.toBe(blob.error);
    const bare = await openXlsx(
      await workbookOf(
        '<sheet name="Bare" sheetId="1" r:id="rId1"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        { "xl/worksheets/sheet1.xml": sheetPart('<row r="1"><c r="A1" t="b"><v>2</v></c></row>') },
      ),
    );
    await expect(readRows(bare, 0)).rejects.toMatchObject({
      code: "INVALID_CELL",
      sheet: "Bare",
      cell: "A1",
      message: "a cell's value is not of its type or names no shared string",
    });
  });

  it("keeps every field and the stack of a refusal it names the sheet of", async () => {
    const book = await openXlsx(fixture("minimal.xlsx"));
    const rows = book.rows(0)[Symbol.asyncIterator]();
    await rows.next();
    const refusal = new XlsxReadError("MISSING_HEADER", "the header row lacks a column", {
      cell: "C1",
      limit: "cellPositions",
      header: "Total",
    });
    const named = await Promise.resolve(rows.throw?.(refusal)).catch((error: unknown) => error);
    const fields = (error: object) => Object.fromEntries(Object.entries(error));
    expect(named).toBeInstanceOf(XlsxReadError);
    expect(fields(named as XlsxReadError)).toEqual({ ...fields(refusal), sheet: book.sheets[0] });
    expect((named as XlsxReadError).message).toBe(refusal.message);
    expect((named as XlsxReadError).stack).toBe(refusal.stack);
  });

  it("reads two sheets at once and one sheet twice at once, each iteration with limits of its own", async () => {
    const book = await openXlsx(fixture("multi-sheet.xlsx"));
    const alone = [await readRows(book, 0), await readRows(book, 1)];
    const positions = alone.map((rows) =>
      rows.reduce((sum, [, ...cells]) => sum + cells.length, 0),
    );
    const limited = await openXlsx(fixture("multi-sheet.xlsx"), {
      limits: { cellPositions: Math.max(...positions) },
    });
    for (const sheets of [
      [0, 1],
      [1, 1],
      [0, 0, 1, 1],
    ]) {
      const iterators = sheets.map((sheet) => limited.rows(sheet)[Symbol.asyncIterator]());
      const read: Rows[] = sheets.map(() => []);
      for (let busy = true; busy;) {
        busy = false;
        for (const [at, iterator] of iterators.entries()) {
          const step = await iterator.next();
          if (step.done) continue;
          busy = true;
          read[at]!.push([step.value.number, ...step.value.cells]);
        }
      }
      expect(read).toEqual(sheets.map((sheet) => alone[sheet]));
    }
  });

  it("stops reading a sheet's part when the loop is left early, and reads nothing for an iteration not started", async () => {
    const rows = Array.from({ length: 40_000 }, (_, at) => [`row ${at}`, at, at % 2 === 0]);
    const bytes = await writeXlsx({ sheets: [{ name: "Long", rows }] }, { compression: "none" });
    const sheet = unzip(bytes).find(({ name }) => name === "xl/worksheets/sheet1.xml")!;
    const ofSheet = (blob: WatchedBlob) =>
      blob.reads.filter(
        ([start]) =>
          start >= sheet.localHeaderOffset &&
          start < sheet.localHeaderOffset + sheet.compressedSize,
      ).length;
    const blob = new WatchedBlob([bytes]);
    const book = await openXlsx(blob);
    const opened = ofSheet(blob);
    book.rows(0)[Symbol.asyncIterator]();
    expect(ofSheet(blob)).toBe(opened);
    await readRows(book, 0);
    expect(ofSheet(blob) - opened).toBeGreaterThan(4);
    for (const leave of ["break", "throw"]) {
      const before = ofSheet(blob);
      try {
        for await (const row of book.rows(0)) {
          if (leave === "throw") throw new Error(`left at row ${row.number}`);
          break;
        }
      } catch (error) {
        expect(error).toEqual(new Error("left at row 1"));
      }
      expect(ofSheet(blob) - before, leave).toBeLessThanOrEqual(2);
    }
    const columns = [{ key: "name", header: "row 0" }];
    const before = ofSheet(blob);
    objectsFromRows(book.rows(0), columns)[Symbol.asyncIterator]();
    expect(ofSheet(blob)).toBe(before);
    for await (const { number } of objectsFromRows(book.rows(0), columns)) {
      expect(number).toBe(2);
      break;
    }
    expect(ofSheet(blob) - before, "objects").toBeLessThanOrEqual(2);
  });
});

async function objectsOutcome(
  rows: AsyncIterable<ReadRow>,
  columns: readonly ReadColumn<string>[],
): Promise<unknown[] | object> {
  const read: unknown[] = [];
  try {
    for await (const { number, values } of objectsFromRows(rows, columns)) {
      read.push([number, values]);
    }
    return read;
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    const { code, sheet, cell, header } = error;
    return { code, sheet, cell, header };
  }
}

describe("objectsFromRows", () => {
  interface Order {
    readonly id: string;
    readonly customer: string;
    readonly total: number;
    readonly placed: Date;
    readonly paid: boolean;
    readonly note: string | null;
    readonly tags: readonly string[];
  }
  const orders: readonly Order[] = [
    {
      id: "A-1",
      customer: "Kari",
      total: 10.5,
      placed: new Date(Date.UTC(2026, 0, 1, 9, 30)),
      paid: true,
      note: null,
      tags: ["new"],
    },
    {
      id: "A-2",
      customer: "Ola",
      total: 1249.9,
      placed: new Date(Date.UTC(2026, 1, 28)),
      paid: false,
      note: "late",
      tags: [],
    },
  ];
  const columns: ObjectColumn<Order>[] = [
    { key: "id", header: "Order", width: 12 },
    { key: "customer", header: "__proto__" },
    { key: "total", header: "Total", style: { numberFormat: '#,##0.00 "kr"' } },
    { key: "placed", header: "2024", style: { numberFormat: "yyyy-mm-dd hh:mm" } },
    { key: "paid", hidden: true },
    { key: "note", header: "Total" },
    { key: "tags", header: "Tags", value: (order) => order.tags.join(", ") },
  ];

  it("reads back the objects sheetFromObjects wrote, by the same column list", async () => {
    const book = await openXlsx(
      await writeXlsx({ sheets: [sheetFromObjects(orders, { name: "Orders", columns })] }),
    );
    expect(await objectsOutcome(book.rows("Orders"), columns)).toEqual([
      [
        2,
        {
          id: "A-1",
          customer: "Kari",
          total: 10.5,
          placed: orders[0]!.placed,
          paid: true,
          note: null,
          tags: "new",
        },
      ],
      [
        3,
        {
          id: "A-2",
          customer: "Ola",
          total: 1249.9,
          placed: orders[1]!.placed,
          paid: false,
          note: "late",
          tags: null,
        },
      ],
    ]);
  });

  it("refuses with MISSING_HEADER a header row that lacks a declared column, a column with an empty header and a sheet without rows, and passes a refusal of the sheet through with the sheet and the cell", async () => {
    const book = await openXlsx(
      await writeXlsx({
        sheets: [
          sheetFromObjects(orders, { name: "Orders", columns }),
          { name: "Empty", rows: [] },
        ],
      }),
    );
    const more = [...columns, { key: "id", header: "Discount" }];
    expect(await objectsOutcome(book.rows("Orders"), more)).toEqual({
      code: "MISSING_HEADER",
      sheet: undefined,
      cell: undefined,
      header: "Discount",
    });
    expect(await objectsOutcome(book.rows("Empty"), columns)).toEqual({
      code: "MISSING_HEADER",
      sheet: undefined,
      cell: undefined,
      header: "Order",
    });
    const blank: ObjectColumn<Order>[] = [{ key: "id", header: "" }];
    const unnamed = await openXlsx(
      await writeXlsx({ sheets: [sheetFromObjects(orders, { name: "Blank", columns: blank })] }),
    );
    expect(await objectsOutcome(unnamed.rows(0), blank)).toEqual({
      code: "MISSING_HEADER",
      sheet: undefined,
      cell: undefined,
      header: "",
    });
    const damaged = await openXlsx(
      await workbookOf(
        '<sheet name="Cells" sheetId="1" r:id="rId1"/>',
        relationship("rId1", "worksheet", "worksheets/sheet1.xml"),
        {
          "xl/worksheets/sheet1.xml": sheetPart(
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Total</t></is></c></row>' +
              '<row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>0x10</v></c></row>',
          ),
        },
      ),
    );
    expect(await objectsOutcome(damaged.rows(0), [{ key: "total", header: "Total" }])).toEqual({
      code: "INVALID_CELL",
      sheet: "Cells",
      cell: "A3",
      header: undefined,
    });
  });
});

class FailingBlob extends Blob {
  failing = false;
  readonly error = new Error("the disk is gone");

  override slice(start?: number, end?: number): Blob {
    if (this.failing) throw this.error;
    return super.slice(start, end);
  }
}

class WatchedBlob extends Blob {
  readonly reads: (readonly [start: number, end: number])[] = [];

  override slice(start = 0, end = this.size): Blob {
    this.reads.push([start, end]);
    return super.slice(start, end);
  }
}
