import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { type Workbook, writeXlsx } from "../src/index.js";
import { createPackage, store } from "../src/opc/index.js";
import { type OpenedWorkbook, XlsxReadError, openXlsx } from "../src/read.js";
import {
  appearanceWorkbook,
  autoWidthWorkbook,
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

async function rowsOf(book: OpenedWorkbook, sheet: string | number): Promise<unknown> {
  let rows = 0;
  try {
    for await (const _ of book.rows(sheet)) rows++;
  } catch (error) {
    return error;
  }
  return rows;
}

const NOT_BUILT = new Error("reading the rows of a worksheet is not built yet");

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A package of the three parts the sheet list is read from. */
function workbookOf(sheets: string, relationships: string): Promise<Uint8Array> {
  const parts = {
    "_rels/.rels": `<Relationships><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns:r="${RELATIONSHIPS}"><sheets>${sheets}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships>${relationships}</Relationships>`,
  };
  const encoder = new TextEncoder();
  return createPackage(
    Object.entries(parts).map(([path, text]) => ({ path, data: encoder.encode(text) })),
    store,
  );
}

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

  it("charges the directory and the three parts of the sheet list to an eighth of inflatedBytes, and nothing more", async () => {
    const bytes = await writeXlsx(imagesWorkbook);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directory = view.getUint32(bytes.length - 22 + 12, true);
    const listParts = ["_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"];
    const charged = unzip(bytes)
      .filter(({ name }) => listParts.includes(name))
      .reduce((total, entry) => total + entry.uncompressedSize, directory);
    const opened = await openXlsx(bytes, { limits: { inflatedBytes: 8 * charged } });
    expect(opened.sheets).toEqual(["Catalogue"]);
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
      expect(thrown).toMatchObject({ code: "UNKNOWN_SHEET" });
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

  it("does not end the iteration of a worksheet as if it were empty", async () => {
    const book = await openXlsx(fixture("multi-sheet.xlsx"));
    for (const sheet of [0, 1, "Empty"]) {
      const ended = await rowsOf(book, sheet);
      expect(ended, String(sheet)).toEqual(NOT_BUILT);
      expect(ended).not.toBeInstanceOf(XlsxReadError);
    }
  });

  it("gives a chart sheet no rows, every time it is iterated", async () => {
    const book = await openXlsx(
      await workbookOf(
        '<sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/>',
        `<Relationship Id="rId1" Type="${RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${RELATIONSHIPS}/chartsheet" Target="chartsheets/sheet1.xml"/>`,
      ),
    );
    expect(book.sheets).toEqual(["Data", "Chart"]);
    expect(await rowsOf(book, "Chart")).toBe(0);
    expect(await rowsOf(book, 1)).toBe(0);
    expect(await rowsOf(book, "Data")).toEqual(NOT_BUILT);
  });
});

class WatchedBlob extends Blob {
  readonly reads: (readonly [start: number, end: number])[] = [];

  override slice(start = 0, end = this.size): Blob {
    this.reads.push([start, end]);
    return super.slice(start, end);
  }
}
