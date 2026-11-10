import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import {
  type Row,
  type StreamOptions,
  type StreamedWorkbook,
  type Workbook,
  streamXlsx,
  writeXlsx,
} from "../src/index.js";
import {
  appearanceWorkbook,
  autoWidthWorkbook,
  cellTypesOptions,
  cellTypesWorkbook,
  formulasWorkbook,
  hyperlinksWorkbook,
  layoutWorkbook,
  minimalWorkbook,
  multiSheetWorkbook,
  printWorkbook,
  stylesWorkbook,
  validationWorkbook,
} from "./fixtures/workbooks.js";
import { readPart, unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** The parts by name, with the one element the byte path writes and the stream cannot removed. */
function comparable(bytes: Uint8Array): [string, string][] {
  return unzip(bytes)
    .map(
      (e) => [e.name, decode(e.data).replace(/<dimension ref="[^"]*"\/>/, "")] as [string, string],
    )
    .sort((a, b) => a[0].localeCompare(b[0]));
}

const fixtures: readonly (readonly [string, Workbook, StreamOptions])[] = [
  ["minimal", minimalWorkbook, {}],
  ["cell types", cellTypesWorkbook, cellTypesOptions],
  ["multi-sheet", multiSheetWorkbook, {}],
  ["styles", stylesWorkbook, {}],
  ["layout", layoutWorkbook, {}],
  ["appearance", appearanceWorkbook, {}],
  ["formulas", formulasWorkbook, {}],
  ["hyperlinks", hyperlinksWorkbook, {}],
  ["validation", validationWorkbook, {}],
  ["print", printWorkbook, {}],
  ["auto-width", autoWidthWorkbook, {}],
];

describe("streamXlsx", () => {
  it.each(fixtures)(
    "streams the %s workbook with shared strings to the same parts writeXlsx writes",
    async (_name, workbook, options) => {
      const streamed = await collect(
        streamXlsx(workbook, { ...options, strings: "shared", compression: "none" }),
      );
      const written = await writeXlsx(workbook, { ...options, compression: "none" });
      expect(comparable(streamed)).toEqual(comparable(written));
      const entries = unzip(streamed);
      expect(entries.every((e) => e.method === 8)).toBe(true);
      expect(entries.map((e) => e.name).slice(0, 3)).toEqual([
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
      ]);
    },
  );

  it("writes inline strings by default, with an empty shared-strings part the package still declares", async () => {
    const bytes = await collect(streamXlsx(minimalWorkbook, { compression: "none" }));
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('t="inlineStr"><is><t>');
    expect(sheet).not.toContain('t="s"');
    expect(readPart(bytes, "xl/sharedStrings.xml")).toContain('count="0" uniqueCount="0"');
    expect(readPart(bytes, "[Content_Types].xml")).toContain("/xl/sharedStrings.xml");
    expect(readPart(bytes, "xl/_rels/workbook.xml.rels")).toContain("sharedStrings.xml");
  });

  it("is byte-identical to the committed streamed fixture", async () => {
    const bytes = await collect(streamXlsx(appearanceWorkbook, { compression: "none" }));
    const file = join(fixturesDir, "streamed.xlsx");
    if (process.env["UPDATE_FIXTURES"] === "1" || !existsSync(file)) writeFileSync(file, bytes);
    expect(bytes).toEqual(new Uint8Array(readFileSync(file)));
  });

  it("compresses by default to the same parts as the stored-block form", async () => {
    const deflated = unzip(await collect(streamXlsx(appearanceWorkbook)));
    const stored = unzip(await collect(streamXlsx(appearanceWorkbook, { compression: "none" })));
    expect(deflated.map((e) => [e.name, decode(e.data)])).toEqual(
      stored.map((e) => [e.name, decode(e.data)]),
    );
    expect(deflated.reduce((n, e) => n + e.compressedSize, 0)).toBeLessThan(
      stored.reduce((n, e) => n + e.compressedSize, 0),
    );
  });

  it("takes rows from async iterables and names the filter range after the last row", async () => {
    async function* rows(): AsyncGenerator<Row> {
      yield ["Name", "Amount"];
      for (let i = 0; i < 5; i++) yield [`row ${i}`, i];
    }
    const workbook: StreamedWorkbook = {
      sheets: [{ name: "Orders", rows: rows(), autoFilter: true }],
    };
    const bytes = await collect(streamXlsx(workbook, { compression: "none" }));
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain('<autoFilter ref="A1:B6"/>');
    expect(readPart(bytes, "xl/workbook.xml")).toContain("$A$1:$B$6</definedName>");
  });

  it("throws before the stream exists when the workbook itself is invalid", () => {
    expect(() => streamXlsx({ sheets: [] })).toThrow(
      expect.objectContaining({ code: "NO_SHEETS" }),
    );
    expect(() =>
      streamXlsx({
        sheets: [
          { name: "a", rows: [] },
          { name: "A", rows: [] },
        ],
      }),
    ).toThrow(expect.objectContaining({ code: "DUPLICATE_SHEET_NAME" }));
  });

  it("errors the stream with the cell's error when a row cannot be written", async () => {
    const stream = streamXlsx({
      sheets: [{ name: "S", rows: [["ok"], [{ value: 1, style: { fill: "red" } }]] }],
    });
    await expect(collect(stream)).rejects.toMatchObject({
      code: "INVALID_STYLE",
      path: "sheets[0].rows[1][0].style.fill",
    });
  });

  it("reads no row before the first pull and closes the row source when the consumer cancels", async () => {
    const log: string[] = [];
    async function* cursor(): AsyncGenerator<Row> {
      try {
        log.push("opened");
        for (let i = 0; ; i++) yield [`row ${i}`];
      } finally {
        log.push("closed");
      }
    }
    const stream = streamXlsx({ sheets: [{ name: "S", rows: cursor() }] });
    await Promise.resolve();
    expect(log).toEqual([]);
    const reader = stream.getReader();
    // The three parts before the sheet go out without touching a row.
    for (let i = 0; i < 100 && log.length === 0; i++) await reader.read();
    expect(log).toEqual(["opened"]);
    await reader.cancel("client went away");
    expect(log).toEqual(["opened", "closed"]);
  });

  it("streams formulas, links, validations and print setup from an async generator, in one pass", async () => {
    async function* rows(): AsyncGenerator<Row> {
      yield ["Item", "Total", "Link"];
      for (let i = 2; i <= 3; i++) {
        yield [
          `item ${i}`,
          { formula: `A${i}*2` },
          { value: "docs", href: `https://example.com/${i}` },
        ];
      }
      yield ["Sum", { formula: "SUM(B2:B3)", value: 10 }, { value: "top", cell: "A1" }];
    }
    const workbook: StreamedWorkbook = {
      sheets: [
        {
          name: "Stream",
          rows: rows(),
          validations: [{ range: "A2:A100", list: ["item 2", "item 3"] }],
          print: { orientation: "landscape", repeatRows: 1 },
        },
      ],
    };
    const bytes = await collect(streamXlsx(workbook, { compression: "none" }));
    const names = unzip(bytes).map((e) => e.name);
    expect(names.indexOf("xl/worksheets/_rels/sheet1.xml.rels")).toBe(
      names.indexOf("xl/worksheets/sheet1.xml") + 1,
    );
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="B2"><f>A2*2</f></c>');
    expect(sheet).toContain('<c r="B4"><f>SUM(B2:B3)</f><v>10</v></c>');
    expect(sheet).toMatch(
      /<\/sheetData><dataValidations count="1">.*<\/dataValidations><hyperlinks><hyperlink ref="C2" r:id="rId1"\/><hyperlink ref="C3" r:id="rId2"\/><hyperlink ref="C4" location="&apos;Stream&apos;!A1"\/><\/hyperlinks><pageMargins [^>]*\/><pageSetup orientation="landscape"\/><\/worksheet>$/,
    );
    expect(readPart(bytes, "xl/worksheets/_rels/sheet1.xml.rels")).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/3"',
    );
    const workbookPart = readPart(bytes, "xl/workbook.xml");
    expect(workbookPart).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Stream&apos;!$1:$1</definedName>',
    );
    expect(workbookPart).toContain('<calcPr fullCalcOnLoad="1"/>');
  });

  it("sizes auto columns from the first 1 000 rows, holding the head until they have been read", async () => {
    async function* rows(wideAt: number): AsyncGenerator<Row> {
      yield ["Name"];
      for (let i = 1; i < 1500; i++)
        yield [i === wideAt ? "a much longer text past the sample" : "ab"];
    }
    const stream = (wideAt: number) =>
      collect(
        streamXlsx(
          { sheets: [{ name: "S", columns: [{ width: "auto" }], rows: rows(wideAt) }] },
          { compression: "none" },
        ),
      );
    const colsOf = (bytes: Uint8Array) =>
      /<cols>.*?<\/cols>/.exec(readPart(bytes, "xl/worksheets/sheet1.xml"))?.[0];
    // Inside the sample the wide value counts; past it the first thousand rows have set the width.
    expect(colsOf(await stream(999))).toBe(
      '<cols><col min="1" max="1" width="32.29" customWidth="1"/></cols>',
    );
    expect(colsOf(await stream(1200))).toBe(
      '<cols><col min="1" max="1" width="7" customWidth="1"/></cols>',
    );
    const sheet = readPart(await stream(1200), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>');
    expect(sheet).toContain('<row r="1500">');
  });

  it("streams an auto-width sheet without rows as an empty sheet, and one whose sample is empty rows", async () => {
    const empty = await collect(
      streamXlsx({ sheets: [{ name: "S", autoWidth: true, rows: [] }] }, { compression: "none" }),
    );
    expect(readPart(empty, "xl/worksheets/sheet1.xml")).toContain("<sheetData/></worksheet>");
    const late = await collect(
      streamXlsx(
        {
          sheets: [
            {
              name: "S",
              autoWidth: true,
              rows: [...Array.from({ length: 1200 }, () => []), ["ab"]],
            },
          ],
        },
        { compression: "none" },
      ),
    );
    const sheet = readPart(late, "xl/worksheets/sheet1.xml");
    expect(sheet).not.toContain("<cols>");
    expect(sheet).toContain('<row r="1201">');
  });

  it("refuses async rows on the byte path with a coded error", async () => {
    async function* rows(): AsyncGenerator<Row> {
      yield ["a"];
    }
    const workbook = { sheets: [{ name: "S", rows: rows() }] } as unknown as Workbook;
    await expect(writeXlsx(workbook)).rejects.toMatchObject({
      code: "ASYNC_ROWS",
      path: "sheets[0].rows",
    });
  });
});
