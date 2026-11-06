import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { type WriteOptions, type Workbook, writeXlsx } from "../src/index.js";
import {
  appearanceWorkbook,
  cellTypesOptions,
  cellTypesWorkbook,
  minimalWorkbook,
  multiSheetWorkbook,
  stylesWorkbook,
  layoutWorkbook,
  ordersWorkbook,
} from "./fixtures/workbooks.js";
import { readPart, unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

/** Golden fixtures are written with STORE entries: byte-deterministic on every platform. */
const STORE: WriteOptions = { compression: "none" };

async function golden(
  name: string,
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  const bytes = await writeXlsx(workbook, { ...options, ...STORE });
  const file = join(fixturesDir, name);
  if (process.env["UPDATE_FIXTURES"] === "1" || !existsSync(file)) {
    writeFileSync(file, bytes);
  }
  const expected = new Uint8Array(readFileSync(file));
  expect(bytes).toEqual(expected);
  return bytes;
}

describe("minimal workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("minimal.xlsx", minimalWorkbook);
  });

  it("contains exactly the parts a one-sheet workbook needs, content types first", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    expect(unzip(bytes).map((e) => e.name)).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/sharedStrings.xml",
      "xl/worksheets/sheet1.xml",
    ]);
  });

  it("declares every part in the content types and links the sheet from the workbook", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    const contentTypes = readPart(bytes, "[Content_Types].xml");
    for (const part of [
      "/xl/workbook.xml",
      "/xl/worksheets/sheet1.xml",
      "/xl/styles.xml",
      "/xl/sharedStrings.xml",
    ]) {
      expect(contentTypes).toContain(`PartName="${part}"`);
    }
    expect(readPart(bytes, "xl/workbook.xml")).toContain(
      '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>',
    );
    expect(readPart(bytes, "xl/_rels/workbook.xml.rels")).toContain(
      'Target="worksheets/sheet1.xml"',
    );
    expect(readPart(bytes, "_rels/.rels")).toContain('Target="xl/workbook.xml"');
  });

  it("round-trips every string through the shared-strings table, escaped", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    const sst = readPart(bytes, "xl/sharedStrings.xml");
    expect(sst).toContain("<t>Ola &lt;Nordmann&gt; &amp; co</t>");
    expect(sst).toContain("<t>Ærlig «ekte» 🎉</t>");
    expect(sst).toContain('<t xml:space="preserve"> padded </t>');
    expect(sst).toContain("<t></t>");
    expect(sst).toContain('count="10" uniqueCount="10"');
  });

  it("writes numbers inline with full precision", async () => {
    const sheet = readPart(await writeXlsx(minimalWorkbook), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="C2"><v>1234.5</v></c>');
    expect(sheet).toContain('<c r="C3"><v>-42</v></c>');
    expect(sheet).toContain('<c r="C4"><v>1e+21</v></c>');
    expect(sheet).toContain('<c r="C5"><v>0</v></c>');
  });
});

describe("cell types workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("cell-types.xlsx", cellTypesWorkbook, cellTypesOptions);
  });

  it("writes every supported kind of value", async () => {
    const sheet = readPart(
      await writeXlsx(cellTypesWorkbook, cellTypesOptions),
      "xl/worksheets/sheet1.xml",
    );
    expect(sheet).toContain('<c r="B3"><v>3.14</v></c>');
    expect(sheet).toContain('<c r="B4" t="b"><v>1</v></c>');
    expect(sheet).toContain('<c r="B5" t="b"><v>0</v></c>');
    expect(sheet).toContain('<c r="B6" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B7" s="2"><v>46330</v></c>');
    expect(sheet).toContain('<c r="B8" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B9" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B10" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B11"><v>42</v></c>');
    expect(sheet).not.toContain('r="B12"');
    expect(sheet).toContain('<c r="B15"><v>12345678901234</v></c>');
  });

  it("strips control characters and keeps newlines in text", async () => {
    const sst = readPart(
      await writeXlsx(cellTypesWorkbook, cellTypesOptions),
      "xl/sharedStrings.xml",
    );
    expect(sst).toContain("<t>abc</t>");
    expect(sst).toContain("<t>line1\nline2</t>");
  });

  it("declares the date styles the cells reference", async () => {
    const styles = readPart(await writeXlsx(cellTypesWorkbook, cellTypesOptions), "xl/styles.xml");
    expect(styles).toContain('<cellXfs count="3">');
    expect(styles).toContain('numFmtId="14"');
    expect(styles).toContain('numFmtId="22"');
  });
});

describe("compression", () => {
  afterEach(() => vi.unstubAllGlobals());

  const partsOf = (bytes: Uint8Array) =>
    unzip(bytes).map((e) => [e.name, e.method, new TextDecoder().decode(e.data)] as const);

  it("is on by default and decompresses part by part to the STORE output", async () => {
    const stored = await writeXlsx(minimalWorkbook, STORE);
    const deflated = await writeXlsx(minimalWorkbook);
    expect(unzip(deflated).every((e) => e.method === 8)).toBe(true);
    expect(partsOf(deflated).map(([name, , text]) => [name, text])).toEqual(
      partsOf(stored).map(([name, , text]) => [name, text]),
    );
    expect(deflated.byteLength).toBeLessThan(stored.byteLength * 0.6);
  });

  it("matches the committed deflate fixture part by part (bytes differ by compressor)", async () => {
    const file = join(fixturesDir, "minimal-deflate.xlsx");
    if (process.env["UPDATE_FIXTURES"] === "1" || !existsSync(file)) {
      writeFileSync(file, await writeXlsx(minimalWorkbook));
    }
    const committed = new Uint8Array(readFileSync(file));
    expect(unzip(committed).every((e) => e.method === 8)).toBe(true);
    expect(partsOf(committed).map(([name, , text]) => [name, text])).toEqual(
      partsOf(await writeXlsx(minimalWorkbook, STORE)).map(([name, , text]) => [name, text]),
    );
  });

  it("falls back to STORE, byte-identical to the golden file, when the runtime cannot deflate", async () => {
    vi.stubGlobal("CompressionStream", undefined);
    const bytes = await writeXlsx(minimalWorkbook);
    expect(bytes).toEqual(new Uint8Array(readFileSync(join(fixturesDir, "minimal.xlsx"))));
  });
});

describe("multi-sheet workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("multi-sheet.xlsx", multiSheetWorkbook);
  });

  it("has one worksheet part, one relationship and one content-type override per sheet, in tab order", async () => {
    const bytes = await writeXlsx(multiSheetWorkbook, STORE);
    const names = unzip(bytes).map((e) => e.name);
    expect(names.filter((n) => n.startsWith("xl/worksheets/"))).toEqual([
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
    ]);
    const contentTypes = readPart(bytes, "[Content_Types].xml");
    expect(contentTypes.match(/worksheets\/sheet\d\.xml/g)).toEqual([
      "worksheets/sheet1.xml",
      "worksheets/sheet2.xml",
      "worksheets/sheet3.xml",
    ]);
    const relationships = readPart(bytes, "xl/_rels/workbook.xml.rels");
    for (let i = 1; i <= 3; i++) {
      expect(relationships).toContain(
        `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`,
      );
    }
    expect(relationships).toContain(
      'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"',
    );
    expect(relationships).toContain(
      'Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"',
    );
  });

  it("escapes sheet names in the workbook part and keeps their order", async () => {
    const workbook = readPart(await writeXlsx(multiSheetWorkbook, STORE), "xl/workbook.xml");
    expect(workbook).toContain(
      '<sheets><sheet name="Orders &amp; Returns" sheetId="1" r:id="rId1"/><sheet name="Ærlig &lt;oversikt&gt;" sheetId="2" r:id="rId2"/><sheet name="Empty" sheetId="3" r:id="rId3"/></sheets>',
    );
  });

  it("shares one strings table across sheets", async () => {
    const bytes = await writeXlsx(multiSheetWorkbook, STORE);
    const sst = readPart(bytes, "xl/sharedStrings.xml");
    expect(sst.match(/<t>shipped<\/t>/g)).toHaveLength(1);
    expect(sst).toContain('count="10" uniqueCount="7"');
    expect(readPart(bytes, "xl/worksheets/sheet3.xml")).toContain("<sheetData/>");
  });

  it("rejects a duplicate sheet name through the public function, at the second sheet", async () => {
    const duplicate = {
      sheets: [multiSheetWorkbook.sheets[0]!, { name: "orders & RETURNS", rows: [] }],
    };
    await expect(writeXlsx(duplicate)).rejects.toMatchObject({
      code: "DUPLICATE_SHEET_NAME",
      path: "sheets[1].name",
    });
  });
});

describe("styles workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("styles.xlsx", stylesWorkbook);
  });

  it("declares exactly one font, number format and cell format per distinct style used", async () => {
    const styles = readPart(await writeXlsx(stylesWorkbook, STORE), "xl/styles.xml");
    expect(styles).toContain('<numFmts count="3">');
    expect(styles).toContain('<numFmt numFmtId="164" formatCode="#,##0.00 &quot;kr&quot;"/>');
    expect(styles).toContain('<numFmt numFmtId="165" formatCode="0.0%"/>');
    expect(styles).toContain('<numFmt numFmtId="166" formatCode="yyyy-mm-dd"/>');
    expect(styles).toContain('<fonts count="2">');
    // default, bold, bold+kroner, bold+percent, bold+iso, kroner, percent, iso, date-time
    expect(styles).toContain('<cellXfs count="9">');
  });

  it("references the right format from each cell", async () => {
    const sheet = readPart(await writeXlsx(stylesWorkbook, STORE), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="A1" s="1" t="s">');
    expect(sheet).toContain('<c r="B1" s="2" t="s">');
    expect(sheet).toContain('<c r="B2" s="5"><v>1234.5</v></c>');
    expect(sheet).toContain('<c r="C2" s="6"><v>0.1</v></c>');
    expect(sheet).toContain('<c r="D2" s="7"><v>46327</v></c>');
    expect(sheet).toContain('<c r="E2" s="8"><v>46329.604166666664</v></c>');
    expect(sheet).toContain('<c r="A5" s="1" t="s">');
    expect(sheet).toContain('<c r="B5" s="2"><v>101333.5</v></c>');
  });

  it("keeps one cell format for the same style object on ten thousand cells", async () => {
    const money = { numberFormat: '#,##0.00 "kr"' };
    const rows = Array.from({ length: 10_000 }, (_, i) => [{ value: i, style: money }]);
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows }] }, STORE);
    expect(readPart(bytes, "xl/styles.xml")).toContain('<cellXfs count="2">');
  });
});

describe("layout workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("layout.xlsx", layoutWorkbook);
  });

  it("declares the hidden filter name for the filtered sheet only, with the sheet name quoted", async () => {
    const workbook = readPart(await writeXlsx(layoutWorkbook, STORE), "xl/workbook.xml");
    expect(workbook).toContain(
      '</sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">&apos;Orders 2026&apos;!$A$1:$C$4</definedName></definedNames></workbook>',
    );
    expect(workbook.match(/definedName /g)).toHaveLength(1);
  });

  it("doubles apostrophes in a sheet name inside the filter reference", async () => {
    const workbook = readPart(
      await writeXlsx(
        { sheets: [{ name: "Per O'Brien", autoFilter: true, rows: [["a"]] }] },
        STORE,
      ),
      "xl/workbook.xml",
    );
    expect(workbook).toContain(">&apos;Per O&apos;&apos;Brien&apos;!$A$1:$A$1</definedName>");
  });

  it("writes no definedNames element when nothing is filtered", async () => {
    expect(readPart(await writeXlsx(minimalWorkbook, STORE), "xl/workbook.xml")).not.toContain(
      "definedNames",
    );
  });

  it("carries widths, the frozen pane and the autofilter in the sheet part", async () => {
    const sheet = readPart(await writeXlsx(layoutWorkbook, STORE), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<col min="2" max="2" width="32" customWidth="1"/>');
    expect(sheet).toContain(
      '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
    );
    expect(sheet).toContain('<autoFilter ref="A1:C4"/>');
  });
});

describe("orders workbook from objects (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("orders.xlsx", ordersWorkbook);
  });

  it("carries the bold header, widths, formats, frozen header row and autofilter", async () => {
    const bytes = await writeXlsx(ordersWorkbook, STORE);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
    );
    expect(sheet).toContain('<col min="4" max="4" width="18" customWidth="1"/>');
    expect(sheet).toContain('<autoFilter ref="A1:D4"/>');
    expect(sheet).toContain('<c r="A1" s="1" t="s"><v>0</v></c>');
    expect(sheet).toContain('<c r="D2" s="5"><v>46327.395833333336</v></c>');
    const styles = readPart(bytes, "xl/styles.xml");
    expect(styles).toContain('formatCode="yyyy-mm-dd hh:mm"');
    expect(styles).toContain('<fonts count="2">');
  });
});

describe("appearance workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("appearance.xlsx", appearanceWorkbook);
  });

  it("declares fonts, fills, borders and alignments once each and references them from the cells", async () => {
    const bytes = await writeXlsx(appearanceWorkbook, STORE);
    const styles = readPart(bytes, "xl/styles.xml");
    // default, title (bold 16 white), bold, italic grey
    expect(styles).toContain('<fonts count="4">');
    expect(styles).toContain(
      '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',
    );
    expect(styles).toContain(
      '<font><i/><sz val="11"/><color rgb="FF808080"/><name val="Calibri"/><family val="2"/></font>',
    );
    // none, gray125, title blue, heading light blue
    expect(styles).toContain('<fills count="4">');
    // empty, thin all round, medium top with thick bottom
    expect(styles).toContain('<borders count="3">');
    expect(styles).toContain('<alignment horizontal="center" vertical="center"/>');
    expect(styles).toContain('<alignment vertical="top" wrapText="1"/>');
    expect(styles).toContain('<alignment horizontal="left" indent="1"/>');
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain(
      '<row r="1" ht="28" customHeight="1"><c r="A1" s="1" t="s"><v>0</v></c></row>',
    );
    expect(sheet).toContain('<row r="3" ht="48" customHeight="1">');
    expect(sheet).toContain(
      '<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells></worksheet>',
    );
  });

  it("writes a cell with a style and no value, so a border can run across an empty cell", async () => {
    const sheet = readPart(await writeXlsx(appearanceWorkbook, STORE), "xl/worksheets/sheet1.xml");
    expect(sheet).toMatch(/<c r="B5" s="\d+"\/>/);
  });
});
