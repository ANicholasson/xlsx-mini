import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { writeXlsx, type Workbook } from "../src/index.js";
import { readPart, unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

/** The workbook behind fixtures/minimal.xlsx. Change it only together with the fixture. */
export const minimalWorkbook: Workbook = {
  sheets: [
    {
      name: "Sheet1",
      rows: [
        ["Name", "City", "Amount"],
        ["Kari Nordmann", "Oslo", 1234.5],
        ["Ola <Nordmann> & co", "Tromsø", -42],
        ["Ærlig «ekte» 🎉", null, 1e21],
        [" padded ", "", 0],
      ],
    },
  ],
};

async function golden(name: string, workbook: Workbook): Promise<Uint8Array> {
  const bytes = await writeXlsx(workbook);
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
