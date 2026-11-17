import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { createMediaTable } from "../../drawing/index.js";
import type { XlsxMiniError } from "../../errors/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import { type Cell, type Sheet, renderWorksheet } from "../index.js";

const images = join(dirname(fileURLToPath(import.meta.url)), "../../../test/fixtures/images");
const stripes = new Uint8Array(readFileSync(join(images, "stripes.png")));
const badge = new Uint8Array(readFileSync(join(images, "badge.png")));

function render(sheet: Sheet) {
  const rendered = renderWorksheet(sheet, {
    strings: createStringTable(),
    styles: createStyleTable(),
    media: createMediaTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
  });
  return { ...rendered, xml: new TextDecoder().decode(rendered.data) };
}

const failure = (sheet: Sheet): XlsxMiniError => {
  try {
    render(sheet);
  } catch (error) {
    return error as XlsxMiniError;
  }
  throw new Error("expected a throw");
};

describe("image cells", () => {
  it("anchors the picture, grows the row to it and the column to it, and writes no cell beneath", () => {
    const { xml, relationships, drawing } = render({
      name: "S",
      rows: [
        ["Product", "Picture"],
        ["Desk", { image: stripes }],
      ],
    });
    expect(xml).toContain('<cols><col min="2" max="2" width="18.86" customWidth="1"/></cols>');
    expect(xml).toContain('<row r="2" ht="33" customHeight="1"><c r="A2" t="s"><v>2</v></c></row>');
    expect(xml).toContain('</sheetData><drawing r:id="rId1"/></worksheet>');
    expect(relationships).toContain(
      'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>',
    );
    expect(drawing?.drawing).toContain("<xdr:col>1</xdr:col>");
    expect(drawing?.drawing).toContain("<xdr:row>1</xdr:row>");
    expect(drawing?.relationships).toContain('Target="../media/image1.png"');
  });

  it("writes the styled cell beneath, and nothing for a cell whose style is empty", () => {
    const { xml } = render({
      name: "S",
      rows: [
        [
          { image: badge, style: { fill: "#d9e1f2" } },
          { image: badge, style: {} },
        ],
      ],
    });
    expect(xml).toContain('<row r="1" ht="39" customHeight="1"><c r="A1" s="1"/></row>');
  });

  it("never shrinks a row below the default height for a small picture", () => {
    const { xml } = render({ name: "S", rows: [["Text", { image: stripes, width: 12 }]] });
    expect(xml).toContain('<row r="1" ht="15" customHeight="1">');
  });

  it("counts picture cells in the used range, so the dimension and the filter reach them", () => {
    const rendered = render({
      name: "S",
      autoFilter: true,
      rows: [["Picture"], [{ image: stripes }], [null, { image: badge }]],
    });
    expect(rendered.xml).toContain('<dimension ref="A1:B3"/>');
    expect(rendered.xml).toContain('<autoFilter ref="A1:B3"/>');
    expect(rendered.xml).toContain(
      '<row r="2" ht="33" customHeight="1"/><row r="3" ht="39" customHeight="1"/>',
    );
    expect(rendered.filterRange).toBe("A1:B3");
  });

  it("takes the tallest picture in a row, and the caller's row height over its own", () => {
    const { xml } = render({
      name: "S",
      rowHeights: { 2: 60, 3: 409 },
      rows: [
        [{ image: badge, height: 24 }, { image: stripes }],
        [{ image: stripes }],
        [{ image: stripes, height: 600 }],
        [{ image: stripes, height: 541 }],
      ],
    });
    expect(xml).toContain('<row r="1" ht="33" customHeight="1"/>');
    expect(xml).toContain('<row r="2" ht="60" customHeight="1"/>');
    expect(xml).toContain('<row r="3" ht="409" customHeight="1"/>');
    expect(xml).toContain('<row r="4" ht="408.75" customHeight="1"/>');
  });

  it("refuses a picture no row can hold unless the caller sizes the row, at the tallest picture's cell", () => {
    expect(
      failure({ name: "S", rows: [["a"], [{ image: badge }, { image: stripes, height: 542 }]] }),
    ).toMatchObject({ code: "INVALID_IMAGE", path: "sheets[0].rows[1][1]" });
    const { xml } = render({
      name: "S",
      rowHeights: { 1: 100 },
      rows: [[{ image: stripes, height: 600 }]],
    });
    expect(xml).toContain('<row r="1" ht="100" customHeight="1"/>');
  });

  it("widens a column to its widest picture unless it has a numeric width, and counts it under auto width", () => {
    const { xml } = render({
      name: "S",
      autoWidth: true,
      columns: [{ width: 5 }, undefined, { width: "auto" }],
      rows: [
        [{ image: stripes }, { image: badge }, "A rather long text in the third column"],
        [null, { image: stripes, width: 60 }, { image: stripes, width: 400 }],
      ],
    });
    expect(xml).toContain(
      '<cols><col min="1" max="1" width="5" customWidth="1"/><col min="2" max="2" width="10.29" customWidth="1"/><col min="3" max="3" width="58.86" customWidth="1"/></cols>',
    );
    const wider = render({
      name: "S",
      autoWidth: true,
      rows: [["A text wider than the badge beneath it"], [{ image: badge }]],
    });
    expect(wider.xml).toContain('<col min="1" max="1" width="35.57" customWidth="1"/>');
    const plain = render({ name: "S", rows: [[{ image: stripes, width: 2000, height: 40 }]] });
    expect(plain.xml).toContain('<col min="1" max="1" width="255" customWidth="1"/>');
  });

  it("grows the row and the column to a picture and its padding, so the picture sits in the middle of its row", () => {
    const { xml, drawing } = render({
      name: "S",
      rows: [[{ image: stripes, height: 30, padding: 10 }], [{ image: badge, padding: 0 }]],
    });
    // 90 by 30 pixels and 10 on every side: 110 pixels across, 50 down.
    expect(xml).toContain('<cols><col min="1" max="1" width="15.71" customWidth="1"/></cols>');
    expect(xml).toContain('<row r="1" ht="37.5" customHeight="1"/>');
    expect(drawing?.drawing).toContain(
      '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:from><xdr:ext cx="857250" cy="285750"/>',
    );
    // No padding at all: the row is the picture's height and the picture starts at the corner.
    expect(xml).toContain('<row r="2" ht="36" customHeight="1"/>');
    expect(drawing?.drawing).toContain(
      "<xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>",
    );
    const exact = render({ name: "S", rows: [[{ image: badge, padding: 0 }]] });
    expect(exact.xml).toContain('<col min="1" max="1" width="6.86" customWidth="1"/>');
  });

  it("sizes a row for the picture that needs the most with its padding, and keeps a caller's height and width", () => {
    const { xml, drawing } = render({
      name: "S",
      columns: [{ width: 30 }],
      rows: [
        [{ image: stripes, padding: 10 }, { image: badge }],
        [
          { image: stripes, height: 20 },
          { image: badge, height: 20, padding: 20 },
        ],
        { cells: [{ image: stripes, padding: 10 }], height: 100 },
      ],
    });
    expect(xml).toContain('<row r="1" ht="45" customHeight="1"/>');
    expect(xml).toContain('<row r="2" ht="45" customHeight="1"/>');
    expect(xml).toContain('<row r="3" ht="100" customHeight="1"/>');
    // Column A keeps its width; column B takes the padded badge, 20 and 40 pixels, over the plain one's 48 and 12.
    expect(xml).toContain(
      '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="8.57" customWidth="1"/></cols>',
    );
    expect(drawing?.drawing).toContain(
      "<xdr:col>0</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>95250</xdr:rowOff>",
    );
  });

  it("refuses a picture whose padding takes it past what a row can hold, and a padding that is not a number from 0", () => {
    const fits = render({ name: "S", rows: [[{ image: stripes, height: 345, padding: 100 }]] });
    expect(fits.xml).toContain('<row r="1" ht="408.75" customHeight="1"/>');
    const error = failure({
      name: "S",
      rows: [[{ image: stripes, height: 346, padding: 100 }]],
    });
    expect(error).toMatchObject({ code: "INVALID_IMAGE", path: "sheets[0].rows[0][0]" });
    expect(error.message).toContain("546 pixels");
    for (const padding of [-1, null, "10", Number.NaN]) {
      expect(
        failure({ name: "S", rows: [[{ image: badge, padding: padding as never }]] }),
      ).toMatchObject({ code: "INVALID_IMAGE", path: "sheets[0].rows[0][0].padding" });
    }
  });

  it("writes one media part per distinct image and numbers the drawing's relationships by first use", () => {
    const media = createMediaTable();
    const rendered = renderWorksheet(
      { name: "S", rows: [[{ image: badge }, { image: stripes }, { image: badge }]] },
      {
        strings: createStringTable(),
        styles: createStyleTable(),
        media,
        coerce: undefined,
        sheetIndex: 0,
        sheetNames: ["S"],
      },
    );
    expect(media.takeParts().map((part) => part.path)).toEqual([
      "xl/media/image1.png",
      "xl/media/image2.png",
    ]);
    expect(rendered.drawing?.drawing.match(/r:embed="rId1"/g)).toHaveLength(2);
  });

  it("numbers the drawing relationship after the link relationships", () => {
    const { xml, relationships } = render({
      name: "S",
      rows: [[{ value: "Docs", href: "https://example.com" }, { image: badge }]],
    });
    expect(xml).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(xml).toContain('<drawing r:id="rId2"/>');
    expect(relationships).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"',
    );
  });

  it("refuses an image cell that also holds a value, formula or link, and bytes that are not a PNG", () => {
    const unsafe = (cell: unknown): Cell => cell as Cell;
    for (const extra of [
      { value: "x" },
      { formula: "A1" },
      { href: "https://e.com" },
      { cell: "A1" },
    ]) {
      expect(failure({ name: "S", rows: [[unsafe({ image: badge, ...extra })]] })).toMatchObject({
        code: "INVALID_IMAGE",
        path: "sheets[0].rows[0][0]",
      });
    }
    expect(
      failure({ name: "S", rows: [[], [null, unsafe({ image: "badge.png" })]] }),
    ).toMatchObject({
      code: "INVALID_IMAGE",
      path: "sheets[0].rows[1][1].image",
    });
    expect(failure({ name: "S", rows: [[{ image: badge, width: -1 }]] })).toMatchObject({
      code: "INVALID_IMAGE",
      path: "sheets[0].rows[0][0].width",
    });
    // Bare bytes are not an image: the kind is named.
    expect(failure({ name: "S", rows: [[unsafe(badge)]] }).code).toBe("UNSUPPORTED_VALUE");
  });

  it("reads rows from a generator once, pictures included", () => {
    let pulls = 0;
    function* rows(): Generator<Cell[]> {
      for (let i = 0; i < 3; i++) {
        pulls++;
        yield [`Row ${i + 1}`, { image: i === 1 ? badge : stripes }];
      }
    }
    const { xml, drawing } = render({ name: "S", rows: rows() });
    expect(pulls).toBe(3);
    expect(xml).toContain('<dimension ref="A1:B3"/>');
    expect(drawing?.drawing.match(/<xdr:oneCellAnchor>/g)).toHaveLength(3);
  });

  it("leaves a sheet without images without a drawing", () => {
    const { xml, relationships, drawing } = render({ name: "S", rows: [["a"]] });
    expect(xml).not.toContain("<drawing");
    expect(relationships).toBeUndefined();
    expect(drawing).toBeUndefined();
  });
});
