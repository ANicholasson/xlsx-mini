import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { createMediaTable } from "../../drawing/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import { type Sheet, type WorksheetContext, renderWorksheet } from "../index.js";

function context(): WorksheetContext {
  return {
    strings: createStringTable(),
    media: createMediaTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
    inlineStrings: true,
  };
}

const render = (sheet: Omit<Sheet, "name">) => renderWorksheet({ name: "S", ...sheet }, context());
const xml = (sheet: Omit<Sheet, "name">) => new TextDecoder().decode(render(sheet).data);
const cols = (sheet: Omit<Sheet, "name">) => /<cols>(.*)<\/cols>/.exec(xml(sheet))?.[1];
const data = (sheet: Omit<Sheet, "name">) => /<sheetData>(.*)<\/sheetData>/.exec(xml(sheet))![1]!;

const images = join(dirname(fileURLToPath(import.meta.url)), "../../../test/fixtures/images");
const stripes = new Uint8Array(readFileSync(join(images, "stripes.png")));

describe("a hidden column", () => {
  it("is written without a width, and keeps the width it is given", () => {
    expect(
      cols({ rows: [], columns: [{ hidden: true }, undefined, { width: 20, hidden: true }] }),
    ).toBe(
      '<col min="1" max="1" hidden="1"/><col min="3" max="3" width="20" customWidth="1" hidden="1"/>',
    );
  });

  it("is fitted to its cells like any other, for when it is shown again", () => {
    const rows = [["a long header", "a long header"]];
    const fitted = cols({ rows, columns: [{ width: "auto" }, { width: "auto", hidden: true }] })!;
    const width = /width="([\d.]+)"/.exec(fitted)![1]!;
    expect(fitted).toBe(
      `<col min="1" max="1" width="${width}" customWidth="1"/>` +
        `<col min="2" max="2" width="${width}" customWidth="1" hidden="1"/>`,
    );
  });

  it("writes nothing when hidden is false or left out", () => {
    expect(cols({ rows: [["a"]], columns: [{ hidden: false }, {}] })).toBeUndefined();
  });
});

describe("a hidden row", () => {
  it("keeps its cells, its style and its height", () => {
    expect(
      data({
        rows: [[1], { cells: [2], hidden: true }, { cells: [3], hidden: true, height: 30 }, [4]],
      }),
    ).toBe(
      '<row r="1"><c r="A1"><v>1</v></c></row>' +
        '<row r="2" hidden="1"><c r="A2"><v>2</v></c></row>' +
        '<row r="3" ht="30" customHeight="1" hidden="1"><c r="A3"><v>3</v></c></row>' +
        '<row r="4"><c r="A4"><v>4</v></c></row>',
    );
  });

  it("is written also without cells, and does not stretch the used range", () => {
    const rendered = render({ rows: [[1], { cells: [], hidden: true }, { cells: [null] }] });
    const text = new TextDecoder().decode(rendered.data);
    expect(text).toContain('<dimension ref="A1:A1"/>');
    expect(text).toContain('</row><row r="2" hidden="1"/></sheetData>');
  });

  it("writes what the same row writes when hidden is false", () => {
    expect(data({ rows: [{ cells: [1], hidden: false }] })).toBe(data({ rows: [[1]] }));
  });

  it("counts for an auto width, as it does when Excel fits a column", () => {
    const columns = [{ width: "auto" as const }];
    const long = "a text much longer than the others in its column";
    expect(cols({ columns, rows: [["a"], { cells: [long], hidden: true }] })).toBe(
      cols({ columns, rows: [["a"], [long]] }),
    );
  });
});

describe("a picture in a hidden row or column", () => {
  it("is anchored to fold away with it, and any other picture keeps its size", () => {
    const image = { image: stripes };
    const rendered = render({
      columns: [undefined, { hidden: true }],
      rows: [[image, image], { cells: [image], hidden: true }, [image]],
    });
    const anchors = [
      ...rendered.drawing!.drawing.matchAll(/<xdr:(\w+)(?: editAs="twoCell")?><xdr:from>/g),
    ];
    expect(anchors.map((anchor) => anchor[1])).toEqual([
      "oneCellAnchor",
      "twoCellAnchor",
      "twoCellAnchor",
      "oneCellAnchor",
    ]);
    // The row and the column still make room for the picture, for when they are shown again.
    const text = new TextDecoder().decode(rendered.data);
    expect(text).toContain('<col min="2" max="2" width="18.86" customWidth="1" hidden="1"/>');
    expect(text).toContain('<row r="2" ht="33" customHeight="1" hidden="1"/>');
  });
});
