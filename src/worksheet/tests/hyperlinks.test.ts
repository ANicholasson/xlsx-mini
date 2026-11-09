import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import {
  type RenderedWorksheet,
  type Sheet,
  type WorksheetContext,
  renderWorksheet,
  streamWorksheet,
} from "../index.js";

function context(overrides: Partial<WorksheetContext> = {}): WorksheetContext {
  return {
    strings: createStringTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["Links", "Per O'Brien"],
    ...overrides,
  };
}

const render = (sheet: Sheet, overrides?: Partial<WorksheetContext>): RenderedWorksheet =>
  renderWorksheet(sheet, context(overrides));
const xml = (sheet: Sheet, overrides?: Partial<WorksheetContext>) =>
  new TextDecoder().decode(render(sheet, overrides).data);

describe("hyperlink cells", () => {
  it("writes external targets as relationships of the sheet, numbered in order and escaped", () => {
    const rendered = render({
      name: "Links",
      rows: [
        [{ value: "Docs", href: "https://example.com/?a=1&b=2" }, "plain"],
        [{ value: "Mail", href: "mailto:hello@example.com" }],
      ],
    });
    const sheet = new TextDecoder().decode(rendered.data);
    expect(sheet).toContain(
      '</sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" r:id="rId2"/></hyperlinks></worksheet>',
    );
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(rendered.relationships).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:hello@example.com" TargetMode="External"/>' +
        "</Relationships>",
    );
  });

  it("writes a jump within the workbook as a location, on this sheet by default, with the name quoted", () => {
    const rendered = render(
      {
        name: "Per O'Brien",
        rows: [[{ value: "Top", cell: "A1" }], [{ value: "Links", sheet: "links", cell: "A1:B5" }]],
      },
      { sheetIndex: 1 },
    );
    const sheet = new TextDecoder().decode(rendered.data);
    expect(sheet).toContain(
      '<hyperlinks><hyperlink ref="A1" location="&apos;Per O&apos;&apos;Brien&apos;!A1"/><hyperlink ref="A2" location="&apos;links&apos;!A1:B5"/></hyperlinks>',
    );
    expect(rendered.relationships).toBeUndefined();
  });

  it("accepts the schemes applications open: https, mailto, file, tel and an app's own", () => {
    const rendered = render({
      name: "Links",
      rows: [
        [
          "https://a.b",
          "mailto:x@y.z",
          "file:///share/report.pdf",
          "tel:+4712345678",
          "obsidian://open?vault=v",
        ].map((href) => ({ value: "x", href })),
      ],
    });
    expect(rendered.relationships?.match(/TargetMode="External"/g)).toHaveLength(5);
  });

  it("writes no hyperlinks element and no relationships part for a sheet without links", () => {
    const rendered = render({ name: "Links", rows: [["a"]] });
    expect(new TextDecoder().decode(rendered.data)).not.toContain("hyperlinks");
    expect(rendered.relationships).toBeUndefined();
  });

  it("keeps a link on a styled empty cell and on a number", () => {
    const sheet = xml({
      name: "Links",
      rows: [
        [
          { value: 42, href: "https://example.com" },
          { value: null, cell: "B9", style: { bold: true } },
        ],
      ],
    });
    expect(sheet).toContain('<c r="A1"><v>42</v></c><c r="B1" s="1"/>');
    expect(sheet).toContain('<hyperlink ref="B1" location="&apos;Links&apos;!B9"/>');
  });

  it("refuses a target that is not an absolute URL, or whose scheme runs code, at href", () => {
    for (const href of [
      "/docs",
      "example.com",
      "",
      42,
      "javascript:alert(1)",
      "VBScript:x",
      "data:text/html,x",
    ]) {
      expect(() => xml({ name: "Links", rows: [[{ value: "x", href }]] } as Sheet)).toThrow(
        expect.objectContaining({
          code: "INVALID_HYPERLINK",
          path: "sheets[0].rows[0][0].href",
          message: expect.stringContaining("absolute URL"),
        }),
      );
    }
  });

  it("refuses a link with both an href and a cell, at the cell", () => {
    expect(() =>
      xml({ name: "Links", rows: [[{ value: "x", href: "https://a.b", cell: "A1" } as never]] }),
    ).toThrow(expect.objectContaining({ code: "INVALID_HYPERLINK", path: "sheets[0].rows[0][0]" }));
  });

  it("refuses a sheet without a cell to jump to", () => {
    expect(() => xml({ name: "Links", rows: [[{ value: "x", sheet: "Links" } as never]] })).toThrow(
      expect.objectContaining({ code: "INVALID_HYPERLINK", path: "sheets[0].rows[0][0].cell" }),
    );
  });

  it("refuses a cell that is not in A1 form or lies outside the sheet, and an unknown sheet", () => {
    expect(() => xml({ name: "Links", rows: [[{ value: "x", cell: "$A$1" }]] })).toThrow(
      expect.objectContaining({
        code: "INVALID_HYPERLINK",
        path: "sheets[0].rows[0][0].cell",
        message: expect.stringContaining('such as "B2", not "$A$1"'),
      }),
    );
    expect(() => xml({ name: "Links", rows: [[{ value: "x", cell: "A1048577" }]] })).toThrow(
      expect.objectContaining({
        code: "INVALID_HYPERLINK",
        path: "sheets[0].rows[0][0].cell",
        message: expect.stringContaining("outside the sheet"),
      }),
    );
    expect(() =>
      xml({ name: "Links", rows: [[{ value: "x", sheet: "Summary", cell: "A1" }]] }),
    ).toThrow(
      expect.objectContaining({
        code: "INVALID_HYPERLINK",
        path: "sheets[0].rows[0][0].sheet",
        message: expect.stringContaining('no sheet named "Summary"'),
      }),
    );
  });

  it("reports the relationships of a streamed sheet once its body has been read", async () => {
    const streamed = streamWorksheet(
      { name: "Links", rows: [[{ value: "Docs", href: "https://example.com" }]] },
      context(),
    );
    expect(streamed.relationships()).toBeUndefined();
    const body = await new Response(streamed.body).text();
    expect(body).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(streamed.relationships()).toContain('Target="https://example.com"');
  });
});
