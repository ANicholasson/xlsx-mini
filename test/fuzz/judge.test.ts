import { describe, expect, it } from "vite-plus/test";
import { type Outcome } from "../helpers/outcome.ts";
import { bytesOf, packageOf } from "../helpers/package.ts";
import { type Book, bookOf, sheetEntry, workbookPart, worksheet } from "../hostile/book.ts";
import { judge } from "./judge.ts";
import { sheetHoldsValues, sheetListed } from "./mutate.ts";

const ROW = '<row r="1"><c r="A1"><v>1</v></c></row>';
const pack = (book: Book = {}) => packageOf(bookOf(book));
const original: Outcome = { sheets: [{ name: "Sheet1", rows: [[1, "a", 1]] }] };
const without: Outcome = { sheets: [] };
const emptied: Outcome = { sheets: [{ name: "Sheet1", rows: [] }] };
const renamed = pack({ workbook: workbookPart(sheetEntry("Other", "rId1")) });

describe("the judgment of a mutated package's read", () => {
  it("passes a read of the original's values and a refusal with a code", () => {
    expect(judge(original, original, { bytes: pack(), same: true })).toEqual({
      ended: "read",
      excused: 0,
      failures: [],
    });
    const refused: Outcome = { refused: { code: "INVALID_XML" } };
    expect(judge(original, refused, { bytes: pack(), same: true })).toEqual({
      ended: "INVALID_XML",
      excused: 0,
      failures: [],
    });
    const sheetRefused: Outcome = {
      sheets: [{ name: "Sheet1", refused: { code: "INVALID_CELL", sheet: "Sheet1", cell: "A1" } }],
    };
    expect(judge(original, sheetRefused, { bytes: pack(), same: false }).ended).toBe(
      "INVALID_CELL",
    );
  });

  it("fails other values from a package whose parts hold what the original's hold", () => {
    const other: Outcome = { sheets: [{ name: "Sheet1", rows: [[1, "a", 2]] }] };
    expect(judge(original, other, { bytes: pack(), same: false }).failures).toEqual([]);
    expect(judge(original, other, { bytes: pack(), same: true }).failures).toEqual([
      expect.stringMatching(/^read values the original does not hold: /),
    ]);
  });

  it("fails a workbook read without a sheet that had rows, unless the workbook part lists it no more", () => {
    expect(judge(original, without, { bytes: pack(), same: false })).toEqual({
      ended: "read other values",
      excused: 0,
      failures: ['sheet "Sheet1" had rows and is not read, though the workbook part lists it'],
    });
    expect(judge(original, without, { bytes: renamed, same: false })).toEqual({
      ended: "read other values",
      excused: 1,
      failures: [],
    });
    expect(
      judge(original, without, { bytes: bytesOf("not a package"), same: false }).failures,
    ).toEqual([
      'sheet "Sheet1" had rows and is not read, and the workbook part cannot be followed',
    ]);
  });

  it("fails a sheet that had rows and reads with none, unless its part holds no value any more", () => {
    expect(judge(original, emptied, { bytes: pack(), same: false }).failures).toEqual([
      'sheet "Sheet1" had rows and reads with none, though its part holds a value',
    ]);
    const hollow = pack({ sheets: [["Sheet1", worksheet("")]] });
    expect(judge(original, emptied, { bytes: hollow, same: false })).toEqual({
      ended: "read other values",
      excused: 1,
      failures: [],
    });
    const unrelated = pack({ workbook: workbookPart(sheetEntry("Sheet1", "rId9")) });
    expect(judge(original, emptied, { bytes: unrelated, same: false }).failures).toEqual([
      'sheet "Sheet1" had rows and reads with none, and its part cannot be followed',
    ]);
  });

  it("asks nothing of a sheet that had no rows", () => {
    const bare: Outcome = { sheets: [{ name: "Sheet1", rows: [] }] };
    expect(judge(bare, without, { bytes: pack(), same: false }).failures).toEqual([]);
  });
});

describe("the search of a package's text", () => {
  it("follows a sheet whose name holds a `>` as Excel writes it, and attributes in either quote", () => {
    const name = "Ærlig <oversikt>";
    const excel = pack({
      sheets: [["Ærlig &lt;oversikt>", worksheet(ROW)]],
    });
    expect(sheetListed(excel, name)).toBe(true);
    expect(sheetHoldsValues(excel, name)).toBe(true);
    const single = pack({
      workbook: workbookPart("<sheet name='Sheet1' sheetId='1' r:id='rId1'/>"),
      sheets: [
        ["Sheet1", worksheet("<row r='1'><c r='A1' t='inlineStr'><is><t>a</t></is></c></row>")],
      ],
    });
    expect(sheetListed(single, "Sheet1")).toBe(true);
    expect(sheetHoldsValues(single, "Sheet1")).toBe(true);
    const blank = worksheet("<row r='1'><c r='A1' t='inlineStr'><is><t></t></is></c></row>");
    expect(sheetHoldsValues(pack({ sheets: [["Sheet1", blank]] }), "Sheet1")).toBe(false);
  });

  it("finds no sheet and no value inside a comment, and no sheet outside the sheet list", () => {
    const commented = pack({
      workbook: workbookPart(
        `<!--${sheetEntry("Sheet1", "rId1")}-->${sheetEntry("Other", "rId1")}`,
      ),
      sheets: [["Other", worksheet(`<!--${ROW}-->`)]],
    });
    expect(sheetListed(commented, "Sheet1")).toBe(false);
    expect(sheetListed(commented, "Other")).toBe(true);
    expect(sheetHoldsValues(commented, "Other")).toBe(false);
    const outside = pack({ workbook: workbookPart("", sheetEntry("Sheet1", "rId1")) });
    expect(sheetListed(outside, "Sheet1")).toBe(false);
  });

  it("reads a reference to a character past the last one as written", () => {
    const far = pack({ workbook: workbookPart(sheetEntry("&#x110000;&#99999999999;", "rId1")) });
    expect(sheetListed(far, "Sheet1")).toBe(false);
    expect(sheetListed(far, "&#x110000;&#99999999999;")).toBe(true);
  });

  it("follows nothing that is not a package", () => {
    expect(sheetListed(bytesOf("not a package"), "Sheet1")).toBeUndefined();
    expect(sheetHoldsValues(bytesOf("not a package"), "Sheet1")).toBeUndefined();
  });
});
