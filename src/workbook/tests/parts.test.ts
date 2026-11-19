import { describe, expect, it } from "vite-plus/test";
import { type SheetEntry, renderWorkbook } from "../index.js";

const plain = (name: string): SheetEntry => ({
  name,
  hidden: false,
  filterRange: undefined,
  printTitles: undefined,
  hasFormulas: false,
});

describe("renderWorkbook", () => {
  it("writes no defined names and no calcPr for plain sheets", () => {
    const xml = renderWorkbook([plain("A"), plain("B")]);
    expect(xml).toContain(
      '<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets></workbook>',
    );
  });

  it("writes a carriage return in a sheet name as itself, in the name and in a reference to the sheet", () => {
    const xml = renderWorkbook([{ ...plain("a\rb"), printTitles: "$1:$1" }]);
    expect(xml).toContain('<sheet name="a\rb" sheetId="1" r:id="rId1"/>');
    expect(xml).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;a\rb&apos;!$1:$1</definedName>',
    );
  });

  it("marks a hidden sheet, and names the first sheet that is shown when a hidden one precedes it", () => {
    const hidden = (name: string): SheetEntry => ({ ...plain(name), hidden: true });
    expect(renderWorkbook([hidden("A"), hidden("B"), plain("C"), plain("D")])).toContain(
      'relationships"><bookViews><workbookView activeTab="2"/></bookViews>' +
        '<sheets><sheet name="A" sheetId="1" state="hidden" r:id="rId1"/>' +
        '<sheet name="B" sheetId="2" state="hidden" r:id="rId2"/>' +
        '<sheet name="C" sheetId="3" r:id="rId3"/><sheet name="D" sheetId="4" r:id="rId4"/></sheets>',
    );
    // The first sheet is shown: the workbook opens on it as it always did.
    expect(renderWorkbook([plain("A"), hidden("B")])).toContain(
      'relationships"><sheets><sheet name="A" sheetId="1" r:id="rId1"/>' +
        '<sheet name="B" sheetId="2" state="hidden" r:id="rId2"/></sheets>',
    );
  });

  it("names the print titles next to the filter database, per sheet, and asks Excel to calculate once any sheet has a formula", () => {
    const xml = renderWorkbook([
      { ...plain("Per O'Brien"), filterRange: "A1:C4", printTitles: "$1:$2" },
      { ...plain("Totals"), hasFormulas: true },
    ]);
    expect(xml).toContain(
      "<definedNames>" +
        '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">&apos;Per O&apos;&apos;Brien&apos;!$A$1:$C$4</definedName>' +
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Per O&apos;&apos;Brien&apos;!$1:$2</definedName>' +
        "</definedNames>" +
        '<calcPr fullCalcOnLoad="1"/></workbook>',
    );
  });
});
