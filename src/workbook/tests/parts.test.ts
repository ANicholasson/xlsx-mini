import { describe, expect, it } from "vite-plus/test";
import { type SheetEntry, renderWorkbook } from "../index.js";

const plain = (name: string): SheetEntry => ({
  name,
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
