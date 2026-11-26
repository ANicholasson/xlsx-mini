import { describe, expect, it } from "vite-plus/test";
import { type ReadBack, differences, readBackOf } from "./corpus/read-back.js";

const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(" ");

const fods = (tables: string, namespaces = NAMESPACES) =>
  `<?xml version="1.0" encoding="UTF-8"?>\n<office:document ${namespaces}><office:meta/><!-- a comment --><office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document>`;

const cell = (attributes: string, content = "") =>
  `<table:table><table:table-row><table:table-cell ${attributes}>${content}</table:table-cell></table:table-row></table:table>`;

describe("readBackOf", () => {
  it("reads each kind of value, each repeat and each piece of text as LibreOffice writes them", () => {
    const tables = `
      <table:table table:name="Kinds &amp; &lt;more&gt;">
        <table:table-column table:number-columns-repeated="3"/>
        <table:table-header-rows><table:table-row>
          <table:table-cell office:value-type="string"><text:p>a<text:s text:c="2"/>b<text:s/>c<text:tab/>d<text:line-break/>e<text:span text:style-name="T1">f</text:span></text:p><text:p>g &#13;h&#x41;\r\ni<![CDATA[<j>]]></text:p></table:table-cell>
          <table:table-cell office:value-type="float" office:value="0.30000000000000004"/>
          <table:table-cell office:value-type="percentage" office:value="0.125"/>
          <table:table-cell table:formula="of:=TRUE()" office:value-type="boolean" office:boolean-value="true"/>
          <table:table-cell office:value-type="boolean" office:boolean-value="false"/>
        </table:table-row></table:table-header-rows>
        <table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
        <table:table-row-group><table:table-rows><table:table-row table:number-rows-repeated="2">
          <table:table-cell table:number-columns-repeated="2"/>
          <table:table-cell office:value-type="date" office:date-value="2026-11-17T12:34:56.4996"/>
          <table:table-cell office:value-type="time" office:time-value="PT24H17M42.71999979S"/>
          <table:table-cell table:formula="of:=1/0" office:value-type="string" office:string-value="" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>
          <table:covered-table-cell office:value-type="string"><text:p>covered</text:p></table:covered-table-cell>
          <table:table-cell office:value-type="string"><text:p></text:p></table:table-cell>
          <table:table-cell table:formula='of:="x"' office:value-type="string" office:string-value="x"><text:p>shown</text:p><office:annotation><text:p>a note</text:p></office:annotation></table:table-cell>
          <table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="2"/>
          <table:table-cell table:number-columns-repeated="1000"/>
        </table:table-row></table:table-rows></table:table-row-group>
        <table:table-row>
          <table:table-cell office:value-type="date" office:date-value="1899-12-31"/>
          <table:table-cell office:value-type="currency" office:currency="SEK" office:value="1234.5"/>
          <table:table-cell office:value-type="time" office:time-value="-PT0H0M1.5S"/>
          <table:table-cell office:value-type="date" office:date-value="1999-12-31T23:59:59.9996"/>
        </table:table-row>
      </table:table>
      <table:table table:name="Empty"><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384"/></table:table-row></table:table>`;
    const row5 = [
      null,
      null,
      { date: "2026-11-17T12:34:56.500" },
      { duration: "PT24H17M42.720S" },
      { result: { error: "#DIV/0!" } },
      "covered",
      null,
      { result: "x" },
      7,
      7,
    ];
    expect(readBackOf(fods(tables))).toEqual({
      sheets: [
        {
          name: "Kinds & <more>",
          rows: [
            [1, "a  b c\td\nef\ng \rhA\ni<j>", 0.3, 0.125, { result: true }, false],
            [5, ...row5],
            [6, ...row5],
            [
              7,
              { date: "1899-12-31T00:00:00.000" },
              1234.5,
              { duration: "-PT0H00M01.500S" },
              { date: "2000-01-01T00:00:00.000" },
            ],
          ],
        },
        { name: "Empty", rows: [] },
      ],
    });
  });

  it.each([
    [
      "a namespace of another URI",
      fods("", NAMESPACES.replace("table:1.0", "table:2.0")),
      "xmlns:table",
    ],
    [
      "a namespace left out",
      fods("", NAMESPACES.replace(/xmlns:calcext="[^"]*"/, "")),
      "xmlns:calcext",
    ],
    [
      "two bodies",
      fods("").replace("<office:body>", "<office:body></office:body><office:body>"),
      "not one <office:body>",
    ],
    ["another kind of value", fods(cell('office:value-type="void"')), "a value of type void"],
    ["a number without its value", fods(cell('office:value-type="float"')), "has no office:value"],
    [
      "another form of date",
      fods(cell('office:value-type="date" office:date-value="17/11/2026"')),
      "a date of 17/11/2026",
    ],
    [
      "another form of duration",
      fods(cell('office:value-type="time" office:time-value="18:00"')),
      "a duration of 18:00",
    ],
    ["a sheet without its name", fods("<table:table/>"), "has no table:name"],
    [
      "an element not closed",
      fods("").replace("</office:document>", ""),
      "<office:document> is not closed",
    ],
    [
      "an element closed by another",
      fods("").replace("</office:spreadsheet>", ""),
      "</office:body> closes another element",
    ],
    ["a broken tag", fods("<table:table"), "not XML at offset"],
    ["no element", "<?xml version='1.0'?>\n", "no document element"],
  ])("refuses %s", (_, text, message) => {
    expect(() => readBackOf(text)).toThrow(message);
  });
});

describe("differences", () => {
  const readBack = (name: string, ...rows: ReadBack["sheets"][number]["rows"]): ReadBack => ({
    sheets: [{ name, rows }],
  });

  it("finds none between equal read-backs", () => {
    expect(differences("w.xlsx", readBack("S", [1, "a", 1]), readBack("S", [1, "a", 1]))).toEqual(
      [],
    );
  });

  it("names the workbook, the sheet and the cell of each difference", () => {
    const committed = readBack("S", [1, "a", 1], [3, null, { date: "2026-11-17T00:00:00.000" }]);
    const made = readBack("S", [1, "a", 2], [2, true], [3]);
    expect(differences("w.xlsx", committed, made)).toEqual([
      'w.xlsx, sheet "S", B1: 1 committed, 2 now',
      'w.xlsx, sheet "S", A2: null committed, true now',
      'w.xlsx, sheet "S", B3: {"date":"2026-11-17T00:00:00.000"} committed, null now',
    ]);
    const wide = (column: number) => [1, ...Array<null>(column).fill(null), 1] as const;
    expect(differences("w.xlsx", readBack("S", wide(27)), readBack("S", wide(701)))).toEqual([
      'w.xlsx, sheet "S", AB1: 1 committed, null now',
      'w.xlsx, sheet "S", ZZ1: null committed, 1 now',
    ]);
  });

  it("names the sheets when the lists differ", () => {
    expect(differences("w.xlsx", readBack("S"), readBack("T"))).toEqual([
      'w.xlsx: sheets ["S"] committed, ["T"] now',
    ]);
  });
});
