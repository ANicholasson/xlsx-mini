import { readFileSync } from "node:fs";
import { describe, expect, it } from "vite-plus/test";
import { XlsxReadError, resolveLimits } from "../../errors/read.js";
import { createPackage, store } from "../../opc/index.js";
import { openPackage, partName } from "../../opc/read.js";
import { type ListedSheet, readWorkbook } from "../read.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MACRO_SHEET = "http://schemas.microsoft.com/office/2006/relationships";
const PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships";
const COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORKSHEET = `${TRANSITIONAL}/worksheet`;
const CHART_SHEET = `${TRANSITIONAL}/chartsheet`;

const relationship = (id: string, type: string, target: string, more = "") =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"${more}/>`;
const relationships = (...items: string[]) =>
  `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="${PACKAGE}">${items.join("")}</Relationships>`;
const sheet = (name: string, id: string, more = "") =>
  `<sheet name="${name}" sheetId="1" r:id="${id}"${more}/>`;
const workbook = (inner: string, main = MAIN, related = TRANSITIONAL) =>
  `<workbook xmlns="${main}" xmlns:r="${related}">${inner}</workbook>`;
const alternate = (requires: string, choice: string, fallback: string) =>
  `<mc:AlternateContent xmlns:mc="${COMPATIBILITY}"><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

type Parts = Readonly<Record<string, string | undefined>>;

const ROOT = {
  "_rels/.rels": relationships(
    relationship("rId1", `${TRANSITIONAL}/officeDocument`, "xl/workbook.xml"),
  ),
};

function book(changes: Parts = {}): Parts {
  return {
    ...ROOT,
    "xl/workbook.xml": workbook(`<sheets>${sheet("Orders", "rId1")}</sheets>`),
    "xl/_rels/workbook.xml.rels": relationships(
      relationship("rId1", WORKSHEET, "worksheets/sheet1.xml"),
    ),
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`,
    ...changes,
  };
}

const shaped = (inner: string, rels = relationship("rId1", WORKSHEET, "worksheets/sheet1.xml")) =>
  book({
    "xl/workbook.xml": workbook(inner),
    "xl/_rels/workbook.xml.rels": relationships(rels),
  });

const listing = (sheets: string, rels?: string) => shaped(`<sheets>${sheets}</sheets>`, rels);

const targeting = (target: string, more = "") =>
  book({
    "xl/_rels/workbook.xml.rels": relationships(relationship("rId1", WORKSHEET, target, more)),
  });

interface Listed {
  readonly sheets: readonly (readonly [string, string | null])[];
  readonly opened: readonly string[];
}

async function read(
  parts: Parts | Uint8Array,
): Promise<{ readonly sheets: readonly ListedSheet[]; readonly opened: readonly string[] }> {
  const encoder = new TextEncoder();
  const bytes =
    parts instanceof Uint8Array
      ? parts
      : await createPackage(
          Object.entries(parts).flatMap(([path, text]) =>
            text === undefined ? [] : [{ path, data: encoder.encode(text) }],
          ),
          store,
        );
  const { book: meters } = resolveLimits(undefined);
  const pkg = await openPackage(bytes, meters.parts, meters.structure);
  const opened: string[] = [];
  const sheets = await readWorkbook((name) => {
    opened.push(name);
    return pkg.open(partName(name, "structure"), meters.structure);
  });
  return { sheets, opened };
}

async function list(parts: Parts | Uint8Array): Promise<Listed> {
  const { sheets, opened } = await read(parts);
  return { sheets: sheets.map(({ name, worksheet }) => [name, worksheet?.name ?? null]), opened };
}

async function outcome(parts: Parts): Promise<string> {
  try {
    await list(parts);
    return "listed";
  } catch (error) {
    if (error instanceof XlsxReadError) return error.code;
    throw error;
  }
}

describe("the sheet list", () => {
  it("names every sheet in the order of the workbook part, hidden ones too, each with the part of its cells", async () => {
    const parts = listing(
      sheet("Lists", "rId3", ' state="hidden"') +
        sheet("Orders &amp; Returns", "rId1") +
        sheet(" Secret ", "rId2", ' state="veryHidden"'),
      relationship("rId1", WORKSHEET, "worksheets/sheet1.xml") +
        relationship("rId2", WORKSHEET, "worksheets/sheet2.xml") +
        relationship("rId3", WORKSHEET, "worksheets/sheet3.xml"),
    );
    expect((await list(parts)).sheets).toEqual([
      ["Lists", "xl/worksheets/sheet3.xml"],
      ["Orders & Returns", "xl/worksheets/sheet1.xml"],
      [" Secret ", "xl/worksheets/sheet2.xml"],
    ]);
  });

  it("opens the package's relationships, the workbook part and its relationships, and nothing else", async () => {
    const parts = book({
      "xl/media/image1.png": "a picture",
      "xl/drawings/drawing1.xml": "<xdr:wsDr/>",
      "xl/sharedStrings.xml": "<sst/>",
      "xl/styles.xml": "<styleSheet/>",
    });
    expect((await list(parts)).opened).toEqual([
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
    ]);
  });

  it("lists a chart sheet, a dialog sheet and a macro sheet without a part of cells, in both families", async () => {
    const types = [
      `${TRANSITIONAL}/worksheet`,
      `${STRICT}/worksheet`,
      `${TRANSITIONAL}/chartsheet`,
      `${STRICT}/chartsheet`,
      `${TRANSITIONAL}/dialogsheet`,
      `${STRICT}/dialogsheet`,
      `${MACRO_SHEET}/xlMacrosheet`,
      `${MACRO_SHEET}/xlIntlMacrosheet`,
    ];
    const parts = listing(
      types.map((_, index) => sheet(`Sheet ${index}`, `rId${index}`)).join(""),
      types.map((type, index) => relationship(`rId${index}`, type, `sheets/${index}.xml`)).join(""),
    );
    expect((await list(parts)).sheets).toEqual([
      ["Sheet 0", "xl/sheets/0.xml"],
      ["Sheet 1", "xl/sheets/1.xml"],
      ["Sheet 2", null],
      ["Sheet 3", null],
      ["Sheet 4", null],
      ["Sheet 5", null],
      ["Sheet 6", null],
      ["Sheet 7", null],
    ]);
  });

  it("reads a workbook of the Strict family", async () => {
    const parts = {
      "_rels/.rels": relationships(
        relationship("rId1", `${STRICT}/officeDocument`, "xl/workbook.xml"),
      ),
      "xl/workbook.xml": workbook(
        `<sheets>${sheet("Strict", "rId1")}</sheets>`,
        STRICT_MAIN,
        STRICT,
      ),
      "xl/_rels/workbook.xml.rels": relationships(
        relationship("rId1", `${STRICT}/worksheet`, "worksheets/sheet1.xml"),
      ),
    };
    expect((await list(parts)).sheets).toEqual([["Strict", "xl/worksheets/sheet1.xml"]]);
  });

  it("refuses a sheet whose relationship is to anything but a kind of sheet", async () => {
    const types = [
      `${TRANSITIONAL}/styles`,
      `${TRANSITIONAL}/officeDocument`,
      `${TRANSITIONAL}/Worksheet`,
      `${TRANSITIONAL}/worksheet `,
      `${MACRO_SHEET}/xlmacrosheet`,
      "worksheet",
      "",
    ];
    for (const type of types) {
      const parts = listing(sheet("Orders", "rId1"), relationship("rId1", type, "sheet1.xml"));
      expect(await outcome(parts), type).toBe("INVALID_WORKBOOK");
    }
    const untyped = listing(sheet("Orders", "rId1"), '<Relationship Id="rId1" Target="a.xml"/>');
    expect(await outcome(untyped)).toBe("INVALID_WORKBOOK");
  });

  it("refuses a worksheet's relationship type with any one character changed, taken out or added", async () => {
    for (const type of [WORKSHEET, `${STRICT}/worksheet`]) {
      const changed = new Set<string>();
      for (let at = 0; at < type.length; at++) {
        const next = String.fromCharCode(type.charCodeAt(at) + 1);
        changed.add(type.slice(0, at) + next + type.slice(at + 1));
        changed.add(type.slice(0, at) + type.slice(at + 1));
        changed.add(type.slice(0, at) + "x" + type.slice(at));
      }
      changed.add(`${type}x`);
      for (const damaged of changed) {
        const parts = listing(
          sheet("Orders", "rId1"),
          relationship("rId1", damaged, "worksheets/sheet1.xml"),
        );
        expect(await outcome(parts), damaged).toBe("INVALID_WORKBOOK");
      }
    }
  });

  it("refuses two sheets with one name, character for character, and lists names that differ only by case", async () => {
    const twice = listing(sheet("Orders", "rId1") + sheet("Orders", "rId1"));
    expect(await outcome(twice)).toBe("INVALID_WORKBOOK");
    const cased = listing(sheet("Orders", "rId1") + sheet("orders", "rId1"));
    expect((await list(cased)).sheets.map(([name]) => name)).toEqual(["Orders", "orders"]);
  });

  it("lists the sheets of a workbook LibreOffice wrote with names that differ only by case or by normalisation", async () => {
    // libreoffice-names.xlsx is what LibreOffice 26.8 writes from libreoffice-names.fods with
    // `soffice --headless --convert-to xlsx`; the names are those of the .fods, written by hand.
    const names = ["\u212a", "k", "\u212b", "\u00e5", "\u0130", "i\u0307", "\u00e9", "e\u0301"];
    const file = readFileSync(new URL("fixtures/libreoffice-names.xlsx", import.meta.url));
    const { sheets } = await list(new Uint8Array(file));
    expect(sheets.map(([name]) => name)).toEqual(names);
    for (let at = 0; at < names.length; at += 2) {
      const [first = "", second = ""] = names.slice(at, at + 2);
      expect(first).not.toBe(second);
      expect(first.toLowerCase().normalize("NFC")).toBe(second.toLowerCase().normalize("NFC"));
    }
  });

  it("names the sheet a refusal is about, and none when it is about the workbook", async () => {
    const aboutOrders = [
      listing(sheet("Orders", "rId1") + sheet("Orders", "rId1")),
      listing(sheet("Orders", "rId2")),
      listing(sheet("Orders", "rId1"), relationship("rId1", `${TRANSITIONAL}/styles`, "s.xml")),
      targeting("../../sheet1.xml"),
    ];
    for (const parts of aboutOrders) {
      await expect(list(parts)).rejects.toMatchObject({
        code: "INVALID_WORKBOOK",
        sheet: "Orders",
      });
    }
    const workbookOutside = book({
      "_rels/.rels": relationships(
        relationship("rId1", `${TRANSITIONAL}/officeDocument`, "../xl/workbook.xml"),
      ),
    });
    await expect(list(workbookOutside)).rejects.toMatchObject({
      code: "INVALID_WORKBOOK",
      sheet: undefined,
    });
  });

  it("refuses a sheet without a name, with an empty one, or without a relationship", async () => {
    const sheets = [
      '<sheet sheetId="1" r:id="rId1"/>',
      '<sheet name="" sheetId="1" r:id="rId1"/>',
      '<sheet name="Orders" sheetId="1"/>',
      sheet("Orders", "rId2"),
    ];
    for (const listed of sheets) {
      expect(await outcome(listing(listed)), listed).toBe("INVALID_WORKBOOK");
    }
  });

  it("refuses a sheet list that is missing, empty or stated twice", async () => {
    const workbooks = [
      "",
      "<sheets/>",
      `<sheets>${sheet("Orders", "rId1")}</sheets><sheets>${sheet("Notes", "rId1")}</sheets>`,
      `<sheets/><sheets>${sheet("Orders", "rId1")}</sheets>`,
      `<book><sheets>${sheet("Orders", "rId1")}</sheets></book>`,
    ];
    for (const inner of workbooks) {
      expect(await outcome(shaped(inner)), inner).toBe("INVALID_WORKBOOK");
    }
  });

  it("takes a sheet only from directly inside the sheet list", async () => {
    const inner =
      `<bookViews>${sheet("Before", "rId1")}</bookViews>` +
      `<sheets>${sheet("Orders", "rId1")}</sheets>` +
      `<extLst>${sheet("After", "rId1")}</extLst>`;
    expect((await list(shaped(inner))).sheets).toEqual([["Orders", "xl/worksheets/sheet1.xml"]]);
  });

  it("refuses a sheet list that holds anything but sheets, rather than list fewer sheets", async () => {
    const [beta, gamma] = [sheet("Beta", "rId1"), sheet("Gamma", "rId1")];
    const held = [
      alternate("r", beta, gamma),
      `<mc:AlternateContent xmlns:mc="${COMPATIBILITY}"><mc:Choice xmlns:zz="urn:zz" Requires="zz">${beta}</mc:Choice><mc:Fallback>${gamma}</mc:Fallback></mc:AlternateContent>`,
      alternate("x14", beta, gamma),
      `<group>${beta}</group>`,
      `<sheetx name="Beta" sheetId="2" r:id="rId1"/>`,
      `<Sheet name="Beta" sheetId="2" r:id="rId1"/>`,
      `<sheet name="Beta" sheetId="2" r:id="rId1"><extLst/></sheet>`,
      `<sheet name="Beta" sheetId="2" r:id="rId1">${gamma}</sheet>`,
    ];
    for (const inner of held) {
      expect(await outcome(listing(sheet("Alpha", "rId1") + inner)), inner).toBe(
        "INVALID_WORKBOOK",
      );
    }
  });

  it("refuses a sheet list or a sheet that stands below an AlternateContent anywhere in the part", async () => {
    const [alpha, beta, gamma] = [
      sheet("Alpha", "rId1"),
      sheet("Beta", "rId1"),
      sheet("Gamma", "rId1"),
    ];
    const workbooks = [
      `<sheets>${alpha}</sheets>${alternate("r", `<sheets>${beta}</sheets>`, `<sheets>${gamma}</sheets>`)}`,
      `${alternate("r", `<sheets>${beta}</sheets>`, `<sheets>${gamma}</sheets>`)}<sheets>${alpha}</sheets>`,
      `<sheets>${alpha}</sheets>${alternate("r", beta, gamma)}`,
      `<sheets>${alpha}</sheets><extLst><ext uri="x">${alternate("r", "", beta)}</ext></extLst>`,
      `<sheets>${alpha}</sheets>${alternate("r", alternate("r", "", ""), `<sheets>${gamma}</sheets>`)}`,
    ];
    for (const inner of workbooks) {
      expect(await outcome(shaped(inner)), inner).toBe("INVALID_WORKBOOK");
    }
  });

  it("gives sheets that share one relationship one part, resolved once", async () => {
    const sheets = sheet("Alpha", "rId1") + sheet("Beta", "rId1") + sheet("Gamma", "rId2");
    const rels =
      relationship("rId1", WORKSHEET, "worksheets/sheet1.xml") +
      relationship("rId2", WORKSHEET, "worksheets/sheet1.xml");
    const [alpha, beta, gamma] = (await read(listing(sheets, rels))).sheets;
    expect(alpha!.worksheet).toBeDefined();
    expect(beta!.worksheet).toBe(alpha!.worksheet);
    expect(gamma!.worksheet).not.toBe(alpha!.worksheet);
  });

  it("lists the sheets of a workbook whose AlternateContent holds no sheet, as Excel writes one", async () => {
    const path = `<x15ac:absPath xmlns:x15ac="urn:x15ac" url="/books/"/>`;
    const inner = `<workbookPr/>${alternate("x15", path, "")}<sheets>${sheet("Orders", "rId1")}</sheets>`;
    expect((await list(shaped(inner))).sheets).toEqual([["Orders", "xl/worksheets/sheet1.xml"]]);
  });

  it("refuses a package without its own relationships, or without a workbook, as no workbook", async () => {
    const packages: Parts[] = [
      book({ "_rels/.rels": undefined }),
      book({
        "_rels/.rels": relationships(relationship("rId1", `${TRANSITIONAL}/thumbnail`, "a.png")),
      }),
      book({
        "_rels/.rels": relationships(
          relationship("rId1", "http://example.invalid/officeDocument", "xl/workbook.xml"),
        ),
      }),
      book({
        "xl/workbook.xml": `<w:document xmlns:w="urn:word"><w:body/></w:document>`,
      }),
      book({
        "xl/workbook.xml": `<Workbook xmlns="${MAIN}" xmlns:r="${TRANSITIONAL}"><sheets>${sheet("Orders", "rId1")}</sheets></Workbook>`,
      }),
    ];
    for (const parts of packages) expect(await outcome(parts)).toBe("NOT_XLSX");
  });

  it("refuses a package with two workbooks, or a workbook part or its relationships missing", async () => {
    const workbookTarget = (target: string) =>
      relationship("rId2", `${STRICT}/officeDocument`, target);
    const packages: Parts[] = [
      book({
        "_rels/.rels": relationships(
          relationship("rId1", `${TRANSITIONAL}/officeDocument`, "xl/workbook.xml"),
          workbookTarget("xl/workbook.xml"),
        ),
      }),
      book({ "xl/workbook.xml": undefined }),
      book({ "xl/_rels/workbook.xml.rels": undefined }),
      book({ "_rels/.rels": relationships(workbookTarget("../xl/workbook.xml")) }),
    ];
    for (const parts of packages) expect(await outcome(parts)).toBe("INVALID_WORKBOOK");
  });

  it("finds the workbook through the package's relationships, by any name and in any case", async () => {
    const parts = {
      "_rels/.rels": relationships(
        relationship("rId1", `${TRANSITIONAL}/officeDocument`, "/Book/Main.XML"),
      ),
      "book/main.xml": workbook(`<sheets>${sheet("Orders", "rId1")}</sheets>`),
      "book/_rels/main.xml.rels": relationships(relationship("rId1", WORKSHEET, "cells.xml")),
    };
    expect(await list(parts)).toEqual({
      sheets: [["Orders", "Book/cells.xml"]],
      opened: ["_rels/.rels", "Book/Main.XML", "Book/_rels/Main.XML.rels"],
    });
  });
});

describe("a relationships part", () => {
  const worksheet = relationship("rId1", WORKSHEET, "worksheets/sheet1.xml");

  it("is refused when its root is anything but Relationships", async () => {
    for (const root of ["Relations", "relationships"]) {
      const parts = book({
        "xl/_rels/workbook.xml.rels": `<${root} xmlns="${PACKAGE}">${worksheet}</${root}>`,
      });
      expect(await outcome(parts), root).toBe("INVALID_WORKBOOK");
    }
  });

  it("is refused when it holds anything but relationships directly inside its root", async () => {
    const parts = [
      `<group>${worksheet}</group>`,
      '<Relation Id="rId1" Type="' + WORKSHEET + '" Target="worksheets/sheet1.xml"/>',
      `${worksheet}<extra/>`,
      worksheet.replace("/>", "><extra/></Relationship>"),
      worksheet.replace(
        "/>",
        `>${relationship("rId2", WORKSHEET, "worksheets/sheet2.xml")}</Relationship>`,
      ),
      worksheet + alternate("r", relationship("rId2", WORKSHEET, "worksheets/sheet2.xml"), ""),
    ];
    for (const inner of parts) {
      const rels = `<Relationships xmlns="${PACKAGE}">${inner}</Relationships>`;
      expect(await outcome(book({ "xl/_rels/workbook.xml.rels": rels })), inner).toBe(
        "INVALID_WORKBOOK",
      );
    }
    const root = relationships(
      relationship("rId1", `${TRANSITIONAL}/officeDocument`, "xl/workbook.xml"),
      "<extra/>",
    );
    expect(await outcome(book({ "_rels/.rels": root }))).toBe("INVALID_WORKBOOK");
  });

  it("is refused when a relationship has no id, or the id of another", async () => {
    const withoutId = book({
      "_rels/.rels": relationships(
        `<Relationship Type="${TRANSITIONAL}/officeDocument" Target="xl/workbook.xml"/>`,
      ),
    });
    expect(await outcome(withoutId)).toBe("INVALID_WORKBOOK");
    const twice = book({
      "xl/_rels/workbook.xml.rels": relationships(
        worksheet,
        relationship("rId1", WORKSHEET, "worksheets/sheet2.xml"),
      ),
    });
    expect(await outcome(twice)).toBe("INVALID_WORKBOOK");
  });
});

describe("a relationship's target", () => {
  it("is resolved from the folder of the part that states it, or from the root after a slash", async () => {
    const targets = [
      "worksheets/sheet1.xml",
      "/xl/worksheets/sheet1.xml",
      "./worksheets/sheet1.xml",
      "../xl/worksheets/sheet1.xml",
      "worksheets/../worksheets/./sheet1.xml",
    ];
    for (const target of targets) {
      expect((await list(targeting(target))).sheets, target).toEqual([
        ["Orders", "xl/worksheets/sheet1.xml"],
      ]);
    }
  });

  it("keeps its case and its other characters as written", async () => {
    const targets = ["Worksheets/SHEET1.xml", "sheets/Ærlig [1] @+=.xml", "worksheets/.rels"];
    for (const target of targets) {
      expect((await list(targeting(target))).sheets, target).toEqual([["Orders", `xl/${target}`]]);
    }
  });

  it("is internal when its mode says so or says nothing, and is refused in any other mode", async () => {
    expect(
      (await list(targeting("worksheets/sheet1.xml", ' TargetMode="Internal"'))).sheets,
    ).toEqual([["Orders", "xl/worksheets/sheet1.xml"]]);
    for (const mode of ["External", "external", "internal", "INTERNAL", ""]) {
      const parts = targeting("worksheets/sheet1.xml", ` TargetMode="${mode}"`);
      expect(await outcome(parts), mode).toBe("INVALID_WORKBOOK");
    }
  });

  it("is refused when it is missing, climbs out of the package, names a folder or has an empty step", async () => {
    const targets = [
      "",
      "../../sheet1.xml",
      "/../xl/worksheets/sheet1.xml",
      "worksheets/../../../sheet1.xml",
      "worksheets/",
      "worksheets/.",
      "worksheets/..",
      ".",
      "worksheets//sheet1.xml",
      "//host/sheet1.xml",
      "/",
    ];
    for (const target of targets)
      expect(await outcome(targeting(target)), target).toBe("INVALID_WORKBOOK");
    const untargeted = book({
      "xl/_rels/workbook.xml.rels": relationships(`<Relationship Id="rId1" Type="${WORKSHEET}"/>`),
    });
    expect(await outcome(untargeted)).toBe("INVALID_WORKBOOK");
  });

  it("is refused when another reader could take it another way: a scheme, an escape, a query, a fragment, a backslash", async () => {
    const targets = [
      "http://example.invalid/sheet1.xml",
      "file:///etc/passwd",
      "urn:sheet",
      "worksheets/sheet%31.xml",
      "worksheets/sheet1.xml?x",
      "worksheets/sheet1.xml#x",
      "worksheets\\sheet1.xml",
    ];
    for (const target of targets)
      expect(await outcome(targeting(target)), target).toBe("INVALID_WORKBOOK");
  });

  it("is followed for a sheet without cells too", async () => {
    const parts = listing(sheet("Chart", "rId1"), relationship("rId1", CHART_SHEET, "../../c.xml"));
    expect(await outcome(parts)).toBe("INVALID_WORKBOOK");
  });
});
