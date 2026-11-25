import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { type ReadValue, openXlsx } from "../src/read.js";
import {
  type ReadBack,
  type ReadBackCell,
  type ReadBackValue,
  reference,
} from "./corpus/read-back.js";

const corpus = join(dirname(fileURLToPath(import.meta.url)), "corpus");

const workbooks = readdirSync(corpus, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .flatMap(({ name }) =>
    readdirSync(join(corpus, name))
      .filter((file) => file.endsWith(".xlsx"))
      .map((file) => `${name}/${file}`),
  )
  .sort();

interface Sheet<Cell> {
  readonly name: string;
  readonly rows: readonly (readonly [number: number, ...cells: Cell[]])[];
}

interface Values<Cell> {
  readonly sheets: readonly Sheet<Cell>[];
}

async function readerValues(workbook: string): Promise<Values<ReadValue>> {
  const book = await openXlsx(new Uint8Array(readFileSync(join(corpus, workbook))));
  const sheets: Sheet<ReadValue>[] = [];
  for (const name of book.sheets) {
    const rows: [number, ...ReadValue[]][] = [];
    for await (const row of book.rows(name)) rows.push([row.number, ...row.cells]);
    sheets.push({ name, rows });
  }
  return { sheets };
}

const readBackOf = (workbook: string): ReadBack =>
  JSON.parse(readFileSync(join(corpus, workbook.replace(/\.xlsx$/, ".json")), "utf8")) as ReadBack;

const isFormula = (cell: ReadBackCell) =>
  cell !== null && typeof cell === "object" && "result" in cell;

const stored = (cell: ReadBackCell): ReadBackValue | null =>
  cell !== null && typeof cell === "object" && "result" in cell ? cell.result : cell;

function typed<Kind extends "error" | "date" | "duration">(cell: ReadBackCell, kind: Kind) {
  const shown = stored(cell);
  if (shown === null || typeof shown !== "object" || !(kind in shown)) return undefined;
  return (shown as Record<Kind, string>)[kind];
}

const DAY = 86_400_000;
const iso = (date: Date) => date.toISOString().slice(0, 23);
const instant = (date: string) => Date.parse(`${date}Z`);

function milliseconds(duration: string): number {
  const [, sign, hours, minutes, seconds] = /^(-?)PT(\d+)H(\d+)M(\d+\.\d+)S$/.exec(duration)!;
  const total = Math.round((Number(hours) * 3600 + Number(minutes) * 60 + Number(seconds)) * 1000);
  return sign === "-" ? -total : total;
}

/** Text with each `_xHHHH_` replaced by the character it names, left to right (ECMA-376). */
const decodeEscapes = (text: string) =>
  text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );

function same(value: ReadValue, cell: ReadBackCell): boolean {
  const shown = stored(cell);
  if (shown === null || typeof shown !== "object") return value === shown;
  const date = typed(cell, "date");
  return date !== undefined && value instanceof Date && iso(value) === date;
}

interface Allowance {
  readonly name: string;
  readonly why: string;
  /** Every cell the rule excuses, in corpus order, with the value the reader gives there. */
  readonly cells: readonly string[];
  readonly excuses: (value: ReadValue, cell: ReadBackCell) => boolean;
}

const ALLOWANCES: readonly Allowance[] = [
  {
    name: "a formula LibreOffice cannot read",
    why: "LibreOffice shows #NAME? for a function written without the prefix Excel writes. Numbers wrote _xlfn.FILTER for _xlfn._xlws.FILTER, and Google Sheets dropped _xlfn. from IFS, TEXTJOIN, CONCAT and _xlws.FILTER. The reader gives the stored text or number, which LibreOffice cannot confirm.",
    cells: [
      'google-sheets/future-functions.xlsx, sheet "Functions", B3: "mid"',
      'google-sheets/future-functions.xlsx, sheet "Functions", B4: "a-b-c"',
      'google-sheets/future-functions.xlsx, sheet "Functions", B5: "#N/A"',
      'google-sheets/future-functions.xlsx, sheet "Functions", B8: "#NAME?"',
      'numbers/future-functions.xlsx, sheet "Functions", B8: 26',
    ],
    excuses: (value, cell) =>
      isFormula(cell) &&
      typed(cell, "error") === "#NAME?" &&
      (typeof value === "string" || typeof value === "number"),
  },
  {
    name: "an error as its text",
    why: "The reader gives an error cell as its text (ADR 0032). LibreOffice keeps it typed as an error.",
    cells: [
      'excel/type-matrix.xlsx, sheet "Formulas", B5: "#DIV/0!"',
      'excel/type-matrix.xlsx, sheet "Formulas", B6: "#N/A"',
      'libreoffice/type-matrix.xlsx, sheet "Formulas", B5: "#DIV/0!"',
      'libreoffice/type-matrix.xlsx, sheet "Formulas", B6: "#N/A"',
    ],
    excuses: (value, cell) => typed(cell, "error") === value,
  },
  {
    name: "a time as a number",
    why: "The reader gives a time of day or an elapsed time as a number of days (ADR 0032). LibreOffice types it as a duration. Both name the same millisecond.",
    cells: [
      'excel/type-matrix.xlsx, sheet "Numbers", B17: 1.5',
      'excel/type-matrix.xlsx, sheet "Numbers", B18: 0.75',
      'excel/type-matrix.xlsx, sheet "Numbers", B19: 0.0123',
      'libreoffice/type-matrix.xlsx, sheet "Numbers", B17: 1.5',
      'libreoffice/type-matrix.xlsx, sheet "Numbers", B18: 0.75',
      'libreoffice/type-matrix.xlsx, sheet "Numbers", B19: 0.0123',
      'numbers/type-matrix.xlsx, sheet "Numbers", B17: 1.5',
      'numbers/type-matrix.xlsx, sheet "Numbers", B18: 1.75',
      'numbers/type-matrix.xlsx, sheet "Numbers", B19: 1.012299999997572',
    ],
    excuses: (value, cell) => {
      const duration = typed(cell, "duration");
      return (
        duration !== undefined &&
        typeof value === "number" &&
        Math.round(value * DAY) === milliseconds(duration)
      );
    },
  },
  {
    name: "a day early before March 1900",
    why: "LibreOffice has no 29 February 1900, so it shows a serial below 60 as the day before the one Excel shows. The reader gives Excel's day.",
    cells: [
      'excel/type-matrix.xlsx, sheet "Dates", B10: the date 1900-01-01T00:00:00.000',
      'excel/type-matrix.xlsx, sheet "Dates", B11: the date 1900-02-28T00:00:00.000',
      'libreoffice/type-matrix.xlsx, sheet "Dates", B10: the date 1900-01-01T00:00:00.000',
      'libreoffice/type-matrix.xlsx, sheet "Dates", B11: the date 1900-02-28T00:00:00.000',
      'numbers/type-matrix.xlsx, sheet "Dates", B10: the date 1900-01-02T00:00:00.000',
    ],
    excuses: (value, cell) => {
      const date = typed(cell, "date");
      return (
        date !== undefined &&
        date < "1900-03-01" &&
        value instanceof Date &&
        value.getTime() - instant(date) === DAY
      );
    },
  },
  {
    name: "29 February 1900 as a number",
    why: "Serial 60 names 29 February 1900, a day that never was. The reader gives the serial (ADR 0032). LibreOffice shows 28 February.",
    cells: ['numbers/type-matrix.xlsx, sheet "Dates", B11: 60'],
    excuses: (value, cell) => {
      const date = typed(cell, "date");
      return (
        date !== undefined &&
        typeof value === "number" &&
        value >= 60 &&
        value < 61 &&
        instant(date) === Date.UTC(1899, 11, 30) + Math.round(value * DAY)
      );
    },
  },
  {
    name: "fifteen significant digits",
    why: "LibreOffice shows 15 significant digits of a number, and the read-back keeps those. The reader gives the double the file states.",
    cells: ['excel/type-matrix.xlsx, sheet "Numbers", B6: 0.30000000000000004'],
    excuses: (value, cell) =>
      typeof value === "number" && Number(value.toPrecision(15)) === stored(cell),
  },
  {
    name: "a carriage return beside a line feed",
    why: "In a text that holds a line feed, LibreOffice reads each carriage return as a line break, and a carriage return before a line feed as one break with it. The reader gives the text as stored.",
    cells: [
      'excel/escapes.xlsx, sheet "Escapes", B13: "a\\r\\nb"',
      'numbers/escapes.xlsx, sheet "Escapes", B13: "a\\r\\nb"',
    ],
    excuses: (value, cell) =>
      typeof value === "string" &&
      value.includes("\n") &&
      value.replaceAll("\r\n", "\n").replaceAll("\r", "\n") === stored(cell),
  },
  {
    name: "an escape in a formula's stored text",
    why: "LibreOffice shows a formula's stored text with its _xHHHH_ escapes as written. The reader decodes them, and Excel shows the decoded text.",
    cells: [
      'excel/escapes.xlsx, sheet "Escapes", B15: "_x0041_"',
      'libreoffice/escapes.xlsx, sheet "Escapes", B15: "_x0041_"',
    ],
    excuses: (value, cell) => {
      const text = stored(cell);
      return isFormula(cell) && typeof text === "string" && value === decodeEscapes(text);
    },
  },
  {
    name: "a stored boolean of an array formula as a number",
    why: "Numbers writes each formula as an array formula, and LibreOffice shows the boolean such a formula stores as 1 or 0.",
    cells: [
      'numbers/formulas.xlsx, sheet "Budget", B8: true',
      'numbers/type-matrix.xlsx, sheet "Formulas", B4: true',
    ],
    excuses: (value, cell) =>
      isFormula(cell) && typeof value === "boolean" && Number(value) === stored(cell),
  },
  {
    name: "a formula without a stored result",
    why: "Numbers wrote a formula whose result is an error without that result. LibreOffice calculates it, and the reader gives null for a formula without a stored result (ADR 0032).",
    cells: ['numbers/type-matrix.xlsx, sheet "Formulas", B5: null'],
    excuses: (value, cell) =>
      value === null && isFormula(cell) && typed(cell, "error") !== undefined,
  },
];

const show = (value: ReadValue) =>
  value instanceof Date ? `the date ${iso(value)}` : JSON.stringify(value);

interface Comparison {
  readonly mismatches: string[];
  readonly excused: { readonly allowance: string; readonly cell: string }[];
}

function compare(
  workbook: string,
  read: Values<ReadValue>,
  readBack: ReadBack,
  allowances: readonly Allowance[],
): Comparison {
  const outcome: Comparison = { mismatches: [], excused: [] };
  const names = (sheets: readonly { readonly name: string }[]) =>
    JSON.stringify(sheets.map(({ name }) => name));
  if (names(read.sheets) !== names(readBack.sheets)) {
    outcome.mismatches.push(
      `${workbook}: the reader lists the sheets ${names(read.sheets)}, LibreOffice ${names(readBack.sheets)}`,
    );
    return outcome;
  }
  read.sheets.forEach((sheet, index) => {
    let last = 0;
    for (const [number] of sheet.rows) {
      if (number <= last) {
        outcome.mismatches.push(
          `${workbook}, sheet ${JSON.stringify(sheet.name)}: the reader gives row ${number} after row ${last}`,
        );
      }
      last = number;
    }
    const ours = new Map(sheet.rows.map(([number, ...cells]) => [number, cells]));
    const theirs = new Map(
      readBack.sheets[index]!.rows.map(([number, ...cells]) => [number, cells]),
    );
    for (const number of [...new Set([...ours.keys(), ...theirs.keys()])].sort((a, b) => a - b)) {
      const values = ours.get(number) ?? [];
      const cells = theirs.get(number) ?? [];
      for (let column = 0; column < Math.max(values.length, cells.length); column++) {
        const value = values[column] ?? null;
        const cell = cells[column] ?? null;
        if (same(value, cell)) continue;
        const place = `${workbook}, sheet ${JSON.stringify(sheet.name)}, ${reference(number, column)}`;
        const allowance = allowances.find(({ excuses }) => excuses(value, cell));
        if (allowance !== undefined) {
          outcome.excused.push({ allowance: allowance.name, cell: `${place}: ${show(value)}` });
          continue;
        }
        outcome.mismatches.push(
          `${place}: the reader gives ${show(value)}, LibreOffice ${JSON.stringify(cell)}`,
        );
      }
    }
  });
  return outcome;
}

function withFirstCellChanged<Cell>(
  sheetsOf: (workbook: string) => readonly Sheet<Cell>[],
  change: (cell: Cell) => Cell | undefined,
): { workbook: string; place: string; sheets: Sheet<Cell>[] } {
  for (const workbook of workbooks) {
    const sheets = sheetsOf(workbook);
    for (const [index, { name, rows }] of sheets.entries()) {
      for (const [row, [number, ...cells]] of rows.entries()) {
        const column = cells.findIndex((cell) => change(cell) !== undefined);
        if (column < 0) continue;
        const changed: [number, ...Cell[]] = [
          number,
          ...cells.map((cell, at) => (at === column ? change(cell)! : cell)),
        ];
        return {
          workbook,
          place: `${workbook}, sheet ${JSON.stringify(name)}, ${reference(number, column)}: `,
          sheets: sheets.map((sheet, at) =>
            at === index
              ? { name, rows: rows.map((old, at) => (at === row ? changed : old)) }
              : sheet,
          ),
        };
      }
    }
  }
  throw new Error("no such cell in the corpus");
}

describe("the corpus reads as LibreOffice reads it", () => {
  const reads = new Map<string, Values<ReadValue>>();
  const readBacks = new Map<string, ReadBack>();

  beforeAll(async () => {
    for (const workbook of workbooks) {
      reads.set(workbook, await readerValues(workbook));
      readBacks.set(workbook, readBackOf(workbook));
    }
  });

  const compareAll = (allowances: readonly Allowance[]) =>
    workbooks.map((workbook) =>
      compare(workbook, reads.get(workbook)!, readBacks.get(workbook)!, allowances),
    );

  it("holds a read-back beside every workbook and a workbook beside every read-back", () => {
    const readBackFiles = readdirSync(corpus, { recursive: true, encoding: "utf8" })
      .map((name) => name.replaceAll("\\", "/"))
      .filter((name) => name.endsWith(".json"))
      .map((name) => name.replace(/\.json$/, ".xlsx"))
      .sort();
    expect(readBackFiles).toEqual(workbooks);
    expect(workbooks.length).toBeGreaterThan(80);
  });

  it.each(workbooks)("%s matches cell for cell", (workbook) => {
    expect(
      compare(workbook, reads.get(workbook)!, readBacks.get(workbook)!, ALLOWANCES).mismatches,
    ).toEqual([]);
  });

  it.each(ALLOWANCES)(
    "excuses as $name its listed cells and no other, each with the value listed, and the corpus fails without it",
    (allowance) => {
      const excused = compareAll(ALLOWANCES)
        .flatMap(({ excused }) => excused)
        .filter((one) => one.allowance === allowance.name)
        .map(({ cell }) => cell);
      expect(excused).toEqual(allowance.cells);
      const without = compareAll(ALLOWANCES.filter((other) => other !== allowance));
      expect(without.flatMap(({ mismatches }) => mismatches).length).toBeGreaterThan(0);
    },
  );

  const NOT_EXCUSED: readonly [string, ReadValue, ReadBackCell][] = [
    ["a formula LibreOffice cannot read", new Date(0), { result: { error: "#NAME?" } }],
    ["a formula LibreOffice cannot read", true, { result: { error: "#NAME?" } }],
    ["a formula LibreOffice cannot read", 26, { error: "#NAME?" }],
    ["an error as its text", "#N/A", { error: "#DIV/0!" }],
    ["an error as its text", 0, { result: { error: "#DIV/0!" } }],
    ["a time as a number", 0.75 + 2 / DAY, { duration: "PT18H00M00.000S" }],
    ["a time as a number", "0.75", { duration: "PT18H00M00.000S" }],
    [
      "a day early before March 1900",
      new Date(Date.UTC(1900, 0, 2)),
      { date: "1899-12-31T00:00:00.000" },
    ],
    [
      "a day early before March 1900",
      new Date(Date.UTC(1900, 2, 2)),
      { date: "1900-03-01T00:00:00.000" },
    ],
    ["29 February 1900 as a number", 61, { date: "1900-03-01T00:00:00.000" }],
    ["29 February 1900 as a number", 60.5, { date: "1900-02-28T00:00:00.000" }],
    ["29 February 1900 as a number", 59, { date: "1900-02-27T00:00:00.000" }],
    ["fifteen significant digits", 0.3000000000001, 0.3],
    ["a carriage return beside a line feed", "a\rb", "a\nb"],
    ["a carriage return beside a line feed", "a\r\nc", "a\nb"],
    ["an escape in a formula's stored text", "_x0041_", "_x005F_x0041_"],
    ["an escape in a formula's stored text", "A", { result: "_x005F_x0041_" }],
    ["a stored boolean of an array formula as a number", true, { result: 0 }],
    ["a stored boolean of an array formula as a number", true, 1],
    ["a stored boolean of an array formula as a number", "1", { result: 1 }],
    ["a formula without a stored result", null, { result: 5 }],
    ["a formula without a stored result", null, { error: "#DIV/0!" }],
  ];

  it.each(NOT_EXCUSED)("does not excuse by %s the reader's %o for %o", (_, value, cell) => {
    const one = (row: readonly [number, ...ReadValue[]]): Values<ReadValue> => ({
      sheets: [{ name: "S", rows: [row] }],
    });
    const outcome = compare(
      "w.xlsx",
      one([1, value]),
      { sheets: [{ name: "S", rows: [[1, cell]] }] },
      ALLOWANCES,
    );
    expect(outcome.mismatches).toHaveLength(1);
    expect(outcome.mismatches[0]).toMatch(/^w\.xlsx, sheet "S", A1: /);
  });

  it("fails on another value in an excused cell, and names the cell", () => {
    const workbook = "google-sheets/future-functions.xlsx";
    const { sheets } = reads.get(workbook)!;
    const wrong = {
      sheets: sheets.map((sheet) => ({
        name: sheet.name,
        rows: sheet.rows.map(([number, first, ...rest]): [number, ...ReadValue[]] =>
          number === 3 ? [number, first!, "WRONG", ...rest.slice(1)] : [number, first!, ...rest],
        ),
      })),
    };
    const { mismatches, excused } = compare(workbook, wrong, readBacks.get(workbook)!, ALLOWANCES);
    expect(mismatches).toEqual([]);
    const cells = excused.map(({ cell }) => cell);
    expect(cells).toContain(`${workbook}, sheet "Functions", B3: "WRONG"`);
    expect(ALLOWANCES[0]!.cells).not.toContain(`${workbook}, sheet "Functions", B3: "WRONG"`);
  });

  it("fails on rows the reader gives out of order or twice", () => {
    const workbook = workbooks[0]!;
    const [first, ...others] = reads.get(workbook)!.sheets;
    const [one, two, ...rows] = first!.rows;
    const place = `${workbook}, sheet ${JSON.stringify(first!.name)}: the reader gives row `;
    for (const [shuffled, line] of [
      [[two!, one!, ...rows], `${place}${one![0]} after row ${two![0]}`],
      [[one!, one!, two!, ...rows], `${place}${one![0]} after row ${one![0]}`],
    ] as const) {
      const { mismatches } = compare(
        workbook,
        { sheets: [{ name: first!.name, rows: shuffled }, ...others] },
        readBacks.get(workbook)!,
        ALLOWANCES,
      );
      expect(mismatches).toEqual([line]);
    }
  });

  it("covers each allowance with a value it does not excuse", () => {
    expect(new Set(NOT_EXCUSED.map(([name]) => name))).toEqual(
      new Set(ALLOWANCES.map(({ name }) => name)),
    );
  });

  const READ_BACK_CHANGES: readonly [string, (cell: ReadBackCell) => ReadBackCell | undefined][] = [
    ["text", (cell) => (typeof cell === "string" ? `${cell}?` : undefined)],
    ["number", (cell) => (typeof cell === "number" ? cell + 1 : undefined)],
    [
      "boolean",
      (cell) => {
        const value = stored(cell);
        if (typeof value !== "boolean") return undefined;
        return isFormula(cell) ? { result: !value } : !value;
      },
    ],
    [
      "date",
      (cell) => {
        const date = typed(cell, "date");
        return date === undefined ? undefined : { date: iso(new Date(instant(date) + 7 * DAY)) };
      },
    ],
    [
      "duration",
      (cell) =>
        typed(cell, "duration") === undefined ? undefined : { duration: "PT99H00M00.000S" },
    ],
    [
      "error",
      (cell) => {
        const error = typed(cell, "error");
        if (error === undefined) return undefined;
        const other = { error: error === "#REF!" ? "#NULL!" : "#REF!" };
        return isFormula(cell) ? { result: other } : other;
      },
    ],
    [
      "formula",
      (cell) => {
        const value = stored(cell);
        return isFormula(cell) && typeof value === "number" ? { result: value + 1 } : undefined;
      },
    ],
    ["empty cell", (cell) => (cell === null ? "?" : undefined)],
  ];

  it.each(READ_BACK_CHANGES)("names the place of a changed %s in a read-back", (_, change) => {
    const { workbook, place, sheets } = withFirstCellChanged(
      (name) => readBacks.get(name)!.sheets,
      change,
    );
    const { mismatches } = compare(workbook, reads.get(workbook)!, { sheets }, ALLOWANCES);
    expect(mismatches).toHaveLength(1);
    expect(mismatches[0]!.startsWith(place)).toBe(true);
  });

  const READER_CHANGES: readonly [string, (value: ReadValue) => ReadValue | undefined][] = [
    ["text", (value) => (typeof value === "string" ? `${value}?` : undefined)],
    ["number", (value) => (typeof value === "number" ? value + 1 : undefined)],
    ["boolean", (value) => (typeof value === "boolean" ? !value : undefined)],
    ["date", (value) => (value instanceof Date ? new Date(value.getTime() + 7 * DAY) : undefined)],
    ["empty cell", (value) => (value === null ? "?" : undefined)],
  ];

  it.each(READER_CHANGES)("names the place of a changed %s from the reader", (_, change) => {
    const { workbook, place, sheets } = withFirstCellChanged(
      (name) => reads.get(name)!.sheets,
      change,
    );
    const { mismatches } = compare(workbook, { sheets }, readBacks.get(workbook)!, ALLOWANCES);
    expect(mismatches).toHaveLength(1);
    expect(mismatches[0]!.startsWith(place)).toBe(true);
  });

  it("names the workbook of a sheet list that differs and each cell of a row left out", () => {
    const workbook = workbooks[0]!;
    const readBack = readBacks.get(workbook)!;
    const [first, ...others] = reads.get(workbook)!.sheets;
    const [left, ...rows] = first!.rows;
    const renamed = {
      sheets: [{ ...readBack.sheets[0]!, name: "Renamed" }, ...readBack.sheets.slice(1)],
    };
    expect(compare(workbook, reads.get(workbook)!, renamed, ALLOWANCES).mismatches).toEqual([
      expect.stringMatching(
        new RegExp(`^${workbook}: the reader lists the sheets .*, LibreOffice \\["Renamed"`),
      ),
    ]);
    const shorter = compare(
      workbook,
      { sheets: [{ name: first!.name, rows }, ...others] },
      readBack,
      ALLOWANCES,
    );
    const [number, ...cells] = left!;
    expect(shorter.mismatches).toHaveLength(cells.filter((cell) => cell !== null).length);
    expect(shorter.mismatches[0]).toMatch(
      `${workbook}, sheet ${JSON.stringify(first!.name)}, A${number}: the reader gives null`,
    );
  });
});
