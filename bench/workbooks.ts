import type { Cell, CellObject, StreamedWorkbook, Workbook } from "../src/index.ts";

/**
 * The three benchmark suites. Every suite is built from a seeded generator, so the same workbook
 * comes out on every machine and every run; building happens before timing starts.
 */

/** mulberry32: a small seeded generator, enough to make repeatable data. */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const FIRST_NAMES = [
  "Kari",
  "Ola",
  "Ingrid",
  "Lars",
  "Astrid",
  "Magnus",
  "Solveig",
  "Henrik",
  "Liv",
  "Erik",
  "Sigrid",
  "Anders",
  "Maren",
  "Jonas",
  "Thea",
  "Mats",
  "Nora",
  "Sander",
  "Ida",
  "Tobias",
];
const LAST_NAMES = [
  "Nordmann",
  "Hansen",
  "Johansen",
  "Olsen",
  "Larsen",
  "Andersen",
  "Pedersen",
  "Nilsen",
  "Kristiansen",
  "Jensen",
  "Karlsen",
  "Johnsen",
  "Pettersen",
  "Eriksen",
  "Berg",
  "Haugen",
  "Hagen",
  "Johannessen",
  "Andreassen",
  "Jacobsen",
  "Dahl",
  "Jørgensen",
  "Halvorsen",
  "Lund",
  "Strand",
  "Solberg",
  "Moen",
  "Bakke",
  "Lie",
  "Amundsen",
  "Knutsen",
  "Kristoffersen",
  "Sørensen",
  "Henriksen",
  "Martinsen",
  "Iversen",
  "Rasmussen",
  "Bøe",
  "Myhre",
  "Gundersen",
  "Holm",
  "Eide",
  "Sandvik",
  "Aas",
  "Ødegård",
  "Thorsen",
  "Birkeland",
  "Fosse",
  "Vik",
  "Ruud",
];
const CITIES = [
  "Oslo",
  "Bergen",
  "Trondheim",
  "Stavanger",
  "Drammen",
  "Fredrikstad",
  "Kristiansand",
  "Tromsø",
  "Sandnes",
  "Sarpsborg",
  "Bodø",
  "Ålesund",
  "Sandefjord",
  "Haugesund",
  "Tønsberg",
  "Moss",
  "Porsgrunn",
  "Arendal",
  "Hamar",
  "Larvik",
  "Halden",
  "Lillehammer",
  "Molde",
  "Harstad",
  "Kongsberg",
  "Gjøvik",
  "Askøy",
  "Horten",
  "Skien",
  "Narvik",
  "Jessheim",
  "Elverum",
  "Steinkjer",
  "Alta",
  "Hønefoss",
  "Kristiansund",
  "Mo i Rana",
  "Leirvik",
  "Grimstad",
  "Drøbak",
  "Florø",
  "Egersund",
  "Levanger",
  "Namsos",
  "Verdal",
  "Kirkenes",
  "Notodden",
  "Orkanger",
  "Førde",
  "Voss",
];
const STATUSES = ["new", "paid", "shipped", "delivered", "returned"];
const WORDS = [
  "order",
  "invoice",
  "customer",
  "delivery",
  "&",
  "<urgent>",
  "warehouse",
  "pallet",
  "return",
  "credit",
  "note",
  "approved",
  "pending",
  "partial",
  "express",
  "standard",
  "weekend",
  "fragile",
  "signature",
  "left at door",
  "call first",
  "gate code",
  "second floor",
  "no lift",
  "«ekte»",
  "🎉",
];

const pick = <T>(next: () => number, items: readonly T[]): T =>
  items[Math.floor(next() * items.length)]!;

/** A phrase of two to six words from the list; some carry characters that need escaping. */
function phrase(next: () => number): string {
  const length = 2 + Math.floor(next() * 5);
  let text = pick(next, WORDS);
  for (let i = 1; i < length; i++) text += ` ${pick(next, WORDS)}`;
  return text;
}

function pool(next: () => number, size: number): string[] {
  return Array.from({ length: size }, () => phrase(next));
}

function header(titles: readonly string[]): CellObject[] {
  const bold = { bold: true };
  return titles.map((value) => ({ value, style: bold }));
}

const DAY = 86_400_000;
const START = Date.UTC(2026, 0, 1);

/** 100 000 rows by 10 mixed columns: ids, names, cities, money, counts, dates, flags, unique SKUs, notes, statuses. */
/** The wide table's rows, one at a time: a header, then `count` rows of ten mixed columns. */
function* wideRows(count: number): Generator<Cell[]> {
  const next = random(1);
  const notes = pool(next, 200);
  yield header([
    "Id",
    "Name",
    "City",
    "Amount",
    "Quantity",
    "When",
    "Flag",
    "SKU",
    "Note",
    "Status",
  ]);
  for (let i = 0; i < count; i++) {
    yield [
      i + 1,
      `${pick(next, FIRST_NAMES)} ${pick(next, LAST_NAMES)}`,
      pick(next, CITIES),
      Math.round(next() * 1_000_000) / 100,
      1 + Math.floor(next() * 500),
      new Date(START + Math.floor(next() * 365 * DAY)),
      next() < 0.5,
      `SKU-${String(100_000 + i)}`,
      pick(next, notes),
      pick(next, STATUSES),
    ];
  }
}

const wideColumns = [
  { width: 10 },
  { width: 24 },
  { width: 16 },
  { width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
  { width: 10 },
  { width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
  { width: 8 },
  { width: 12 },
  { width: 32 },
  { width: 12 },
];

/** 100 000 rows by 10 mixed columns, built in memory as `writeXlsx` takes it. */
function wideWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: "Orders",
        columns: wideColumns,
        freeze: { rows: 1 },
        autoFilter: true,
        rows: [...wideRows(100_000)],
      },
    ],
  };
}

/** The wide table with `count` rows produced as they are pulled, as `streamXlsx` takes it. */
export function wideStreamed(count: number): StreamedWorkbook {
  return {
    sheets: [
      {
        name: "Orders",
        columns: wideColumns,
        freeze: { rows: 1 },
        autoFilter: true,
        rows: wideRows(count),
      },
    ],
  };
}

/** 40 sheets of 2 500 rows by 6 columns: the same shape repeated, shared strings across sheets. */
function manySheetsWorkbook(): Workbook {
  const next = random(2);
  const sheets = Array.from({ length: 40 }, (_, s) => {
    const rows: Cell[][] = Array.from({ length: 2_500 }, () => [
      `${pick(next, FIRST_NAMES)} ${pick(next, LAST_NAMES)}`,
      pick(next, CITIES),
      Math.round(next() * 100_000) / 100,
      1 + Math.floor(next() * 50),
      new Date(START + Math.floor(next() * 365 * DAY)),
      pick(next, STATUSES),
    ]);
    return {
      name: `Region ${String(s + 1).padStart(2, "0")}`,
      columns: [
        { width: 24 },
        { width: 16 },
        { width: 14, style: { numberFormat: "#,##0.00" } },
        { width: 10 },
        { width: 14, style: { numberFormat: "yyyy-mm-dd" } },
        { width: 12 },
      ],
      freeze: { rows: 1 },
      rows: [header(["Name", "City", "Amount", "Quantity", "Date", "Status"]), ...rows],
    };
  });
  return { sheets };
}

/** 200 000 rows by 5 text columns drawn from 2 000 phrases: interning hits far more than misses. */
function stringHeavyWorkbook(): Workbook {
  const next = random(3);
  const phrases = pool(next, 2_000);
  const rows: Cell[][] = Array.from({ length: 200_000 }, () => [
    pick(next, phrases),
    pick(next, phrases),
    pick(next, phrases),
    pick(next, phrases),
    pick(next, phrases),
  ]);
  return { sheets: [{ name: "Text", rows: [header(["A", "B", "C", "D", "E"]), ...rows] }] };
}

/**
 * The same workbook with every column sized to its content instead of a fixed width. The columns
 * are named one by one: a build from before auto widths refuses `width: "auto"`, where it would
 * ignore a sheet's `autoWidth` and write the suite unmeasured, which is not the same work.
 */
export function withAutoWidths(workbook: Workbook): Workbook {
  return {
    sheets: workbook.sheets.map((sheet) => {
      const [first] = sheet.rows;
      // The suites give every row as an array.
      const width = Array.isArray(first) ? first.length : 0;
      const count = Math.max(sheet.columns?.length ?? 0, width);
      return {
        ...sheet,
        columns: Array.from({ length: count }, (_, index) => {
          const style = sheet.columns?.[index]?.style;
          return style === undefined ? { width: "auto" } : { width: "auto", style };
        }),
      };
    }),
  };
}

export type SuiteName = "wide table" | "many sheets" | "string-heavy";

/** The suites in report order. */
export const suites: ReadonlyArray<readonly [SuiteName, () => Workbook]> = [
  ["wide table", wideWorkbook],
  ["many sheets", manySheetsWorkbook],
  ["string-heavy", stringHeavyWorkbook],
];
