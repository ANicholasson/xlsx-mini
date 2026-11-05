/**
 * The three benchmark workbooks. Plain JavaScript so the memory script can load them in a bare
 * Node process next to `dist/`. Every suite is built from a seeded generator, so the same bytes
 * come out on every machine and every run; building is done before timing starts.
 */

/** mulberry32: a small seeded generator, enough to make repeatable data. */
function random(seed) {
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

/** A phrase of two to six words from the list; some carry characters that need escaping. */
function phrase(next) {
  const length = 2 + Math.floor(next() * 5);
  let text = WORDS[Math.floor(next() * WORDS.length)];
  for (let i = 1; i < length; i++) text += " " + WORDS[Math.floor(next() * WORDS.length)];
  return text;
}

function pool(next, size) {
  const items = new Array(size);
  for (let i = 0; i < size; i++) items[i] = phrase(next);
  return items;
}

const DAY = 86_400_000;
const START = Date.UTC(2026, 0, 1);

/** 100 000 rows by 10 mixed columns: ids, names, cities, money, counts, dates, flags, unique SKUs, notes, statuses. */
export function wideWorkbook() {
  const next = random(1);
  const notes = pool(next, 200);
  const rows = new Array(100_000);
  for (let i = 0; i < rows.length; i++) {
    const name = `${FIRST_NAMES[Math.floor(next() * FIRST_NAMES.length)]} ${LAST_NAMES[Math.floor(next() * LAST_NAMES.length)]}`;
    rows[i] = [
      i + 1,
      name,
      CITIES[Math.floor(next() * CITIES.length)],
      Math.round(next() * 1_000_000) / 100,
      1 + Math.floor(next() * 500),
      new Date(START + Math.floor(next() * 365 * DAY)),
      next() < 0.5,
      `SKU-${String(100_000 + i)}`,
      notes[Math.floor(next() * notes.length)],
      STATUSES[Math.floor(next() * STATUSES.length)],
    ];
  }
  return {
    sheets: [
      {
        name: "Orders",
        columns: [
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
        ],
        freeze: { rows: 1 },
        autoFilter: true,
        rows: [
          header([
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
          ]),
          ...rows,
        ],
      },
    ],
  };
}

/** 40 sheets of 2 500 rows by 6 columns: the same shape repeated, shared strings across sheets. */
export function manySheetsWorkbook() {
  const next = random(2);
  const sheets = new Array(40);
  for (let s = 0; s < sheets.length; s++) {
    const rows = new Array(2_500);
    for (let i = 0; i < rows.length; i++) {
      rows[i] = [
        `${FIRST_NAMES[Math.floor(next() * FIRST_NAMES.length)]} ${LAST_NAMES[Math.floor(next() * LAST_NAMES.length)]}`,
        CITIES[Math.floor(next() * CITIES.length)],
        Math.round(next() * 100_000) / 100,
        1 + Math.floor(next() * 50),
        new Date(START + Math.floor(next() * 365 * DAY)),
        STATUSES[Math.floor(next() * STATUSES.length)],
      ];
    }
    sheets[s] = {
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
  }
  return { sheets };
}

/** 200 000 rows by 5 text columns drawn from 2 000 phrases: interning hits far more than misses. */
export function stringHeavyWorkbook() {
  const next = random(3);
  const phrases = pool(next, 2_000);
  const rows = new Array(200_000);
  for (let i = 0; i < rows.length; i++) {
    rows[i] = [
      phrases[Math.floor(next() * phrases.length)],
      phrases[Math.floor(next() * phrases.length)],
      phrases[Math.floor(next() * phrases.length)],
      phrases[Math.floor(next() * phrases.length)],
      phrases[Math.floor(next() * phrases.length)],
    ];
  }
  return {
    sheets: [{ name: "Text", rows: [header(["A", "B", "C", "D", "E"]), ...rows] }],
  };
}

function header(titles) {
  const bold = { bold: true };
  return titles.map((value) => ({ value, style: bold }));
}

/** The suites by name, for scripts that pick one from the command line. */
export const suites = {
  wide: wideWorkbook,
  "many-sheets": manySheetsWorkbook,
  strings: stringHeavyWorkbook,
};
