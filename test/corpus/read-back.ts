// node test/corpus/read-back.ts [--check]
//
// LibreOffice's typed read-back of every workbook under test/corpus/<application>/: one headless
// conversion of them all to flat OpenDocument, then one JSON file beside each workbook. With
// --check it writes nothing and fails with every cell that differs from the committed read-back.
//
// The read-back keeps nothing a LibreOffice release or a locale can move: numbers to 15
// significant digits, dates and durations to the millisecond, the locale set to en-US, and a
// formula's stored result, since the profile tells LibreOffice never to recalculate on load.
import { execFileSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

export type ReadBackValue =
  | string
  | number
  | boolean
  | { readonly date: string }
  | { readonly duration: string }
  | { readonly error: string };

export type ReadBackCell = null | ReadBackValue | { readonly result: ReadBackValue };

type ReadBackRow = readonly [number: number, ...cells: ReadBackCell[]];

export interface ReadBack {
  readonly sheets: readonly { readonly name: string; readonly rows: readonly ReadBackRow[] }[];
}

interface Element {
  readonly name: string;
  readonly attributes: ReadonlyMap<string, string>;
  readonly children: (Element | string)[];
}

const NAMESPACES: Readonly<Record<string, string>> = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
};

const TOKEN =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<!\[CDATA\[([\s\S]*?)\]\]>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/y;
const ATTRIBUTE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const REFERENCE = /&(?:#x([0-9a-f]+)|#([0-9]+)|(amp|lt|gt|quot|apos));/gi;
const NAMED: Readonly<Record<string, string>> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

function decode(text: string): string {
  return text
    .replaceAll("\r\n", "\n")
    .replaceAll("\r", "\n")
    .replace(REFERENCE, (_, hex?: string, decimal?: string, name?: string) =>
      name === undefined
        ? String.fromCodePoint(hex === undefined ? Number(decimal) : Number.parseInt(hex, 16))
        : NAMED[name.toLowerCase()]!,
    );
}

function parseLibreOfficeXml(xml: string): Element {
  const root: Element = { name: "", attributes: new Map(), children: [] };
  const open: Element[] = [root];
  TOKEN.lastIndex = 0;
  while (TOKEN.lastIndex < xml.length) {
    const at = TOKEN.lastIndex;
    const token = TOKEN.exec(xml);
    if (token === null) throw new Error(`not XML at offset ${at}`);
    const [, cdata, close, name, attributes, empty, text] = token;
    const children = open.at(-1)!.children;
    if (cdata !== undefined) children.push(cdata);
    else if (text !== undefined) children.push(decode(text));
    else if (close !== undefined) {
      if (open.pop()?.name !== close) throw new Error(`</${close}> closes another element`);
    } else if (name !== undefined) {
      const element: Element = {
        name,
        attributes: new Map(
          [...(attributes ?? "").matchAll(ATTRIBUTE)].map(([, key, double, single]) => [
            key!,
            decode(double ?? single ?? ""),
          ]),
        ),
        children: [],
      };
      children.push(element);
      if (empty === "") open.push(element);
    }
  }
  if (open.length !== 1) throw new Error(`<${open.at(-1)!.name}> is not closed`);
  const document = root.children.find((child) => typeof child !== "string");
  if (document === undefined) throw new Error("no document element");
  return document;
}

const elements = (parent: Element, name: string) =>
  parent.children.filter(
    (child): child is Element => typeof child !== "string" && child.name === name,
  );

function only(parent: Element, name: string): Element {
  const [found, ...more] = elements(parent, name);
  if (found === undefined || more.length > 0)
    throw new Error(`not one <${name}> in <${parent.name}>`);
  return found;
}

function attribute(element: Element, name: string): string {
  const value = element.attributes.get(name);
  if (value === undefined) throw new Error(`<${element.name}> has no ${name}`);
  return value;
}

function paragraph(element: Element): string {
  let text = "";
  for (const child of element.children) {
    if (typeof child === "string") text += child;
    else if (child.name === "text:s")
      text += " ".repeat(Number(child.attributes.get("text:c") ?? 1));
    else if (child.name === "text:tab") text += "\t";
    else if (child.name === "text:line-break") text += "\n";
    else text += paragraph(child);
  }
  return text;
}

const textOf = (cell: Element) => elements(cell, "text:p").map(paragraph).join("\n");

const DATE = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d):(\d\d)(\.\d+)?)?$/;
const DURATION = /^(-?)PT(\d+)H(\d+)M(\d+(?:\.\d+)?)S$/;

function dateOf(value: string): string {
  const parts = DATE.exec(value);
  if (parts === null) throw new Error(`a date of ${value}`);
  const [, year, month, day, hour = "0", minute = "0", second = "0", fraction = ""] = parts;
  const time = new Date(0);
  time.setUTCFullYear(Number(year), Number(month) - 1, Number(day));
  time.setUTCHours(Number(hour), Number(minute), Number(second));
  return new Date(time.getTime() + Math.round(Number(`0${fraction}`) * 1000))
    .toISOString()
    .slice(0, 23);
}

function durationOf(value: string): string {
  const parts = DURATION.exec(value);
  if (parts === null) throw new Error(`a duration of ${value}`);
  const [, sign, hours, minutes, seconds] = parts;
  const total = Math.round((Number(hours) * 3600 + Number(minutes) * 60 + Number(seconds)) * 1000);
  const pad = (number: number, width: number) => String(number).padStart(width, "0");
  const second = total % 60_000;
  return `${sign}PT${Math.floor(total / 3_600_000)}H${pad(Math.floor(total / 60_000) % 60, 2)}M${pad(Math.floor(second / 1000), 2)}.${pad(second % 1000, 3)}S`;
}

function cellValue(cell: Element): ReadBackValue | null {
  if (cell.attributes.get("calcext:value-type") === "error") return { error: textOf(cell) };
  const type = cell.attributes.get("office:value-type");
  switch (type) {
    case undefined:
      return null;
    case "float":
    case "percentage":
    case "currency":
      return Number(Number(attribute(cell, "office:value")).toPrecision(15));
    case "boolean":
      return attribute(cell, "office:boolean-value") === "true";
    case "date":
      return { date: dateOf(attribute(cell, "office:date-value")) };
    case "time":
      return { duration: durationOf(attribute(cell, "office:time-value")) };
    case "string": {
      const text = cell.attributes.get("office:string-value") ?? textOf(cell);
      return text === "" ? null : text;
    }
    default:
      throw new Error(`a value of type ${type}`);
  }
}

function cellOf(cell: Element): ReadBackCell {
  const value = cellValue(cell);
  return value !== null && cell.attributes.has("table:formula") ? { result: value } : value;
}

const ROW_GROUPS = new Set([
  "table:table-header-rows",
  "table:table-row-group",
  "table:table-rows",
]);
const CELLS = new Set(["table:table-cell", "table:covered-table-cell"]);

function* rowElements(parent: Element): Generator<Element> {
  for (const child of parent.children) {
    if (typeof child === "string") continue;
    if (child.name === "table:table-row") yield child;
    else if (ROW_GROUPS.has(child.name)) yield* rowElements(child);
  }
}

const repeated = (element: Element, name: string) => Number(element.attributes.get(name) ?? 1);

export function readBackOf(fods: string): ReadBack {
  const document = parseLibreOfficeXml(fods);
  for (const [prefix, uri] of Object.entries(NAMESPACES)) {
    if (document.attributes.get(`xmlns:${prefix}`) !== uri)
      throw new Error(`xmlns:${prefix} is not ${uri}`);
  }
  const spreadsheet = only(only(document, "office:body"), "office:spreadsheet");
  return {
    sheets: elements(spreadsheet, "table:table").map((table) => {
      const rows: ReadBackRow[] = [];
      let number = 1;
      for (const row of rowElements(table)) {
        const cells: ReadBackCell[] = [];
        let empty = 0;
        for (const cell of row.children) {
          if (typeof cell === "string" || !CELLS.has(cell.name)) continue;
          const value = cellOf(cell);
          const count = repeated(cell, "table:number-columns-repeated");
          if (value === null) {
            empty += count;
            continue;
          }
          cells.push(...Array<null>(empty).fill(null), ...Array<ReadBackCell>(count).fill(value));
          empty = 0;
        }
        const count = repeated(row, "table:number-rows-repeated");
        if (cells.length > 0) {
          for (let index = 0; index < count; index++) rows.push([number + index, ...cells]);
        }
        number += count;
      }
      return { name: attribute(table, "table:name"), rows };
    }),
  };
}

function serialize(readBack: ReadBack): string {
  const sheets = readBack.sheets.map(
    ({ name, rows }) =>
      `    {\n      "name": ${JSON.stringify(name)},\n      "rows": [\n${rows.map((row) => `        ${JSON.stringify(row)}`).join(",\n")}\n      ]\n    }`,
  );
  return `{\n  "sheets": [\n${sheets.join(",\n")}\n  ]\n}\n`;
}

export const reference = (rowNumber: number, columnIndex: number): string => {
  let letters = "";
  for (let rest = columnIndex + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  }
  return `${letters}${rowNumber}`;
};

export function differences(workbook: string, committed: ReadBack, made: ReadBack): string[] {
  const names = (readBack: ReadBack) => readBack.sheets.map(({ name }) => name);
  if (JSON.stringify(names(committed)) !== JSON.stringify(names(made))) {
    return [
      `${workbook}: sheets ${JSON.stringify(names(committed))} committed, ${JSON.stringify(names(made))} now`,
    ];
  }
  const lines: string[] = [];
  committed.sheets.forEach((sheet, index) => {
    const before = new Map(sheet.rows.map(([number, ...cells]) => [number, cells]));
    const after = new Map(made.sheets[index]!.rows.map(([number, ...cells]) => [number, cells]));
    for (const number of [...new Set([...before.keys(), ...after.keys()])].sort((a, b) => a - b)) {
      const old = before.get(number) ?? [];
      const now = after.get(number) ?? [];
      for (let column = 0; column < Math.max(old.length, now.length); column++) {
        const was = JSON.stringify(old[column] ?? null);
        const is = JSON.stringify(now[column] ?? null);
        if (was !== is) {
          lines.push(
            `${workbook}, sheet ${JSON.stringify(sheet.name)}, ${reference(number, column)}: ${was} committed, ${is} now`,
          );
        }
      }
    }
  });
  return lines;
}

function findSoffice(): string {
  const candidates = [
    process.env["SOFFICE"],
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/lib/libreoffice/program/soffice",
  ];
  const found = candidates.find((path) => path !== undefined && existsSync(path));
  if (found === undefined) throw new Error("LibreOffice was not found; set SOFFICE to its path");
  return found;
}

/** A user profile that pins the locale and stops LibreOffice from recalculating on load. */
const PROFILE = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>en-US</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>en-US</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>1</value></prop></item>
</oor:items>
`;

function convert(workbooks: readonly string[]): Map<string, string> {
  const temporary = mkdtempSync(join(tmpdir(), "xlsx-mini-read-back-"));
  try {
    mkdirSync(join(temporary, "profile", "user"), { recursive: true });
    writeFileSync(join(temporary, "profile", "user", "registrymodifications.xcu"), PROFILE);
    mkdirSync(join(temporary, "in"));
    mkdirSync(join(temporary, "out"));
    const inputs = workbooks.map((workbook) => {
      const input = join(temporary, "in", workbook.replace("/", "--"));
      copyFileSync(join(corpus, workbook), input);
      return input;
    });
    execFileSync(
      findSoffice(),
      [
        `-env:UserInstallation=${pathToFileURL(join(temporary, "profile")).href}`,
        "--headless",
        "--convert-to",
        "fods",
        "--outdir",
        join(temporary, "out"),
        ...inputs,
      ],
      { stdio: "pipe", timeout: 600_000 },
    );
    const made = new Map<string, string>();
    for (const workbook of workbooks) {
      const output = join(
        temporary,
        "out",
        workbook.replace("/", "--").replace(/\.xlsx$/, ".fods"),
      );
      if (existsSync(output)) made.set(workbook, readFileSync(output, "utf8"));
    }
    return made;
  } finally {
    rmSync(temporary, { recursive: true, force: true });
  }
}

const corpus = dirname(fileURLToPath(import.meta.url));

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const check = process.argv.includes("--check");
  const workbooks = readdirSync(corpus, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .flatMap(({ name }) =>
      readdirSync(join(corpus, name))
        .filter((file) => file.endsWith(".xlsx"))
        .map((file) => `${name}/${file}`),
    )
    .sort();
  const started = performance.now();
  const made = convert(workbooks);
  const lines: string[] = [];
  const written: string[] = [];
  for (const workbook of workbooks) {
    const fods = made.get(workbook);
    if (fods === undefined) {
      lines.push(`${workbook}: LibreOffice did not convert it`);
      continue;
    }
    let readBack: ReadBack;
    try {
      readBack = readBackOf(fods);
    } catch (error) {
      lines.push(`${workbook}: ${(error as Error).message}`);
      continue;
    }
    const path = join(corpus, workbook.replace(/\.xlsx$/, ".json"));
    if (!check) {
      writeFileSync(path, serialize(readBack));
      written.push(path);
    } else if (!existsSync(path)) lines.push(`${workbook}: no committed read-back`);
    else
      lines.push(
        ...differences(workbook, JSON.parse(readFileSync(path, "utf8")) as ReadBack, readBack),
      );
  }
  for (const application of new Set(workbooks.map((workbook) => workbook.split("/")[0]!))) {
    for (const file of readdirSync(join(corpus, application))) {
      const workbook = `${application}/${file.replace(/\.json$/, ".xlsx")}`;
      if (file.endsWith(".json") && !workbooks.includes(workbook))
        lines.push(`${workbook}: a read-back without its workbook`);
    }
  }
  if (written.length > 0) {
    execFileSync(
      join(corpus, "..", "..", "node_modules", ".bin", "vp"),
      ["fmt", "--write", ...written],
      {
        stdio: "ignore",
      },
    );
  }
  const seconds = ((performance.now() - started) / 1000).toFixed(1);
  console.log(`${workbooks.length} workbooks read back by LibreOffice in ${seconds} s`);
  if (lines.length > 0) {
    console.error(lines.join("\n"));
    console.error(
      `${lines.length} differences; node test/corpus/read-back.ts writes the read-back again`,
    );
    process.exitCode = 1;
  }
}
