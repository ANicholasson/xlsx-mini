import { getHeapStatistics } from "node:v8";
import { type StreamedPart, selectStreamCompression, streamPackage } from "../src/opc/index.js";
import { XlsxReadError, openXlsx } from "../src/read.js";

interface Reading {
  readonly row: string;
  readonly sheetBytes: number;
  readonly fileBytes: number;
  readonly outcomes: readonly string[];
  readonly milliseconds: number;
}

export interface Report {
  readonly heapLimit: number;
  readonly readings: readonly Reading[];
}

const encoder = new TextEncoder();
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function columnName(column: number): string {
  let name = "";
  for (let left = column; left > 0; left = Math.floor((left - 1) / 26)) {
    name = String.fromCharCode(65 + ((left - 1) % 26)) + name;
  }
  return name;
}

/**
 * A text of `length` characters, one of them outside Latin-1 so that an engine holds it at two
 * bytes a character.
 */
const text = (length: number) => encoder.encode(`Ā${"a".repeat(length - 1)}`);

function sheet(pieces: () => Generator<Uint8Array>): {
  readonly body: () => ReadableStream<Uint8Array>;
  readonly bytes: number;
} {
  let bytes = 0;
  for (const piece of pieces()) bytes += piece.length;
  return {
    bytes,
    body() {
      const made = pieces();
      return new ReadableStream({
        pull(controller) {
          const next = made.next();
          if (next.done) controller.close();
          else controller.enqueue(next.value);
        },
      });
    },
  };
}

const head = encoder.encode(`<worksheet xmlns="${MAIN}"><sheetData><row r="1">`);
const tail = encoder.encode("</row></sheetData></worksheet>");

const LAST_COLUMN = 16_384;
const CELL_CHARACTERS = 7_600;
const RUN_CHARACTERS = 4_000_000;
const RUNS = 31;

function* cells(): Generator<Uint8Array> {
  const value = text(CELL_CHARACTERS);
  yield head;
  for (let column = 1; column <= LAST_COLUMN; column++) {
    yield encoder.encode(`<c r="${columnName(column)}1" t="str"><v>`);
    yield value;
    yield encoder.encode("</v></c>");
  }
  yield tail;
}

function* runs(): Generator<Uint8Array> {
  const value = text(RUN_CHARACTERS);
  yield head;
  yield encoder.encode('<c r="A1" t="inlineStr"><is>');
  for (let run = 0; run < RUNS; run++) {
    yield encoder.encode("<r><t>");
    yield value;
    yield encoder.encode("</t></r>");
  }
  yield encoder.encode("</is></c>");
  yield tail;
}

async function packageOf(body: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Record<string, string> = {
    "_rels/.rels": `<Relationships><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns:r="${RELATIONSHIPS}"><sheets>${[1, 2, 3].map((at) => `<sheet name="Sheet${at}" sheetId="${at}" r:id="rId1"/>`).join("")}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="${RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  };
  async function* streamed(): AsyncGenerator<StreamedPart> {
    for (const [path, xml] of Object.entries(parts)) {
      yield { path, body: new Blob([encoder.encode(xml)]).stream() };
    }
    yield { path: "xl/worksheets/sheet1.xml", body };
  }
  const written: Uint8Array[] = [];
  let length = 0;
  for await (const chunk of streamPackage(streamed(), selectStreamCompression("deflate"))) {
    written.push(chunk);
    length += chunk.length;
  }
  const file = new Uint8Array(length);
  let at = 0;
  for (const chunk of written) {
    file.set(chunk, at);
    at += chunk.length;
  }
  return file;
}

async function outcomeOf(rows: AsyncIterable<unknown>): Promise<string> {
  try {
    for await (const _ of rows);
    return "read";
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    return `${error.code} ${error.limit} ${error.sheet}`;
  }
}

const readings: Reading[] = [];
for (const [row, pieces] of [
  ["cells", cells],
  ["runs", runs],
] as const) {
  const made = sheet(pieces);
  const file = await packageOf(made.body());
  const book = await openXlsx(file);
  const started = performance.now();
  const outcomes = await Promise.all(book.sheets.map((name) => outcomeOf(book.rows(name))));
  readings.push({
    row,
    sheetBytes: made.bytes,
    fileBytes: file.length,
    outcomes,
    milliseconds: performance.now() - started,
  });
}

const report: Report = { heapLimit: getHeapStatistics().heap_size_limit, readings };
process.stdout.write(`${JSON.stringify(report)}\n`);
