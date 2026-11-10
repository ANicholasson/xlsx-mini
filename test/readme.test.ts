import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, afterEach, describe, expect, it, vi } from "vite-plus/test";
import { XLSX_MIME_TYPE } from "../src/index.js";
import { readPart, unzip } from "./helpers/unzip.js";

/**
 * Every TypeScript block in the README follows a `<!-- snippet: name -->` marker, invisible to
 * readers, that names its file under examples/snippets.
 * The block must equal the file, and the file must run: a README cannot rot that way.
 */
const root = join(dirname(fileURLToPath(import.meta.url)), "..");
/** Git on Windows may check text out with CRLF; blocks and files compare on LF. */
const text = (file: string) => readFileSync(file, "utf8").replaceAll("\r\n", "\n");
const readme = text(join(root, "README.md"));
const scratch = mkdtempSync(join(tmpdir(), "xlsx-mini-readme-"));

interface Block {
  readonly file: string;
  readonly body: string;
}

function blocks(): Block[] {
  const found: Block[] = [];
  for (const match of readme.matchAll(/<!-- snippet: ([\w-]+) -->\n\n```ts\n([\s\S]*?)```/g)) {
    found.push({ file: `examples/snippets/${match[1]!}.ts`, body: match[2]! });
  }
  return found;
}

type Snippet = Record<string, unknown>;
type Fn<T> = (...args: never[]) => Promise<T>;

function isWorkbook(bytes: unknown): asserts bytes is Uint8Array {
  expect(bytes).toBeInstanceOf(Uint8Array);
  expect(unzip(bytes as Uint8Array).map((e) => e.name)).toContain("xl/workbook.xml");
}

/** The parts of a document the download touches, recorded instead of rendered. */
function stubDocument() {
  const anchor = { href: "", download: "", click: vi.fn(), remove: vi.fn() };
  vi.stubGlobal("document", { createElement: () => anchor, body: { append: vi.fn() } });
  vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:readme");
  vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
  return anchor;
}

async function asResponse(snippet: Snippet, name: string): Promise<Response> {
  const response = await (snippet[name] as Fn<Response>)();
  expect(response.headers.get("Content-Type")).toBe(XLSX_MIME_TYPE);
  isWorkbook(new Uint8Array(await response.arrayBuffer()));
  return response;
}

/** How each snippet runs, and what proves it did its job. */
const runs: Record<string, (snippet: Snippet) => Promise<void>> = {
  "quick-start": async (snippet) => {
    const path = join(scratch, "orders.xlsx");
    await (snippet["default"] as (path: string) => Promise<void>)(path);
    const bytes = new Uint8Array(readFileSync(path));
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/sharedStrings.xml")).toContain("<t>Kari Nordmann</t>");
  },
  response: async (snippet) => {
    const response = await asResponse(snippet, "default");
    expect(response.headers.get("Content-Disposition")).toContain("report.xlsx");
  },
  express: async (snippet) => {
    const headers: Record<string, string> = {};
    let body: Uint8Array | undefined;
    await (snippet["default"] as (reply: unknown) => Promise<void>)({
      setHeader: (name: string, value: string) => {
        headers[name] = value;
      },
      end: (bytes: Uint8Array) => {
        body = bytes;
      },
    });
    expect(headers["Content-Type"]).toBe(XLSX_MIME_TYPE);
    isWorkbook(body);
  },
  browser: async (snippet) => {
    const anchor = stubDocument();
    await (snippet["default"] as Fn<void>)();
    expect(anchor.download).toBe("report.xlsx");
    expect(anchor.click).toHaveBeenCalledOnce();
  },
  "next-route": async (snippet) => {
    await asResponse(snippet, "GET");
  },
  streaming: async (snippet) => {
    const pages = [[["A-1", "Kari", 10.5]], [["A-2", "Ola", 20]], []];
    let fetched = 0;
    const response = (
      snippet["default"] as (fetchPage: (page: number) => Promise<unknown[]>) => Response
    )((page) => {
      fetched++;
      return Promise.resolve(pages[page] ?? []);
    });
    expect(response.headers.get("Content-Type")).toBe(XLSX_MIME_TYPE);
    expect(response.headers.get("Content-Disposition")).toContain("orders.xlsx");
    // The handler returned before a single page was fetched: rows are pulled as the body is read.
    expect(fetched).toBe(0);
    const bytes = new Uint8Array(await response.arrayBuffer());
    isWorkbook(bytes);
    expect(fetched).toBe(3);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="B3" t="inlineStr"><is><t>Ola</t></is></c>');
    expect(sheet).toContain('<autoFilter ref="A1:C3"/>');
  },
  "next-edge": async (snippet) => {
    expect(snippet["runtime"]).toBe("edge");
    await asResponse(snippet, "GET");
  },
  actions: async (snippet) => {
    isWorkbook(await (snippet["exportReport"] as Fn<Uint8Array>)());
  },
  "next-client": async (snippet) => {
    const anchor = stubDocument();
    await (snippet["onExportClick"] as Fn<void>)();
    expect(anchor.download).toBe("report.xlsx");
  },
  objects: async (snippet) => {
    const orders = [
      { id: "A-1", customer: "Kari", total: 10.5, placed: new Date(Date.UTC(2026, 0, 1)) },
    ];
    const bytes = await (snippet["default"] as (orders: unknown[]) => Promise<Uint8Array>)(orders);
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/sharedStrings.xml")).toContain("<t>Customer</t>");
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain('<autoFilter ref="A1:D2"/>');
  },
  formulas: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="D2"><f>B2*C2</f></c>');
    expect(sheet).toContain('<c r="B4"><f>SUM(B2:B3)</f><v>6</v></c>');
    expect(sheet).toContain('<hyperlink ref="A7" location="&apos;Summary&apos;!A1"/>');
    expect(readPart(bytes, "xl/worksheets/_rels/sheet1.xml.rels")).toContain(
      'Target="https://example.com/prices"',
    );
    expect(readPart(bytes, "xl/worksheets/sheet2.xml")).toContain("<f>Budget!D4</f>");
    expect(readPart(bytes, "xl/workbook.xml")).toContain('<calcPr fullCalcOnLoad="1"/>');
  },
  validation: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<dataValidations count="4">');
    expect(sheet).toContain("<formula1>&quot;Open,Closed,On hold&quot;</formula1>");
    expect(sheet).toContain("<formula1>&apos;People&apos;!$A$2:$A$50</formula1>");
    expect(sheet).toContain(
      'type="date" operator="greaterThanOrEqual" allowBlank="1" showErrorMessage="1" errorStyle="warning"',
    );
  },
  print: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain(
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>',
    );
    expect(sheet).toContain("<oddFooter>&amp;L&amp;A&amp;RPage &amp;P of &amp;N</oddFooter>");
    expect(readPart(bytes, "xl/workbook.xml")).toContain("_xlnm.Print_Titles");
  },
  "auto-width": async (snippet) => {
    const products = [
      {
        sku: "SKU-100001",
        name: "Height-adjustable desk",
        price: 7490,
        restocked: new Date(0),
        note: "",
      },
      { sku: "SKU-7", name: "Chair", price: 1299.5, restocked: new Date(0), note: "Delivered" },
    ];
    const bytes = await (snippet["default"] as (products: unknown[]) => Promise<Uint8Array>)(
      products,
    );
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    const widths = [...sheet.matchAll(/<col min="\d+" max="\d+" width="([\d.]+)"/g)].map((m) =>
      Number(m[1]),
    );
    expect(widths).toHaveLength(5);
    expect(widths[4]).toBe(40);
    expect(widths.slice(0, 4).every((w) => w > 3 && w < 60)).toBe(true);
  },
  formats: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    const styles = readPart(bytes, "xl/styles.xml");
    expect(styles).toContain('formatCode="0.0%"');
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
  },
  appearance: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/styles.xml")).toContain(
      '<alignment horizontal="center" vertical="center"/>',
    );
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain('<mergeCell ref="A1:D1"/>');
  },
  dates: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/styles.xml")).toContain('numFmtId="22"');
  },
  register: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain("<v>12500</v>");
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain("<v>1234567890123</v>");
  },
  errors: async (snippet) => {
    expect(await (snippet["default"] as Fn<string>)()).toBe("INVALID_SHEET_NAME at sheets[0].name");
  },
};

describe("README snippets", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });
  afterAll(() => rmSync(scratch, { recursive: true, force: true }));

  const found = blocks();

  it("has one block per snippet file and one file per block", () => {
    const expected = Object.keys(runs).map((name) => `examples/snippets/${name}.ts`);
    expect(found.map((b) => b.file).sort()).toEqual(expected.sort());
  });

  it.each(found)("$file equals its README block and runs", async ({ file, body }) => {
    expect(body).toBe(text(join(root, file)));
    const name = file.slice("examples/snippets/".length, -".ts".length);
    const snippet = (await import(join(root, file))) as Snippet;
    await runs[name]!(snippet);
  });
});
