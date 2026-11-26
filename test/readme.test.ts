import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, afterEach, describe, expect, it, vi } from "vite-plus/test";
import { resolveLimits } from "../src/errors/read.js";
import { type Row, XLSX_MIME_TYPE, writeXlsx } from "../src/index.js";
import { badgePng, stripesPng } from "./fixtures/workbooks.js";
import { packageOf } from "./helpers/package.js";
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
  images: async (snippet) => {
    const products = [
      { name: "Desk", code: "D-1001", picture: stripesPng },
      { name: "Chair", code: "C-2002", picture: badgePng },
    ];
    const bytes = await (snippet["default"] as (products: unknown[]) => Promise<Uint8Array>)(
      products,
    );
    isWorkbook(bytes);
    expect(unzip(bytes).map((e) => e.name)).toContain("xl/media/image2.png");
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<drawing r:id="rId1"/>');
    expect(sheet).toContain('ht="33" customHeight="1"');
  },
  "image-padding": async (snippet) => {
    const items = [
      { name: "Desk", label: stripesPng },
      { name: "Chair", label: stripesPng },
    ];
    const bytes = await (snippet["default"] as (items: unknown[]) => Promise<Uint8Array>)(items);
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    // 40 pixels of picture and 10 above and below: 60 pixels, 45 points.
    expect(sheet.match(/ht="45" customHeight="1"/g)).toHaveLength(2);
    expect(readPart(bytes, "xl/drawings/drawing1.xml")).toContain(
      "<xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>95250</xdr:rowOff>",
    );
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
  "row-styles": async (snippet) => {
    const orders = [
      ["A-1", "Kari", 10.5],
      ["A-2", "Ola", 20],
    ];
    const bytes = await (snippet["default"] as (orders: unknown[]) => Promise<Uint8Array>)(orders);
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<row r="1" ht="24" customHeight="1"><c r="A1" s="1"');
    // The band on the second order, the rule across the total row's empty cell.
    expect(sheet).toMatch(/<c r="A3" s="(\d+)" t="s"><v>\d+<\/v><\/c><c r="B3" s="\1"/);
    expect(sheet).toMatch(/<c r="B4" s="\d+"\/><c r="C4" s="\d+"><v>30.5<\/v><\/c>/);
  },
  hidden: async (snippet) => {
    const orders = [
      { id: 1, number: "A-1", status: "Open", amount: 10.5, cancelled: false },
      { id: 2, number: "A-2", status: "Closed", amount: 20, cancelled: true },
    ];
    const bytes = await (snippet["default"] as (orders: unknown[]) => Promise<Uint8Array>)(orders);
    isWorkbook(bytes);
    const sheet = readPart(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<cols><col min="1" max="1" hidden="1"/>');
    expect(sheet).toContain('<row r="3" hidden="1"><c r="A3"><v>2</v></c>');
    expect(sheet).toContain("<formula1>&apos;Lists&apos;!$A$1:$A$3</formula1>");
    // The first sheet is shown, so the workbook opens on it without being told to.
    expect(readPart(bytes, "xl/workbook.xml")).toContain(
      '<sheets><sheet name="Orders" sheetId="1" r:id="rId1"/>' +
        '<sheet name="Lists" sheetId="2" state="hidden" r:id="rId2"/></sheets>',
    );
  },
  "sheet-view": async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/worksheets/sheet1.xml")).toContain(
      '<sheetView showGridLines="0" zoomScale="120" workbookViewId="0"/>',
    );
    expect(readPart(bytes, "xl/worksheets/sheet2.xml")).toContain(
      '<sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1"',
    );
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
  "read-objects": async (snippet) => {
    const exportOrders = snippet["exportOrders"] as (orders: unknown[]) => Promise<Uint8Array>;
    const importOrders = snippet["importOrders"] as (file: Uint8Array) => Promise<unknown[]>;
    const orders = [
      { id: "A-1", total: 10.5, placed: new Date(Date.UTC(2026, 0, 1)) },
      { id: "A-2", total: 1249.9, placed: new Date(Date.UTC(2026, 1, 28, 14, 30)) },
    ];
    const bytes = await exportOrders(orders);
    isWorkbook(bytes);
    expect(readPart(bytes, "xl/sharedStrings.xml")).toContain("<t>Placed</t>");
    expect(await importOrders(bytes)).toEqual(orders);
    const free = { id: "A-3", total: "free", placed: new Date(Date.UTC(2026, 2, 1)) };
    await expect(importOrders(await exportOrders([...orders, free]))).rejects.toThrow(
      "Row 4 is not an order",
    );
  },
  "read-rows": async (snippet) => {
    const path = join(scratch, "totals.xlsx");
    const orders: Row[] = [
      ["Order", "Customer", "Total"],
      ["A-1", "C-1", 10.5],
      [],
      ["A-2", "C-2", 20],
      ["A-3", "C-3", "free"],
      [null, null, 1.25],
    ];
    const notes: Row[] = [["Total"], [1000]];
    const sheets = [
      { name: "Notes", rows: notes },
      { name: "Orders", rows: orders },
    ];
    writeFileSync(path, await writeXlsx({ sheets }));
    expect(await (snippet["default"] as (path: string) => Promise<number>)(path)).toBe(31.75);
  },
  "read-errors": async (snippet) => {
    const countRows = snippet["default"] as (upload: Blob) => Promise<number | string>;
    const rows: Row[] = [["Order", "Total"], ["A-1", 10], [], ["A-2", 20]];
    const stored = await writeXlsx({ sheets: [{ name: "Orders", rows }] }, { compression: "none" });
    expect(await countRows(new Blob([stored]))).toBe(3);
    expect(await countRows(new Blob(["Order,Total\nA-1,10\n"]))).toBe("This is not an .xlsx file.");

    const value = Buffer.from(stored).indexOf("<v>20</v>");
    expect(value, "the stored value in the package").toBeGreaterThan(0);
    const changed = stored.slice();
    changed[value + 3] = "3".charCodeAt(0);
    expect(await countRows(new Blob([changed]))).toBe("This file is damaged.");

    const parts = unzip(stored).map(({ name, data }) => ({
      path: name,
      data:
        name === "xl/worksheets/sheet1.xml"
          ? new TextDecoder().decode(data).replace('r="B2"', 'r="A2"')
          : data,
    }));
    expect(await countRows(new Blob([packageOf(parts)]))).toBe(
      "This file is damaged at cell A2 of Orders.",
    );

    const lastColumn = 16_384;
    const defaultCellPositions = 5_000_000;
    const far: Row[] = Array.from(
      { length: Math.floor(defaultCellPositions / lastColumn) + 1 },
      () => [...Array<null>(lastColumn - 1).fill(null), 1],
    );
    expect(
      await countRows(new Blob([await writeXlsx({ sheets: [{ name: "Far", rows: far }] })])),
    ).toBe("This file asks for more than limits.cellPositions allows.");

    class Unreadable extends Blob {
      override arrayBuffer(): Promise<ArrayBuffer> {
        return Promise.reject(new Error("the disk failed"));
      }
      override slice(): Blob {
        return new Unreadable();
      }
    }
    await expect(countRows(new Unreadable([stored]))).rejects.toThrow("the disk failed");
  },
  "read-upload": async (snippet) => {
    const post = snippet["POST"] as (request: Request) => Promise<Response>;
    const send = (file?: Blob) => {
      const form = new FormData();
      if (file !== undefined) form.set("file", file);
      return post(new Request("https://example.com/counts", { method: "POST", body: form }));
    };
    const count = async (rows: Row[]) =>
      send(new Blob([await writeXlsx({ sheets: [{ name: "Count", rows }] })]));

    const counted = await count([
      ["SKU", "Counted"],
      ["S-1", 4],
      ["S-2", 0],
      ["S-1", 5],
    ]);
    expect(counted.status).toBe(200);
    expect(await counted.json()).toEqual({ counted: 2 });

    const wrong = await count([
      ["SKU", "Counted"],
      ["S-1", 4],
      ["S-2", "many"],
    ]);
    expect([wrong.status, await wrong.text()]).toEqual([422, "Row 3 needs a SKU and a count."]);

    const snippetCellPositions = 50_000;
    const large = Array.from({ length: snippetCellPositions / 2 }, (_, index): Row => [
      `S-${index}`,
      index,
    ]);
    const refused = await count([["SKU", "Counted"], ...large]);
    expect([refused.status, await refused.text()]).toEqual([
      422,
      "The file was refused: LIMIT_EXCEEDED.",
    ]);

    const renamed = await count([
      ["Item", "Counted"],
      ["S-1", 4],
    ]);
    expect(await renamed.text()).toBe("The file was refused: MISSING_HEADER.");
    expect(await (await send(new Blob(["SKU,Counted\nS-1,4\n"]))).text()).toBe(
      "The file was refused: NOT_XLSX.",
    );
    expect((await send()).status).toBe(400);
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

function limitsTable(): [limit: string, allowance: number][] {
  const table = /^\| Limit +\| Default +\| What it bounds +\|\n\|[- |]+\n((?:\|[^\n]*\n)+)/m.exec(
    readme,
  );
  expect(table, "the README's table of reading limits").not.toBeNull();
  return table![1]!
    .trimEnd()
    .split("\n")
    .map((line) => {
      const row = /^\| `(\w+)` +\| (\d[\d ]*?)( MiB)? +\|/.exec(line);
      expect(row, line).not.toBeNull();
      const [, limit, digits, mebibytes] = row!;
      return [limit!, Number(digits!.replaceAll(" ", "")) * (mebibytes ? 1_048_576 : 1)];
    });
}

describe("the README's table of reading limits", () => {
  const meters = () => {
    const limits = resolveLimits(undefined);
    return [...Object.values(limits.book), ...Object.values(limits.rows())];
  };
  const takes = (meter: { charge(amount: number): void }, amount: number): boolean => {
    try {
      meter.charge(amount);
      return true;
    } catch (error) {
      expect(error).toMatchObject({ code: "LIMIT_EXCEEDED" });
      return false;
    }
  };

  it("names every limit once", () => {
    const named = limitsTable().map(([limit]) => limit);
    expect(new Set(named).size, "a limit named twice").toBe(named.length);
    expect(new Set(named)).toEqual(new Set(meters().map(({ limit }) => limit)));
  });

  it("gives each limit the default the read side holds it to", () => {
    for (const [limit, allowance] of limitsTable()) {
      const of = (amount: number) =>
        meters()
          .filter((meter) => meter.limit === limit)
          .map((meter) => takes(meter, amount));
      expect(of(allowance), `a meter of ${limit} takes ${allowance}`).toContain(true);
      expect(of(allowance + 1), `no meter of ${limit} takes more`).not.toContain(true);
    }
  });
});
