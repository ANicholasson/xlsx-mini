import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, afterEach, describe, expect, it, vi } from "vite-plus/test";
import { XLSX_MIME_TYPE } from "../src/index.js";
import { readPart, unzip } from "./helpers/unzip.js";

/**
 * Every TypeScript block in the README names its file under examples/snippets on its first line.
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
  for (const match of readme.matchAll(
    /```ts\n(\/\/ (examples\/snippets\/[\w-]+\.ts)\n[\s\S]*?)```/g,
  )) {
    found.push({ file: match[2]!, body: match[1]! });
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
  "next-edge": async (snippet) => {
    expect(snippet["runtime"]).toBe("edge");
    await asResponse(snippet, "GET");
  },
  "next-action": async (snippet) => {
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
  formats: async (snippet) => {
    const bytes = await (snippet["default"] as Fn<Uint8Array>)();
    isWorkbook(bytes);
    const styles = readPart(bytes, "xl/styles.xml");
    expect(styles).toContain('formatCode="0.0%"');
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
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
