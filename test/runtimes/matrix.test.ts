import { execFile as execFileCallback, execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { EdgeVM } from "@edge-runtime/vm";
import { Miniflare } from "miniflare";
import { rolldown } from "rolldown";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { XLSX_MIME_TYPE, writeXlsx } from "../../src/index.js";
import { unzip } from "../helpers/unzip.js";
import { type HostReport, workbook } from "./scenario.mjs";

const execFile = promisify(execFileCallback);
const root = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const binary = (name: string) => join(root, "node_modules", ".bin", name);

/** A host of the matrix: how the scenario runs there, and the zip method its default output uses. */
interface Host {
  readonly name: string;
  readonly run: () => Promise<HostReport>;
  /** 8 where the host implements raw deflate, 0 where the writer falls back to STORE. */
  readonly method: 0 | 8;
}

const hosts: readonly Host[] = [
  {
    name: "Node ESM",
    method: 8,
    run: () => spawn(process.execPath, ["test/runtimes/entry.mjs"]),
  },
  {
    name: "Node CommonJS",
    method: 8,
    run: () => spawn(process.execPath, ["test/runtimes/entry.cjs"]),
  },
  {
    name: "Bun",
    method: 8,
    run: () => spawn(binary("bun"), ["run", "test/runtimes/entry.mjs"]),
  },
  {
    name: "Deno",
    method: 8,
    run: () => spawn(binary("deno"), ["run", "--quiet", "--no-lock", "test/runtimes/entry.mjs"]),
  },
  { name: "Cloudflare workerd", method: 8, run: runInWorkerd },
  { name: "Vercel edge runtime", method: 0, run: runInEdgeVM },
];

/** Runs an entry in a separate process and parses the one JSON line it prints. */
async function spawn(command: string, args: readonly string[]): Promise<HostReport> {
  const { stdout } = await execFile(command, [...args], { cwd: root, encoding: "utf8" });
  return JSON.parse(stdout) as HostReport;
}

/** Bundles an entry with everything it imports from dist/ into one file, as a deploy tool would. */
async function bundle(entry: string, format: "esm" | "iife"): Promise<string> {
  const build = await rolldown({ input: join(root, entry), platform: "browser" });
  try {
    const { output } = await build.generate({ format });
    return output[0].code;
  } finally {
    await build.close();
  }
}

async function runInWorkerd(): Promise<HostReport> {
  const contents = await bundle("test/runtimes/worker.mjs", "esm");
  const miniflare = new Miniflare({
    modules: [{ type: "ESModule", path: "worker.mjs", contents }],
    compatibilityDate: "2026-07-01",
  });
  try {
    const response = await miniflare.dispatchFetch("http://runtime.local/");
    return (await response.json()) as HostReport;
  } finally {
    await miniflare.dispose();
  }
}

async function runInEdgeVM(): Promise<HostReport> {
  const vm = new EdgeVM();
  vm.evaluate(await bundle("test/runtimes/edge.mjs", "iife"));
  return await vm.evaluate<Promise<HostReport>>("xlsxMiniReport");
}

const decode = (base64: string) => new Uint8Array(Buffer.from(base64, "base64"));
const parts = (bytes: Uint8Array) =>
  unzip(bytes).map((e) => [e.name, new TextDecoder().decode(e.data)] as const);

describe("runtime matrix", () => {
  let expected: Uint8Array;

  beforeAll(async () => {
    execFileSync(binary("tsc"), ["-p", "tsconfig.build.json"], { cwd: root, stdio: "inherit" });
    expected = await writeXlsx(workbook, { compression: "none" });
  }, 60_000);

  describe.each(hosts)("$name", (host) => {
    let report: HostReport;

    beforeAll(async () => {
      report = await host.run();
    }, 60_000);

    it("writes the STORE package byte for byte as the source does in this Node", () => {
      expect(decode(report.stored)).toEqual(expected);
    });

    it(`writes the default package with zip method ${host.method} and the same parts`, () => {
      const entries = unzip(decode(report.deflated));
      expect(entries.map((e) => e.method)).toEqual(entries.map(() => host.method));
      expect(parts(decode(report.deflated))).toEqual(parts(expected));
    });

    it("loads the browser entry, which refuses to download without a document", () => {
      expect(report.mime).toBe(XLSX_MIME_TYPE);
      expect(report.noDom).toBe("NO_DOM");
    });
  });
});
