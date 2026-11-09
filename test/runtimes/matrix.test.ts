import { execFile as execFileCallback, execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { EdgeVM } from "@edge-runtime/vm";
import { Miniflare } from "miniflare";
import { rolldown } from "rolldown";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { XLSX_MIME_TYPE, streamXlsx, writeXlsx } from "../../src/index.js";
import { type Command, binary as binaryIn, direct } from "../helpers/bin.js";
import { readPart as readPartOf, unzip } from "../helpers/unzip.js";
import { type HostReport, workbook } from "./scenario.mjs";

const execFile = promisify(execFileCallback);
const root = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const binary = (name: string) => binaryIn(root, name);
const node = direct(process.execPath);

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
    run: () => spawn(node, ["test/runtimes/entry.mjs"]),
  },
  {
    name: "Node CommonJS",
    method: 8,
    run: () => spawn(node, ["test/runtimes/entry.cjs"]),
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
async function spawn(command: Command, args: readonly string[]): Promise<HostReport> {
  const { stdout } = await execFile(command.file, [...args], {
    cwd: root,
    encoding: "utf8",
    shell: command.shell,
  });
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
  let expectedStreamed: Uint8Array;

  beforeAll(async () => {
    const tsc = binary("tsc");
    execFileSync(tsc.file, ["-p", "tsconfig.build.json"], {
      cwd: root,
      stdio: "inherit",
      shell: tsc.shell,
    });
    expected = await writeXlsx(workbook, { compression: "none" });
    expectedStreamed = new Uint8Array(
      await new Response(
        streamXlsx(workbook, { strings: "shared", compression: "none" }),
      ).arrayBuffer(),
    );
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

    it("streams the stored-block package byte for byte as the source does, from async rows", () => {
      expect(decode(report.streamed)).toEqual(expectedStreamed);
      const withoutDimension = (list: readonly (readonly [string, string])[]) =>
        list.map(([name, xml]) => [name, xml.replace(/<dimension ref="[^"]*"\/>/, "")]);
      const byName = (list: readonly (readonly [string, string])[]) =>
        [...list].sort((a, b) => a[0].localeCompare(b[0]));
      expect(withoutDimension(byName(parts(decode(report.streamed))))).toEqual(
        withoutDimension(byName(parts(expected))),
      );
    });

    it("streams the default package through the host's own deflate, or stored blocks, to the same parts", () => {
      const entries = unzip(decode(report.streamedDeflated));
      expect(entries.every((e) => e.method === 8)).toBe(true);
      for (const entry of entries) {
        // Stored blocks cost five bytes per 65 535-byte block plus the empty final block; raw
        // deflate must beat that on every part of a workbook this repetitive.
        const storedSize =
          entry.uncompressedSize + 5 * Math.ceil(entry.uncompressedSize / 65_535) + 5;
        if (host.method === 8) expect(entry.compressedSize).toBeLessThan(storedSize);
        else expect(entry.compressedSize).toBe(storedSize);
      }
      expect(readPartOf(decode(report.streamedDeflated), "xl/worksheets/sheet1.xml")).toContain(
        't="inlineStr"',
      );
    });

    it("loads the browser entry, which refuses to download without a document", () => {
      expect(report.mime).toBe(XLSX_MIME_TYPE);
      expect(report.noDom).toBe("NO_DOM");
    });
  });
});
