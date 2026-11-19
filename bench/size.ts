// Measures what the built package adds to an application's bundle: the package is bundled from
// its `dist/` folder the way a deploy tool would, minified and gzipped, once for each way an
// application imports it. The size test holds the result to its budgets and the writer's bundles
// to their committed digests, and `--dist <dist> --base <dist> --json <file>` writes both builds'
// sizes for the benchmark report.
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { gzipSync } from "node:zlib";
import { rolldown } from "rolldown";

/** One bundle: the SHA-256 of its minified bytes, their number, and their number gzipped. */
interface BundleSize {
  readonly digest: string;
  readonly minified: number;
  readonly gzip: number;
}

/**
 * The bundles measured. The writer's four: an application importing only `writeXlsx`, only
 * `streamXlsx`, every export, and the browser entry. Then the reader's, which a build from before
 * the reader does not have.
 */
export interface PackageSize {
  readonly writeXlsx: BundleSize;
  readonly streamXlsx: BundleSize;
  readonly everything: BundleSize;
  readonly browser: BundleSize;
  readonly read?: BundleSize;
}

/** What the benchmark report shows: the build under test and, on a pull request, its base. */
export interface SizeResult {
  readonly head: PackageSize;
  readonly base?: PackageSize;
}

/**
 * The most gzipped bytes each bundle may take. Only a feature that needs the bytes raises it, in its
 * own pull request and in ADR 0021.
 */
export const SIZE_BUDGET = { writeXlsx: 19_000, everything: 22_000, read: 8_000 } as const;

const ENTRY = "\0entry";

/** Bundles a module made of one statement, minified, and measures the one file that comes out. */
async function measureBundle(statement: string): Promise<BundleSize> {
  const build = await rolldown({
    input: "entry",
    platform: "browser",
    logLevel: "silent",
    // What the package's own `sideEffects: false` tells a bundler, said here so that a build
    // measured from a folder without that package.json is shaken the same way.
    treeshake: { moduleSideEffects: false },
    plugins: [
      {
        name: "entry",
        resolveId: (id) => (id === "entry" ? ENTRY : null),
        load: (id) => (id === ENTRY ? statement : null),
      },
    ],
  });
  try {
    const { output } = await build.generate({ format: "esm", minify: true });
    const code = new TextEncoder().encode(output[0].code);
    return {
      digest: createHash("sha256").update(code).digest("hex"),
      minified: code.length,
      gzip: gzipSync(code, { level: 9 }).length,
    };
  } finally {
    await build.close();
  }
}

/** The sizes of the package built into `dist`. */
export async function measurePackage(dist: string): Promise<PackageSize> {
  const url = (file: string) => JSON.stringify(pathToFileURL(resolve(dist, file)).href);
  const index = url("index.js");
  const writer = {
    writeXlsx: await measureBundle(`export { writeXlsx } from ${index};`),
    streamXlsx: await measureBundle(`export { streamXlsx } from ${index};`),
    everything: await measureBundle(`export * from ${index};`),
    browser: await measureBundle(`export * from ${url("browser.js")};`),
  };
  if (!existsSync(resolve(dist, "read.js"))) return writer;
  return { ...writer, read: await measureBundle(`export * from ${url("read.js")};`) };
}

function argument(argv: readonly string[], name: string): string | undefined {
  const index = argv.indexOf(name);
  return index === -1 ? undefined : argv[index + 1];
}

/** `--dist <dist>` measures a build; `--base <dist>` a second one; `--json <file>` keeps both. */
async function main(argv: readonly string[]): Promise<void> {
  const base = argument(argv, "--base");
  const result: SizeResult = {
    head: await measurePackage(argument(argv, "--dist") ?? "dist"),
    ...(base === undefined ? {} : { base: await measurePackage(base) }),
  };
  console.table(
    Object.fromEntries(
      Object.entries(result).flatMap(([build, size]: [string, PackageSize]) =>
        Object.entries(size).map(([bundle, bytes]) => [`${build}, ${bundle}`, bytes]),
      ),
    ),
  );
  const out = argument(argv, "--json");
  if (out === undefined) return;
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, `${JSON.stringify(result, null, 2)}\n`);
  console.log(`written ${out}`);
}

if ((import.meta as { main?: boolean }).main === true) await main(process.argv.slice(2));
