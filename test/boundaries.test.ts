import { spawnSync } from "node:child_process";
import {
  cpSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { parseAst } from "rolldown/parseAst";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { binary } from "./helpers/bin.js";
import { readSideSources } from "./helpers/read-side.ts";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

const PLANTED: Readonly<Record<string, readonly string[]>> = {
  "src/index.ts": ["xlsx-mini/read"],
  "src/read.ts": ["./errors/read.ts", "./opc/read.ts"],
  "src/errors/read.ts": ["./index.ts"],
  "src/errors/index.ts": [],
  "src/write-xlsx.ts": ["./errors/read.ts", "./opc/reader.ts"],
  "src/opc/index.ts": ["./read.ts", "./lib/zip.ts"],
  "src/opc/reader.ts": [],
  "src/opc/read.ts": ["./read/part.ts", "../errors/index.ts"],
  "src/opc/read/part.ts": ["../lib/crc32.ts"],
  "src/opc/lib/crc32.ts": [],
  "src/opc/lib/zip.ts": ["./crc32.ts", "../read/part.ts"],
  "src/opc/tests/read.test.ts": ["../read.ts"],
  "test/read.test.ts": ["../src/read.ts"],
  "bench/read.ts": ["../src/read.ts"],
};

interface Violation {
  readonly from: string;
  readonly to: string;
  readonly rule: { readonly name: string };
}

describe("the boundary between the write side and the read side", () => {
  let dir: string;
  let violations: readonly Violation[];

  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "xlsx-mini-boundaries-"));
    for (const config of [".dependency-cruiser.cjs", "tsconfig.json"]) {
      cpSync(join(root, config), join(dir, config));
    }
    for (const [file, imports] of Object.entries(PLANTED)) {
      mkdirSync(dirname(join(dir, file)), { recursive: true });
      writeFileSync(
        join(dir, file),
        `${imports.map((target) => `import "${target}";\n`).join("")}export {};\n`,
      );
    }
    const depcruise = binary(root, "depcruise");
    const result = spawnSync(depcruise.file, ["src", "test", "bench", "--output-type", "json"], {
      cwd: dir,
      encoding: "utf8",
      shell: depcruise.shell,
    });
    violations = (JSON.parse(result.stdout) as { summary: { violations: Violation[] } }).summary
      .violations;
  }, 30_000);

  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  it("refuses an import into the read side from a writer file, by path or by the package's name, and allows tests and benchmarks", () => {
    expect(violations.map(({ rule, from, to }) => `${rule.name}: ${from} -> ${to}`).sort()).toEqual(
      [
        "write-side-never-imports-read-side: src/index.ts -> src/read.ts",
        "write-side-never-imports-read-side: src/opc/index.ts -> src/opc/read.ts",
        "write-side-never-imports-read-side: src/opc/lib/zip.ts -> src/opc/read/part.ts",
        "write-side-never-imports-read-side: src/write-xlsx.ts -> src/errors/read.ts",
      ],
    );
  });
});

/** The one read-side file that may forge: it builds the meters, and only a cast makes one. */
const BUILDS_METERS = "src/errors/read.ts";

const FORGERIES = `
const cast = { limit: "parts", charge(amount: number) {} } as Meter<"parts">;
const through = handmade as unknown as Meter<"parts">;
const angled = <Meter<"parts">>handmade;
const spread = { ...meter, charge() {} };
const { ...rest } = meter;
function forge<T>(value: T): T { return { ...value, charge() {} }; }
const assigned = Object.assign({}, meter, { charge() {} });
const created = Object["create"](meter);
const cloned = globalThis.structuredClone(meter);
const { assign } = Object;
const reached = globalThis.Object.assign({}, meter);
const kept = [...meters, ...Object.keys(meter), ["a"] as const] satisfies unknown[];
use(...meters, Object.freeze(meter), meters[0]!);
const inherited = Object.setPrototypeOf({ charge() {} }, meter);
function isMeter(value: unknown): value is Meter<"parts"> { return true; }
function assertMeter(value: unknown): asserts value is Meter<"parts"> {}
const indexed = globalThis["structuredClone"](meter);
const reflected = Reflect.get(meter, "charge");
const everything = globalThis;
const inflate = globalThis.DecompressionStream;
`;

/** What read-side code may call on `Object`: the functions that read one, and `freeze`. */
const OBJECT_READS: ReadonlySet<string> = new Set([
  "entries",
  "freeze",
  "hasOwn",
  "is",
  "keys",
  "values",
]);

type Node = { readonly type: string; readonly start: number } & Readonly<Record<string, unknown>>;

function memberName({ computed, property }: Node): unknown {
  const { name, value } = property as Node;
  return computed === true ? value : name;
}

function forgeryAt(
  node: Node,
  parent: Node | undefined,
  key: string | undefined,
): string | undefined {
  if (node.type === "TSTypeAssertion") return "cast";
  if (node.type === "TSTypePredicate") return "type predicate";
  if (node.type === "TSAsExpression") {
    const { typeName } = node["typeAnnotation"] as Node;
    return (typeName as Node | undefined)?.["name"] === "const" ? undefined : "cast";
  }
  if (node.type === "SpreadElement" && parent?.type === "ObjectExpression") return "object spread";
  if (node.type === "RestElement" && parent?.type === "ObjectPattern") return "object rest";
  if (node.type !== "Identifier") return undefined;
  if (node["name"] === "structuredClone") return "structuredClone";
  if (node["name"] === "Reflect") return "Reflect";
  const named =
    parent?.type === "MemberExpression" && key === "object" && parent["computed"] !== true;
  if (node["name"] === "globalThis") return named ? undefined : "globalThis";
  if (node["name"] !== "Object") return undefined;
  const member =
    parent?.type === "MemberExpression" && key === "object" ? memberName(parent) : undefined;
  if (typeof member !== "string") return "Object";
  return OBJECT_READS.has(member) ? undefined : `Object.${member}`;
}

function forgeriesIn(source: string): string[] {
  const forgeries: string[] = [];
  const visit = (value: unknown, parent?: Node, key?: string): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item, parent, key);
      return;
    }
    if (typeof value !== "object" || value === null || !("type" in value)) return;
    const node = value as Node;
    const forgery = forgeryAt(node, parent, key);
    if (forgery !== undefined) {
      forgeries.push(`${source.slice(0, node.start).split("\n").length}: ${forgery}`);
    }
    for (const [childKey, child] of Object.entries(node)) visit(child, node, childKey);
  };
  visit(parseAst(source, { lang: "ts" }));
  return forgeries;
}

describe("read-side code", () => {
  it("is searched for a cast, a type predicate, a copied object, and Object, Reflect or globalThis used to make one", () => {
    expect(forgeriesIn(FORGERIES)).toEqual([
      "2: cast",
      "3: cast",
      "3: cast",
      "4: cast",
      "5: object spread",
      "6: object rest",
      "7: object spread",
      "8: Object.assign",
      "9: Object.create",
      "10: structuredClone",
      "11: Object",
      "12: Object",
      "15: Object.setPrototypeOf",
      "16: type predicate",
      "17: type predicate",
      "18: globalThis",
      "19: Reflect",
      "20: globalThis",
    ]);
  });

  it("holds none of them, except in the file that builds the meters", () => {
    const sources = readSideSources(root);
    expect(sources).toContain(BUILDS_METERS);
    const forgeries = sources
      .filter((file) => file !== BUILDS_METERS)
      .flatMap((file) =>
        forgeriesIn(readFileSync(join(root, file), "utf8")).map((how) => `${file}:${how}`),
      );
    expect(forgeries).toEqual([]);
  });
});

/** The one file that may name the platform's inflate as a global: it constructs it. */
const INFLATES = "src/opc/read/part.ts";

const USES_OF_INFLATE = `export const uses = [
  typeof DecompressionStream,
  new DecompressionStream("gzip"),
  globalThis.DecompressionStream,
  globalThis["DecompressionStream"],
  self.DecompressionStream,
  typeof process,
];
`;

interface Diagnostic {
  readonly code: string;
  readonly filename: string;
  readonly labels: readonly { readonly span: { readonly line: number } }[];
}

describe("the platform's inflate", () => {
  it("is refused by the linter as a global outside the one file that constructs it, and as a property of the global object in every file, and that file is refused a Node global as every core file is", () => {
    const dir = mkdtempSync(join(tmpdir(), "xlsx-mini-inflate-"));
    try {
      for (const config of ["package.json", "tsconfig.json", "vite.config.ts"]) {
        cpSync(join(root, config), join(dir, config));
      }
      symlinkSync(join(root, "node_modules"), join(dir, "node_modules"), "junction");
      const files = [INFLATES, "src/opc/read/directory.ts", "src/opc/lib/zip.ts", "src/browser.ts"];
      for (const file of files) {
        mkdirSync(dirname(join(dir, file)), { recursive: true });
        writeFileSync(join(dir, file), USES_OF_INFLATE);
      }
      const vp = binary(root, "vp");
      const result = spawnSync(vp.file, ["lint", "src", "--format", "json"], {
        cwd: dir,
        encoding: "utf8",
        shell: vp.shell,
      });
      const { diagnostics } = JSON.parse(result.stdout) as { diagnostics: Diagnostic[] };
      const refused = (file: string) =>
        diagnostics
          .filter(({ filename }) => filename.replaceAll("\\", "/") === file)
          .map(({ code, labels }) => `${labels[0]?.span.line}: ${code}`)
          .sort();
      const asProperty = [4, 5, 6].map((line) => `${line}: eslint(no-restricted-properties)`);
      const asGlobal = [2, 3].map((line) => `${line}: eslint(no-restricted-globals)`);
      const nodeGlobal = "7: eslint(no-restricted-globals)";

      expect(refused(INFLATES)).toEqual([...asProperty, nodeGlobal]);
      for (const file of ["src/opc/read/directory.ts", "src/opc/lib/zip.ts"]) {
        expect(refused(file), file).toEqual([...asGlobal, ...asProperty, nodeGlobal]);
      }
      expect(refused("src/browser.ts"), "the browser entry may still use a Node global").toEqual([
        ...asGlobal,
        ...asProperty,
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 30_000);
});
