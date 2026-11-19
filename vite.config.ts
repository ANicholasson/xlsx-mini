import { fileURLToPath } from "node:url";
import { type ViteUserConfig, defineConfig } from "vite-plus";

const CORE_ONLY_MESSAGE =
  "The core runs everywhere: web-standard APIs only. Node and DOM globals belong in runtime-specific entries or tests.";

const CORE_ONLY_GLOBALS = ["process", "Buffer", "require", "document", "window", "navigator"].map(
  (name) => ({ name, message: CORE_ONLY_MESSAGE }),
);

const ONE_INFLATE_MESSAGE =
  "The platform's inflate is constructed in src/opc/read/part.ts and nowhere else: what it gives out is sized and checksummed there (ADR 0032).";

const INFLATE_GLOBAL = { name: "DecompressionStream", message: ONE_INFLATE_MESSAGE };

const config: ViteUserConfig = {
  // Snippets and examples import the package by name; tests run them against the source.
  resolve: {
    alias: [
      {
        find: "xlsx-mini/browser",
        replacement: fileURLToPath(new URL("./src/browser.ts", import.meta.url)),
      },
      {
        find: "xlsx-mini/read",
        replacement: fileURLToPath(new URL("./src/read.ts", import.meta.url)),
      },
      { find: "xlsx-mini", replacement: fileURLToPath(new URL("./src/index.ts", import.meta.url)) },
    ],
  },
  lint: {
    options: { typeAware: true, typeCheck: true },
    rules: {
      "no-restricted-globals": ["error", ...CORE_ONLY_GLOBALS, INFLATE_GLOBAL],
      // The same name as a property of the global object, which the rule above does not see.
      "no-restricted-properties": [
        "error",
        ...["globalThis", "self", "window"].map((object) => ({
          object,
          property: INFLATE_GLOBAL.name,
          message: ONE_INFLATE_MESSAGE,
        })),
      ],
      "no-restricted-imports": [
        "error",
        { patterns: [{ group: ["node:*"], message: CORE_ONLY_MESSAGE }] },
      ],
      "typescript/no-explicit-any": "error",
    },
    overrides: [
      {
        // The browser entry is the one module allowed to touch the DOM (ADR 0004).
        files: ["src/browser.ts"],
        rules: { "no-restricted-globals": ["error", INFLATE_GLOBAL] },
      },
      {
        files: ["src/opc/read/part.ts"],
        rules: { "no-restricted-globals": ["error", ...CORE_ONLY_GLOBALS] },
      },
      {
        files: ["src/**/tests/**", "test/**", "bench/**", "examples/**", "*.config.ts", "*.cjs"],
        rules: {
          "no-restricted-globals": "off",
          "no-restricted-imports": "off",
        },
      },
    ],
  },
  staged: {
    // Lint only code: the linter fails when every staged file is JSON, Markdown or YAML.
    "*.{ts,cts,mts,js,cjs,mjs}": ["vp fmt --write", "vp lint --fix"],
    "*.{json,md,yml,yaml}": ["vp fmt --write"],
  },
  test: {
    include: ["src/**/tests/**/*.test.ts", "test/**/*.test.ts"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      exclude: ["src/**/tests/**"],
      reporter: ["text", "lcov"],
      thresholds: { statements: 100, branches: 100, functions: 100, lines: 100 },
    },
  },
};

export default defineConfig(config);
