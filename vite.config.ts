import codspeedPlugin from "@codspeed/vitest-plugin";
import { type ViteUserConfig, defineConfig } from "vite-plus";

type VitePlugin = NonNullable<ViteUserConfig["plugins"]>[number];

const CORE_ONLY_MESSAGE =
  "The core runs everywhere: web-standard APIs only. Node and DOM globals belong in runtime-specific entries or tests.";

// Active only under the CodSpeed runner in CI; locally the benchmarks run through tinybench.
// Cast: the plugin carries the `Plugin` type of the Vite it was built against, which TypeScript
// cannot finish comparing with Vite+'s own.
const plugins = [codspeedPlugin() as unknown as VitePlugin];

const config: ViteUserConfig = {
  plugins,
  lint: {
    options: { typeAware: true, typeCheck: true },
    rules: {
      "no-restricted-globals": [
        "error",
        { name: "process", message: CORE_ONLY_MESSAGE },
        { name: "Buffer", message: CORE_ONLY_MESSAGE },
        { name: "require", message: CORE_ONLY_MESSAGE },
        { name: "document", message: CORE_ONLY_MESSAGE },
        { name: "window", message: CORE_ONLY_MESSAGE },
        { name: "navigator", message: CORE_ONLY_MESSAGE },
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
        rules: { "no-restricted-globals": "off" },
      },
      {
        // The benchmark fixtures preallocate their arrays on purpose.
        files: ["bench/**"],
        rules: { "unicorn/no-new-array": "off" },
      },
      {
        files: ["src/**/tests/**", "test/**", "bench/**", "*.config.ts", "*.cjs"],
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
    benchmark: { include: ["bench/**/*.bench.ts"] },
    // The benchmarks load the built package natively: through the module runner every export
    // becomes a getter, which a million-cell write would pay on each hot call.
    server: { deps: { external: [/\/node_modules\//, /\/dist\//] } },
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
