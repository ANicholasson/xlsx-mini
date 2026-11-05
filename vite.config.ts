import { type ViteUserConfig, defineConfig } from "vite-plus";

const CORE_ONLY_MESSAGE =
  "The core runs everywhere: web-standard APIs only. Node and DOM globals belong in runtime-specific entries or tests.";

const config: ViteUserConfig = {
  // Snippets and examples import the package by name; tests run them against the source.
  resolve: {
    alias: [
      {
        find: "xlsx-mini/browser",
        replacement: new URL("./src/browser.ts", import.meta.url).pathname,
      },
      { find: "xlsx-mini", replacement: new URL("./src/index.ts", import.meta.url).pathname },
    ],
  },
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
