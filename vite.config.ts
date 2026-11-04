import { defineConfig } from "vite-plus";

const CORE_ONLY_MESSAGE =
  "The core runs everywhere: web-standard APIs only. Node and DOM globals belong in runtime-specific entries or tests.";

export default defineConfig({
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
        files: ["src/**/tests/**", "test/**", "*.config.ts", "*.cjs"],
        rules: {
          "no-restricted-globals": "off",
          "no-restricted-imports": "off",
        },
      },
    ],
  },
  staged: {
    "*.{ts,cts,mts,js,cjs,mjs,json,md,yml,yaml}": ["vp fmt --write", "vp lint --fix"],
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
});
