import { defineConfig } from "vite-plus";

export default defineConfig({
  test: {
    include: ["src/**/tests/**/*.test.ts", "test/**/*.test.ts"],
  },
});
