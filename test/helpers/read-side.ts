import { readdirSync } from "node:fs";
import { createRequire } from "node:module";
import { join } from "node:path";

interface Rule {
  readonly name: string;
  readonly to: { readonly path?: string | readonly string[] };
}

export function readSideSources(root: string): string[] {
  const config = createRequire(import.meta.url)(join(root, ".dependency-cruiser.cjs")) as {
    readonly forbidden: readonly Rule[];
  };
  const rule = config.forbidden.find(({ name }) => name === "write-side-never-imports-read-side");
  if (rule?.to.path === undefined) throw new Error("no read side in .dependency-cruiser.cjs");
  const patterns = [rule.to.path].flat().map((pattern) => new RegExp(pattern));
  return readdirSync(join(root, "src"), { recursive: true, encoding: "utf8" })
    .map((name) => `src/${name.replaceAll("\\", "/")}`)
    .filter((file) => /\.[cm]?ts$/.test(file) && patterns.some((pattern) => pattern.test(file)))
    .sort();
}
