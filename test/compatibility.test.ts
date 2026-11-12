import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const text = (file: string) => readFileSync(join(root, file), "utf8").replaceAll("\r\n", "\n");

/**
 * The row the README's compatibility table gives each golden fixture of the compatibility log, in
 * the order of both. A fixture added to the log needs its row here and there.
 */
const ROWS: Readonly<Record<string, string>> = {
  "minimal.xlsx": "Strings and numbers",
  "cell-types.xlsx": "Every kind of cell value",
  "minimal-deflate.xlsx": "Deflated entries",
  "multi-sheet.xlsx": "Several sheets, escaped names",
  "styles.xlsx": "Styles and number formats",
  "layout.xlsx": "Widths, frozen panes, autofilter",
  "orders.xlsx": "Built from objects",
  "appearance.xlsx": "Fonts, fills, borders, merges",
  "streamed.xlsx": "Streamed, descriptors, inline",
  "formulas.xlsx": "Formulas, cached and not",
  "hyperlinks.xlsx": "Hyperlinks, external and within",
  "validation.xlsx": "Data validation",
  "auto-width.xlsx": "Auto column widths",
  "images.xlsx": "Image cells",
  "print.xlsx": "Print setup",
};

describe("the README's compatibility table", () => {
  it("has a row for every golden fixture in the compatibility log", () => {
    const logged = [...text("docs/compatibility.md").matchAll(/^## ([\w-]+\.xlsx) /gm)].map(
      (match) => match[1]!,
    );
    expect(logged).toEqual(Object.keys(ROWS));

    const table = /\| Fixture [^\n]*\n\|[- :|]+\n((?:\|[^\n]*\n)+)/.exec(text("README.md"));
    const rows = table![1]!
      .trimEnd()
      .split("\n")
      .map((line) => line.split("|")[1]!.trim());
    expect(rows).toEqual(Object.values(ROWS));
  });
});
