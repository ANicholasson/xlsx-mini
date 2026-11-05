import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";

/**
 * LibreOffice, an independent reader, converts every golden fixture to CSV headlessly, and the
 * result must match the committed CSV cell for cell. Runs with `XLSX_MINI_VALIDATE=1`, as the
 * validate job does; `pnpm validate` sets it.
 */
const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");
const expectedDir = join(fixturesDir, "csv");

/** Comma, double quotes, UTF-8, formatted as shown, and every sheet as `<fixture>-<sheet>.csv`. */
const CSV_FILTER = "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1";

function findSoffice(): string | undefined {
  const candidates = [
    process.env["SOFFICE"],
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/lib/libreoffice/program/soffice",
  ];
  return candidates.find((path): path is string => path !== undefined && existsSync(path));
}

const enabled = process.env["XLSX_MINI_VALIDATE"] === "1";
const soffice = enabled ? findSoffice() : undefined;
if (enabled && soffice === undefined) {
  throw new Error("XLSX_MINI_VALIDATE=1 but LibreOffice was not found; set SOFFICE to its path");
}

/** LibreOffice names a CSV after the sheet; characters Windows forbids in a file name become `_`. */
function portable(name: string): string {
  return name.replace(/[<>:"/\\|?*]/g, "_");
}

const expectedFiles = readdirSync(expectedDir)
  .filter((name) => name.endsWith(".csv"))
  .sort();

describe.skipIf(!enabled)("LibreOffice reads every golden fixture", () => {
  let outDir: string;
  let conversion: string | undefined;

  beforeAll(() => {
    outDir = mkdtempSync(join(tmpdir(), "xlsx-mini-validate-"));
    const fixtures = readdirSync(fixturesDir)
      .filter((name) => name.endsWith(".xlsx"))
      .map((name) => join(fixturesDir, name));
    try {
      conversion = execFileSync(
        soffice!,
        ["--headless", "--convert-to", CSV_FILTER, "--outdir", outDir, ...fixtures],
        { encoding: "utf8", stdio: "pipe", timeout: 180_000 },
      );
    } catch (error) {
      conversion = `conversion failed: ${(error as { stderr?: string; message: string }).stderr ?? (error as Error).message}`;
    }
  }, 200_000);

  afterAll(() => rmSync(outDir, { recursive: true, force: true }));

  it("converts every fixture without an error", () => {
    expect(conversion).not.toMatch(/^conversion failed|Error/);
  });

  /** The produced file for an expected name, through the portable mapping. */
  const producedFor = (expected: string): string | undefined =>
    readdirSync(outDir).find((name) => portable(name) === expected);

  it("produces one CSV per sheet of every fixture", () => {
    const produced = readdirSync(outDir)
      .filter((name) => name.endsWith(".csv"))
      .map(portable)
      .sort();
    expect(produced).toEqual(expectedFiles);
  });

  it.each(expectedFiles)("%s matches cell for cell", (name) => {
    const produced = producedFor(name);
    expect(produced, `no CSV produced for ${name}`).toBeDefined();
    expect(readFileSync(join(outDir, produced!), "utf8")).toBe(
      readFileSync(join(expectedDir, name), "utf8"),
    );
  });
});
