import { XlsxMiniError } from "../../errors/index.js";
import type { StreamedWorkbook } from "./model.js";

const MAX_NAME_LENGTH = 31;
const FORBIDDEN_CHARACTERS = /[[\]:*?/\\]/;
const RESERVED_NAME = "history";
/** A zip archive counts its entries in 16 bits, and six parts are not sheets. */
const MAX_SHEETS = 0xffff - 6;

/**
 * Throws the first problem Excel would refuse: no sheets, more sheets than the package can hold,
 * a bad sheet name, a duplicate name.
 */
export function validateWorkbook(workbook: StreamedWorkbook): void {
  const sheets = workbook.sheets;
  if (sheets.length === 0) {
    throw new XlsxMiniError("NO_SHEETS", "sheets", "A workbook needs at least one sheet");
  }
  if (sheets.length > MAX_SHEETS) {
    throw new XlsxMiniError(
      "PACKAGE_TOO_LARGE",
      "sheets",
      `A package holds at most ${MAX_SHEETS} sheets`,
    );
  }
  const seen = new Set<string>();
  for (let i = 0; i < sheets.length; i++) {
    const path = `sheets[${i}].name`;
    const name = sheets[i]!.name;
    const problem = sheetNameProblem(name);
    if (problem !== undefined) throw new XlsxMiniError("INVALID_SHEET_NAME", path, problem);
    const key = name.toLowerCase();
    if (seen.has(key)) {
      throw new XlsxMiniError(
        "DUPLICATE_SHEET_NAME",
        path,
        `Sheet name "${name}" is already used; Excel compares names case-insensitively`,
      );
    }
    seen.add(key);
  }
}

function sheetNameProblem(name: unknown): string | undefined {
  if (typeof name !== "string" || name.length === 0) return "Sheet name must be a non-empty string";
  if (name.length > MAX_NAME_LENGTH)
    return `Sheet name "${name}" is longer than ${MAX_NAME_LENGTH} characters`;
  if (FORBIDDEN_CHARACTERS.test(name)) return `Sheet name "${name}" contains one of [ ] : * ? / \\`;
  if (name.startsWith("'") || name.endsWith("'"))
    return `Sheet name "${name}" cannot start or end with an apostrophe`;
  if (name.toLowerCase() === RESERVED_NAME) return `Sheet name "${name}" is reserved by Excel`;
  return undefined;
}
