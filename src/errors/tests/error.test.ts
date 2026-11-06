import { describe, expect, it } from "vite-plus/test";
import { XlsxMiniError, atPath } from "../index.js";

describe("XlsxMiniError", () => {
  it("carries a code and a path and reads like an Error", () => {
    const error = new XlsxMiniError(
      "INVALID_NUMBER",
      "sheets[0].rows[3][2]",
      "NaN cannot be written",
    );
    expect(error).toBeInstanceOf(Error);
    expect(error).toBeInstanceOf(XlsxMiniError);
    expect(error.name).toBe("XlsxMiniError");
    expect(error.code).toBe("INVALID_NUMBER");
    expect(error.path).toBe("sheets[0].rows[3][2]");
    expect(error.message).toBe("NaN cannot be written at sheets[0].rows[3][2]");
  });

  it("moves an error to a longer path, keeping its code and bare message", () => {
    const local = new XlsxMiniError(
      "INVALID_STYLE",
      "border.top",
      "Border styles are thin, medium or thick",
    );
    const moved = atPath(local, "sheets[2].columns[0].style");
    expect(moved).toBeInstanceOf(XlsxMiniError);
    expect(moved.code).toBe("INVALID_STYLE");
    expect(moved.path).toBe("sheets[2].columns[0].style.border.top");
    expect(moved.message).toBe(
      "Border styles are thin, medium or thick at sheets[2].columns[0].style.border.top",
    );
    const rootless = atPath(
      new XlsxMiniError("NO_SHEETS", "", "A workbook needs at least one sheet"),
      "workbook",
    );
    expect(rootless.path).toBe("workbook");
    expect(rootless.message).toBe("A workbook needs at least one sheet at workbook");
    const altered = new XlsxMiniError("NO_DOM", "x", "No document");
    altered.message = "rewritten";
    expect(atPath(altered, "y").message).toBe("rewritten at y.x");
  });
});
