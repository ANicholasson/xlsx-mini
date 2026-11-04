import { describe, expect, it } from "vite-plus/test";
import { XlsxMiniError } from "../index.js";

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
});
