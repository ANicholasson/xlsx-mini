import { describe, expect, it } from "vite-plus/test";
import { validateWorkbook } from "../index.js";

const sheet = (name: string) => ({ name, rows: [] });

describe("validateWorkbook", () => {
  it("accepts ordinary names", () => {
    expect(() =>
      validateWorkbook({ sheets: [sheet("Orders 2026"), sheet("Ærlig"), sheet("a".repeat(31))] }),
    ).not.toThrow();
  });

  it("requires at least one sheet", () => {
    expect(() => validateWorkbook({ sheets: [] })).toThrow(
      expect.objectContaining({ code: "NO_SHEETS", path: "sheets" }),
    );
  });

  it("rejects names Excel refuses, with the path of the offending sheet", () => {
    const bad = [
      "",
      "a".repeat(32),
      "a/b",
      "a\\b",
      "a:b",
      "a*b",
      "a?b",
      "[a]",
      "'leading",
      "trailing'",
      "History",
      "history",
    ];
    for (const name of bad) {
      expect(() => validateWorkbook({ sheets: [sheet("ok"), sheet(name)] })).toThrow(
        expect.objectContaining({ code: "INVALID_SHEET_NAME", path: "sheets[1].name" }),
      );
    }
  });

  it("rejects duplicate names case-insensitively, at the second occurrence", () => {
    expect(() =>
      validateWorkbook({ sheets: [sheet("Data"), sheet("Other"), sheet("data")] }),
    ).toThrow(expect.objectContaining({ code: "DUPLICATE_SHEET_NAME", path: "sheets[2].name" }));
  });

  it("rejects more sheets than a zip archive can count", () => {
    const sheets = Array.from({ length: 0xffff - 5 }, (_, i) => ({ name: `S${i}`, rows: [] }));
    expect(() => validateWorkbook({ sheets })).toThrow(
      expect.objectContaining({ code: "PACKAGE_TOO_LARGE", path: "sheets" }),
    );
    expect(() => validateWorkbook({ sheets: sheets.slice(1) })).not.toThrow();
  });
});
