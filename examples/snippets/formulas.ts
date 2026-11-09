import { type CellStyle, writeXlsx } from "xlsx-mini";

const link: CellStyle = { fontColor: "#0563c1", underline: true };

export default function formulasAndLinks(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Budget",
        rows: [
          ["Item", "Qty", "Price", "Total"],
          ["Desk", 2, 4999, { formula: "B2*C2" }],
          ["Chair", 4, 1299.5, { formula: "B3*C3" }],
          // A cached value shows until the application recalculates; it is optional.
          ["Sum", { formula: "SUM(B2:B3)", value: 6 }, null, { formula: "SUM(D2:D3)" }],
          [],
          [{ value: "Price list", href: "https://example.com/prices", style: link }],
          [{ value: "Go to the summary", sheet: "Summary", cell: "A1", style: link }],
        ],
      },
      { name: "Summary", rows: [["Grand total", { formula: "Budget!D4" }]] },
    ],
  });
}
