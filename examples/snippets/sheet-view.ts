import { writeXlsx } from "xlsx-mini";

export default function report(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Report",
        // A report reads better without the grid, a little larger.
        view: { gridlines: false, zoom: 120 },
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 120000],
          ["Q2", 135500],
        ],
      },
      {
        name: "فواتير",
        // Column A sits at the right edge; the frozen header row is part of the same view.
        view: { rightToLeft: true },
        freeze: { rows: 1 },
        rows: [
          ["الطلب", "المبلغ"],
          ["A-1001", 1234.5],
        ],
      },
    ],
  });
}
