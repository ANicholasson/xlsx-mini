import { type CellStyle, writeXlsx } from "xlsx-mini";

const title: CellStyle = {
  bold: true,
  fontSize: 16,
  fontColor: "#ffffff",
  fill: "#1f4e79",
  align: "center",
  verticalAlign: "middle",
};
const heading: CellStyle = { bold: true, fill: "#d9e1f2", border: "thin" };
const total: CellStyle = { bold: true, border: { top: "medium", bottom: "thick" } };

export default function appearance(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Invoice",
        columns: [
          { width: 14 },
          { width: 36 },
          { width: 8, style: { align: "right" } },
          { width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
        ],
        // The title spans four columns; its row and the wrapped row get a height in points.
        merges: ["A1:D1"],
        rowHeights: { 1: 28, 3: 48 },
        rows: [
          [{ value: "Invoice 2026-1042", style: title }],
          [
            { value: "Item", style: heading },
            { value: "Description", style: heading },
            { value: "Qty", style: heading },
            { value: "Amount", style: heading },
          ],
          [
            "Desk",
            {
              value: "Height-adjustable desk, delivered and assembled on site",
              style: { wrap: true, verticalAlign: "top" },
            },
            2,
            9998,
          ],
          ["Chair", { value: "Mesh office chair", style: { indent: 1 } }, 4, 5198],
          [
            { value: "Total", style: total },
            // A styled cell without a value: the rule runs across the blank.
            { value: null, style: total },
            { value: 6, style: total },
            { value: 15196, style: total },
          ],
          [],
          [{ value: "Payment within 30 days.", style: { italic: true, fontColor: "#808080" } }],
        ],
      },
    ],
  });
}
