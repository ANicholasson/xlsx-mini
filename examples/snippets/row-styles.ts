import { type CellStyle, type Row, writeXlsx } from "xlsx-mini";

const header: CellStyle = { bold: true, fill: "#1f2937", fontColor: "#ffffff" };
const band: CellStyle = { fill: "#f3f4f6" };
const total: CellStyle = { bold: true, border: { top: "thin" } };

type Order = [id: string, customer: string, amount: number];

export default function report(orders: Order[]): Promise<Uint8Array> {
  const sum = orders.reduce((soFar, [, , amount]) => soFar + amount, 0);
  const rows: Row[] = [
    { cells: ["Order", "Customer", "Amount"], style: header, height: 24 },
    // Every second order on a band; the others stay plain arrays.
    ...orders.map((order, index): Row => (index % 2 === 1 ? { cells: order, style: band } : order)),
    // The rule above the total runs across the empty cell.
    { cells: ["Total", null, sum], style: total },
  ];
  return writeXlsx({
    sheets: [
      {
        name: "Report",
        columns: [{ width: 12 }, { width: 24 }, { width: 14, style: { numberFormat: "#,##0.00" } }],
        rows,
      },
    ],
  });
}
