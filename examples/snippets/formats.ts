// examples/snippets/formats.ts
import { type CellStyle, writeXlsx } from "xlsx-mini";

const header: CellStyle = { bold: true };
const money: CellStyle = { numberFormat: '#,##0.00 "kr"' };
const percent: CellStyle = { numberFormat: "0.0%" };
const isoDate: CellStyle = { numberFormat: "yyyy-mm-dd" };

export default function formats(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Invoice",
        // A column style applies to every cell in the column; a cell's own style wins.
        columns: [undefined, { style: money }, { style: percent }, { style: isoDate }],
        rows: [
          [
            { value: "Item", style: header },
            { value: "Price", style: header },
            { value: "VAT", style: header },
            { value: "Due", style: header },
          ],
          ["Desk", 4999, 0.25, new Date(Date.UTC(2026, 11, 1))],
          [
            "Chair",
            { value: 1299.5, style: { numberFormat: "0" } },
            0.25,
            new Date(Date.UTC(2026, 11, 15)),
          ],
        ],
      },
    ],
  });
}
