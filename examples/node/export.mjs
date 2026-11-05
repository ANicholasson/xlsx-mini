// A plain Node script: `pnpm build && node examples/node/export.mjs` writes orders.xlsx here.
import { writeFile } from "node:fs/promises";
import { sheetFromObjects, writeXlsx } from "xlsx-mini";

const orders = [
  {
    id: "A-1001",
    customer: "Kari Nordmann",
    total: 1234.5,
    placed: new Date("2026-11-01T09:30:00Z"),
  },
  { id: "A-1002", customer: "Ola Nordmann", total: 99, placed: new Date("2026-11-02T13:00:00Z") },
  { id: "A-1003", customer: "Ingrid Hansen", total: 250, placed: new Date("2026-11-03T16:45:00Z") },
];

const sheet = sheetFromObjects(orders, {
  name: "Orders",
  columns: [
    { key: "id", header: "Order", width: 12 },
    { key: "customer", header: "Customer", width: 24 },
    { key: "total", header: "Total", width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
    { key: "placed", header: "Placed", width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
  ],
});

const bytes = await writeXlsx({ sheets: [sheet] });
await writeFile("orders.xlsx", bytes);
console.log(`orders.xlsx: ${bytes.byteLength} bytes, ${orders.length} orders`);
