import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Order {
  id: string;
  customer: string;
  total: number;
  placed: Date;
}

export default function fromObjects(orders: Order[]): Promise<Uint8Array> {
  const sheet = sheetFromObjects(orders, {
    name: "Orders",
    columns: [
      { key: "id", header: "Order", width: 12 },
      { key: "customer", header: "Customer", width: 24 },
      { key: "total", header: "Total", width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
      { key: "placed", header: "Placed", width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
