import { type ObjectColumn, sheetFromObjects, writeXlsx } from "xlsx-mini";
import { objectsFromRows, openXlsx } from "xlsx-mini/read";

interface Order {
  id: string;
  total: number;
  placed: Date;
}

const columns = [
  { key: "id", header: "Order" },
  { key: "total", header: "Total", style: { numberFormat: '#,##0.00 "kr"' } },
  { key: "placed", header: "Placed", style: { numberFormat: "yyyy-mm-dd" } },
] as const satisfies readonly ObjectColumn<Order>[];

export function exportOrders(orders: Order[]): Promise<Uint8Array> {
  return writeXlsx({ sheets: [sheetFromObjects(orders, { name: "Orders", columns })] });
}

export async function importOrders(file: Uint8Array | Blob): Promise<Order[]> {
  const book = await openXlsx(file);
  const orders: Order[] = [];
  for await (const { number, values } of objectsFromRows(book.rows("Orders"), columns)) {
    const { id, total, placed } = values;
    if (typeof id !== "string" || typeof total !== "number" || !(placed instanceof Date)) {
      throw new Error(`Row ${number} is not an order`);
    }
    orders.push({ id, total, placed });
  }
  return orders;
}
