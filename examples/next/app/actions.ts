"use server";

// A server action: the workbook is built on the server and the bytes go to the client.
import { sheetFromObjects, writeXlsx } from "xlsx-mini";
import { columns, orders } from "./orders";

export async function exportOrders(): Promise<Uint8Array> {
  return writeXlsx({ sheets: [sheetFromObjects(orders, { name: "Orders", columns })] });
}
