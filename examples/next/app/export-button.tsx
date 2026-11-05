"use client";

// A client component: calls the server action, then saves the bytes through the browser.
import { downloadXlsx } from "xlsx-mini/browser";
import { exportOrders } from "./actions";

export function ExportButton() {
  return (
    <button type="button" onClick={async () => downloadXlsx(await exportOrders(), "orders.xlsx")}>
      Export orders (server action)
    </button>
  );
}
