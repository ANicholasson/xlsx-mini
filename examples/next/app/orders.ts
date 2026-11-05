// The data both exports share; in an application this comes from a database.
import type { ObjectColumn } from "xlsx-mini";

export interface Order {
  id: string;
  customer: string;
  total: number;
  placed: Date;
}

export const orders: Order[] = [
  {
    id: "A-1001",
    customer: "Kari Nordmann",
    total: 1234.5,
    placed: new Date("2026-11-01T09:30:00Z"),
  },
  { id: "A-1002", customer: "Ola Nordmann", total: 99, placed: new Date("2026-11-02T13:00:00Z") },
  { id: "A-1003", customer: "Ingrid Hansen", total: 250, placed: new Date("2026-11-03T16:45:00Z") },
];

export const columns: ObjectColumn<Order>[] = [
  { key: "id", header: "Order", width: 12 },
  { key: "customer", header: "Customer", width: 24 },
  { key: "total", header: "Total", width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
  { key: "placed", header: "Placed", width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
];
