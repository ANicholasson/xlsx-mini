import { type Row, writeXlsx } from "xlsx-mini";

type Order = { id: number; number: string; status: string; amount: number; cancelled: boolean };

const statuses = ["Open", "Shipped", "Closed"];

export default function orderSheet(orders: Order[]): Promise<Uint8Array> {
  const rows: Row[] = [
    ["Id", "Order", "Status", "Amount"],
    ...orders.map((order): Row => {
      const cells = [order.id, order.number, order.status, order.amount];
      // A cancelled order stays in the file and in the sums, out of sight.
      return order.cancelled ? { cells, hidden: true } : cells;
    }),
  ];
  return writeXlsx({
    sheets: [
      {
        name: "Orders",
        // The id is kept for a later import, not for the reader.
        columns: [{ hidden: true }, { width: 12 }, { width: 12 }, { width: 14 }],
        rows,
        validations: [{ range: "C2:C1000", list: `Lists!A1:A${statuses.length}` }],
      },
      // The dropdown reads its values from a sheet nobody needs to see.
      { name: "Lists", hidden: true, rows: statuses.map((status) => [status]) },
    ],
  });
}
