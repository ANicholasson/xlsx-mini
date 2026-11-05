// examples/snippets/quick-start.ts
import { writeFile } from "node:fs/promises";
import { writeXlsx } from "xlsx-mini";

export default async function quickStart(path = "orders.xlsx"): Promise<void> {
  const bytes = await writeXlsx({
    sheets: [
      {
        name: "Orders",
        rows: [
          ["Order", "Customer", "Total", "Paid"],
          ["A-1001", "Kari Nordmann", 1234.5, true],
          ["A-1002", "Ola Nordmann", 99, false],
        ],
      },
    ],
  });
  await writeFile(path, bytes);
}
