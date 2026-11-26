import { openAsBlob } from "node:fs";
import { openXlsx } from "xlsx-mini/read";

export default async function sumOfTotals(path = "orders.xlsx"): Promise<number> {
  const book = await openXlsx(await openAsBlob(path));
  let sum = 0;
  for await (const { number, cells } of book.rows("Orders")) {
    const total = cells[2];
    if (number > 1 && typeof total === "number") sum += total;
  }
  return sum;
}
