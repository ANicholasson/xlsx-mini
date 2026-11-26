import { openAsBlob } from "node:fs";
import { openXlsx } from "xlsx-mini/read";

export default async function sumOfTotals(path = "orders.xlsx"): Promise<number> {
  // A Blob is read in slices, never whole.
  const book = await openXlsx(await openAsBlob(path));
  let sum = 0;
  // Each row that holds a value: its number in the sheet from 1, its cells from column A.
  for await (const { number, cells } of book.rows("Orders")) {
    const total = cells[2];
    if (number > 1 && typeof total === "number") sum += total;
  }
  return sum;
}
