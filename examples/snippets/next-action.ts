// examples/snippets/next-action.ts
// app/actions.ts
"use server";

import { writeXlsx } from "xlsx-mini";

export async function exportReport(): Promise<Uint8Array> {
  return writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
}
