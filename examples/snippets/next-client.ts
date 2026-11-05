// examples/snippets/next-client.ts
// app/export-button.tsx: wire this to a button's onClick
"use client";

import { downloadXlsx } from "xlsx-mini/browser";
import { exportReport } from "./next-action.js";

export async function onExportClick(): Promise<void> {
  downloadXlsx(await exportReport(), "report.xlsx");
}
