// app/export-button.tsx: wire this to a button's onClick
"use client";

import { downloadXlsx } from "xlsx-mini/browser";
import { exportReport } from "./actions";

export async function onExportClick(): Promise<void> {
  downloadXlsx(await exportReport(), "report.xlsx");
}
