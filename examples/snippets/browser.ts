import { writeXlsx } from "xlsx-mini";
import { downloadXlsx } from "xlsx-mini/browser";

export default async function exportInBrowser(): Promise<void> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  downloadXlsx(bytes, "report.xlsx");
}
