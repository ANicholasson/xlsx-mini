// examples/snippets/next-route.ts
// app/api/report/route.ts
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export async function GET(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="report.xlsx"',
    },
  });
}
