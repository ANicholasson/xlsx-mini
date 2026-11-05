// examples/snippets/response.ts
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export default async function report(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="report.xlsx"',
    },
  });
}
