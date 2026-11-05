// examples/snippets/next-edge.ts
// app/api/report/route.ts on the edge runtime: the same handler, one line more
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export const runtime = "edge";

export async function GET(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, { headers: { "Content-Type": XLSX_MIME_TYPE } });
}
