import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

/** Any response with setHeader and end: Express, Node's http, Fastify's raw reply. */
interface Reply {
  setHeader(name: string, value: string): void;
  end(body: Uint8Array): void;
}

export default async function handler(reply: Reply): Promise<void> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  reply.setHeader("Content-Type", XLSX_MIME_TYPE);
  reply.setHeader("Content-Disposition", 'attachment; filename="report.xlsx"');
  reply.end(bytes);
}
