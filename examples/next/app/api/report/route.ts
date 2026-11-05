// A route handler: GET /api/report returns the workbook as a download.
import { XLSX_MIME_TYPE, sheetFromObjects, writeXlsx } from "xlsx-mini";
import { columns, orders } from "../../orders";

export async function GET(): Promise<Response> {
  const bytes = await writeXlsx({
    sheets: [sheetFromObjects(orders, { name: "Orders", columns })],
  });
  return new Response(bytes, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="orders.xlsx"',
    },
  });
}
