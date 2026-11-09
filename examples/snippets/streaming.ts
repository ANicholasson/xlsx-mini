import { type Row, XLSX_MIME_TYPE, streamXlsx } from "xlsx-mini";

/** Rows arrive page by page from wherever they live; nothing is held beyond the page in hand. */
async function* orders(fetchPage: (page: number) => Promise<Row[]>): AsyncGenerator<Row> {
  yield ["Order", "Customer", "Total"];
  for (let page = 0; ; page++) {
    const rows = await fetchPage(page);
    if (rows.length === 0) return;
    yield* rows;
  }
}

/** The response a route handler returns; `fetchPage` is however the application reads a page. */
export default function ordersResponse(fetchPage: (page: number) => Promise<Row[]>): Response {
  const stream = streamXlsx({
    sheets: [{ name: "Orders", rows: orders(fetchPage), autoFilter: true }],
  });
  return new Response(stream, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="orders.xlsx"',
    },
  });
}
