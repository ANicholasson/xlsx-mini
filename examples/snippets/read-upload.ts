import { XlsxReadError, objectsFromRows, openXlsx } from "xlsx-mini/read";

const columns = [
  { key: "sku", header: "SKU" },
  { key: "counted", header: "Counted" },
] as const;

/** A route handler that takes a stock count of a few thousand rows, uploaded from a form. */
export async function POST(request: Request): Promise<Response> {
  const file = (await request.formData()).get("file");
  if (!(file instanceof Blob)) return new Response("Choose a file.", { status: 400 });
  const counts = new Map<string, number>();
  try {
    // Far below the defaults, which are sized for any file a person uploads.
    const limits = { inflatedBytes: 8 * 1024 * 1024, cellPositions: 50_000 };
    const book = await openXlsx(file, { limits });
    for await (const { number, values } of objectsFromRows(book.rows(0), columns)) {
      if (typeof values.sku !== "string" || typeof values.counted !== "number") {
        return new Response(`Row ${number} needs a SKU and a count.`, { status: 422 });
      }
      counts.set(values.sku, values.counted);
    }
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    return new Response(`The file was refused: ${error.code}.`, { status: 422 });
  }
  return Response.json({ counted: counts.size });
}
