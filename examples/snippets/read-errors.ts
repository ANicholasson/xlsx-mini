import { XlsxReadError, openXlsx } from "xlsx-mini/read";

export default async function countRows(upload: Blob): Promise<number | string> {
  let count = 0;
  try {
    const book = await openXlsx(upload);
    for await (const _row of book.rows(0)) count++;
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    switch (error.code) {
      case "NOT_XLSX":
        return "This is not an .xlsx file.";
      case "LIMIT_EXCEEDED":
        return `This file asks for more than limits.${error.limit} allows.`;
      case "INVALID_PACKAGE":
      case "INVALID_XML":
      case "INVALID_WORKBOOK":
      case "INVALID_CELL":
        // A damaged file. Another runtime can refuse it with another of these four codes.
        return error.cell === undefined
          ? "This file is damaged."
          : `This file is damaged at cell ${error.cell} of ${error.sheet}.`;
      default:
        // NO_INFLATE, UNKNOWN_SHEET and INVALID_LIMIT are about the runtime or the code.
        throw error;
    }
  }
  return count;
}
