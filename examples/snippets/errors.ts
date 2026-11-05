// examples/snippets/errors.ts
import { XlsxMiniError, writeXlsx } from "xlsx-mini";

export default async function safely(): Promise<string> {
  try {
    await writeXlsx({ sheets: [{ name: "Bad/Name", rows: [] }] });
    return "written";
  } catch (error) {
    // code says what, path says where: "INVALID_SHEET_NAME at sheets[0].name"
    if (error instanceof XlsxMiniError) return `${error.code} at ${error.path}`;
    throw error;
  }
}
