/**
 * The one workbook every runtime writes. Plain JavaScript, so Bun, Deno, workerd and the edge VM
 * run it as it is. The hosts hand in the package as they loaded it: by name through the exports
 * map in Node, Bun and Deno, bundled from `dist/` for workerd and the edge VM.
 */
export const workbook = {
  sheets: [
    {
      name: "Runtime",
      columns: [
        { width: 24 },
        { width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
        { width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
      ],
      freeze: { rows: 1 },
      autoFilter: true,
      rows: [
        [
          { value: "Name", style: { bold: true } },
          { value: "Amount", style: { bold: true } },
          { value: "When", style: { bold: true } },
        ],
        ["Kari Nordmann", 1234.5, new Date(Date.UTC(2026, 10, 5, 9, 30))],
        ["Ola <Nordmann> & co", -42, new Date(Date.UTC(2026, 0, 1))],
        ["Ærlig «ekte» 🎉", true, null],
      ],
    },
  ],
};

async function* later(rows) {
  for (const row of rows) {
    await Promise.resolve();
    yield row;
  }
}

/** A value as JSON carries it: a `Date` as its instant, so that one read as a number is told apart. */
function plain(value) {
  return value instanceof Date ? { date: value.toISOString() } : value;
}

/** The sheets and the rows of every sheet, or the code of the refusal and whether it is an `XlsxReadError`. */
async function readBack(read, file) {
  try {
    const book = await read.openXlsx(file);
    const rows = [];
    for (let sheet = 0; sheet < book.sheets.length; sheet++) {
      for await (const { number, cells } of book.rows(sheet)) {
        rows.push([number, ...cells.map(plain)]);
      }
    }
    return { sheets: [...book.sheets], rows };
  } catch (error) {
    return { refused: error.code, coded: error instanceof read.XlsxReadError };
  }
}

/**
 * Writes the workbook twice, with STORE entries and with the default compression, streams it
 * twice, and probes the browser entry, which must load anywhere and refuse to download where
 * there is no document. Then reads every package back with the read entry, and the default one
 * from a `Blob` as well.
 */
export async function run(lib, browser, read) {
  const stored = await lib.writeXlsx(workbook, { compression: "none" });
  const deflated = await lib.writeXlsx(workbook);
  // The same workbook streamed with shared strings and stored blocks, its rows arriving late,
  // collected through a Response as a route handler's client would.
  const streamed = new Uint8Array(
    await new Response(
      lib.streamXlsx(
        { sheets: [{ ...workbook.sheets[0], rows: later(workbook.sheets[0].rows) }] },
        { strings: "shared", compression: "none" },
      ),
    ).arrayBuffer(),
  );
  const streamedDeflated = new Uint8Array(
    await new Response(lib.streamXlsx(workbook)).arrayBuffer(),
  );
  let noDom;
  try {
    browser.downloadXlsx(stored, "runtime.xlsx");
  } catch (error) {
    noDom = error instanceof lib.XlsxMiniError ? error.code : String(error);
  }
  return {
    stored: base64(stored),
    deflated: base64(deflated),
    streamed: base64(streamed),
    streamedDeflated: base64(streamedDeflated),
    mime: browser.XLSX_MIME_TYPE,
    noDom,
    read: {
      stored: await readBack(read, stored),
      deflated: await readBack(read, deflated),
      streamed: await readBack(read, streamed),
      streamedDeflated: await readBack(read, streamedDeflated),
      blob: await readBack(read, new Blob([deflated])),
    },
  };
}

function base64(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}
