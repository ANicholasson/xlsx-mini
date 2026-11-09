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

/**
 * Writes the workbook twice, with STORE entries and with the default compression, streams it
 * twice, and probes the browser entry, which must load anywhere and refuse to download where
 * there is no document.
 */
export async function run(lib, browser) {
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
  };
}

function base64(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}
