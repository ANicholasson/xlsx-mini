# Bytes out as Uint8Array; downloading is a separate browser entry

The writer returns a `Uint8Array`, which feeds `fs.writeFile`, a `Response` body, `new Blob([bytes])` and `Buffer.from` without conversion. A `downloadXlsx(bytes, filename)` helper lives in the `xlsx-mini/browser` entry, the only module allowed to touch the DOM; the core never auto-downloads and there is no Node file-system helper. The MIME type is exported as `XLSX_MIME_TYPE` for response headers.
