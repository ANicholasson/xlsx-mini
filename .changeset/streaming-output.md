---
"xlsx-mini": minor
---

`streamXlsx` writes a workbook as a web `ReadableStream` of the file's bytes, produced as the consumer reads them: a `Response` body, a file, or `for await` over the chunks. Rows may be an async iterable, so a database cursor can feed a sheet page by page, and peak memory stays flat however many rows go out. Text is written inline by default; `strings: "shared"` keeps the shared-strings table instead. The streamed package uses data descriptors and, where the runtime has no raw deflate, stored deflate blocks, so sequential zip readers can open it too. `writeXlsx` is unchanged, byte for byte, and now reports `ASYNC_ROWS` instead of a `TypeError` when given an async iterable.
