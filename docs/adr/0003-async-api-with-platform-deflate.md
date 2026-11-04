# One async entry point, DEFLATE through the platform, STORE as the deterministic option

`writeXlsx` is asynchronous because the only zero-dependency way to compress on every runtime is the web `CompressionStream`, which is async. Compression uses `deflate-raw` by default and is feature-detected by constructing the stream inside a try/catch, since some runtimes list a format they do not implement; when unavailable the writer falls back to STORE entries. `compression: "none"` writes STORE on purpose, which is byte-deterministic and what golden fixtures use.

## Considered options

- Sync-only STORE: simplest, but files five to ten times larger.
- A Node-only sync DEFLATE through `node:zlib`: would need a separate Node build and contradicts the single web-standard core.
