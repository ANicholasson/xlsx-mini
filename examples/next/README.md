# xlsx-mini in Next.js

Two exports of the same orders: a route handler that returns the workbook as a response, and a
server action whose bytes a client button saves through the browser.

```sh
pnpm build                 # at the repository root: the example links to the local package
cd examples/next
pnpm install
pnpm dev                   # http://localhost:3000
```

In your own application, depend on `xlsx-mini` from npm instead of the local link.
