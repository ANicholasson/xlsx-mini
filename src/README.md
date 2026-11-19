# Features are deep modules

Every folder directly under `src/` is a feature: a lot of behaviour behind a small interface. Its public surface is its root files; everything in a subfolder is private.

```
src/
  index.ts              public API of the package
  write-xlsx.ts         application service: render parts, then package them
  browser.ts            browser entry: downloadXlsx, the one file allowed to touch the DOM
  read.ts               reader entry: xlsx-mini/read
  <feature>/
    index.ts            entry point: the only way in from outside
    read.ts             the feature's read side: what the reader needs of it
    lib/                implementation, free to import its siblings
    read/               the read side's implementation
    tests/              the feature's tests and fixtures
```

Five rules, enforced by `pnpm lint:boundaries` (dependency-cruiser, config in `.dependency-cruiser.cjs`):

1. Code outside a feature imports only that feature's root files, never anything in its subfolders.
2. A feature's own files import each other freely, but reach other features only through their root files.
3. A feature's tests import it through its root files, like everyone else; they may also use their own `tests/` fixtures.
4. No dependency cycles. Type-only imports count.
5. Nothing outside the read side imports it. The read side is `src/read.ts` and every feature's `read.ts` and `read/`. Tests and benchmarks may import it, and it may import the write side under rules 1 and 2. So an application that imports only the writer bundles the same bytes whatever the reader holds, which the size test also checks (ADR 0032).

Layering on top: `xml/` and `opc/` are kernels that know nothing about spreadsheets and may not import a spreadsheet feature.

Prefer several small root files (`index.ts`, `browser.ts`) over one barrel that re-exports a subtree. Adding an entry point is adding a root file; adding a private file needs no configuration.
