# Features are deep modules

Every folder directly under `src/` is a feature: a lot of behaviour behind a small interface. Its public surface is its root files; everything in a subfolder is private.

```
src/
  index.ts              public API of the package
  write-xlsx.ts         application service: render parts, then package them
  browser.ts            browser entry: downloadXlsx, the one file allowed to touch the DOM
  <feature>/
    index.ts            entry point: the only way in from outside
    lib/                implementation, free to import its siblings
    tests/              the feature's tests and fixtures
```

Four rules, enforced by `pnpm lint:boundaries` (dependency-cruiser, config in `.dependency-cruiser.cjs`):

1. Code outside a feature imports only that feature's root files, never anything in its subfolders.
2. A feature's own files import each other freely, but reach other features only through their root files.
3. A feature's tests import it through its root files, like everyone else; they may also use their own `tests/` fixtures.
4. No dependency cycles. Type-only imports count.

Layering on top: `xml/` and `opc/` are kernels that know nothing about spreadsheets and may not import a spreadsheet feature.

Prefer several small root files (`index.ts`, `browser.ts`) over one barrel that re-exports a subtree. Adding an entry point is adding a root file; adding a private file needs no configuration.
