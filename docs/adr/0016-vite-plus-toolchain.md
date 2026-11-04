# Vite+ as the single development toolchain, TypeScript 7 for emit

Formatting, linting with type-aware rules, and tests come from one dev dependency, Vite+ (`vp`), which bundles oxfmt, oxlint and vitest; `tsc` from TypeScript 7 emits the package and runs the typecheck. knip finds dead exports and dependencies, dependency-cruiser enforces boundaries, fast-check drives property tests, changesets manages releases. A pre-commit hook through Vite+'s dispatcher formats and lints staged files and typechecks; CI is the real gate. Biome is the fallback if Vite+ proves too young.
