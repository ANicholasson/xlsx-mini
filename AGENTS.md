# xlsx-mini

Zero-dependency TypeScript library that writes `.xlsx` files. Write-only; reading is out of scope.

## Never compromise

- Zero runtime dependencies and no install scripts.
- The repair dialog never appears. Every change keeps the golden fixtures opening clean in Excel, Numbers, Google Sheets and LibreOffice; `docs/compatibility.md` records each check.
- The core runs everywhere: web-standard APIs only. Node and DOM globals belong in runtime-specific entries and tests; lint enforces it.
- Output is deterministic: identical input, identical bytes.
- Performance is a design input, measured by benchmarks, with the hot-path rules in `CODING_STANDARDS.md`.

## Read before changing code

- `GLOSSARY.md`: the vocabulary for names, tests and issues.
- `docs/adr/`: decisions with their reasons. Contradicting one is a proposal to discuss, not a change to make.
- `CODING_STANDARDS.md`: the judgement calls reviewers apply.
- `src/README.md`: features are deep modules; import only through a feature's root files.

## Verify

- `pnpm verify` runs format, type-aware lint, typecheck, boundaries, knip, the tests with the coverage floor, and the package checks (build, publint, are-the-types-wrong). The pre-commit hook runs the staged subset and blocks what it cannot fix. The pull request checks run the same, split into jobs, on Node 20, 22 and 24 and on Windows.
- `pnpm validate` converts every golden fixture to CSV with LibreOffice and compares it with `test/fixtures/csv/`; the validate job does the same. A changed fixture regenerates its CSV in the same change.
- A change users notice gets a changeset (`pnpm changeset`). Merging the version pull request publishes through npm trusted publishing.
- README code lives in `examples/snippets/`, one file per block; an HTML comment `<!-- snippet: name -->` before the block names the file. `test/readme.test.ts` runs every file and fails when a block and its file differ. Edit the file, then paste it into the README.
- The runtime matrix in `test/runtimes/` builds `dist/` and writes a workbook in Node (ESM and CommonJS), Bun, Deno, workerd and the Vercel edge VM on every test run. `pnpm install` provides every host; nothing else to install.
- `pnpm bench` runs the three benchmark suites in `bench/` (wide table, many sheets, string-heavy) on Node; `pnpm bench:all` also on Bun and Deno; `--base <dist>` times a second build in turns with the first. On every pull request the workflow builds `main` alongside and posts one comment with the changes, measured on the same runner. `pnpm bench:memory` prints peak memory per suite. The budgets live in `docs/adr/0021-performance-baseline-and-budgets.md`. Time and memory budgets only ratchet down. The bundle-size budget is a ceiling: only a feature that needs the bytes raises it, in its own pull request, and records in that ADR what the bytes buy.
- Test at the lowest layer that fails; extend an existing case before adding one; a bug fix ships its failing test first.
- Changing a golden fixture means regenerating it from its input, opening it in the four applications, and updating the compatibility log in the same change.

## Agent skills

### Issue tracker

GitHub Issues in this repository, with native blocking edges between tickets. See `docs/agents/issue-tracker.md`.

### Triage labels

The default five-label vocabulary. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: `GLOSSARY.md` at the root and `docs/adr/`. See `docs/agents/domain.md`.
