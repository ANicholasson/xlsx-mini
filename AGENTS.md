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

- `pnpm verify` runs format, type-aware lint, typecheck, boundaries, knip and the tests with the coverage floor. The pre-commit hook runs the staged subset and blocks what it cannot fix.
- Test at the lowest layer that fails; extend an existing case before adding one; a bug fix ships its failing test first.
- Changing a golden fixture means regenerating it from its input, opening it in the four applications, and updating the compatibility log in the same change.

## Agent skills

### Issue tracker

GitHub Issues in this repository, with native blocking edges between tickets. See `docs/agents/issue-tracker.md`.

### Triage labels

The default five-label vocabulary. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: `GLOSSARY.md` at the root and `docs/adr/`. See `docs/agents/domain.md`.
