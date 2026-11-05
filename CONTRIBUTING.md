# Contributing

Thank you for considering it. The project is small on purpose, and so is this guide.

## Before you start

- A new feature: open an issue first and say what you want to write to a workbook; a short discussion saves a long pull request.
- A bug: a pull request with a failing test is welcome straight away. If an application showed a repair dialog, the bug-report form asks for what we need.

## Working on it

```sh
pnpm install
pnpm verify
```

`pnpm verify` runs everything the pull request checks run: format, lint, typecheck, boundaries, knip, the tests with the coverage floor, and the package checks. The pre-commit hook runs the staged subset. `AGENTS.md` and `CODING_STANDARDS.md` say how the code is shaped; `GLOSSARY.md` gives the words.

## Opening a pull request

- Title in Conventional Commits form, such as `fix: keep a trailing space in cell text`; it becomes the commit on `main`.
- A change users notice gets a changeset: `pnpm changeset`, pick patch or minor, write one line a reader of the changelog understands.
- A change to a golden fixture opens it in Excel, Numbers, Google Sheets and LibreOffice and records the result in `docs/compatibility.md`.
- The benchmark comment on the pull request shows what the change did to performance; changes inside ten percent are noise.

Reviews are kind and concrete. If something is unclear, ask in the pull request.
