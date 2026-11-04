# Coding standards

Judgement calls a reviewer applies. Anything a tool can check lives in the tool instead: formatting and lint in `vp check`, module boundaries in `pnpm lint:boundaries`, dead code in `pnpm knip`, the coverage floor in `pnpm test:coverage`.

## Deep modules

A feature's public surface is its root files; everything under `lib/` and `tests/` is private. Prefer one deep function behind a small type to several shallow helpers that leak their mechanics. A new export must carry its weight: a test, a use, and a sentence in the README if it is public.

## Names come from the glossary

Use the terms in `GLOSSARY.md` for types, functions, tests and issues; when a concept is missing there, add it before naming code after it. No abbreviations in public names.

## No speculative generality

An option, hook, type parameter or extension point exists only once a test and a README example use it. Reserve nothing "for later"; later adds it in one commit.

## Hot path

Cell normalisation, escaping, interning, style lookup and part assembly run once per cell or once per row. On that path: one pass over the rows, no regular expressions, no closures created per cell, no intermediate objects per cell, strings built in chunks and encoded once per part. A change that touches it states its benchmark delta in the PR.

## Errors

Every error thrown to a caller is the project's one error class, with a code from the documented union and a path such as `sheets[1].rows[40][2]`. Validation fails on the first problem. The only silent fix is stripping characters XML cannot carry; everything else throws.

## Types

Public contracts are interfaces so consumers can augment them. Discriminated unions over boolean flags. Inputs are `readonly`. Literal tables use `satisfies`. No enums, no classes except the error, no `any`. A non-null assertion is acceptable only on an index the surrounding loop proves in range, with the proof in a comment.

## Tests

Test at the lowest layer that fails. Extend an existing case before adding a new one. Every bug fix ships the test that failed first. Golden fixtures change only together with their input and an entry in the compatibility log. Expected values come from an independent source (a published check value, a reader that is not ours), never from the code under test.

## Documentation

Every public symbol has a TSDoc comment with one example. README snippets are executed by a test. Internal docs record decisions and constraints that code cannot show; they do not restate code.

## Commits and pull requests

Conventional commit subjects in plain language, imperative mood, intent over mechanics. One underlying problem per pull request. The PR body says what was verified and how, and names what could not be verified.
