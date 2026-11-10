# Performance baseline and budgets

The first measured baseline is the budget, and a budget only ratchets down. Numbers come from `pnpm bench` (the three suites in `bench/`, deflate and STORE each, seven runs after one warm-up against the built package loaded by the host itself) and `pnpm bench:memory` (one write per suite in a fresh process). On pull requests the same script runs on GitHub's runners in Node, Bun and Deno with `main` built alongside, and the report compares the two on that one machine.

## Baseline, 2026-11-05, Apple M5 Pro, 48 GiB, Node 24.22.1

| Suite                                                       | Output (DEFLATE) | Mean                   | Margin of error | Budget                 |
| ----------------------------------------------------------- | ---------------- | ---------------------- | --------------- | ---------------------- |
| Wide table: 100 000 rows by 10 mixed columns                | 6.4 MiB          | 514 ms (STORE: 177 ms) | ±3 % (±14 %)    | 520 ms (STORE: 180 ms) |
| Many sheets: 40 sheets of 2 500 rows by 6 columns           | 3.9 MiB          | 245 ms                 | ±6 %            | 250 ms                 |
| String-heavy: 200 000 rows by 5 text columns, 2 000 phrases | 5.1 MiB          | 389 ms                 | ±5 %            | 390 ms                 |

Peak memory is the growth of the process's resident set during the write, measured after the input workbook is built and garbage has been collected, so it is the writer's own cost:

| Suite        | Output  | Peak growth | Growth over output | Budget |
| ------------ | ------- | ----------- | ------------------ | ------ |
| Wide table   | 6.4 MiB | 159 MiB     | 25×                | 30×    |
| Many sheets  | 3.9 MiB | 95 MiB      | 24×                | 30×    |
| String-heavy | 5.1 MiB | 132 MiB     | 26×                | 30×    |

## Streamed baseline, 2026-11-06, Apple M5 Pro, 48 GiB, Node 24.22.1

`streamXlsx` with its defaults, inline strings and raw deflate, into a sink that discards the bytes; `pnpm bench` reports it as the `stream` variant next to `deflate` and `store`, and the pull-request comment shows it without a base until `main` streams too.

| Suite                  | Output  | Minimum of 7 | Against `deflate` |
| ---------------------- | ------- | ------------ | ----------------- |
| Wide table, streamed   | 7.2 MiB | 522 ms       | same              |
| Many sheets, streamed  | 4.0 MiB | 248 ms       | 12 % faster       |
| String-heavy, streamed | 5.0 MiB | 913 ms       | 2.2× slower       |

Inline strings cost one escape and one write per text cell instead of one per distinct string, and the deflater sees every repetition; on the string-heavy suite, 1 000 000 cells drawn from 2 000 phrases, that doubles the time. `strings: "shared"` restores the table for such data when memory allows.

Peak memory when streaming is the heap still in use after a collection, sampled every 8 MiB of output, because the resident set grows with V8's appetite for heap rather than with what is held; `pnpm bench:memory` prints both:

| Suite                      | Output   | Retained heap, peak | Resident growth | Budget |
| -------------------------- | -------- | ------------------- | --------------- | ------ |
| Wide table, 100 000 rows   | 7.2 MiB  | 9.0 MiB             | 39 MiB          | 12 MiB |
| Wide table, 1 000 000 rows | 72.1 MiB | 10.1 MiB            | 124 MiB         | 12 MiB |

Ten times the rows leave the retained heap where it was: the rows are produced as they are pulled, the sheet goes out in chunks of at most 65 535 bytes, and only the style table and the zip's central directory grow, with the vocabulary rather than the data.

Budgets are the means rounded up to the next ten milliseconds and the growth ratios rounded up to a round number. A local run that lands above a budget by more than its own margin of error is a regression to explain in the pull request; the pull-request report is where to look, since local numbers move with the machine, and it informs rather than gates. A change that lowers a number by more than the margin lowers the budget in the same change.

## Auto width, 2026-11-10, Apple M5 Pro, 48 GiB, Node 24.22.1

The `auto` variant is each suite with every column sized to its content (ADR 0024), written as STORE, read against `store` in the same run: the difference is the cost of measuring, since nothing else changes. Minimum of seven, `store` then `auto`, in milliseconds.

| Suite        | Node 24.22.1 | Bun 1.4.4 | Deno 2.9.6 |
| ------------ | ------------ | --------- | ---------- |
| Wide table   | 159 / 153    | 131 / 119 | 102 / 116  |
| Many sheets  | 74 / 79      | 54 / 61   | 52 / 57    |
| String-heavy | 91 / 89      | 91 / 95   | 83 / 84    |

Two `store` runs of the same build in one session differed by fifteen percent on Node, so differences inside ten percent here are noise. The wide table, the suite the issue budgeted at ten percent, measures within that noise on Node and Bun and fourteen percent above it on Deno; the many-sheets suite pays seven to thirteen percent, the string-heavy suite nothing. What remains per cell is the call that asks whether the column is measured: a shared string is scanned once per sheet by its table index, a text whose length cannot beat the column is not scanned, a number whose magnitude cannot beat the column is not counted, and a column at the ceiling is left alone. The three suites without auto width are unchanged, which the pull-request report against `main` confirms; on the report the `auto` row has no `main` side, since a build from before auto widths would write the suite without widths rather than measure it. No budget changes.

## CI baseline, 2026-11-05, GitHub-hosted ubuntu-latest, seven runs

The runner is about twice as slow as the machine above, and machines in the pool differ among themselves by up to forty percent; these numbers are a reference, not budgets and not a comparison base. Deflate first, STORE second, in milliseconds.

| Suite        | Node 24.22.2 | Bun 1.4.4 | Deno 2.9.6 |
| ------------ | ------------ | --------- | ---------- |
| Wide table   | 1105 / 441   | 951 / 433 | 806 / 277  |
| Many sheets  | 543 / 204    | 453 / 167 | 404 / 120  |
| String-heavy | 780 / 236    | 709 / 234 | 666 / 213  |

## What the pull-request report measures

Wall time of the pull request's build and of `main`'s build in the same process on the same runner, taken in turns, seven runs each, compared by their minimum. Taking the minimum discards the noise, which only ever adds time; measuring both sides on one machine discards the difference between runners, which is the larger effect by far. Changes inside five percent are labelled noise, beyond five percent flagged, beyond fifteen percent called a regression or an improvement outright, and the job never gates a merge. The STORE variant is the JavaScript path alone; the deflate variant adds zlib, which differs per runtime and is the number users feel. Cross-runtime differences are shown, not judged.

## What the profiling pass found

A CPU profile of four STORE writes of the wide table, before the pass: CRC-32 took 29 % of the time, garbage collection 23 %, rendering and interning the rest.

- CRC-32 advanced one byte per lookup. It now uses slice-by-8 tables built once: 40 MiB in 43 ms instead of 123 ms, 2.9× faster, bit-identical by the standard check values and a property test against Node's own implementation.
- A worksheet was built as one string, joined from one string per row, then encoded. Rows now go into a byte sink as they are built, so a sheet never exists as one string. STORE writes of the wide table got 36 % faster, peak memory growth fell from 74× to 25× of the output, and garbage collection fell to 11 % of the time.
- An unstyled date cell built a key string per cell to find its cell format. The index per built-in date format is now cached.

What still happens per cell, and why it stays: the cell reference and the cell element are strings because they are the output; a number becomes a string because that is the output; a date cell inside a styled column resolves its merged style through a key string, bounded by such cells and left alone. The shared-strings and styles parts are still built as strings, proportional to distinct strings and formats rather than to cells. No regular expression and no closure runs per cell: the only regular expressions run once per error message and once per filtered sheet.
