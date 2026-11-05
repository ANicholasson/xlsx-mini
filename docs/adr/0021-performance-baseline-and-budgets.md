# Performance baseline and budgets

The first measured baseline is the budget, and a budget only ratchets down. Numbers come from `pnpm bench` (the three suites in `bench/`, deflate and STORE each, seven runs after one warm-up against the built package loaded by the host itself) and `pnpm bench:memory` (one write per suite in a fresh process). On pull requests the same script runs on GitHub's runners in Node, Bun and Deno, and the report compares the medians with the latest run on `main`, kept on the `benchmarks` branch together with the history.

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

Budgets are the means rounded up to the next ten milliseconds and the growth ratios rounded up to a round number. A local run that lands above a budget by more than its own margin of error is a regression to explain in the pull request; the pull-request report is where to look, since local numbers move with the machine, and it informs rather than gates. A change that lowers a number by more than the margin lowers the budget in the same change.

## What the pull-request report measures

Wall time on a shared runner, so run-to-run noise of several percent is normal: changes inside ten percent are labelled noise, beyond ten percent flagged, beyond twenty-five percent called a regression or an improvement outright, and the job never gates a merge. The STORE variant is the JavaScript path alone; the deflate variant adds zlib, which differs per runtime and is the number users feel. Cross-runtime differences are shown, not judged.

## What the profiling pass found

A CPU profile of four STORE writes of the wide table, before the pass: CRC-32 took 29 % of the time, garbage collection 23 %, rendering and interning the rest.

- CRC-32 advanced one byte per lookup. It now uses slice-by-8 tables built once: 40 MiB in 43 ms instead of 123 ms, 2.9× faster, bit-identical by the standard check values and a property test against Node's own implementation.
- A worksheet was built as one string, joined from one string per row, then encoded. Rows now go into a byte sink as they are built, so a sheet never exists as one string. STORE writes of the wide table got 36 % faster, peak memory growth fell from 74× to 25× of the output, and garbage collection fell to 11 % of the time.
- An unstyled date cell built a key string per cell to find its cell format. The index per built-in date format is now cached.

What still happens per cell, and why it stays: the cell reference and the cell element are strings because they are the output; a number becomes a string because that is the output; a date cell inside a styled column resolves its merged style through a key string, bounded by such cells and left alone. The shared-strings and styles parts are still built as strings, proportional to distinct strings and formats rather than to cells. No regular expression and no closure runs per cell: the only regular expressions run once per error message and once per filtered sheet.
