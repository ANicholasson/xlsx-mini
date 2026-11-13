# Performance baseline and budgets

The first measured baseline is the budget, and a time or memory budget only ratchets down; the bundle-size budget has its own rule, in its section below. Numbers come from `pnpm bench` (the three suites in `bench/` against the built package loaded by the host itself; seven runs after one warm-up in one process until 2026-11-11, since then every benchmark in processes of its own, four of them with four runs each after one warm-up) and `pnpm bench:memory` (one write per suite in a fresh process). On pull requests the same script runs on GitHub's runners in Node, Bun and Deno with `main` built alongside, and the report compares the two on that one machine.

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

## The checksum in blocks, and one process per benchmark, 2026-11-11

Two pull-request reports flagged the wide-table STORE row on Node while the code under review had not touched that path, and two copies of one build, timed against each other in the harness, showed the same: 164 ms against 145 ms. Three faults lay behind it, one in the library and two in the harness, and the harness had been hiding the one in the library.

### The library: a loop in a function called too rarely

`crc32Update` held its loop in a function called once per part, a handful of calls per workbook. An engine compiles a function it sees called often; a long loop in a function called rarely is instead replaced on the stack while it runs. On V8, when the first part a process checksummed was a large one, that happened again on every call: each began in the slow tiers, and the function itself was never compiled. A CPU profile of ten STORE writes of the wide table showed the checksum at 430 ms in that state and 222 ms after a small write had come first, three quarters of the difference between the two. Which state a program got depended on what it happened to write first.

The checksum now runs over blocks of 16 KiB, the same slice-by-8 loop called on one slice after another, so the function is called often and compiled on first use everywhere. Twelve mebibytes in one call, first call in the process and then the minimum of eleven more, in milliseconds:

| Runtime      | Before, first / minimum | In blocks, first / minimum |
| ------------ | ----------------------- | -------------------------- |
| Node 24.22.1 | 15.8 / 13.0             | 7.5 / 5.5                  |
| Bun 1.4.4    | 14.5 / 5.7              | 10.4 / 5.6                 |
| Deno 2.9.6   | 17.2 / 12.3             | 7.7 / 5.3                  |

Blocks of 64 KiB and more started slowly on V8 again, and 4 KiB gained nothing over 16. Handing the loop a start and an end within the whole array, in place of a slice, ran 2.4 times slower on Bun, which no longer saw that the indexes stay inside the array; hence the slices, which cost one small object per 16 KiB.

What a program sees when its first workbook is the suite, one build per process with nothing written before, minimum of nine writes after the first, the lower of two such runs, in milliseconds:

| Suite, variant      | Node, before / after | Bun, before / after | Deno, before / after |
| ------------------- | -------------------- | ------------------- | -------------------- |
| Wide table, STORE   | 166 / 142            | 127 / 109           | 109 / 110            |
| Wide table, deflate | 493 / 465            | 359 / 330           | 338 / 339            |
| Many sheets, STORE  | 71 / 71              | 55 / 57             | 52 / 53              |
| String-heavy, STORE | 88 / 89              | 94 / 93             | 87 / 88              |

The wide table as STORE is fourteen percent faster on Node and on Bun, six and eight percent with deflate. Deno reaches the fast state by itself after a few writes, so its minimum is unchanged, and its first write of the wide table falls from 178 ms to 146 ms. The many-sheets suite checksums forty small parts and was never affected. Output is byte for byte what it was.

### The harness: a probe, and two builds in one process

The harness loaded both builds into one process and timed them call by call in turns. Two things about that were wrong.

Before timing, it wrote one tiny workbook with the base build alone, to learn whether that build knew auto widths. That write put the base build's checksum in the fast state and left the other in the slow one. Which build was loaded first, which was warmed first and how many warm-ups each got made no difference; moving that one write to the other build moved the fifteen percent with it.

And two builds in one process are not independent. On the runner, Bun wrote the streamed wide table with `main`'s build at less than half the speed of the pull request's, on three reports in a row, whatever the pull request changed. With the two swapped the slow one was still `main`'s, at −56 % as the base and +139 % as the build under test; two copies of one build read within one percent of each other; and alone in a process the two builds took 761 ms and 762 ms. Neither the order nor the code: one build was slowed by sharing a process with a different one. On Deno the build under test read a few percent slow on every row of two reports.

Every measurement now runs in a process of its own: it loads one build, builds one suite, writes it one way once to warm up and times four more writes. The two builds' processes for a benchmark run one after the other, the order swapping from one benchmark and one round to the next, for four rounds, and the minimum of the sixteen runs is compared. Nothing is written before the suite, so a build meets what a program whose first workbook is that suite meets. A base from before a feature fails its warm-up and has no row; the auto-width suite names its columns one by one for that reason, since such a base refuses `width: "auto"` where it would ignore a sheet's `autoWidth`. Every measuring process is given a collector to call, on Deno too. The three runtimes run as three jobs side by side, about six minutes each.

A first version gave each build one process for all the suites. A row's two timings then lay minutes apart, and on the runner two copies of one build differed by more than the band on Node's stream rows and on several of Deno's: the minima of eight such processes spread by fourteen percent on Node's streamed wide table and by eighteen on Deno's STORE. One process per benchmark puts the two timings seconds apart.

On the runner, two copies of one build, six rounds of six runs, the largest difference on any of the twelve rows, and how often a row lands outside the band at the default, taken over every choice of four of the six rounds:

| Runtime      | Largest difference, 36 runs | Rows outside ±5 % at four rounds of four |
| ------------ | --------------------------- | ---------------------------------------- |
| Node 24.22.2 | 4.2 %                       | none, largest 4.9 %                      |
| Bun 1.4.4    | 1.9 %                       | 0.6 %, one row, largest 7.7 %            |
| Deno 2.9.6   | 2.3 %                       | none, largest 3.2 %                      |

So a flag on a report is seldom noise, and now and then it is: one row in a few hundred. The same run with `main` as the base read the wide table as STORE at −9.2 % on Node and −12.3 % on Bun, the auto variant of it at −6.9 % and −9.5 %, Bun's deflate at −5.2 %, every row on Deno inside one and a half percent, and Bun's streamed wide table at +0.5 %.

The numbers a run prints are not those of the sections above, which came from one long process: locally the streamed string-heavy suite reads a quarter lower. The wide-table budgets ratchet down, from 520 ms to 490 ms and, for STORE, from 180 ms to 160 ms. What the harness still cannot see is the cost of a first call: each benchmark is warmed once before it is timed and the minimum is compared, so a cost paid only once per process is outside the report.

## Bundle size, 2026-11-12

What the built package adds to an application's bundle is a cost its users pay on every page load and every cold start, and it is the plainest difference between this library and the ones that also read: so it is measured and held like the time. `bench/size.ts` bundles the package from its `dist/` folder with the bundler the runtime matrix already uses, for the browser platform, minified, and gzips the one file that comes out at the highest level. It does so twice: for an application that imports only `writeXlsx`, and for every export.

| Bundle            | Minified |  Gzipped | Budget (gzipped) |
| ----------------- | -------: | -------: | ---------------: |
| `writeXlsx` alone | 45 747 B | 16 423 B |         18 000 B |
| Every export      | 52 741 B | 18 648 B |         20 000 B |

The budgets are what the README states, at first under 18 kB and under 20 kB, with a kilobyte as a thousand bytes; a test fails when a bundle passes its budget and when the README states another figure. They sit a tenth above the measurement, which is room for the small features planned and no more. A feature need not cost time or memory, so those budgets only ratchet down; bytes are what a feature is made of, so this budget is a ceiling and not a ratchet. A pull request whose feature needs a higher one raises it in that same pull request, here, with the bytes and what they buy. Nothing else raises it. The two bundles differ by little because the writer is one path through every slice; what `writeXlsx` alone leaves out is streaming and `sheetFromObjects`.

The badge in the README is the same measurement of the version on npm by bundlejs, which minifies with another tool and read 17.1 kB for `writeXlsx` alone on 0.6.1, four percent above the figure here. The README's sentence therefore states the budget, which both are under, and not either reading.

The benchmark report carries the gzipped size of both bundles for the pull request and for `main`, with the difference in bytes. Bytes do not vary between runs, so every byte shows and there is no noise band.

### Raised for row objects, 2026-11-13

| Bundle            | Before row objects | With them | Budget, before | Budget, now |
| ----------------- | -----------------: | --------: | -------------: | ----------: |
| `writeXlsx` alone |           17 608 B |  18 003 B |       18 000 B |    19 000 B |
| Every export      |           19 871 B |  20 290 B |       20 000 B |    21 000 B |

Row objects (ADR 0027) cost about 400 bytes: the second row loop, the merge of a row style over a column's, and three error messages. They buy a style and a height on the row, the most asked-for thing this writer lacked. The functions newer than Excel 2007 (ADR 0023) had taken 1 185 bytes of the first budget's room before them. The new budgets are again round figures the README can state, a kilobyte apart from the old, with room for the hidden rows and the sheet view that are planned and for the four percent by which the badge's minifier reads above this one.

## CI baseline, 2026-11-05, GitHub-hosted ubuntu-latest, seven runs

The runner is about twice as slow as the machine above, and machines in the pool differ among themselves by up to forty percent; these numbers are a reference, not budgets and not a comparison base. Deflate first, STORE second, in milliseconds.

| Suite        | Node 24.22.2 | Bun 1.4.4 | Deno 2.9.6 |
| ------------ | ------------ | --------- | ---------- |
| Wide table   | 1105 / 441   | 951 / 433 | 806 / 277  |
| Many sheets  | 543 / 204    | 453 / 167 | 404 / 120  |
| String-heavy | 780 / 236    | 709 / 234 | 666 / 213  |

## What the pull-request report measures

Wall time of the pull request's build and of `main`'s build on the same runner, every benchmark of each build in a process of its own, the two builds one after the other and in turns, sixteen runs each over four processes, compared by their minimum. Neither build writes anything before the suite it is timed on. Taking the minimum discards the noise, which only ever adds time; measuring both sides on one machine discards the difference between runners, which is the larger effect by far. Changes inside five percent are labelled noise, beyond five percent flagged, beyond fifteen percent called a regression or an improvement outright, and the job never gates a merge. The STORE variant is the JavaScript path alone; the deflate variant adds zlib, which differs per runtime and is the number users feel. Cross-runtime differences are shown, not judged. Below the tables the report states the bundle size of the two builds, measured once, on the Node job.

## What the profiling pass found

A CPU profile of four STORE writes of the wide table, before the pass: CRC-32 took 29 % of the time, garbage collection 23 %, rendering and interning the rest.

- CRC-32 advanced one byte per lookup. It now uses slice-by-8 tables built once: 40 MiB in 43 ms instead of 123 ms, 2.9× faster, bit-identical by the standard check values and a property test against Node's own implementation.
- A worksheet was built as one string, joined from one string per row, then encoded. Rows now go into a byte sink as they are built, so a sheet never exists as one string. STORE writes of the wide table got 36 % faster, peak memory growth fell from 74× to 25× of the output, and garbage collection fell to 11 % of the time.
- An unstyled date cell built a key string per cell to find its cell format. The index per built-in date format is now cached.

What still happens per cell, and why it stays: the cell reference and the cell element are strings because they are the output; a number becomes a string because that is the output; a date cell inside a styled column resolves its merged style through a key string, bounded by such cells and left alone. The shared-strings and styles parts are still built as strings, proportional to distinct strings and formats rather than to cells. No regular expression and no closure runs per cell: the only regular expressions run once per error message and once per filtered sheet.
