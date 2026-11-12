# Non-goals: reading, charts, pivot tables and password encryption

Four things are outside this library and stay outside: reading or editing an existing workbook, charts, pivot tables and password encryption. The README names them, so the answer is there before the question.

Each is large against a writer of this size. Reading is a parser for everything other applications write, where the writer only has to produce what it chooses to (ADR 0001). A chart is a drawing format of its own with a part per chart, dozens of chart kinds and their options. A pivot table needs a cache of its source data, definition and records, in a second format beside the sheet, and applications disagree on how much of it they rebuild. An encrypted workbook is not a package at all but a compound file around one, with key derivation and a block cipher that the web-standard core would have to reach or carry (ADR 0002).

Demand does not change this. In the issue trackers of other JavaScript spreadsheet libraries, read in November 2026, charts were the most supported open request, and encryption and pivot tables were asked for steadily. The same trackers show what those libraries pay for breadth: their most discussed problems are files that open with a repair dialog, memory, and bundle size. This library is the other trade: a small surface, proven to open, at a bundle size it holds to a budget (ADR 0021).

## What would reopen one

A non-goal is reopened by a design, not by a request: one that leaves the bundle of an application that does not use the feature unchanged, byte for byte, and that can be proven in the four applications the way every fixture is. A separate entry point is the likely shape. Until someone brings that, the answer is the README's.
