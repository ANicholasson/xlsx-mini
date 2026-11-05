import { describe, expect, it } from "vite-plus/test";
import {
  type BenchmarkResult,
  COMMENT_MARKER,
  HISTORY_LIMIT,
  HISTORY_POINTS,
  type HistoryEntry,
  type RunResult,
  appendHistory,
  describeChange,
  renderHistoryChart,
  renderReport,
  renderSvg,
} from "../bench/report.ts";

function run(runtime: string, commit: string, scale = 1): RunResult {
  const benchmark = (
    suite: string,
    variant: "deflate" | "store",
    p50Ms: number,
  ): BenchmarkResult => ({
    suite,
    variant,
    samples: 7,
    p50Ms: p50Ms * scale,
    minMs: p50Ms * scale * 0.95,
    meanMs: p50Ms * scale * 1.02,
  });
  return {
    runtime,
    version: runtime === "node" ? "24.22.1" : "1.4.4",
    arch: "x64",
    commit,
    date: `2026-11-05T10:00:0${commit.length % 10}Z`,
    benchmarks: [
      benchmark("wide table", "deflate", 500),
      benchmark("wide table", "store", 180),
      benchmark("many sheets", "deflate", 250),
    ],
  };
}

describe("describeChange", () => {
  it("labels changes by the noise and alert bands", () => {
    expect(describeChange(100, 103)).toBe("+3.0 %");
    expect(describeChange(100, 97)).toBe("-3.0 %");
    expect(describeChange(100, 115)).toBe("⚠️ +15.0 %");
    expect(describeChange(100, 140)).toBe("🔴 +40.0 %");
    expect(describeChange(100, 85)).toBe("🟢 -15.0 %");
  });
});

describe("renderReport", () => {
  const head = [run("bun", "head1234"), run("node", "head1234", 1.2)];
  const base = [run("node", "base5678")];

  it("starts with the marker, orders runtimes, and compares against the base per runtime", () => {
    const report = renderReport(head, base, [], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report.startsWith(`${COMMENT_MARKER}\n## Benchmarks`)).toBe(true);
    expect(report.indexOf("### Node 24.22.1")).toBeLessThan(report.indexOf("### Bun 1.4.4"));
    expect(report).toContain("| wide table | deflate | 500.0 ms | 600.0 ms | ⚠️ +20.0 % |");
    expect(report).toContain("| wide table | store | 180.0 ms | 216.0 ms | ⚠️ +20.0 % |");
    // Bun has no base run: the base and change columns stay empty.
    expect(report).toContain("| wide table | deflate | — | 500.0 ms | — |");
    expect(report).toContain("against main `base567`");
  });

  it("charts the first runtime with the base as bars and the head as a line", () => {
    const report = renderReport(head, base, [], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).toContain("Node 24.22.1, deflate: main as bars, this pull request as the line");
    expect(report).toContain('x-axis ["wide table", "many sheets"]');
    expect(report).toContain("bar [500.0, 250.0]");
    expect(report).toContain("line [600.0, 300.0]");
    expect(report).toContain("Node 24.22.1, store: main as bars");
    expect(report).toContain("bar [180.0]");
    expect(report).toContain("line [216.0]");
    expect(report).toContain('y-axis "ms" 0 --> 700');
  });

  it("says so when the base has no results, and charts the head alone", () => {
    const report = renderReport(head, [], [], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).toContain("No results on `main` yet");
    expect(report).toContain("bar [600.0, 300.0]");
    expect(report).toContain("bar [216.0]");
    expect(report).not.toContain("line [");
  });
});

describe("history", () => {
  it("appends one entry per benchmark, replaces a rerun commit, and caps per benchmark", () => {
    let history: HistoryEntry[] = [];
    for (let i = 0; i < HISTORY_LIMIT + 5; i++) {
      history = appendHistory(history, [run("node", `c${String(i).padStart(4, "0")}`)]);
    }
    const wideDeflate = history.filter((e) => e.suite === "wide table" && e.variant === "deflate");
    expect(wideDeflate).toHaveLength(HISTORY_LIMIT);
    expect(wideDeflate.at(-1)?.commit).toBe(`c${String(HISTORY_LIMIT + 4).padStart(4, "0")}`);

    const rerun = appendHistory(history, [run("node", wideDeflate.at(-1)!.commit, 2)]);
    expect(rerun.filter((e) => e.commit === wideDeflate.at(-1)!.commit)).toHaveLength(3);
    expect(rerun.at(-1)?.p50Ms).toBe(500);
  });

  it("charts the last commits of the deflate variants as one line per suite", () => {
    let history: HistoryEntry[] = [];
    for (let i = 0; i < HISTORY_POINTS + 3; i++) {
      history = appendHistory(history, [run("node", `h${String(i).padStart(6, "0")}`)]);
    }
    const chart = renderHistoryChart(history, "node");
    expect(chart).toContain("History on main: Node, deflate, ms");
    expect(chart).toContain("Lines in order: wide table, many sheets.");
    expect(chart?.match(/"h\d{6}"/g)).toHaveLength(HISTORY_POINTS);
    expect(chart?.match(/^ {4}line /gm)).toHaveLength(2);
    expect(renderHistoryChart(history, "deno")).toBeUndefined();
  });
});

describe("renderSvg", () => {
  it("draws one bar per suite and runtime with a title carrying the value", () => {
    const svg = renderSvg([run("node", "abc"), run("bun", "abc", 0.5)]);
    expect(svg.startsWith("<svg ")).toBe(true);
    expect(svg.match(/<rect [^>]*><title>/g)).toHaveLength(4);
    expect(svg).toContain("<title>wide table: 500.0 ms</title>");
    expect(svg).toContain("<title>wide table: 250.0 ms</title>");
    expect(svg).toContain(">Node 24.22.1</text>");
    expect(svg).toContain(">Bun 1.4.4</text>");
  });
});
