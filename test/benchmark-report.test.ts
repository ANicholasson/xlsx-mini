import { describe, expect, it } from "vite-plus/test";
import {
  type BenchmarkResult,
  COMMENT_MARKER,
  type RunResult,
  describeChange,
  renderReport,
} from "../bench/report.ts";

function run(runtime: string, headScale = 1, withBase = true): RunResult {
  const benchmark = (suite: string, variant: "deflate" | "store", ms: number): BenchmarkResult => ({
    suite,
    variant,
    samples: 7,
    head: { minMs: ms * headScale, p50Ms: ms * headScale * 1.03 },
    ...(withBase ? { base: { minMs: ms, p50Ms: ms * 1.03 } } : {}),
  });
  return {
    runtime,
    version: runtime === "node" ? "24.22.1" : "1.4.4",
    arch: "x64",
    commit: "head1234abcd",
    ...(withBase ? { baseCommit: "base5678abcd" } : {}),
    date: "2026-11-05T10:00:00Z",
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
    expect(describeChange(100, 108)).toBe("⚠️ +8.0 %");
    expect(describeChange(100, 120)).toBe("🔴 +20.0 %");
    expect(describeChange(100, 90)).toBe("🟢 -10.0 %");
  });
});

describe("renderReport", () => {
  it("starts with the marker, orders runtimes, and compares minima per row", () => {
    const report = renderReport([run("bun"), run("node", 1.2)], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report.startsWith(`${COMMENT_MARKER}\n## Benchmarks`)).toBe(true);
    expect(report.indexOf("### Node 24.22.1")).toBeLessThan(report.indexOf("### Bun 1.4.4"));
    expect(report).toContain("| wide table | deflate | 500.0 ms | 600.0 ms | 🔴 +20.0 % |");
    expect(report).toContain("| wide table | store | 180.0 ms | 216.0 ms | 🔴 +20.0 % |");
    expect(report).toContain("| wide table | deflate | 500.0 ms | 500.0 ms | +0.0 % |");
    expect(report).toContain("minimum of 7 runs each");
    expect(report).toContain(
      "this pull request `head123` against main `base567`; 7 runs each, minimum.",
    );
    expect(report).not.toContain("```");
  });

  it("shows dashes when a run has no base build", () => {
    const report = renderReport([run("node", 1, false)], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).toContain("| wide table | deflate | — | 500.0 ms | — |");
    expect(report).toContain("this pull request `head123`; 7 runs each, minimum.");
    expect(report).not.toContain("against main");
  });
});
