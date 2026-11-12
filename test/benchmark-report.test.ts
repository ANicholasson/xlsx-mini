import { describe, expect, it } from "vite-plus/test";
import {
  type BenchmarkResult,
  COMMENT_MARKER,
  type RunResult,
  describeChange,
  describeSizeChange,
  readSize,
  renderReport,
} from "../bench/report.ts";
import type { SizeResult } from "../bench/size.ts";

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

  it("adds the size of both builds in a bundle, gzipped, when it is given one", () => {
    const size: SizeResult = {
      head: {
        writeXlsx: { minified: 45_900, gzip: 16_490 },
        everything: { minified: 52_741, gzip: 18_648 },
      },
      base: {
        writeXlsx: { minified: 45_747, gzip: 16_423 },
        everything: { minified: 52_741, gzip: 18_648 },
      },
    };
    const report = renderReport([run("node")], {
      baseLabel: "main",
      headLabel: "this pull request",
      size,
    });
    expect(report.indexOf("### Bundle size")).toBeGreaterThan(report.indexOf("### Node"));
    expect(report).toContain("| Import | main | this pull request | Change |");
    expect(report).toContain("| `writeXlsx` alone | 16.42 kB | 16.49 kB | +67 B (+0.4 %) |");
    expect(report).toContain("| Every export | 18.65 kB | 18.65 kB | same |");
  });

  it("shows the size without a change when there is no base build", () => {
    const report = renderReport([run("node", 1, false)], {
      baseLabel: "main",
      headLabel: "this pull request",
      size: {
        head: {
          writeXlsx: { minified: 45_747, gzip: 16_423 },
          everything: { minified: 52_741, gzip: 18_648 },
        },
      },
    });
    expect(report).toContain("| `writeXlsx` alone | — | 16.42 kB | — |");
  });

  it("has no size table without sizes", () => {
    const report = renderReport([run("node")], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).not.toContain("Bundle size");
  });
});

describe("describeSizeChange", () => {
  it("shows every byte, with its sign", () => {
    expect(describeSizeChange(16_000, 16_000)).toBe("same");
    expect(describeSizeChange(16_000, 16_160)).toBe("+160 B (+1.0 %)");
    expect(describeSizeChange(16_000, 15_200)).toBe("-800 B (-5.0 %)");
  });
});

describe("readSize", () => {
  it("is nothing for a directory without a size file", () => {
    expect(readSize("bench/no-such-results")).toBeUndefined();
  });
});
