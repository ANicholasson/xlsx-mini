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
import type { PackageSize, SizeResult } from "../bench/size.ts";

function bundle(digest: string, gzip: number): PackageSize["writeXlsx"] {
  return { digest, minified: gzip * 3, gzip };
}

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
    expect(report).not.toContain("`read` reads");
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

  it("states how many tokens per second the tokenizer reads, at its minimum", () => {
    const tokens: BenchmarkResult = {
      suite: "wide table",
      variant: "tokens",
      samples: 7,
      head: { minMs: 120, p50Ms: 125 },
      tokens: 4_200_000,
    };
    const node = run("node", 1, false);
    const report = renderReport([{ ...node, benchmarks: [...node.benchmarks, tokens] }], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).toContain("| wide table | tokens | — | 120.0 ms | — |");
    expect(report).toContain(
      "`tokens` is the reader's tokenizer over the worksheet of the wide table: 4,200,000 tokens, which this pull request reads at 35.0 million tokens per second.",
    );
  });

  it("shows read time beside the write time of a suite, and says once what a read reads", () => {
    const reads: BenchmarkResult[] = [
      {
        suite: "wide table",
        variant: "read",
        samples: 7,
        head: { minMs: 330, p50Ms: 335 },
        base: { minMs: 300, p50Ms: 305 },
      },
      { suite: "wide table", variant: "objects", samples: 7, head: { minMs: 360, p50Ms: 365 } },
    ];
    const withReads = (runtime: string): RunResult => {
      const written = run(runtime);
      return { ...written, benchmarks: [...written.benchmarks, ...reads] };
    };
    const report = renderReport([withReads("node"), withReads("bun")], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report).toContain(
      "| wide table | store | 180.0 ms | 180.0 ms | +0.0 % |\n| wide table | read | 300.0 ms | 330.0 ms | ⚠️ +10.0 % |\n| wide table | objects | — | 360.0 ms | — |\n| many sheets | deflate | 250.0 ms | 250.0 ms | +0.0 % |",
    );
    expect(report.match(/`read` reads/g)).toHaveLength(1);
    expect(report).toContain(
      "`read` reads every row of every sheet of the suite through `rows`, and `objects` through `objectsFromRows` with a column for each header. Both builds read one file: the suite as this pull request writes it by default.",
    );
    expect(report.indexOf("`read` reads")).toBeLessThan(report.indexOf("### Node"));
  });

  it("joins the runs of one runtime into one table, by suite and then by variant", () => {
    const reading: RunResult = {
      ...run("node", 1, false),
      benchmarks: [
        { suite: "wide table", variant: "read", samples: 7, head: { minMs: 330, p50Ms: 335 } },
        { suite: "many sheets", variant: "read", samples: 7, head: { minMs: 200, p50Ms: 205 } },
      ],
    };
    const report = renderReport([reading, run("node")], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(report.match(/### Node/g)).toHaveLength(1);
    expect(report).toContain(
      "| wide table | deflate | 500.0 ms | 500.0 ms | +0.0 % |\n| wide table | store | 180.0 ms | 180.0 ms | +0.0 % |\n| wide table | read | — | 330.0 ms | — |\n| many sheets | deflate | 250.0 ms | 250.0 ms | +0.0 % |\n| many sheets | read | — | 200.0 ms | — |",
    );
    const elsewhere = renderReport([{ ...reading, version: "24.22.2" }, run("node")], {
      baseLabel: "main",
      headLabel: "this pull request",
    });
    expect(elsewhere.match(/### Node/g)).toHaveLength(2);
  });

  it("adds the size of both builds in a bundle, gzipped, and whether the bytes are the same", () => {
    const size: SizeResult = {
      head: {
        writeXlsx: bundle("w2", 16_490),
        streamXlsx: bundle("s2", 17_000),
        everything: bundle("e", 18_648),
        browser: bundle("b", 450),
        read: bundle("r", 5_100),
      },
      base: {
        writeXlsx: bundle("w1", 16_423),
        streamXlsx: bundle("s1", 17_000),
        everything: bundle("e", 18_648),
        browser: bundle("b", 450),
      },
    };
    const report = renderReport([run("node")], {
      baseLabel: "main",
      headLabel: "this pull request",
      size,
    });
    expect(report.indexOf("### Bundle size")).toBeGreaterThan(report.indexOf("### Node"));
    expect(report).toContain("| Import | main | this pull request | Change | Same bytes as main |");
    expect(report).toContain("| `writeXlsx` alone | 16.42 kB | 16.49 kB | +67 B (+0.4 %) | no |");
    expect(report).toContain("| `streamXlsx` alone | 17.00 kB | 17.00 kB | same | no |");
    expect(report).toContain("| Every export | 18.65 kB | 18.65 kB | same | yes |");
    expect(report).toContain("| `xlsx-mini/browser` | 0.45 kB | 0.45 kB | same | yes |");
    expect(report).toContain("| `xlsx-mini/read` | — | 5.10 kB | — | — |");
  });

  it("shows the size without a change when there is no base build", () => {
    const report = renderReport([run("node", 1, false)], {
      baseLabel: "main",
      headLabel: "this pull request",
      size: {
        head: {
          writeXlsx: bundle("w", 16_423),
          streamXlsx: bundle("s", 17_000),
          everything: bundle("e", 18_648),
          browser: bundle("b", 450),
        },
      },
    });
    expect(report).toContain("| `writeXlsx` alone | — | 16.42 kB | — | — |");
    expect(report).not.toContain("xlsx-mini/read");
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
