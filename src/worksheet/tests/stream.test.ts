import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import {
  type Row,
  type Sheet,
  type StreamedSheet,
  type WorksheetContext,
  renderWorksheet,
  streamWorksheet,
} from "../index.js";

const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

function context(overrides: Partial<WorksheetContext> = {}): WorksheetContext {
  return {
    strings: createStringTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
    ...overrides,
  };
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<string> {
  let text = "";
  for await (const chunk of stream) text += decode(chunk);
  return text;
}

const withoutDimension = (xml: string) => xml.replace(/<dimension ref="[^"]*"\/>/, "");

async function* later<T>(items: readonly T[]): AsyncGenerator<T> {
  for (const item of items) {
    await Promise.resolve();
    yield item;
  }
}

describe("streamWorksheet", () => {
  const shapes: readonly (readonly [string, Sheet])[] = [
    ["an empty sheet", { name: "S", rows: [] }],
    [
      "rows with every kind of value",
      { name: "S", rows: [["a", 1, true, new Date(Date.UTC(2026, 0, 2))], [], ["b"]] },
    ],
    ["heights without cells", { name: "S", rows: [], rowHeights: { 1: 28 } }],
    [
      "a height on an empty row before a row with cells",
      { name: "S", rows: [[], ["a"]], rowHeights: { 1: 20 } },
    ],
    [
      "columns, freeze, merges and a filter",
      {
        name: "S",
        columns: [{ width: 12, style: { bold: true } }],
        freeze: { rows: 1 },
        autoFilter: true,
        merges: ["A3:B3"],
        rows: [["h", "i"], ["x", "y"], ["z"]],
      },
    ],
  ];

  it.each(shapes)(
    "writes the same part as renderWorksheet for %s, without the dimension",
    async (_name, sheet) => {
      const rendered = renderWorksheet(sheet, context());
      const streamed = streamWorksheet(sheet, context({ strings: createStringTable() }));
      const body = await collect(streamed.body);
      expect(body).toBe(withoutDimension(decode(rendered.data)));
      expect(body).not.toContain("<dimension");
      expect(streamed.facts().filterRange).toBe(rendered.filterRange);
    },
  );

  it("takes rows from an async iterable and from a sync iterator alike", async () => {
    const rows: Row[] = [["a"], ["b", 2]];
    const sync = await collect(streamWorksheet({ name: "S", rows }, context()).body);
    const async = await collect(streamWorksheet({ name: "S", rows: later(rows) }, context()).body);
    expect(async).toBe(sync);
    expect(sync).toContain('<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2</v></c></row>');
  });

  it("writes text into the cell when the context asks for inline strings", async () => {
    const strings = createStringTable();
    const sheet: StreamedSheet = { name: "S", rows: [[" lead", "a & b"]] };
    const body = await collect(
      streamWorksheet(sheet, context({ strings, inlineStrings: true })).body,
    );
    expect(body).toContain(
      '<c r="A1" t="inlineStr"><is><t xml:space="preserve"> lead</t></is></c>',
    );
    expect(body).toContain('<c r="B1" t="inlineStr"><is><t>a &amp; b</t></is></c>');
    expect(strings.render()).toContain('count="0" uniqueCount="0"');
  });

  it("sends the head before the first row and the filter range only at the end", async () => {
    function* rows(): Generator<Row> {
      for (let i = 0; i < 20_000; i++) yield [`row ${i}`, i];
    }
    const streamed = streamWorksheet({ name: "S", autoFilter: true, rows: rows() }, context());
    const reader = streamed.body.getReader();
    const first = decode((await reader.read()).value!);
    expect(first.startsWith('<?xml version="1.0"')).toBe(true);
    expect(first).toContain("<sheetData>");
    expect(first).not.toContain("</sheetData>");
    expect(streamed.facts().filterRange).toBeUndefined();
    reader.releaseLock();
    await collect(streamed.body);
    expect(streamed.facts().filterRange).toBe("A1:B20000");
  });

  it("reads no row before the consumer pulls, and returns the iterator when cancelled or finished", async () => {
    const log: string[] = [];
    function* rows(): Generator<Row> {
      try {
        log.push("started");
        for (let i = 0; ; i++) yield [`row ${i}`];
      } finally {
        log.push("returned");
      }
    }
    const streamed = streamWorksheet({ name: "S", rows: rows() }, context());
    await Promise.resolve();
    expect(log).toEqual([]);
    const reader = streamed.body.getReader();
    await reader.read();
    expect(log).toEqual(["started"]);
    await reader.cancel();
    expect(log).toEqual(["started", "returned"]);

    // An array iterator has no `return`: cancelling must cope with that too.
    const plain = streamWorksheet({ name: "S", rows: [["a"], ["b"]] }, context());
    const plainReader = plain.body.getReader();
    await plainReader.cancel();
    await expect(plainReader.closed).resolves.toBeUndefined();

    const log2: string[] = [];
    async function* asyncRows(): AsyncGenerator<Row> {
      try {
        yield ["a"];
      } finally {
        log2.push("returned");
      }
    }
    await collect(streamWorksheet({ name: "S", rows: asyncRows() }, context()).body);
    expect(log2).toEqual(["returned"]);
  });

  it("errors the body with the cell's error when a row cannot be written, and closes the row source", async () => {
    const log: string[] = [];
    async function* rows(): AsyncGenerator<Row> {
      try {
        yield ["fine"];
        yield [Number.NaN];
        yield ["never"];
      } finally {
        log.push("returned");
      }
    }
    const streamed = streamWorksheet({ name: "S", rows: rows() }, context({ sheetIndex: 2 }));
    await expect(collect(streamed.body)).rejects.toMatchObject({
      code: "INVALID_NUMBER",
      path: "sheets[2].rows[1][0]",
    });
    expect(log).toEqual(["returned"]);
    // An array iterator has no `return` to call on the way out.
    await expect(
      collect(streamWorksheet({ name: "S", rows: [[Number.NaN]] }, context()).body),
    ).rejects.toMatchObject({ code: "INVALID_NUMBER" });
  });

  it("streams a large sheet in chunks no larger than a stored block", async () => {
    function* rows(): Generator<Row> {
      for (let i = 0; i < 20_000; i++) yield [`row ${i}`, i, "some text that repeats"];
    }
    const chunks: number[] = [];
    for await (const chunk of streamWorksheet({ name: "S", rows: rows() }, context()).body) {
      chunks.push(chunk.length);
    }
    expect(chunks.length).toBeGreaterThan(5);
    expect(Math.max(...chunks)).toBeLessThanOrEqual(65_535);
  });
});
