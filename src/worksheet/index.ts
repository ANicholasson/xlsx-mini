import { XlsxMiniError } from "../errors/index.js";
import { XML_DECLARATION, createChunkSink, createXmlSink } from "../xml/index.js";
import { collectColumns, collectRowHeights, renderFreeze, renderMerges } from "./lib/layout.js";
import type { CellStyle } from "../styles/index.js";
import type { Row, Sheet, StreamedSheet } from "./lib/model.js";
import {
  type WorksheetContext,
  createSheetDataWriter,
  renderSheetData,
} from "./lib/render-rows.js";

export type {
  Cell,
  CellObject,
  CellValue,
  Coerce,
  Column,
  Freeze,
  Register,
  Registered,
  Row,
  Sheet,
  StreamedSheet,
  TemporalInstantLike,
  TemporalLike,
  TemporalPlainDateLike,
  TemporalPlainDateTimeLike,
  TemporalZonedDateTimeLike,
} from "./lib/model.js";
export type { WorksheetContext } from "./lib/render-rows.js";
export { columnLetter } from "./lib/cell-ref.js";
export {
  type ObjectColumn,
  type SheetFromObjectsOptions,
  sheetFromObjects,
} from "./lib/from-objects.js";
export { serialFromCivil, serialFromEpochMillis } from "./lib/serial.js";

/** A rendered worksheet part and what the workbook part needs to know about it. */
export interface RenderedWorksheet {
  readonly data: Uint8Array;
  /** The autofilter range such as `A1:C4`, when the sheet filters a non-empty range. */
  readonly filterRange: string | undefined;
}

const WORKSHEET_OPEN = `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

/** The sheet's settings, validated up front, and the markup around its rows. */
interface Frame {
  readonly columnStyles: readonly (CellStyle | undefined)[];
  readonly rowHeights: Readonly<Record<number, number>> | undefined;
  /** Everything before the rows; the dimension only when the range is known. */
  readonly head: (range: string | undefined, hasRows: boolean) => string;
  /** Everything after the rows. */
  readonly tail: (range: string, hasRows: boolean) => string;
  readonly filterRange: (range: string) => string | undefined;
}

function frame(sheet: StreamedSheet, context: WorksheetContext): Frame {
  const columns = collectColumns(sheet, context.sheetIndex, (style) =>
    context.styles.validate(style),
  );
  const views = renderFreeze(sheet, context.sheetIndex);
  const merges = renderMerges(sheet, context.sheetIndex);
  const rowHeights = collectRowHeights(sheet, context.sheetIndex);
  const filterRange = (range: string): string | undefined =>
    sheet.autoFilter === true && range.length > 0 ? range : undefined;
  return {
    columnStyles: columns.styles,
    rowHeights,
    head: (range, hasRows) =>
      WORKSHEET_OPEN +
      (range === undefined || range.length === 0 ? "" : `<dimension ref="${range}"/>`) +
      views +
      columns.cols +
      (hasRows ? "<sheetData>" : "<sheetData/>"),
    tail: (range, hasRows) => {
      const filter = filterRange(range);
      return (
        (hasRows ? "</sheetData>" : "") +
        (filter === undefined ? "" : `<autoFilter ref="${filter}"/>`) +
        merges +
        "</worksheet>"
      );
    },
    filterRange,
  };
}

/** Renders `xl/worksheets/sheetN.xml`, interning text and registering styles as it goes. */
export function renderWorksheet(sheet: Sheet, context: WorksheetContext): RenderedWorksheet {
  const rows = sheet.rows as Iterable<Row> | AsyncIterable<Row>;
  if (typeof (rows as Iterable<Row>)[Symbol.iterator] !== "function") {
    throw new XlsxMiniError(
      "ASYNC_ROWS",
      `sheets[${context.sheetIndex}].rows`,
      "writeXlsx takes an array or a synchronous iterable of rows; streamXlsx also takes an async iterable",
    );
  }
  const { columnStyles, rowHeights, head, tail, filterRange } = frame(sheet, context);
  const sink = createXmlSink();
  const { range, hasRows } = renderSheetData(sheet.rows, columnStyles, context, sink, rowHeights);
  return {
    data: sink.finish(head(range, hasRows), tail(range, hasRows)),
    filterRange: filterRange(range),
  };
}

/** A worksheet part produced row by row as it is pulled. */
export interface StreamedWorksheet {
  readonly body: ReadableStream<Uint8Array>;
  /** The autofilter range, known once the body has been read to its end. */
  readonly filterRange: () => string | undefined;
}

/**
 * Renders `xl/worksheets/sheetN.xml` as a stream: the head goes out before the first row, without
 * a dimension, and the rows as they are pulled from the sheet, awaiting them only when they are
 * asynchronous. Cancelling the body returns the row iterator.
 */
export function streamWorksheet(
  sheet: StreamedSheet,
  context: WorksheetContext,
): StreamedWorksheet {
  const { columnStyles, rowHeights, head, tail, filterRange } = frame(sheet, context);
  const sink = createChunkSink();
  let headed = false;
  const headedSink = {
    write(text: string) {
      if (!headed) {
        headed = true;
        sink.write(head(undefined, true));
      }
      sink.write(text);
    },
  };
  const writer = createSheetDataWriter(columnStyles, context, headedSink, rowHeights);
  const rows = sheet.rows;
  const asyncRows = (rows as AsyncIterable<Row>)[Symbol.asyncIterator];
  const iterator: Iterator<Row> | AsyncIterator<Row> =
    typeof asyncRows === "function"
      ? asyncRows.call(rows)
      : (rows as Iterable<Row>)[Symbol.iterator]();
  let range: string | undefined;
  let ended = false;

  const finish = (): void => {
    const data = writer.end();
    range = data.range;
    if (headed) sink.write(tail(data.range, true));
    else sink.write(head(undefined, false) + tail(data.range, false));
    sink.end();
    ended = true;
  };

  return {
    body: new ReadableStream<Uint8Array>(
      {
        async pull(controller) {
          for (;;) {
            const chunk = sink.take();
            if (chunk !== undefined) {
              controller.enqueue(chunk);
              return;
            }
            if (ended) {
              controller.close();
              return;
            }
            const next = await iterator.next();
            if (next.done === true) {
              finish();
              continue;
            }
            try {
              writer.row(next.value);
            } catch (error) {
              // An errored stream never has its cancel called: close the row source here.
              await iterator.return?.();
              throw error;
            }
          }
        },
        async cancel() {
          await iterator.return?.();
        },
      },
      { highWaterMark: 0 },
    ),
    filterRange: () => (range === undefined ? undefined : filterRange(range)),
  };
}
