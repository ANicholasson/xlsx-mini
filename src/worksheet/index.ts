import { DRAWING_RELATIONSHIP, createSheetDrawing, drawingPath } from "../drawing/index.js";
import { XlsxMiniError } from "../errors/index.js";
import { XML_DECLARATION, createChunkSink, createXmlSink } from "../xml/index.js";
import { createAutoWidths } from "./lib/auto-width.js";
import { createHyperlinks } from "./lib/hyperlinks.js";
import {
  collectColumns,
  collectMerges,
  collectRowHeights,
  mergeStarts,
  renderFreeze,
  renderMerges,
} from "./lib/layout.js";
import type { CellStyle } from "../styles/index.js";
import type { Row, Sheet, StreamedSheet } from "./lib/model.js";
import { renderPrint } from "./lib/print.js";
import { createSheetRelationships } from "./lib/relationships.js";
import {
  type RowContext,
  type RowRecords,
  type WorksheetContext,
  createSheetDataWriter,
  renderSheetData,
} from "./lib/render-rows.js";
import { renderValidations } from "./lib/validations.js";

export type {
  Cell,
  CellObject,
  CellValue,
  Coerce,
  Column,
  DataValidation,
  DateBounds,
  FormulaCell,
  Freeze,
  ImageCell,
  LinkCell,
  NumberBounds,
  PaperSize,
  PrintMargins,
  PrintSetup,
  Register,
  Registered,
  Row,
  Sheet,
  SheetLinkCell,
  StreamedSheet,
  TemporalInstantLike,
  TemporalLike,
  TemporalPlainDateLike,
  TemporalPlainDateTimeLike,
  TemporalZonedDateTimeLike,
  ValidationBase,
  ValidationError,
  ValidationMessage,
} from "./lib/model.js";
export type { WorksheetContext } from "./lib/render-rows.js";
export {
  absolute,
  boxProblem,
  columnLetter,
  parseReference,
  quoteSheetName,
} from "./lib/cell-ref.js";
export { renderPrint } from "./lib/print.js";
export { renderValidations } from "./lib/validations.js";
export {
  type ObjectColumn,
  type SheetFromObjectsOptions,
  sheetFromObjects,
} from "./lib/from-objects.js";
export { serialFromCivil, serialFromEpochMillis } from "./lib/serial.js";

/** What the workbook part needs to know about a rendered sheet. */
interface WorksheetFacts {
  /** The autofilter range such as `A1:C4`, when the sheet filters a non-empty range. */
  readonly filterRange: string | undefined;
  /** The rows repeated on every printed page as `$1:$N`, when the print setup asks for them. */
  readonly printTitles: string | undefined;
  /** True when a cell holds a formula, so the workbook asks Excel to calculate on open. */
  readonly hasFormulas: boolean;
}

/** The sheet's drawing part and its relationships, when any cell holds an image. */
interface SheetDrawingParts {
  /** `xl/drawings/drawingN.xml`. */
  readonly drawing: string;
  /** `xl/drawings/_rels/drawingN.xml.rels`. */
  readonly relationships: string;
}

/** A rendered worksheet part, its relationships part when it has any, its drawing, and its facts. */
export interface RenderedWorksheet extends WorksheetFacts {
  readonly data: Uint8Array;
  /** `xl/worksheets/_rels/sheetN.xml.rels`, when a cell links to an external target or the sheet has a drawing. */
  readonly relationships: string | undefined;
  readonly drawing: SheetDrawingParts | undefined;
}

const WORKSHEET_OPEN = `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

/** When streaming, the rows measured for auto widths before the head goes out. */
const AUTO_WIDTH_SAMPLE_ROWS = 1000;

/** The sheet's settings, validated up front, and the markup around its rows. */
interface Frame {
  readonly columnStyles: readonly (CellStyle | undefined)[];
  readonly rowHeights: Readonly<Record<number, number>> | undefined;
  /** The context the rows render with, carrying what they record: links, formulas, widths. */
  readonly context: RowContext;
  readonly printTitles: string | undefined;
  /** The sheet's relationships part, complete once the tail has been rendered. */
  readonly relationships: () => string | undefined;
  /** The drawing parts, once the rows have been rendered, or undefined without images. */
  readonly drawing: () => SheetDrawingParts | undefined;
  /** Everything before the rows; the dimension only when the range is known. */
  readonly head: (range: string | undefined, hasRows: boolean) => string;
  /** Everything after the rows, in the order the schema fixes. */
  readonly tail: (range: string, hasRows: boolean) => string;
  readonly filterRange: (range: string) => string | undefined;
}

function frame(sheet: StreamedSheet, context: WorksheetContext): Frame {
  const columns = collectColumns(sheet, context.sheetIndex, (style) =>
    context.styles.validate(style),
  );
  const views = renderFreeze(sheet, context.sheetIndex);
  const boxes = collectMerges(sheet, context.sheetIndex);
  const merges = renderMerges(boxes);
  const rowHeights = collectRowHeights(sheet, context.sheetIndex);
  const validations = renderValidations(sheet.validations, context.sheetIndex, context.sheetNames);
  const print = renderPrint(sheet.print, context.sheetIndex);
  const widths =
    columns.measured === undefined
      ? undefined
      : createAutoWidths(columns.measured, columns.beyond, mergeStarts(boxes));
  const relationships = createSheetRelationships();
  const drawing = createSheetDrawing();
  const records: RowRecords = {
    links: createHyperlinks(context.sheetNames, context.sheetIndex, relationships),
    hasFormulas: false,
    widths,
    drawing,
    rowImagePixels: 0,
    imagePixels: [],
  };
  const filterRange = (range: string): string | undefined =>
    sheet.autoFilter === true && range.length > 0 ? range : undefined;
  return {
    columnStyles: columns.styles,
    rowHeights,
    context: { ...context, records },
    printTitles: print.printTitles,
    relationships: () => relationships.render(),
    drawing: () =>
      drawing.hasImages
        ? { drawing: drawing.render(), relationships: drawing.relationships() }
        : undefined,
    head: (range, hasRows) =>
      WORKSHEET_OPEN +
      print.sheetPr +
      (range === undefined || range.length === 0 ? "" : `<dimension ref="${range}"/>`) +
      views +
      columns.cols(widths?.widths(), records.imagePixels) +
      (hasRows ? "<sheetData>" : "<sheetData/>"),
    tail: (range, hasRows) => {
      const filter = filterRange(range);
      return (
        (hasRows ? "</sheetData>" : "") +
        (filter === undefined ? "" : `<autoFilter ref="${filter}"/>`) +
        merges +
        validations +
        records.links.render() +
        print.tail +
        (drawing.hasImages
          ? `<drawing r:id="${relationships.add(DRAWING_RELATIONSHIP, `../${drawingPath(context.sheetIndex + 1).slice(3)}`, false)}"/>`
          : "") +
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
  const {
    columnStyles,
    rowHeights,
    context: rowContext,
    printTitles,
    relationships,
    drawing,
    head,
    tail,
    filterRange,
  } = frame(sheet, context);
  const sink = createXmlSink();
  const { range, hasRows } = renderSheetData(
    sheet.rows,
    columnStyles,
    rowContext,
    sink,
    rowHeights,
  );
  const records = rowContext.records;
  // The tail adds the drawing relationship, so the relationships part is read after it.
  const data = sink.finish(head(range, hasRows), tail(range, hasRows));
  return {
    data,
    relationships: relationships(),
    drawing: drawing(),
    filterRange: filterRange(range),
    printTitles,
    hasFormulas: records.hasFormulas,
  };
}

/** A worksheet part produced row by row as it is pulled. */
export interface StreamedWorksheet {
  readonly body: ReadableStream<Uint8Array>;
  /** The relationships part for external links and the drawing, known once the body has been read to its end. */
  readonly relationships: () => string | undefined;
  /** The drawing parts, known once the body has been read to its end, or undefined without images. */
  readonly drawing: () => SheetDrawingParts | undefined;
  /** What the workbook part needs, known once the body has been read to its end. */
  readonly facts: () => WorksheetFacts;
}

/**
 * Renders `xl/worksheets/sheetN.xml` as a stream: the head goes out before the first row, without
 * a dimension, and the rows as they are pulled from the sheet, awaiting them only when they are
 * asynchronous. When a column sizes itself the head waits for the first rows, held as rendered
 * markup, since `<cols>` precedes the rows and the widths come from them; the rows after the
 * sample change nothing. Cancelling the body returns the row iterator.
 */
export function streamWorksheet(
  sheet: StreamedSheet,
  context: WorksheetContext,
): StreamedWorksheet {
  const {
    columnStyles,
    rowHeights,
    context: rowContext,
    printTitles,
    relationships,
    drawing,
    head,
    tail,
    filterRange,
  } = frame(sheet, context);
  const records = rowContext.records;
  const sink = createChunkSink();
  let headed = false;
  let holding = records.widths !== undefined;
  const held: string[] = [];
  const headedSink = {
    write(text: string) {
      if (holding) {
        held.push(text);
        return;
      }
      if (!headed) {
        headed = true;
        sink.write(head(undefined, true));
      }
      sink.write(text);
    },
  };
  /** Ends the sample: the widths freeze, and the head and the held rows go out together. */
  const release = (): void => {
    if (!holding) return;
    holding = false;
    records.widths?.stop();
    if (held.length === 0) return;
    headed = true;
    sink.write(head(undefined, true));
    sink.write(held.join(""));
    held.length = 0;
  };
  const writer = createSheetDataWriter(columnStyles, rowContext, headedSink, rowHeights);
  const rows = sheet.rows;
  const asyncRows = (rows as AsyncIterable<Row>)[Symbol.asyncIterator];
  const iterator: Iterator<Row> | AsyncIterator<Row> =
    typeof asyncRows === "function"
      ? asyncRows.call(rows)
      : (rows as Iterable<Row>)[Symbol.iterator]();
  let range: string | undefined;
  let ended = false;
  let sampled = 0;

  const finish = (): void => {
    const data = writer.end();
    release();
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
            if (holding && ++sampled >= AUTO_WIDTH_SAMPLE_ROWS) release();
          }
        },
        async cancel() {
          await iterator.return?.();
        },
      },
      { highWaterMark: 0 },
    ),
    relationships,
    drawing,
    facts: () => ({
      filterRange: range === undefined ? undefined : filterRange(range),
      printTitles,
      hasFormulas: records.hasFormulas,
    }),
  };
}
