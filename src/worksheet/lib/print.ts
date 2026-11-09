import { XlsxMiniError, describeValue } from "../../errors/index.js";
import { escapeXml } from "../../xml/index.js";
import { MAX_ROWS } from "./cell-ref.js";
import type { PrintMargins, PrintSetup } from "./model.js";

/** Excel's paper size codes for the names the setup accepts. */
const PAPER_SIZES: Readonly<Record<string, number>> = {
  letter: 1,
  tabloid: 3,
  legal: 5,
  A3: 8,
  A4: 9,
  A5: 11,
};
const ORIENTATIONS: ReadonlySet<string> = new Set(["portrait", "landscape"]);
/** Excel's defaults, in inches, written whenever the sheet has a print setup. */
const DEFAULT_MARGINS: Readonly<Record<keyof PrintMargins, number>> = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
};
const MARGIN_SIDES = ["left", "right", "top", "bottom", "header", "footer"] as const;
/** A margin wider than any paper Excel prints on. */
const MAX_MARGIN = 49;
const MAX_FIT_PAGES = 32_767;
const MAX_HEADER_LENGTH = 255;

/** The print setup as the worksheet part writes it: before the rows, after them, and in the workbook. */
export interface RenderedPrint {
  /** The `<sheetPr>` element when the sheet fits to pages, else an empty string. */
  readonly sheetPr: string;
  /** Margins, page setup and header and footer, after the hyperlinks; empty without a setup. */
  readonly tail: string;
  /** The rows repeated on every page as `$1:$N`, for the workbook's `_xlnm.Print_Titles` name. */
  readonly printTitles: string | undefined;
}

const NONE: RenderedPrint = { sheetPr: "", tail: "", printTitles: undefined };

/** The feature's error, at the property that is wrong. */
function invalid(path: string, message: string): XlsxMiniError {
  return new XlsxMiniError("INVALID_PRINT_SETUP", path, message);
}

/** Validates the print setup up front and renders its elements. */
export function renderPrint(print: PrintSetup | undefined, sheetIndex: number): RenderedPrint {
  if (print === undefined) return NONE;
  const path = `sheets[${sheetIndex}].print`;
  if (typeof print !== "object" || print === null) {
    throw invalid(path, `A print setup is an object, not ${describeValue(print)}`);
  }
  let setup = "";
  const orientation = print.orientation;
  if (orientation !== undefined) {
    if (!ORIENTATIONS.has(orientation)) {
      throw invalid(
        `${path}.orientation`,
        `Orientation is portrait or landscape, not ${describeValue(orientation)}`,
      );
    }
    setup += ` orientation="${orientation}"`;
  }
  const paperSize = print.paperSize;
  if (paperSize !== undefined) {
    const code = Object.hasOwn(PAPER_SIZES, paperSize) ? PAPER_SIZES[paperSize] : undefined;
    if (code === undefined) {
      throw invalid(
        `${path}.paperSize`,
        `Paper sizes are ${Object.keys(PAPER_SIZES).join(", ")}, not ${describeValue(paperSize)}`,
      );
    }
    setup = ` paperSize="${code}"${setup}`;
  }
  let sheetPr = "";
  const fitTo = print.fitTo;
  if (fitTo !== undefined) {
    if (typeof fitTo !== "object" || fitTo === null) {
      throw invalid(
        `${path}.fitTo`,
        `fitTo is an object with width and height in pages, not ${describeValue(fitTo)}`,
      );
    }
    const width = pages(fitTo.width, `${path}.fitTo.width`);
    const height = pages(fitTo.height, `${path}.fitTo.height`);
    if (width === 0 && height === 0) {
      throw invalid(`${path}.fitTo`, "fitTo needs a width, a height or both");
    }
    sheetPr = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';
    setup += ` fitToWidth="${width}" fitToHeight="${height}"`;
  }
  const margins = print.margins;
  if (margins !== undefined && (typeof margins !== "object" || margins === null)) {
    throw invalid(
      `${path}.margins`,
      `Margins are an object with sides in inches, not ${describeValue(margins)}`,
    );
  }
  let marginsXml = "<pageMargins";
  for (const side of MARGIN_SIDES) {
    const value = margins?.[side];
    if (value !== undefined && (typeof value !== "number" || !(value >= 0) || value > MAX_MARGIN)) {
      throw invalid(
        `${path}.margins.${side}`,
        `A margin is 0 to ${MAX_MARGIN} inches, not ${describeValue(value)}`,
      );
    }
    marginsXml += ` ${side}="${value ?? DEFAULT_MARGINS[side]}"`;
  }
  marginsXml += "/>";
  const header = headerText(print.header, `${path}.header`);
  const footer = headerText(print.footer, `${path}.footer`);
  let headerFooter = "";
  if (header !== undefined || footer !== undefined) {
    headerFooter =
      "<headerFooter>" +
      (header === undefined ? "" : `<oddHeader>${header}</oddHeader>`) +
      (footer === undefined ? "" : `<oddFooter>${footer}</oddFooter>`) +
      "</headerFooter>";
  }
  const repeatRows = print.repeatRows;
  let printTitles: string | undefined;
  if (repeatRows !== undefined) {
    if (!Number.isInteger(repeatRows) || repeatRows < 1 || repeatRows > MAX_ROWS) {
      throw invalid(
        `${path}.repeatRows`,
        `repeatRows is a whole number from 1 to ${MAX_ROWS}, not ${describeValue(repeatRows)}`,
      );
    }
    printTitles = `$1:$${repeatRows}`;
  }
  return {
    sheetPr,
    tail: marginsXml + (setup.length === 0 ? "" : `<pageSetup${setup}/>`) + headerFooter,
    printTitles,
  };
}

/** A page count, or 0 for "as many as needed" when the dimension is left out. */
function pages(value: unknown, path: string): number {
  if (value === undefined) return 0;
  if (!Number.isInteger(value) || (value as number) < 1 || (value as number) > MAX_FIT_PAGES) {
    throw invalid(
      path,
      `Pages to fit is a whole number from 1 to ${MAX_FIT_PAGES}, not ${describeValue(value)}`,
    );
  }
  return value as number;
}

/** Header or footer text escaped for XML, or undefined when none is set. */
function headerText(value: unknown, path: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.length > MAX_HEADER_LENGTH) {
    throw invalid(
      path,
      `Header and footer text is a string of at most ${MAX_HEADER_LENGTH} characters`,
    );
  }
  return escapeXml(value);
}
