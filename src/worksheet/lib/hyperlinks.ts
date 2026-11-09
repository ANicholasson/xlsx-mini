import { XlsxMiniError, describeValue } from "../../errors/index.js";
import { XML_DECLARATION, escapeXml } from "../../xml/index.js";
import { boxProblem, parseReference, quoteSheetName } from "./cell-ref.js";
import type { LinkCell, SheetLinkCell } from "./model.js";

const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/** Collects the sheet's links as its rows are written and renders them once the rows are done. */
export interface Hyperlinks {
  /** Validates the link and remembers it for the cell at `reference`; the path is the cell's. */
  add(reference: string, cell: LinkCell | SheetLinkCell, path: string): void;
  /** The `<hyperlinks>` element, or an empty string when no cell linked anywhere. */
  render(): string;
  /** The sheet's relationships part, or undefined when no link is external. */
  relationships(): string | undefined;
}

interface Link {
  readonly reference: string;
  /** `r:id="rIdN"` for an external target, `location="..."` for a cell in the workbook. */
  readonly target: string;
}

/** Schemes that run code rather than open a document; no spreadsheet link needs them. */
const REFUSED_SCHEMES: ReadonlySet<string> = new Set(["javascript:", "vbscript:", "data:"]);

/** True for an absolute URL whose scheme opens something, such as https, mailto, file or an app's own. */
function isLinkableUrl(href: string): boolean {
  try {
    return !REFUSED_SCHEMES.has(new URL(href).protocol);
  } catch {
    return false;
  }
}

/** The feature's error, at the cell or property that is wrong. */
function invalid(path: string, message: string): XlsxMiniError {
  return new XlsxMiniError("INVALID_HYPERLINK", path, message);
}

/**
 * External targets become relationships of the worksheet part, one per link; in-workbook targets
 * are written as locations and need none. A sheet named in a link must exist in the workbook.
 */
export function createHyperlinks(sheetNames: readonly string[], sheetIndex: number): Hyperlinks {
  const links: Link[] = [];
  const targets: string[] = [];
  let lowerNames: readonly string[] | undefined;
  return {
    add(reference, cell, path) {
      if ("href" in cell) {
        if ("cell" in cell || "sheet" in cell) {
          throw invalid(path, "A link has either href or a cell in the workbook, not both");
        }
        const href = cell.href;
        if (typeof href !== "string" || !isLinkableUrl(href)) {
          throw invalid(
            path + ".href",
            `A link target is an absolute URL such as "https://example.com" that opens a document, not ${describeValue(href)}`,
          );
        }
        targets.push(href);
        links.push({ reference, target: ` r:id="rId${targets.length}"` });
        return;
      }
      const box = parseReference(cell.cell);
      if (box === undefined) {
        throw invalid(
          path + ".cell",
          `A link target is a cell or range in A1 form such as "B2", not ${describeValue(cell.cell)}`,
        );
      }
      const problem = boxProblem(box);
      if (problem !== undefined) throw invalid(path + ".cell", problem);
      let sheet = cell.sheet;
      if (sheet === undefined) {
        sheet = sheetNames[sheetIndex]!;
      } else {
        lowerNames ??= sheetNames.map((name) => name.toLowerCase());
        if (typeof sheet !== "string" || !lowerNames.includes(sheet.toLowerCase())) {
          throw invalid(path + ".sheet", `The workbook has no sheet named ${describeValue(sheet)}`);
        }
      }
      links.push({
        reference,
        target: ` location="${escapeXml(`${quoteSheetName(sheet)}!${box.ref}`)}"`,
      });
    },
    render() {
      if (links.length === 0) return "";
      let xml = "<hyperlinks>";
      for (const link of links) xml += `<hyperlink ref="${link.reference}"${link.target}/>`;
      return `${xml}</hyperlinks>`;
    },
    relationships() {
      if (targets.length === 0) return undefined;
      let xml = `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">`;
      for (let i = 0; i < targets.length; i++) {
        xml += `<Relationship Id="rId${i + 1}" Type="${HYPERLINK_RELATIONSHIP}" Target="${escapeXml(targets[i]!)}" TargetMode="External"/>`;
      }
      return `${xml}</Relationships>`;
    },
  };
}
