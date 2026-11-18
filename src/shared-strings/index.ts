import { XML_DECLARATION, escapeXml, needsSpacePreserve } from "../xml/index.js";

/** Interns cell text into the workbook's shared-strings table and renders that part. */
export interface StringTable {
  /** Returns the zero-based index to reference from a cell, adding the text on first sight. */
  intern(text: string): number;
  /** Renders `xl/sharedStrings.xml`. Call once, after every sheet has been rendered. */
  render(): string;
}

export function createStringTable(): StringTable {
  const indexes = new Map<string, number>();
  const items: string[] = [];
  let references = 0;

  return {
    intern(text) {
      references++;
      const existing = indexes.get(text);
      if (existing !== undefined) return existing;
      const index = items.length;
      indexes.set(text, index);
      items.push(renderItem(text));
      return index;
    },
    render() {
      return `${XML_DECLARATION}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${references}" uniqueCount="${items.length}">${items.join("")}</sst>`;
    },
  };
}

function renderItem(text: string): string {
  return needsSpacePreserve(text)
    ? `<si><t xml:space="preserve">${escapeText(text)}</t></si>`
    : `<si><t>${escapeText(text)}</t></si>`;
}

const UNDERSCORE = 95;
const LOWERCASE_X = 120;

/**
 * Escapes text for XML and then for the format, whose readers take `_xHHHH_` for the character
 * with that code: an underscore that would start one is written `_x005F_`.
 *
 * @example escapeText("_x0041_ & co") // "_x005F_x0041_ &amp; co"
 */
export function escapeText(text: string): string {
  // XML first: a character it drops could otherwise join the pieces of an escape.
  const xml = escapeXml(text);
  let out = "";
  let copiedUpTo = 0;
  let closing = -1;
  for (let index = 0; index < xml.length; index++) {
    if (xml.charCodeAt(index) !== UNDERSCORE || !startsEscape(xml, index)) continue;
    if (index === closing) {
      // Numbers reads `_x005F_xHHHH_` as one piece, this underscore included, so the escape that
      // starts here keeps it and has its `x` escaped instead.
      out += `${xml.slice(copiedUpTo, index + 1)}_x0078_`;
      copiedUpTo = index + 2;
    } else {
      out += `${xml.slice(copiedUpTo, index)}_x005F_`;
      copiedUpTo = index + 1;
      closing = index + 6;
    }
  }
  return copiedUpTo === 0 ? xml : out + xml.slice(copiedUpTo);
}

/** The applications decode an escape with a lowercase `x` only. */
function startsEscape(text: string, index: number): boolean {
  if (text.charCodeAt(index + 1) !== LOWERCASE_X) return false;
  for (let offset = 2; offset < 6; offset++) {
    if (!isHexDigit(text.charCodeAt(index + offset))) return false;
  }
  return text.charCodeAt(index + 6) === UNDERSCORE;
}

function isHexDigit(code: number): boolean {
  return (code >= 48 && code <= 57) || (code >= 65 && code <= 70) || (code >= 97 && code <= 102);
}
