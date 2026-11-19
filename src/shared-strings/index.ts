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

/**
 * Escapes text for XML and then for the format, whose readers take `_xHHHH_` for the character
 * with that code: in text that looks like one, the `x` is written `_x0078_`. Escaping the
 * underscore as `_x005F_` instead reads wrong in Numbers after an `_x` that is not an escape,
 * because Numbers takes that underscore as the end of the `_x` before it. A carriage return is
 * written `&#13;`, because XML reads one written as itself as a line feed. Only text takes that
 * reference: Excel repairs a file whose formula holds it.
 *
 * @example escapeText("_x0041_ & co\r") // "__x0078_0041_ &amp; co&#13;"
 */
export function escapeText(text: string): string {
  const xml = escapeXml(text);
  let out = "";
  let copiedUpTo = 0;
  // The applications decode an escape with a lowercase `x` only.
  for (let index = xml.indexOf("_x"); index !== -1; index = xml.indexOf("_x", index + 2)) {
    if (!startsEscape(xml, index)) continue;
    out += `${xml.slice(copiedUpTo, index + 1)}_x0078_`;
    copiedUpTo = index + 2;
  }
  const escaped = copiedUpTo === 0 ? xml : out + xml.slice(copiedUpTo);
  return escaped.includes("\r") ? escaped.replaceAll("\r", "&#13;") : escaped;
}

function startsEscape(text: string, index: number): boolean {
  for (let offset = 2; offset < 6; offset++) {
    if (!isHexDigit(text.charCodeAt(index + offset))) return false;
  }
  return text.charCodeAt(index + 6) === UNDERSCORE;
}

function isHexDigit(code: number): boolean {
  return (code >= 48 && code <= 57) || (code >= 65 && code <= 70) || (code >= 97 && code <= 102);
}
