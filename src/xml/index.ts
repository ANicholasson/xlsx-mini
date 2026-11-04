/** Declaration every part starts with. Excel writes `standalone="yes"`; some readers expect it. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const AMPERSAND = 38;
const LESS_THAN = 60;
const GREATER_THAN = 62;
const DOUBLE_QUOTE = 34;
const APOSTROPHE = 39;

/**
 * Escapes the five characters XML reserves, for use in both text and attribute values.
 * Scans once and returns the input untouched when nothing needs escaping, which is the
 * common case for cell text.
 *
 * @example escapeXml('Tom & Jerry') // 'Tom &amp; Jerry'
 */
export function escapeXml(text: string): string {
  const length = text.length;
  let index = 0;
  for (; index < length; index++) {
    if (isReserved(text.charCodeAt(index))) break;
  }
  if (index === length) return text;

  let out = "";
  let copiedUpTo = 0;
  for (; index < length; index++) {
    const replacement = replacementFor(text.charCodeAt(index));
    if (replacement === undefined) continue;
    out += text.slice(copiedUpTo, index) + replacement;
    copiedUpTo = index + 1;
  }
  return out + text.slice(copiedUpTo);
}

/** True when leading or trailing whitespace must survive, so the `<t>` needs `xml:space="preserve"`. */
export function needsSpacePreserve(text: string): boolean {
  if (text.length === 0) return false;
  return isWhitespace(text.charCodeAt(0)) || isWhitespace(text.charCodeAt(text.length - 1));
}

function isReserved(code: number): boolean {
  return (
    code === AMPERSAND ||
    code === LESS_THAN ||
    code === GREATER_THAN ||
    code === DOUBLE_QUOTE ||
    code === APOSTROPHE
  );
}

function replacementFor(code: number): string | undefined {
  switch (code) {
    case AMPERSAND:
      return "&amp;";
    case LESS_THAN:
      return "&lt;";
    case GREATER_THAN:
      return "&gt;";
    case DOUBLE_QUOTE:
      return "&quot;";
    case APOSTROPHE:
      return "&apos;";
    default:
      return undefined;
  }
}

function isWhitespace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13;
}
