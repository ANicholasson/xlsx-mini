/** Declaration every part starts with. Excel writes `standalone="yes"`; some readers expect it. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const TAB = 9;
const LINE_FEED = 10;
const CARRIAGE_RETURN = 13;
const SPACE = 32;
const DOUBLE_QUOTE = 34;
const AMPERSAND = 38;
const APOSTROPHE = 39;
const LESS_THAN = 60;
const GREATER_THAN = 62;
const HIGH_SURROGATE_START = 0xd800;
const LOW_SURROGATE_START = 0xdc00;
const LOW_SURROGATE_END = 0xdfff;
const NOT_A_CHARACTER = 0xfffe;

/**
 * Makes text safe to place in XML: escapes the five reserved characters and drops what XML 1.0
 * cannot carry (control characters other than tab, newline and carriage return, the two
 * non-characters U+FFFE and U+FFFF, and lone surrogates). Scans once and returns the input
 * untouched when nothing needs changing, which is the common case for cell text.
 *
 * @example escapeXml("Tom & Jerry\u0000") // "Tom &amp; Jerry"
 */
export function escapeXml(text: string): string {
  const length = text.length;
  let index = 0;
  for (; index < length; index++) {
    const code = text.charCodeAt(index);
    if (isReserved(code) || !isPlainValid(code)) break;
  }
  if (index === length) return text;

  let out = "";
  let copiedUpTo = 0;
  for (; index < length; index++) {
    const code = text.charCodeAt(index);
    if (code >= HIGH_SURROGATE_START && code <= LOW_SURROGATE_END) {
      // A high surrogate followed by a low surrogate is one valid character; anything else is dropped.
      const next = index + 1 < length ? text.charCodeAt(index + 1) : 0;
      if (code < LOW_SURROGATE_START && next >= LOW_SURROGATE_START && next <= LOW_SURROGATE_END) {
        index++;
        continue;
      }
      out += text.slice(copiedUpTo, index);
      copiedUpTo = index + 1;
      continue;
    }
    const replacement = replacementFor(code);
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

/** Valid on its own: not a control character, not a non-character, not a surrogate. */
function isPlainValid(code: number): boolean {
  if (code < SPACE) return code === TAB || code === LINE_FEED || code === CARRIAGE_RETURN;
  if (code >= HIGH_SURROGATE_START && code <= LOW_SURROGATE_END) return false;
  return code < NOT_A_CHARACTER;
}

/** The escape for a reserved character, an empty string for an invalid one, undefined to keep. */
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
      return isPlainValid(code) ? undefined : "";
  }
}

function isWhitespace(code: number): boolean {
  return code === SPACE || code === TAB || code === LINE_FEED || code === CARRIAGE_RETURN;
}
