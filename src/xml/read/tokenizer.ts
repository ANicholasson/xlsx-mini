import type { XmlCursor } from "./contract.js";
import { ATTRIBUTE, CDATA, TEXT, decode, detach, parseIndex, referenced, refuse } from "./text.js";

/** What `Tokenizer.next` returns. Numbers, not strings: compared once per token. */
export const NEED = 0;
export const OPEN = 1;
export const CLOSE = 2;
const END = 3;
type Token = typeof NEED | typeof OPEN | typeof CLOSE | typeof END;

/** Markup that is no token for a visitor has been read: a comment, an instruction, CDATA. */
const SKIPPED = 4;

/** Elements may nest this deep. An open element costs memory that no inflate limit sees. */
const MAX_DEPTH = 256;
/** One tag, one run of text or one comment may be this many characters: 128 times the longest cell text. */
const MAX_TOKEN_LENGTH = 4_194_304;

const MALFORMED = "a part is not well-formed XML";
const MISMATCHED = "a part closes a tag it did not open";
const UNDEFINED_ENTITY = "a part holds an entity XML does not define";
const DOCTYPE = "a part holds a DOCTYPE";
const TOO_DEEP = "a part nests elements deeper than 256";
const TOO_LONG = "a part holds a token longer than 4 194 304 characters";
const AFTER_ROOT = "a part holds content after its root element";

const EXCLAMATION = 33;
const DOUBLE_QUOTE = 34;
const HASH = 35;
const AMPERSAND = 38;
const APOSTROPHE = 39;
const SLASH = 47;
const COLON = 58;
const SEMICOLON = 59;
const LESS_THAN = 60;
const EQUALS = 61;
const GREATER_THAN = 62;
const QUESTION = 63;
const RIGHT_BRACKET = 93;

/** What a character can be, as flags: one lookup per character answers every loop below. */
const SPACE = 1;
const NAME_START = 2;
const NAME = 4;
/** Ends a run of character data, or needs a look: `<`, `&`, `>`, a carriage return, a forbidden character. */
const TEXT_STOP = 8;
/** Ends an attribute value, or needs a look: a quote, `<`, `&`, a forbidden character. */
const VALUE_STOP = 16;
/** XML 1.0 allows no document to hold it: a control character other than tab and the line ends. */
const FORBIDDEN = 32;

const KINDS = new Uint8Array(128);
KINDS.fill(TEXT_STOP | VALUE_STOP | FORBIDDEN, 0, 32);
KINDS[9] = KINDS[10] = KINDS[32] = SPACE;
KINDS[13] = SPACE | TEXT_STOP;
KINDS[DOUBLE_QUOTE] = KINDS[APOSTROPHE] = VALUE_STOP;
KINDS[AMPERSAND] = KINDS[LESS_THAN] = TEXT_STOP | VALUE_STOP;
KINDS[GREATER_THAN] = TEXT_STOP;
KINDS[45] = KINDS[46] = NAME;
KINDS.fill(NAME, 48, 58);
KINDS.fill(NAME_START | NAME, 65, 91);
KINDS[95] = NAME_START | NAME;
KINDS.fill(NAME_START | NAME, 97, 123);

/**
 * The flags of one UTF-16 unit. Outside ASCII a unit is text and nothing else, so a name holds
 * ASCII only; U+FFFE and U+FFFF are forbidden. A surrogate is half of a character XML allows:
 * the decoder gives no other.
 */
function kindOf(code: number): number {
  // The comparison holds the index inside the table.
  return code < 128 ? KINDS[code]! : code < 0xfffe ? 0 : TEXT_STOP | VALUE_STOP | FORBIDDEN;
}

/** The buffer ends in a character XML forbids, so every loop stops there without asking where the end is. */
const SENTINEL = "\0";

/** Below this many attributes a tag's names are compared in pairs; from there on, through a set. */
const PAIRED_ATTRIBUTES = 8;

/**
 * A pull tokenizer over decoded text. It holds the text of at most one unfinished token, the text
 * of the element it is in, and the newest chunk; it checks that tags nest and match, and never
 * recurses.
 *
 * It is the cursor too: after `next` returns `OPEN` or `CLOSE`, its `XmlCursor` members describe
 * that token. No object is created per token and none per attribute.
 */
export interface Tokenizer extends XmlCursor {
  /**
   * The next token. `NEED` means the buffer ends inside a token: `feed` more and ask again, which
   * is the only point where the driver awaits. `END` comes once, after the root has closed and the
   * end of the part has been fed.
   *
   * Text is never a token of its own: it belongs to the `CLOSE` of the element it is directly
   * inside, so a visitor has one path for it.
   */
  next(): Token;
  /**
   * Appends decoded text, or marks the end of the part with `undefined`. An unfinished token is
   * read again only once as much new text has come as the buffer held of it, so a token of any
   * length costs linear time.
   */
  feed(text: string | undefined): void;
}

export function createTokenizer(): Tokenizer {
  /** Where the sentinel stands in the buffer: one past the last character fed. */
  let limit = 0;
  /** The next character to read. Everything before it is read, except the text run of a leaf. */
  let position = 0;
  /** The characters dropped from the front of the buffer so far. */
  let dropped = 0;
  /** Text fed since the buffer was last read. */
  let fed = "";
  let ended = false;
  /** The characters `fed` must hold before an unfinished token is read again. */
  let awaited = 0;

  /** The open elements by their names as written, which a close tag must repeat, and by local name. */
  const names: string[] = [];
  const locals: string[] = [];
  let depth = 0;
  let rooted = false;
  let selfClosing = false;

  /** The open element has no child yet, so the character data read so far is its text. */
  let leaf = false;
  /** The element the last `CLOSE` closed had no child. */
  let textful = false;
  /** The leaf's text before the comment, instruction or CDATA section that split it, decoded. */
  let joined = "";
  /** The leaf's last run of character data, as written, and whether it reads as written. */
  let runStart = 0;
  let runEnd = 0;
  let runPlain = true;

  /** The attributes of the open tag, four numbers each: where its key and its value start and end. */
  const spans: number[] = [];
  let spanCount = 0;
  const seen = new Set<string>();

  /**
   * The buffer cannot hold a token at `at`. When that is its end and more text may come, `NEED`;
   * otherwise the part is refused.
   */
  function stuck(at: number, what: string): typeof NEED {
    if (at < limit || ended) refuse(what);
    const unfinished = limit - position;
    if (unfinished > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
    awaited = Math.max(unfinished, 1);
    return NEED;
  }

  /** Whether `text` holds `word` at `at`. */
  function holds(text: string, at: number, word: string, from: number, length: number): boolean {
    for (let offset = 0; offset < length; offset++) {
      if (text.charCodeAt(at + offset) !== word.charCodeAt(from + offset)) return false;
    }
    return true;
  }

  /** Where the reference that starts at `at` ends, at its `;`, or -1 when the buffer ends first. */
  function reference(text: string, at: number): number {
    let end = at + 1;
    let code = text.charCodeAt(end);
    while (code === HASH || (kindOf(code) & NAME) !== 0) code = text.charCodeAt(++end);
    if (code !== SEMICOLON) return stuck(end, UNDEFINED_ENTITY) - 1;
    if (referenced(text, at + 1, end) < 0) refuse(UNDEFINED_ENTITY);
    return end;
  }

  /** Where `word` next stands from `at` on, past characters XML allows, or -1 when the buffer ends first. */
  function find(text: string, at: number, word: string): number {
    const first = word.charCodeAt(0);
    for (;;) {
      const code = text.charCodeAt(at);
      if (code === first && text.startsWith(word, at)) return at;
      if ((kindOf(code) & FORBIDDEN) !== 0) return stuck(at, MALFORMED) - 1;
      at++;
    }
  }

  /** The leaf's run of character data joins the text before it: what follows splits its text. */
  function join(text: string, more: string): void {
    joined +=
      (runPlain ? text.slice(runStart, runEnd) : decode(text, runStart, runEnd, TEXT)) + more;
    runEnd = runStart;
    if (joined.length > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
  }

  /** The index in `spans` of the attribute with this key, or -1. */
  function attributeAt(key: string): number {
    const text = cursor.buffer;
    const length = key.length;
    // Every index below `spanCount` was written by the scan of the open tag.
    for (let at = 0; at < spanCount; at += 4) {
      const start = spans[at]!;
      if (spans[at + 1]! - start === length && holds(text, start, key, 0, length)) return at;
    }
    return -1;
  }

  /**
   * Reads on from `position`: character data up to the next markup, then that markup. Returns the
   * token it completes, `SKIPPED` for markup a visitor never sees, or `NEED` with nothing read.
   */
  function scan(): Token | typeof SKIPPED {
    const text = cursor.buffer;
    let at = position;
    let code = text.charCodeAt(at);

    if (code !== LESS_THAN) {
      if (depth === 0) {
        // Outside the root only white space is character data.
        while ((kindOf(code) & SPACE) !== 0) code = text.charCodeAt(++at);
        position = at;
        if (code !== LESS_THAN) {
          if (at < limit || !ended) return stuck(at, rooted ? AFTER_ROOT : MALFORMED);
          if (!rooted) refuse(MALFORMED);
          return END;
        }
      } else {
        const start = at;
        let plain = true;
        for (;;) {
          if ((kindOf(code) & TEXT_STOP) === 0) {
            code = text.charCodeAt(++at);
            continue;
          }
          if (code === LESS_THAN) break;
          plain = false;
          if (code === AMPERSAND) {
            at = reference(text, at);
            if (at < 0) return NEED;
          } else if (code === GREATER_THAN) {
            // Character data may not hold the end of a CDATA section.
            if (
              text.charCodeAt(at - 1) === RIGHT_BRACKET &&
              text.charCodeAt(at - 2) === RIGHT_BRACKET
            ) {
              refuse(MALFORMED);
            }
          } else if ((kindOf(code) & FORBIDDEN) !== 0) {
            return stuck(at, MALFORMED);
          }
          code = text.charCodeAt(++at);
        }
        if (at - start > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
        if (leaf) {
          runStart = start;
          runEnd = at;
          runPlain = plain;
        }
        position = at;
      }
    }

    const second = text.charCodeAt(at + 1);
    let next: number;

    if (second === SLASH) {
      if (depth === 0) refuse(MISMATCHED);
      // An element is open, so its names are in the two lists.
      const name = names[depth - 1]!;
      let end = at + 2;
      for (let offset = 0; offset < name.length; offset++, end++) {
        if (text.charCodeAt(end) !== name.charCodeAt(offset)) return stuck(end, MISMATCHED);
      }
      code = text.charCodeAt(end);
      while ((kindOf(code) & SPACE) !== 0) code = text.charCodeAt(++end);
      if (code !== GREATER_THAN) return stuck(end, MISMATCHED);
      next = end + 1;
      if (next - at > MAX_TOKEN_LENGTH || joined.length + runEnd - runStart > MAX_TOKEN_LENGTH) {
        refuse(TOO_LONG);
      }
      cursor.name = locals[depth - 1]!;
      cursor.depth = depth--;
      textful = leaf;
      leaf = false;
      spanCount = 0;
      position = next;
      return CLOSE;
    }

    if (second === EXCLAMATION) {
      const comment = text.charCodeAt(at + 2) !== 91;
      const opening = comment ? "--" : "[CDATA[";
      for (let offset = 0; offset < opening.length; offset++) {
        if (text.charCodeAt(at + 2 + offset) !== opening.charCodeAt(offset)) {
          return stuck(at + 2 + offset, DOCTYPE);
        }
      }
      if (!comment && depth === 0) refuse(rooted ? AFTER_ROOT : MALFORMED);
      const start = at + 2 + opening.length;
      const end = find(text, start, comment ? "--" : "]]>");
      if (end < 0) return NEED;
      // A comment ends at its first two hyphens: XML allows none inside one.
      if (comment && text.charCodeAt(end + 2) !== GREATER_THAN) return stuck(end + 2, MALFORMED);
      next = end + 3;
      if (next - at > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
      if (leaf) join(text, comment ? "" : decode(text, start, end, CDATA));
      position = next;
      return SKIPPED;
    }

    let nameStart = at + 1;
    let declaration = false;
    if (second === QUESTION) {
      // A processing instruction, or at the very start of the part the XML declaration.
      let end = at + 2;
      code = text.charCodeAt(end);
      if ((kindOf(code) & NAME_START) === 0) return stuck(end, MALFORMED);
      do code = text.charCodeAt(++end);
      while ((kindOf(code) & NAME) !== 0);
      if (code !== QUESTION && (kindOf(code) & SPACE) === 0) return stuck(end, MALFORMED);
      const reserved =
        end - at === 5 &&
        (text.charCodeAt(at + 2) | 32) === 120 &&
        (text.charCodeAt(at + 3) | 32) === 109 &&
        (text.charCodeAt(at + 4) | 32) === 108;
      if (!reserved) {
        const close = find(text, end, "?>");
        if (close < 0) return NEED;
        if (code === QUESTION && close !== end) refuse(MALFORMED);
        next = close + 2;
        if (next - at > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
        if (leaf) join(text, "");
        position = next;
        return SKIPPED;
      }
      if (dropped + at !== 0 || !text.startsWith("xml", at + 2)) refuse(MALFORMED);
      declaration = true;
      nameStart = at + 2;
    } else if ((kindOf(second) & NAME_START) === 0) {
      return stuck(nameStart, MALFORMED);
    } else if (depth === 0 && rooted) {
      refuse(AFTER_ROOT);
    }

    // An open tag, an empty element, or the declaration, which is written like one.
    let end = nameStart;
    let colon = 0;
    for (;;) {
      code = text.charCodeAt(++end);
      if ((kindOf(code) & NAME) !== 0) continue;
      if (code !== COLON || colon !== 0) break;
      colon = end;
    }
    const nameEnd = end;
    if (colon === nameEnd - 1) return stuck(end, MALFORMED);

    spanCount = 0;
    let empty = false;
    for (;;) {
      const gap = end;
      while ((kindOf(code) & SPACE) !== 0) code = text.charCodeAt(++end);
      if (code === GREATER_THAN && !declaration) break;
      if (code === (declaration ? QUESTION : SLASH)) {
        code = text.charCodeAt(++end);
        if (code !== GREATER_THAN) return stuck(end, MALFORMED);
        empty = true;
        break;
      }
      // An attribute stands after white space.
      if (end === gap || (kindOf(code) & NAME_START) === 0) return stuck(end, MALFORMED);
      const start = end;
      let prefixEnd = 0;
      for (;;) {
        code = text.charCodeAt(++end);
        if ((kindOf(code) & NAME) !== 0) continue;
        if (code !== COLON || prefixEnd !== 0) break;
        prefixEnd = end;
      }
      const keyEnd = end;
      if (prefixEnd === keyEnd - 1) return stuck(end, MALFORMED);
      // A prefix is dropped, but not the one that makes the name a namespace declaration, and not
      // in the declaration, whose three names have none.
      const key =
        prefixEnd === 0 ||
        declaration ||
        (prefixEnd - start === 5 && text.startsWith("xmlns", start))
          ? start
          : prefixEnd + 1;
      while ((kindOf(code) & SPACE) !== 0) code = text.charCodeAt(++end);
      if (code !== EQUALS) return stuck(end, MALFORMED);
      do code = text.charCodeAt(++end);
      while ((kindOf(code) & SPACE) !== 0);
      if (code !== DOUBLE_QUOTE && code !== APOSTROPHE) return stuck(end, MALFORMED);
      const quote = code;
      const valueStart = end + 1;
      for (;;) {
        code = text.charCodeAt(++end);
        if ((kindOf(code) & VALUE_STOP) === 0) continue;
        if (code === quote) break;
        if (code === AMPERSAND) {
          end = reference(text, end);
          if (end < 0) return NEED;
        } else if (code !== DOUBLE_QUOTE && code !== APOSTROPHE) {
          return stuck(end, MALFORMED);
        }
      }

      // Two attributes of one tag with one key are one too many: a lookup could not tell them apart.
      const length = keyEnd - key;
      if (spanCount < 4 * PAIRED_ATTRIBUTES) {
        for (let other = 0; other < spanCount; other += 4) {
          const otherStart = spans[other]!;
          if (
            spans[other + 1]! - otherStart === length &&
            holds(text, key, text, otherStart, length)
          ) {
            refuse(MALFORMED);
          }
        }
      } else {
        if (spanCount === 4 * PAIRED_ATTRIBUTES) {
          seen.clear();
          for (let other = 0; other < spanCount; other += 4) {
            seen.add(text.slice(spans[other]!, spans[other + 1]!));
          }
        }
        const name = text.slice(key, keyEnd);
        if (seen.has(name)) refuse(MALFORMED);
        seen.add(name);
      }
      spans[spanCount] = key;
      spans[spanCount + 1] = keyEnd;
      spans[spanCount + 2] = valueStart;
      spans[spanCount + 3] = end;
      spanCount += 4;
      code = text.charCodeAt(++end);
    }
    next = end + 1;
    if (next - at > MAX_TOKEN_LENGTH) refuse(TOO_LONG);

    if (declaration) {
      // `version="1.0"`, then an encoding that is UTF-8, then `standalone`: in that order.
      const encoding = attributeAt("encoding");
      const standalone = attributeAt("standalone");
      const value = (span: number): string => text.slice(spans[span + 2]!, spans[span + 3]!);
      if (
        attributeAt("version") !== 0 ||
        value(0) !== "1.0" ||
        (encoding >= 0 && (encoding !== 4 || value(4).toLowerCase() !== "utf-8")) ||
        (standalone >= 0 &&
          (standalone !== spanCount - 4 ||
            (value(standalone) !== "yes" && value(standalone) !== "no"))) ||
        spanCount !== 4 + (encoding >= 0 ? 4 : 0) + (standalone >= 0 ? 4 : 0)
      ) {
        refuse(MALFORMED);
      }
      spanCount = 0;
      position = next;
      return SKIPPED;
    }

    if (depth === MAX_DEPTH) refuse(TOO_DEEP);
    // Siblings repeat their name: the element last opened at this depth lends its strings.
    const sibling = names[depth];
    const length = nameEnd - nameStart;
    if (sibling?.length !== length || !holds(text, nameStart, sibling, 0, length)) {
      const name = detach(text.slice(nameStart, nameEnd));
      names[depth] = name;
      locals[depth] = colon === 0 ? name : detach(text.slice(colon + 1, nameEnd));
    }
    // The two lists were written just now or when that sibling opened.
    cursor.name = locals[depth]!;
    cursor.depth = ++depth;
    rooted = true;
    selfClosing = empty;
    leaf = true;
    joined = "";
    runEnd = runStart;
    position = next;
    return OPEN;
  }

  const cursor = {
    name: "",
    depth: 0,
    valueEnd: 0,
    buffer: SENTINEL,

    attribute(name: string): string | undefined {
      const at = attributeAt(name);
      if (at < 0) return undefined;
      return detach(decode(cursor.buffer, spans[at + 2]!, spans[at + 3]!, ATTRIBUTE));
    },

    index(name: string): number {
      const at = attributeAt(name);
      return at < 0 ? -1 : parseIndex(cursor.buffer, spans[at + 2]!, spans[at + 3]!);
    },

    valueStart(name: string): number {
      const at = attributeAt(name);
      if (at < 0) return -1;
      cursor.valueEnd = spans[at + 3]!;
      return spans[at + 2]!;
    },

    text(): string {
      if (!textful) return "";
      const text = cursor.buffer;
      return detach(
        joined + (runPlain ? text.slice(runStart, runEnd) : decode(text, runStart, runEnd, TEXT)),
      );
    },

    next(): Token {
      if (selfClosing) {
        selfClosing = false;
        cursor.depth = depth--;
        textful = true;
        leaf = false;
        spanCount = 0;
        return CLOSE;
      }
      if (textful) {
        textful = false;
        joined = "";
        runEnd = runStart;
      }
      if (fed.length < awaited && !ended) return NEED;
      awaited = 0;
      if (fed !== "") {
        // The run of a leaf stays until its element closes: its text is cut from the buffer then.
        const kept = runEnd > runStart ? runStart : position;
        // Joined, not added: V8 reads every character of a string made with `+` through the
        // string it was made of, one step more on each character the tokenizer reads.
        cursor.buffer = [cursor.buffer.slice(kept, limit), fed, SENTINEL].join("");
        fed = "";
        dropped += kept;
        position -= kept;
        runStart -= kept;
        runEnd -= kept;
        limit = cursor.buffer.length - 1;
      }
      for (;;) {
        const token = scan();
        if (token !== SKIPPED) return token;
      }
    },

    feed(text: string | undefined): void {
      if (text === undefined) ended = true;
      else fed += text;
    },
  };
  return cursor;
}
