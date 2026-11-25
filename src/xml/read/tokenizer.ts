import type { XmlCursor } from "./contract.js";
import {
  ATTRIBUTE,
  CDATA,
  JOINED_PIECES,
  TEXT,
  decode,
  detach,
  parseIndex,
  referenced,
  refuse,
} from "./text.js";

export const NEED = 0;
export const OPEN = 1;
export const CLOSE = 2;
const END = 3;
type Token = typeof NEED | typeof OPEN | typeof CLOSE | typeof END;

const SKIPPED = 4;

/** An open element costs memory that no inflate limit sees. */
const MAX_DEPTH = 256;
/**
 * About 128 times the longest text a cell holds. The names kept for the elements last opened at
 * each depth, and the declarations in scope, may hold as many characters together.
 */
const MAX_TOKEN_LENGTH = 4_194_304;
/**
 * What a declaration in scope counts toward those characters besides its prefix and namespace
 * name, for the objects that hold it: fewer than 1 024 fit, and no workbook in the corpora holds
 * more than ten.
 */
const DECLARATION_CHARACTERS = 4096;

const MALFORMED = "a part is not well-formed XML";
const MISMATCHED = "a part closes a tag it did not open";
const UNDEFINED_ENTITY = "a part holds an entity XML does not define";
const DOCTYPE = "a part holds a DOCTYPE";
const TOO_DEEP = "a part nests elements deeper than 256";
const TOO_LONG = "a part holds a token longer than 4 194 304 characters";
const TOO_MANY_NAMES = "a part holds names in scope longer than 4 194 304 characters together";
const UNBOUND = "a part uses a prefix no declaration binds";
const AFTER_ROOT = "a part holds content after its root element";

/** Bound by the standard to `xml`, without a declaration, and to no other prefix. */
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

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

const SPACE = 1;
const NAME_START = 2;
const NAME = 4;
const TEXT_STOP = 8;
const VALUE_STOP = 16;
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
 * Outside ASCII a unit is text and nothing else, so a name holds ASCII only. A surrogate is half
 * of a character XML allows: the decoder gives no other.
 */
function kindOf(code: number): number {
  // The comparison holds the index inside the table.
  return code < 128 ? KINDS[code]! : code < 0xfffe ? 0 : TEXT_STOP | VALUE_STOP | FORBIDDEN;
}

/** The buffer ends in a character XML forbids, so every loop stops there without asking where the end is. */
const SENTINEL = "\0";

const PAIRED_ATTRIBUTES = 8;

/**
 * `marks` holds three numbers for each attribute of a tag under a prefix or named `xmlns`: its
 * place in `spans`, where its prefix ends (0 for `xmlns`), and `DECLARED` or, once its prefix is
 * resolved, `RESOLVED` less the number the reader gives its namespace.
 */
const MARK = 3;
const DECLARED = -1;
const RESOLVED = -2;

/** A prefix, `""` for the default namespace, bound by a declaration in scope. */
interface Binding {
  readonly prefix: string;
  readonly uri: string;
  /** The number the visitor's `namespaces` gives `uri`, 0 when it names none. */
  readonly space: number;
  /** The binding of the same prefix that this one hides until its element closes. */
  readonly hidden: Binding | undefined;
  /** Of the element that declares it, from 0, and -1 for none. */
  readonly depth: number;
}

/**
 * A pull tokenizer over decoded text. It holds the text of at most one unfinished token, the text
 * of the element it is in, the newest chunk, and the namespace declarations in scope; it checks
 * that tags nest and match, that every prefix is bound, and never recurses.
 *
 * It is the cursor too: after `next` returns `OPEN` or `CLOSE`, its `XmlCursor` members describe
 * that token. No object is created per token and none per attribute: a declaration makes one.
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

export function createTokenizer(namespaces: ReadonlyMap<string, number>): Tokenizer {
  let limit = 0;
  let position = 0;
  let dropped = 0;
  let fed = "";
  let ended = false;
  let awaited = 0;

  const names: string[] = [];
  const locals: string[] = [];
  let nameCharacters = 0;
  let depth = 0;
  let rooted = false;
  let selfClosing = false;

  const xml: Binding = {
    prefix: "xml",
    uri: XML_NAMESPACE,
    space: 0,
    hidden: undefined,
    depth: -1,
  };
  /** In the order declared, so the last are those of the innermost element that declares any. */
  const bindings: Binding[] = [xml];
  /** The innermost binding of each prefix in scope. */
  const scope = new Map<string, Binding>([["xml", xml]]);
  let defaultSpace = 0;
  /** Of the innermost open element that declares a namespace, or -1. */
  let declaringDepth = -1;
  let lastLookedUp: Binding | undefined;

  let leaf = false;
  let textful = false;
  let copied = "";
  let joined = "";
  let pieces = 0;
  let runStart = 0;
  let runEnd = 0;
  let runPlain = true;

  const spans: number[] = [];
  let spanCount = 0;
  const seen = new Set<string>();
  const marks: number[] = [];
  let markCount = 0;

  function stuck(at: number, what: string): typeof NEED {
    if (at < limit || ended) refuse(what);
    const unfinished = limit - position;
    if (unfinished > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
    awaited = Math.max(unfinished, 1);
    return NEED;
  }

  function holds(text: string, at: number, word: string, from: number, length: number): boolean {
    for (let offset = 0; offset < length; offset++) {
      if (text.charCodeAt(at + offset) !== word.charCodeAt(from + offset)) return false;
    }
    return true;
  }

  function reference(text: string, at: number): number {
    let end = at + 1;
    let code = text.charCodeAt(end);
    while (code === HASH || (kindOf(code) & NAME) !== 0) code = text.charCodeAt(++end);
    if (code !== SEMICOLON) {
      stuck(end, UNDEFINED_ENTITY);
      return -1;
    }
    if (referenced(text, at + 1, end) < 0) refuse(UNDEFINED_ENTITY);
    return end;
  }

  function find(text: string, at: number, word: string): number {
    const first = word.charCodeAt(0);
    for (;;) {
      const code = text.charCodeAt(at);
      if (code === first && text.startsWith(word, at)) return at;
      if ((kindOf(code) & FORBIDDEN) !== 0) {
        stuck(at, MALFORMED);
        return -1;
      }
      at++;
    }
  }

  function join(text: string, more: string): void {
    joined += detach(
      (runPlain ? text.slice(runStart, runEnd) : decode(text, runStart, runEnd, TEXT)) + more,
    );
    if (++pieces === JOINED_PIECES) {
      copied += detach(joined);
      joined = "";
      pieces = 0;
    }
    runEnd = runStart;
    if (copied.length + joined.length > MAX_TOKEN_LENGTH) refuse(TOO_LONG);
  }

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

  function namespacedAt(key: string, space: number): number {
    const text = cursor.buffer;
    const length = key.length;
    let found = -1;
    // Every index below `markCount` was written by the scan of the open tag, and resolved; once
    // the element closes, `spanCount` is 0 and no attribute is found.
    for (let at = 0; spanCount !== 0 && at < markCount; at += MARK) {
      const span = marks[at]!;
      const local = marks[at + 1]! + 1;
      if (
        marks[at + 2] === RESOLVED - space &&
        spans[span + 1]! - local === length &&
        holds(text, local, key, 0, length)
      ) {
        if (found >= 0) refuse(MALFORMED);
        found = span;
      }
    }
    return found;
  }

  /** Compares with the last prefix found before slicing, so a prefix repeated allocates nothing. */
  function lookup(text: string, start: number, end: number): Binding {
    const length = end - start;
    if (
      lastLookedUp?.prefix.length !== length ||
      !holds(text, start, lastLookedUp.prefix, 0, length)
    ) {
      lastLookedUp = scope.get(text.slice(start, end)) ?? refuse(UNBOUND);
    }
    return lastLookedUp;
  }

  /**
   * Binds the declarations of the open tag, in the order written, for the element at `index`; then
   * resolves the prefix of each other attribute under one. Every index below `markCount` was
   * written by the scan of the tag.
   */
  function resolve(text: string, index: number): void {
    let count = 0;
    for (let at = 0; at < markCount; at += MARK) {
      if (marks[at + 2] !== DECLARED) continue;
      const span = marks[at]!;
      const prefixEnd = marks[at + 1]!;
      const prefix = prefixEnd === 0 ? "" : detach(text.slice(prefixEnd + 1, spans[span + 1]));
      const uri = detach(decode(text, spans[span + 2]!, spans[span + 3]!, ATTRIBUTE));
      // XML Namespaces 1.0, section 3: `xml` binds its own namespace and nothing binds `xmlns`.
      if (
        prefix === "xmlns" ||
        (prefix === "xml") !== (uri === XML_NAMESPACE) ||
        (uri === "" && prefix !== "")
      ) {
        refuse(MALFORMED);
      }
      nameCharacters += prefix.length + uri.length + DECLARATION_CHARACTERS;
      if (nameCharacters > MAX_TOKEN_LENGTH) refuse(TOO_MANY_NAMES);
      const binding = {
        prefix,
        uri,
        space: namespaces.get(uri) ?? 0,
        hidden: scope.get(prefix),
        depth: index,
      };
      bindings.push(binding);
      scope.set(prefix, binding);
      if (prefix === "") defaultSpace = binding.space;
      count++;
    }
    if (count !== 0) {
      declaringDepth = index;
      lastLookedUp = undefined;
    }
    for (let at = 0; at < markCount; at += MARK) {
      if (marks[at + 2] !== DECLARED) {
        marks[at + 2] = RESOLVED - lookup(text, spans[marks[at]!]!, marks[at + 1]!).space;
      }
    }
  }

  function release(): void {
    // `xml`, which no element declares, is always the first.
    while (bindings[bindings.length - 1]!.depth === declaringDepth) {
      const { prefix, uri, hidden } = bindings.pop()!;
      if (hidden === undefined) scope.delete(prefix);
      else scope.set(prefix, hidden);
      nameCharacters -= prefix.length + uri.length + DECLARATION_CHARACTERS;
      if (prefix === "") defaultSpace = hidden?.space ?? 0;
    }
    declaringDepth = bindings[bindings.length - 1]!.depth;
    lastLookedUp = undefined;
  }

  function mark(text: string, start: number, prefixEnd: number, length: number): void {
    const declares =
      (prefixEnd === 0 ? length : prefixEnd - start) === 5 && text.startsWith("xmlns", start);
    if (declares || prefixEnd !== 0) {
      marks[markCount] = spanCount;
      marks[markCount + 1] = prefixEnd;
      marks[markCount + 2] = declares ? DECLARED : 0;
      markCount += MARK;
    }
  }

  function scan(): Token | typeof SKIPPED {
    const text = cursor.buffer;
    let at = position;
    let code = text.charCodeAt(at);

    if (code !== LESS_THAN) {
      if (depth === 0) {
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
      if (copied !== "") {
        joined = copied + joined;
        copied = "";
      }
      if (next - at > MAX_TOKEN_LENGTH || joined.length + runEnd - runStart > MAX_TOKEN_LENGTH) {
        refuse(TOO_LONG);
      }
      cursor.name = locals[depth - 1]!;
      cursor.depth = depth--;
      if (depth === declaringDepth) release();
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

    if (spans.length > 4 * PAIRED_ATTRIBUTES) {
      spans.length = marks.length = 0;
      seen.clear();
    }
    spanCount = markCount = 0;
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

      // Two attributes of one tag written with one name are one too many: a lookup could not tell
      // them apart. Every index below `spanCount` was written for an attribute before this one.
      const length = keyEnd - start;
      if (spanCount < 4 * PAIRED_ATTRIBUTES) {
        for (let other = 0; other < spanCount; other += 4) {
          const otherStart = spans[other]!;
          if (
            spans[other + 1]! - otherStart === length &&
            holds(text, start, text, otherStart, length)
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
        const name = text.slice(start, keyEnd);
        if (seen.has(name)) refuse(MALFORMED);
        seen.add(name);
      }
      if (prefixEnd !== 0 || length === 5) mark(text, start, prefixEnd, length);
      spans[spanCount] = start;
      spans[spanCount + 1] = keyEnd;
      spans[spanCount + 2] = valueStart;
      spans[spanCount + 3] = end;
      spanCount += 4;
      code = text.charCodeAt(++end);
    }
    next = end + 1;
    if (next - at > MAX_TOKEN_LENGTH) refuse(TOO_LONG);

    if (declaration) {
      // `version="1.0"`, then an encoding that is UTF-8, then `standalone`: in that order. `value`
      // is asked only for an attribute `attributeAt` found.
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
    // A declaration holds for the tag it stands on, wherever in the tag, and for what it holds.
    if (markCount !== 0) resolve(text, depth);
    // Siblings repeat their name: the element last opened at this depth lends its strings.
    const sibling = names[depth];
    const length = nameEnd - nameStart;
    if (sibling?.length !== length || !holds(text, nameStart, sibling, 0, length)) {
      nameCharacters += length - (sibling?.length ?? 0);
      if (nameCharacters > MAX_TOKEN_LENGTH) refuse(TOO_MANY_NAMES);
      const name = detach(text.slice(nameStart, nameEnd));
      names[depth] = name;
      locals[depth] = colon === 0 ? name : detach(text.slice(colon + 1, nameEnd));
    }
    cursor.namespace = colon === 0 ? defaultSpace : lookup(text, nameStart, colon).space;
    // The two lists were written just now or when that sibling opened.
    cursor.name = locals[depth]!;
    cursor.depth = ++depth;
    rooted = true;
    selfClosing = empty;
    leaf = true;
    copied = joined = "";
    runEnd = runStart;
    position = next;
    return OPEN;
  }

  const cursor = {
    name: "",
    namespace: 0,
    depth: 0,
    valueEnd: 0,
    buffer: SENTINEL,

    // Each lookup gives the first of the four numbers the scan of the open tag wrote, or -1.
    attribute(name: string, namespace = 0): string | undefined {
      const at = namespace === 0 ? attributeAt(name) : namespacedAt(name, namespace);
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
        if (depth === declaringDepth) release();
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
