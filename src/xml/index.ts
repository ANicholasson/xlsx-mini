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

/** The first block of a sink; blocks double up to the cap, so small parts stay small. */
const FIRST_BLOCK_SIZE = 16_384;
const MAX_BLOCK_SIZE = 1_048_576;
/** Pending text is encoded once it reaches this many UTF-16 units: few calls, short-lived ropes. */
const FLUSH_AT = 65_536;

const utf8 = new TextEncoder();

/**
 * Collects the text of one part and encodes it as it comes, so a large part never exists as one
 * JavaScript string. `write` takes whole characters; `finish` wraps the body in the text that is
 * only known at the end, such as a dimension computed from the rows, and returns the bytes.
 */
export interface XmlSink extends TextSink {
  finish(prefix: string, suffix: string): Uint8Array;
}

/** Where rendered text goes; the renderers need nothing more than `write`. */
export interface TextSink {
  write(text: string): void;
}

/** The most bytes one streamed chunk carries: one stored deflate block where that fallback applies. */
export const MAX_CHUNK_SIZE = 65_535;
/** A chunk sink encodes once pending text could fill a chunk: three bytes per UTF-16 unit at most. */
const CHUNK_FLUSH_AT = 21_845;

/**
 * Collects the text of one streamed part and hands it out as encoded chunks of at most
 * `MAX_CHUNK_SIZE` bytes: `take` returns the next ready chunk or `undefined`, `end` encodes what
 * is left so the last `take` calls drain it.
 */
export interface ChunkSink extends TextSink {
  take(): Uint8Array | undefined;
  end(): void;
}

export function createChunkSink(): ChunkSink {
  const ready: Uint8Array[] = [];
  let next = 0;
  let pending = "";
  let chunk = new Uint8Array(MAX_CHUNK_SIZE);
  let used = 0;

  const encode = (): void => {
    let text = pending;
    pending = "";
    while (text.length > 0) {
      const { read, written } = utf8.encodeInto(text, chunk.subarray(used));
      used += written;
      if (read === text.length) break;
      // The chunk cannot take the next character: it is ready, start another.
      ready.push(chunk.subarray(0, used));
      chunk = new Uint8Array(MAX_CHUNK_SIZE);
      used = 0;
      text = text.slice(read);
    }
  };

  return {
    write(text) {
      pending += text;
      if (pending.length >= CHUNK_FLUSH_AT) encode();
    },
    take() {
      if (next < ready.length) {
        const out = ready[next]!;
        ready[next] = undefined as never;
        next++;
        return out;
      }
      return undefined;
    },
    end() {
      encode();
      if (used > 0) ready.push(chunk.subarray(0, used));
      chunk = new Uint8Array(0);
      used = 0;
    },
  };
}

export function createXmlSink(): XmlSink {
  const blocks: Uint8Array[] = [];
  let blockSize = FIRST_BLOCK_SIZE;
  let block = new Uint8Array(blockSize);
  let used = 0;
  let total = 0;
  let pending = "";

  const flush = (): void => {
    let text = pending;
    pending = "";
    while (text.length > 0) {
      const { read, written } = utf8.encodeInto(text, block.subarray(used));
      used += written;
      total += written;
      if (read === text.length) break;
      // The block cannot take the next character: seal it and start a bigger one.
      blocks.push(block.subarray(0, used));
      blockSize = Math.min(blockSize * 2, MAX_BLOCK_SIZE);
      block = new Uint8Array(blockSize);
      used = 0;
      text = text.slice(read);
    }
  };

  return {
    write(text) {
      pending += text;
      if (pending.length >= FLUSH_AT) flush();
    },
    finish(prefix, suffix) {
      flush();
      const head = utf8.encode(prefix);
      const tail = utf8.encode(suffix);
      const out = new Uint8Array(head.length + total + tail.length);
      out.set(head, 0);
      let offset = head.length;
      for (let i = 0; i < blocks.length; i++) {
        const sealed = blocks[i]!;
        out.set(sealed, offset);
        offset += sealed.length;
      }
      out.set(block.subarray(0, used), offset);
      out.set(tail, offset + used);
      return out;
    },
  };
}
