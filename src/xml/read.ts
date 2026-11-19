/**
 * The read side of XML: turns the bytes of one part into open and close tokens for a part reader.
 * Knows nothing about spreadsheets and nothing about packages.
 *
 * A part reader is a synchronous visitor. It never sees bytes, a promise or a chunk edge, so it
 * cannot inflate, cannot wait and cannot be handed half a token. What XML itself can do to a
 * reader is refused here with `INVALID_XML`, for every part alike: a DOCTYPE, text that is not
 * valid UTF-8, nesting deeper than 256, a token longer than 4 194 304 characters, tags that do
 * not match, an entity XML does not define, content after the root. Those two numbers are not
 * limits a caller sets: no workbook comes near them.
 */
import type { ByteChunks, XmlVisitor } from "./read/contract.js";
import { textChunks } from "./read/text.js";
import { CLOSE, NEED, OPEN, createTokenizer } from "./read/tokenizer.js";

export type { ByteChunks, XmlCursor, XmlVisitor } from "./read/contract.js";
export { parseIndex } from "./read/text.js";

/**
 * Drives one part through its visitor, lazily: tokens are read synchronously, one `await` is paid
 * per chunk of the part and one per item. The part is closed when the generator finishes, throws
 * or is returned early, so a consumer that stops reading stops the inflate.
 *
 * The only way from bytes to tokens: a part reader added later gets the refusals above without
 * asking for them. `end` runs only after the last chunk has been read, which is when a package
 * has held the part to its size and its checksum.
 */
export async function* scanXml<Item, Result>(
  bytes: ByteChunks,
  visitor: XmlVisitor<Item, Result>,
): AsyncGenerator<Item, Result, undefined> {
  const text = textChunks(bytes);
  const tokenizer = createTokenizer();
  try {
    for (;;) {
      const token = tokenizer.next();
      if (token === NEED) tokenizer.feed(await text.next());
      else if (token === OPEN) visitor.open(tokenizer);
      else if (token === CLOSE) {
        const item = visitor.close(tokenizer);
        if (item !== undefined) yield item;
      } else return visitor.end();
    }
  } finally {
    bytes.close();
  }
}

/** `scanXml` for a part that yields nothing as it goes: runs it to its end and returns the result. */
export async function readXml<Result>(
  bytes: ByteChunks,
  visitor: XmlVisitor<never, Result>,
): Promise<Result> {
  const step = await scanXml(bytes, visitor).next();
  return step.value;
}
