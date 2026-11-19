/**
 * The contract between the tokenizer and a part reader. Its three types are the public surface of
 * `xml/read.ts`, which exports them again; they sit here so the tokenizer can implement the cursor
 * without importing the file that imports it.
 */

/** One part's bytes as chunks. A package's part is one; a test passes its own. */
export interface ByteChunks {
  /**
   * The next chunk, or `undefined` after the last. A chunk is read before the next is asked for,
   * which may overwrite it.
   */
  read(): Promise<Uint8Array | undefined>;
  /**
   * Called once when the part has been read to its end, refused, or left early. Not called when
   * the scan was never started.
   */
  close(): void;
}

/**
 * The token a visitor is looking at. One object for the whole part, valid only during the call it
 * is passed to: a visitor copies out what it keeps.
 */
export interface XmlCursor {
  /**
   * The local name of the element this token opens or closes: a prefix is dropped, so `x:row` and
   * `row` read alike in either namespace family.
   */
  readonly name: string;
  /** How deep that element sits: 1 for the root. */
  readonly depth: number;

  /**
   * An attribute of the open tag by local name, or `undefined`. Its value is as XML normalizes it:
   * references decoded, and a tab or a line end written as itself read as a space. A namespace
   * declaration is found only by its whole name. The value keeps no chunk of the part alive.
   * Allocates: for parts read once. Only valid in `open`.
   */
  attribute(name: string): string | undefined;

  /**
   * An attribute of the open tag as a non-negative decimal integer, without allocating: -1 when
   * the tag has no such attribute, `NaN` when its value is anything but one to fifteen digits (a
   * sign, a fraction, an exponent, a reference, an empty value). Only valid in `open`.
   */
  index(name: string): number;

  /**
   * Where the value of an attribute of the open tag starts in `buffer`, as written, or -1;
   * `valueEnd` is then one past its end. For the one value read once per cell that is neither text
   * nor an index: the cell reference. Only valid in `open`.
   */
  valueStart(name: string): number;
  readonly valueEnd: number;
  /** The decoded text the two spans index into. Never kept: it is replaced as the part is read. */
  readonly buffer: string;

  /**
   * The character data directly inside the element being closed, when it had no child element;
   * `""` otherwise. References and CDATA are decoded, CR LF and a lone CR read as LF. The text keeps
   * no chunk of the part alive. Only valid in `close`.
   */
  text(): string;
}

/**
 * What a part reader is: three synchronous functions over one part's tokens.
 *
 * `end` is not optional. It runs once the part has ended well formed and must say what the part
 * amounts to, or refuse: a reader of a part cannot leave "what I was looking for never came"
 * undecided, which is how a damaged workbook reads as an empty one.
 */
export interface XmlVisitor<Item, Result> {
  /** An element opens. Its attributes are on the cursor. */
  open(xml: XmlCursor): void;
  /**
   * An element closes (a self-closing tag opens and closes). Returning an item hands it to the
   * consumer before another token is read; returning `undefined` goes on.
   */
  close(xml: XmlCursor): Item | undefined;
  /** The part ended. Returns its result or refuses. */
  end(): Result;
}
