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
 *
 * An element is what its namespace and its local name say together, so `x:row` under
 * `xmlns:x="…/spreadsheetml/2006/main"` is a row and `row` under `xmlns="urn:other"` is not: a
 * visitor that matches by `name` alone reads another vocabulary as its own. An attribute without a
 * prefix is in no namespace, and one with a prefix is never it: `x:r` is not `r`. Each declaration
 * is resolved once, where it stands, to the number the visitor's `namespaces` gives it: an element
 * without a prefix costs one read of the default namespace in scope, and one with a prefix a
 * lookup as long as the prefix.
 *
 * As XML Namespaces 1.0 has it, a declaration holds for its own tag and what that holds, and an
 * inner one hides an outer one until its element closes; `xmlns=""` leaves no default namespace;
 * `xml` is bound without a declaration. Refused as `INVALID_XML`: a prefix that no declaration in
 * scope binds, on an element or an attribute; a declaration of `xmlns`, of `xml` to another
 * namespace, of another prefix to that of `xml`, and of a prefix to no namespace; two attributes
 * of a tag written with one name. Two written with two prefixes bound alike are refused only when
 * asked for by namespace, as no other lookup can find them by their local name. The declarations
 * in scope count toward the characters the names of the open elements may hold together.
 */
export interface XmlCursor {
  /** The local name of the element this token opens or closes, its prefix dropped. */
  readonly name: string;
  /**
   * The namespace of the element as the visitor's `namespaces` numbers it, and 0 for one the
   * visitor does not name or for none. Only valid in `open`.
   */
  readonly namespace: number;
  /** How deep that element sits: 1 for the root. */
  readonly depth: number;

  /**
   * An attribute of the open tag, or `undefined`. Without `namespace`, or with 0, it is the
   * attribute written as `name`, so a name without a prefix finds none under one, and a
   * declaration is found only by its whole name. With another number, it is the one whose local
   * name is `name` under a prefix bound to a namespace the visitor numbers so, refused when the tag
   * has two. Its value is as XML normalizes it: references decoded, and a tab or a line end written
   * as itself read as a space. The value keeps no chunk of the part alive. Allocates: for parts
   * read once. Only valid in `open`.
   */
  attribute(name: string, namespace?: number): string | undefined;

  /**
   * The attribute of the open tag written as `name`, as a non-negative decimal integer, without
   * allocating: -1 when the tag has no such attribute, `NaN` when its value is anything but one to
   * fifteen digits (a sign, a fraction, an exponent, a reference, an empty value). Only valid in
   * `open`.
   */
  index(name: string): number;

  /**
   * Where the value of the attribute of the open tag written as `name` starts in `buffer`, as
   * written, or -1; `valueEnd` is then one past its end. For the one value read once per cell that
   * is neither text nor an index: the cell reference. Only valid in `open`.
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
  /**
   * The namespaces this part's reader asks for, each by its name with the number the cursor gives
   * it; two names may share a number, as the two families of one vocabulary do. A reader that
   * names none sees every element as foreign.
   */
  readonly namespaces: ReadonlyMap<string, number>;
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
