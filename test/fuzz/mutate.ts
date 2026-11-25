import { type Part, bytesOf, packageOf } from "../helpers/package.ts";
import { unzip } from "../helpers/unzip.ts";

/**
 * Where a mutation changes a workbook: the bytes of its package; the records of the package, the
 * rest kept consistent; or the text of one part, packed again with its right checksum.
 */
export type Depth = "bytes" | "records" | "text";

export interface Mutation {
  readonly depth: Depth;
  /** What was changed, where, for the log. */
  readonly description: string;
  readonly bytes: Uint8Array;
  /**
   * Whether the package states the same workbook as the original: its parts in another order or
   * another layout. A clean read of it, or of damaged bytes, must give the original's values.
   */
  readonly same: boolean;
}

/**
 * The numbers of mutation `index` of a run from `seed`, from 0 up to 1: sfc32 over the seed's two
 * halves and the index, so any one mutation is drawn again without the ones before it.
 */
export function randomOf(seed: number, index: number): () => number {
  let a = seed >>> 0;
  let b = Math.floor(seed / 2 ** 32) >>> 0;
  let c = index >>> 0;
  let d = 0x9e3779b9;
  const next = () => {
    let t = (a + b) | 0;
    a = b ^ (b >>> 9);
    b = (c + (c << 3)) | 0;
    c = (c << 21) | (c >>> 11);
    d = (d + 1) | 0;
    t = (t + d) | 0;
    c = (c + t) | 0;
    return (t >>> 0) / 2 ** 32;
  };
  for (let warm = 0; warm < 16; warm++) next();
  return next;
}

type Random = () => number;

const below = (random: Random, count: number) => Math.floor(random() * count);

function pick<T>(random: Random, items: readonly T[]): T {
  const item = items[below(random, items.length)];
  if (item === undefined) throw new Error("nothing to pick from");
  return item;
}

/** Draws one mutation of `original`, a package the test's own unzip reads. */
export function mutate(original: Uint8Array, random: Random): Mutation {
  const draw = random();
  if (draw < 0.25) return mutateBytes(original, random);
  if (draw < 0.5) return mutateRecords(original, random);
  return mutateText(original, random);
}

function directoryStart(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let at = bytes.length - 22; at >= 0; at--) {
    if (view.getUint32(at, true) === 0x06054b50) return view.getUint32(at + 16, true);
  }
  throw new Error("no end record");
}

/** One to three changes to the bytes, half of them in the directory and the end record. */
function mutateBytes(original: Uint8Array, random: Random): Mutation {
  let bytes = Array.from(original);
  const directory = directoryStart(original);
  const changes: string[] = [];
  for (let count = 1 + below(random, 3); count > 0 && bytes.length > 0; count--) {
    const at =
      random() < 0.5
        ? below(random, bytes.length)
        : Math.min(bytes.length - 1, directory + below(random, original.length - directory));
    const length = 1 + below(random, 64);
    const kind = pick(random, ["flip", "flip", "zero", "delete", "duplicate", "insert", "cut"]);
    if (kind === "flip") {
      const bits = 1 + below(random, 8);
      for (let flip = 0; flip < bits; flip++) {
        const where = Math.min(bytes.length - 1, at + below(random, length));
        bytes[where] = bytes[where]! ^ (1 << below(random, 8));
      }
      changes.push(`flip ${bits} bits from byte ${at}`);
    } else if (kind === "zero") {
      bytes.fill(0, at, at + length);
      changes.push(`zero ${length} bytes at ${at}`);
    } else if (kind === "delete") {
      bytes.splice(at, length);
      changes.push(`delete ${length} bytes at ${at}`);
    } else if (kind === "duplicate") {
      bytes.splice(at, 0, ...bytes.slice(at, at + length));
      changes.push(`repeat ${length} bytes at ${at}`);
    } else if (kind === "insert") {
      bytes.splice(at, 0, ...Array.from({ length }, () => below(random, 256)));
      changes.push(`insert ${length} bytes at ${at}`);
    } else {
      bytes = bytes.slice(0, at);
      changes.push(`cut at ${at}`);
    }
  }
  return {
    depth: "bytes",
    description: changes.join(", "),
    bytes: new Uint8Array(bytes),
    same: true,
  };
}

const SAME_LAYOUTS = [
  "reorder",
  "stored",
  "descriptors",
  "zip64",
  "extra",
  "comment",
  "front",
] as const;
const OTHER_PACKAGES = [
  "rename",
  "case",
  "delete",
  "duplicate",
  "swap",
  "method",
  "size",
  "checksum",
  "compressed",
  "local",
] as const;

/** The parts of a package, packed again with records that state the truth unless told not to. */
function mutateRecords(original: Uint8Array, random: Random): Mutation {
  const parts: Part[] = unzip(original).map(({ name, method, data }) => ({
    path: name,
    data,
    method,
  }));
  const kind = pick(random, [...SAME_LAYOUTS, ...OTHER_PACKAGES]);
  const at = below(random, parts.length);
  const part = parts[at]!;
  const other = parts[below(random, parts.length)]!;
  const change = (description: string, changed: Part[], layout = {}): Mutation => ({
    depth: "records",
    description,
    bytes: packageOf(changed, layout),
    same: (SAME_LAYOUTS as readonly string[]).includes(kind),
  });
  const replaced = (with_: Part) => parts.map((each, index) => (index === at ? with_ : each));
  switch (kind) {
    case "reorder": {
      const order = parts.map((each) => [random(), each] as const).sort(([x], [y]) => x - y);
      return change(
        "parts in another order",
        order.map(([, each]) => each),
      );
    }
    case "stored":
      return change(
        "every part stored",
        parts.map((each) => ({ ...each, method: 0 })),
      );
    case "descriptors":
      return change(
        "data descriptors",
        parts.map((each) => ({ ...each, descriptor: true })),
      );
    case "zip64":
      return change(
        "zip64 records",
        parts.map((each) => ({ ...each, zip64: true })),
        {
          zip64: true,
        },
      );
    case "extra":
      return change(
        "an extra field on every part",
        parts.map((each) => ({ ...each, extra: new Uint8Array([0x55, 0x54, 1, 0, 1]) })),
      );
    case "comment":
      return change("a comment", parts, { comment: bytesOf("a comment") });
    case "front":
      return change("bytes in front that the offsets count", parts, {
        before: new Uint8Array(1 + below(random, 512)),
        offsetsCountBefore: true,
      });
    case "rename": {
      const path = renamed(part.path, random);
      return change(`${part.path} renamed ${path}`, replaced({ ...part, path }));
    }
    case "case": {
      const path = part.path.toUpperCase();
      return change(`${part.path} renamed ${path}`, replaced({ ...part, path }));
    }
    case "delete":
      return change(
        `${part.path} left out`,
        parts.filter((_, index) => index !== at),
      );
    case "duplicate":
      return change(`${part.path} twice`, [...parts, { ...other, path: part.path }]);
    case "swap":
      return change(
        `${part.path} and ${other.path} named each other`,
        parts.map((each) =>
          each === part
            ? { ...each, path: other.path }
            : each === other
              ? { ...each, path: part.path }
              : each,
        ),
      );
    case "method": {
      const method = pick(random, [1, 9, 12, 14, 99]);
      return change(`${part.path} under method ${method}`, replaced({ ...part, method }));
    }
    case "size": {
      const size = Math.max(0, part.data.length + pick(random, [-1, 1, 1_000_000]));
      return change(`${part.path} states size ${size}`, replaced({ ...part, size }));
    }
    case "checksum":
      return change(`${part.path} states a wrong CRC-32`, replaced({ ...part, crc32: 1 }));
    case "compressed": {
      const more = 1 + below(random, 16);
      const alone = packageOf([part]);
      const central = alone.length - 22 - 46 - bytesOf(part.path).length;
      const compressed = new DataView(alone.buffer).getUint32(central + 20, true);
      return change(
        `${part.path} states ${more} more compressed bytes`,
        replaced({ ...part, compressedSize: compressed + more }),
      );
    }
    case "local": {
      const localPath = renamed(part.path, random);
      return change(
        `${part.path} named ${localPath} in its local header`,
        replaced({ ...part, localPath }),
      );
    }
  }
}

function renamed(path: string, random: Random): string {
  const at = below(random, path.length);
  const character = pick(random, ["x", "/", ".", "%", "Å", ""]);
  return path.slice(0, at) + character + path.slice(at + 1);
}

/** A tag of a part's text: `<name attributes>`, `<name/>` or `</name>`. */
interface Tag {
  readonly start: number;
  readonly end: number;
  readonly name: string;
  readonly closing: boolean;
  readonly empty: boolean;
}

function tagsOf(text: string): Tag[] {
  const tags: Tag[] = [];
  for (const match of text.matchAll(/<(\/?)([^\s/>!?]+)[^>]*?(\/?)>/g)) {
    tags.push({
      start: match.index,
      end: match.index + match[0].length,
      name: match[2]!,
      closing: match[1] === "/",
      empty: match[3] === "/",
    });
  }
  return tags;
}

/** Where the element that `open` starts ends: after its own closing tag, or after itself. */
function elementEnd(tags: readonly Tag[], open: number): number {
  const tag = tags[open]!;
  if (tag.empty) return tag.end;
  let depth = 0;
  for (let at = open; at < tags.length; at++) {
    const each = tags[at]!;
    if (each.name !== tag.name || each.empty) continue;
    depth += each.closing ? -1 : 1;
    if (depth === 0) return each.end;
  }
  return tag.end;
}

const EDGES = [
  "",
  " ",
  "0",
  "-1",
  "1.5",
  "1048576",
  "1048577",
  "16384",
  "4294967296",
  "99999999999999999999",
  "1e308",
  "NaN",
  "XFD1048576",
  "A0",
  "x",
];

const NAMESPACES = [
  [
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
  ],
  [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
  ],
] as const;

const TOKENS = [
  "<",
  ">",
  "&",
  "</row>",
  '<row r="5">',
  '<c r="XFD1048576"><v>1</v></c>',
  '<!DOCTYPE x [<!ENTITY a "b">]>',
  "&#xFFFFFFFF;",
  "&#0;",
  "&a;",
  "<![CDATA[x]]>",
  "<!-- x -->",
  "<?x y?>",
  "\u0000",
  "_x0000_",
  "_xD83D_",
  "<a>".repeat(300),
  "<si/>",
  '<x:a xmlns:x="urn:x">1</x:a>',
  "<extLst><v>9</v></extLst>",
  "\r\n",
];

const TEXT_KINDS = [
  "delete",
  "double",
  "rename",
  "attribute",
  "no attribute",
  "namespace",
  "number",
  "relationship",
  "insert",
] as const;

/** One change to the text of one part the reader reads, packed again as a valid package. */
function mutateText(original: Uint8Array, random: Random): Mutation {
  const entries = unzip(original);
  const readable = entries.filter(
    ({ name }) =>
      name === "_rels/.rels" ||
      name === "xl/workbook.xml" ||
      name === "xl/_rels/workbook.xml.rels" ||
      name === "xl/sharedStrings.xml" ||
      name === "xl/styles.xml" ||
      /^xl\/worksheets\/[^/]+\.xml$/.test(name),
  );
  const target = pick(random, readable);
  const text = new TextDecoder().decode(target.data);
  const [changed, description] = changeText(text, target.name, random);
  return {
    depth: "text",
    description: `${target.name}: ${description}`,
    bytes: packageOf(
      entries.map(({ name, method, data }) => ({
        path: name,
        data: name === target.name ? bytesOf(changed) : data,
        method,
      })),
    ),
    same: false,
  };
}

function changeText(text: string, path: string, random: Random): readonly [string, string] {
  const tags = tagsOf(text);
  const opening = tags.map((tag, at) => [tag, at] as const).filter(([tag]) => !tag.closing);
  const kind = pick(random, TEXT_KINDS);
  const replace = (start: number, end: number, by: string) =>
    text.slice(0, start) + by + text.slice(end);
  if ((kind === "delete" || kind === "double" || kind === "rename") && opening.length > 0) {
    const [tag, at] = pick(random, opening);
    const end = elementEnd(tags, at);
    if (kind === "delete") {
      return [replace(tag.start, end, ""), `delete <${tag.name}> at ${tag.start}`];
    }
    if (kind === "double") {
      return [
        replace(end, end, text.slice(tag.start, end)),
        `double <${tag.name}> at ${tag.start}`,
      ];
    }
    const name = pick(random, [
      `${tag.name}x`,
      `x:${tag.name}`,
      ...new Set(tags.map((each) => each.name)),
    ]);
    const element = text.slice(tag.start, end);
    const closed = element.endsWith(`</${tag.name}>`);
    const inner = element.slice(1 + tag.name.length, closed ? -3 - tag.name.length : undefined);
    return [
      replace(tag.start, end, `<${name}${inner}${closed ? `</${name}>` : ""}`),
      `rename <${tag.name}> at ${tag.start} <${name}>`,
    ];
  }
  const attributes = [...text.matchAll(/\s([^\s=<>"']+)="([^"]*)"/g)];
  if ((kind === "attribute" || kind === "no attribute") && attributes.length > 0) {
    const attribute = pick(random, attributes);
    const start = attribute.index;
    const end = start + attribute[0].length;
    if (kind === "no attribute") {
      return [replace(start, end, ""), `delete ${attribute[1]} at ${start}`];
    }
    const value = pick(random, [...EDGES, pick(random, attributes)[2]!]);
    return [
      replace(start, end, ` ${attribute[1]}="${value}"`),
      `${attribute[1]}="${attribute[2]}" at ${start} set to "${value}"`,
    ];
  }
  const declarations = attributes.filter(
    ([, name]) => name === "xmlns" || name!.startsWith("xmlns:"),
  );
  if (kind === "namespace" && declarations.length > 0) {
    const declaration = pick(random, declarations);
    const uri = declaration[2]!;
    const family = NAMESPACES.find((pair) => (pair as readonly string[]).includes(uri));
    const value = pick(random, [
      family?.find((each) => each !== uri) ?? "urn:x",
      "urn:x",
      "",
      `${uri}x`,
      uri.slice(0, -1),
    ]);
    const start = declaration.index;
    return [
      replace(start, start + declaration[0].length, ` ${declaration[1]}="${value}"`),
      `${declaration[1]}="${uri}" set to "${value}"`,
    ];
  }
  const numbers = [...text.matchAll(/\d+(\.\d+)?/g)];
  if (kind === "number" && numbers.length > 0) {
    const number = pick(random, numbers);
    const value = pick(random, EDGES.slice(2, 12));
    return [
      replace(number.index, number.index + number[0].length, value),
      `${number[0]} at ${number.index} set to ${value}`,
    ];
  }
  const relationships = [...text.matchAll(/<Relationship\s[^>]*>/g)];
  if (kind === "relationship" && path.endsWith(".rels") && relationships.length > 1) {
    const relationship = pick(random, relationships);
    const another = pick(random, relationships);
    const field = pick(random, ["Target", "Type", "Id"]);
    const value = new RegExp(`${field}="([^"]*)"`).exec(another[0])?.[1] ?? "";
    const changed = relationship[0].replace(new RegExp(`${field}="[^"]*"`), `${field}="${value}"`);
    return [
      replace(relationship.index, relationship.index + relationship[0].length, changed),
      `${field} at ${relationship.index} set to "${value}"`,
    ];
  }
  const boundary = tags.length === 0 ? 0 : pick(random, tags).end;
  const token = pick(random, TOKENS);
  return [
    replace(boundary, boundary, token),
    `insert ${JSON.stringify(token.slice(0, 40))} at ${boundary}`,
  ];
}

/** Folds the letters A to Z, as the reader folds part names. */
const folded = (name: string) => name.replace(/[A-Z]/g, (letter) => letter.toLowerCase());

const ENTITIES: Readonly<Record<string, string>> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

/** An element by its local name: its attributes, where a quoted `>` does not end the tag, and its content. */
const element = (name: string) =>
  new RegExp(
    `<(?:[\\w.-]+:)?${name}\\b((?:"[^"]*"|'[^']*'|[^>"'])*?)(?:/>|>([\\s\\S]*?)</(?:[\\w.-]+:)?${name}>)`,
    "g",
  );

const TEXT = /<(?:[\w.-]+:)?t(?:\s[^>]*)?>[^<]/;

/** A part's text without its comments; a comment that never closes takes the rest. */
function withoutComments(text: string): string {
  let kept = "";
  for (let at = 0; ;) {
    const open = text.indexOf("<!--", at);
    if (open < 0) return kept + text.slice(at);
    kept += text.slice(at, open);
    const close = text.indexOf("-->", open + 4);
    if (close < 0) return kept;
    at = close + 3;
  }
}

interface Followed {
  readonly textOf: (path: string) => string | undefined;
  /** The `sheet` elements inside the workbook part's sheet list, as written. */
  readonly listed: readonly string[];
  readonly related: readonly { id: string | undefined; type: string; target: string }[];
}

function attribute(tag: string, name: string): string | undefined {
  const found = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  const value = found?.[1] ?? found?.[2];
  return value?.replace(/&(?:#x([\da-f]+)|#(\d+)|(\w+));/gi, (reference, hex, decimal, named) => {
    if (!hex && !decimal) return ENTITIES[named] ?? reference;
    const point = hex ? Number.parseInt(hex, 16) : Number(decimal);
    return point <= 0x10ffff ? String.fromCodePoint(point) : reference;
  });
}

/**
 * The workbook part of a package and what it relates to, found by searching the parts' text with
 * comments left out, which no reader of ours does. `undefined` when the package or its
 * relationships cannot be followed that way.
 */
function follow(file: Uint8Array): Followed | undefined {
  let entries: ReturnType<typeof unzip>;
  try {
    entries = unzip(file);
  } catch {
    return undefined;
  }
  const textOf = (path: string) => {
    const entry = entries.find(({ name }) => folded(name) === folded(path));
    return entry === undefined ? undefined : withoutComments(new TextDecoder().decode(entry.data));
  };
  const resolve = (from: string, target: string) => {
    const steps = target.startsWith("/") ? [] : from.split("/").slice(0, -1);
    for (const step of target.replace(/^\//, "").split("/")) {
      if (step === "..") steps.pop();
      else if (step !== ".") steps.push(step);
    }
    return steps.join("/");
  };
  const relationshipsOf = (part: string) =>
    [
      ...(textOf(part.replace(/([^/]*)$/, "_rels/$1.rels")) ?? "").matchAll(
        element("Relationship"),
      ),
    ].map(([tag]) => ({
      id: attribute(tag, "Id"),
      type: attribute(tag, "Type") ?? "",
      target: resolve(part, attribute(tag, "Target") ?? ""),
    }));
  const workbook = relationshipsOf("").find(({ type }) => type.endsWith("/officeDocument"))?.target;
  const text = workbook === undefined ? undefined : textOf(workbook);
  if (workbook === undefined || text === undefined) return undefined;
  const listed = [...text.matchAll(element("sheets"))].flatMap(([, , content = ""]) =>
    [...content.matchAll(element("sheet"))].map(([tag]) => tag),
  );
  return { textOf, listed, related: relationshipsOf(workbook) };
}

/**
 * Whether the workbook part of a package lists a sheet named `sheet`. `undefined` when the package
 * cannot be followed.
 */
export function sheetListed(file: Uint8Array, sheet: string): boolean | undefined {
  return follow(file)?.listed.some((tag) => attribute(tag, "name") === sheet);
}

/**
 * Whether the sheet named `sheet` holds a value in a package: a cell whose `v` has text, whose
 * inline string has a `t` with text, or that names an entry of the shared-string table that has
 * one. `undefined` when the package or the sheet's relationship cannot be followed.
 */
export function sheetHoldsValues(file: Uint8Array, sheet: string): boolean | undefined {
  const followed = follow(file);
  if (followed === undefined) return undefined;
  const { textOf, listed, related } = followed;
  const tag = listed.find((each) => attribute(each, "name") === sheet);
  const id = tag && /\s[\w.-]+:id\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(tag)?.slice(1).find(Boolean);
  const part = textOf(related.find((each) => each.id === id)?.target ?? "");
  if (part === undefined) return undefined;
  const table = related.find(({ type }) => type.endsWith("/sharedStrings"));
  const items = [...(textOf(table?.target ?? "") ?? "").matchAll(element("si"))];
  return [...part.matchAll(element("c"))].some(([, tag = "", content = ""]) => {
    const value = /<(?:[\w.-]+:)?v(?:\s[^>]*)?>([^<]+)</.exec(content)?.[1];
    if (attribute(tag, "t") === "inlineStr") return TEXT.test(content);
    if (attribute(tag, "t") !== "s") return value !== undefined;
    return value !== undefined && TEXT.test(items[Number(value)]?.[0] ?? "");
  });
}
