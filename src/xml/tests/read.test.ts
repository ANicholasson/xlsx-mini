import { describe, expect, it } from "vite-plus/test";
import { XlsxReadError } from "../../errors/read.js";
import { XML_DECLARATION } from "../index.js";
import { type XmlCursor, type XmlVisitor, parseIndex, readXml, scanXml } from "../read.js";
import {
  NAMESPACES,
  type Token,
  chunked,
  collector,
  outcomeOf,
  tokensOf,
  utf8,
} from "./reading.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

async function textOf(content: string): Promise<string> {
  const tokens = await tokensOf(`<t>${content}</t>`);
  expect(tokens).toHaveLength(2);
  return tokens[1]![3] as string;
}

async function valueOf(quoted: string): Promise<string | undefined> {
  const tokens = await tokensOf(`<e a=${quoted}/>`, { asked: ["a"] });
  return (tokens[0]![3] as Record<string, string>)["a"];
}

async function onOpen<Seen>(part: string, look: (xml: XmlCursor) => Seen): Promise<Seen> {
  const seen: Seen[] = [];
  await readXml(chunked(part), {
    namespaces: NAMESPACES,
    open(xml) {
      if (seen.length === 0) seen.push(look(xml));
    },
    close: () => undefined,
    end: () => undefined,
  });
  return seen[0]!;
}

async function expectRefused(part: string | Uint8Array, why?: string): Promise<void> {
  const refusal = await readXml(chunked(part), {
    namespaces: NAMESPACES,
    open() {},
    close: () => undefined,
    end: () => "read",
  }).catch((error: unknown) => error);
  expect(refusal, why ?? String(part)).toBeInstanceOf(XlsxReadError);
  expect(refusal).toMatchObject({ code: "INVALID_XML" });
}

async function expectRead(part: string | Uint8Array, why?: string): Promise<void> {
  expect(await outcomeOf(part), why ?? String(part)).not.toBe("INVALID_XML");
}

async function namespacesOf(part: string, cuts?: readonly number[]): Promise<string[]> {
  const seen: string[] = [];
  await readXml(chunked(part, cuts), {
    namespaces: NAMESPACES,
    open(xml) {
      seen.push(`<${xml.name} ${xml.namespace}`);
    },
    close(xml) {
      seen.push(`/${xml.name}`);
      return undefined;
    },
    end: () => undefined,
  });
  return seen;
}

describe("the tokens of a part", () => {
  it("are every open and close tag in order, with the depth and the local name of its element", async () => {
    const sheet =
      `${XML_DECLARATION}<worksheet xmlns="${MAIN}"><sheetData><row r="1">` +
      `<c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>`;
    expect(await tokensOf(sheet, { asked: ["r", "t"] })).toEqual([
      ["open", "worksheet", 1, {}],
      ["open", "sheetData", 2, {}],
      ["open", "row", 3, { r: "1" }],
      ["open", "c", 4, { r: "A1", t: "s" }],
      ["open", "v", 5, {}],
      ["close", "v", 5, "0"],
      ["close", "c", 4, ""],
      ["open", "c", 4, { r: "B1" }],
      ["open", "v", 5, {}],
      ["close", "v", 5, "42"],
      ["close", "c", 4, ""],
      ["close", "row", 3, ""],
      ["close", "sheetData", 2, ""],
      ["close", "worksheet", 1, ""],
    ] satisfies Token[]);
  });

  it("open and close an empty element, with and without white space before its end", async () => {
    expect(await tokensOf(`<a><b/><c r="1" /><d\n/></a>`, { asked: ["r"] })).toEqual([
      ["open", "a", 1, {}],
      ["open", "b", 2, {}],
      ["close", "b", 2, ""],
      ["open", "c", 2, { r: "1" }],
      ["close", "c", 2, ""],
      ["open", "d", 2, {}],
      ["close", "d", 2, ""],
      ["close", "a", 1, ""],
    ]);
  });

  it("give the local name, its prefix dropped, so that both ways of writing a namespace read alike", async () => {
    const prefixed = `<x:sst xmlns:x="${MAIN}"><x:si><x:t>a</x:t></x:si></x:sst>`;
    const plain = `<sst xmlns="${MAIN}"><si><t>a</t></si></sst>`;
    expect(await tokensOf(prefixed)).toEqual(await tokensOf(plain));
    expect((await tokensOf(prefixed)).map(([, name]) => name)).toEqual([
      "sst",
      "si",
      "t",
      "t",
      "si",
      "sst",
    ]);
  });

  it("allow what XML allows around the root: a byte order mark, the declaration, comments, instructions, white space", async () => {
    const around = [
      `\uFEFF<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a/>`,
      `<?xml version='1.0' encoding='utf-8'?><a/>`,
      `<?xml version = "1.0" standalone = 'no' ?><a/>`,
      `<?xml\tversion="1.0"\nencoding="Utf-8"\r\nstandalone="yes"?><a/>`,
      `<a/>`,
      ` \t\r\n<!-- before --><?before it?>\n<a/>\n<!-- after --><?after?> \n`,
      `<?xml-stylesheet href="a.css"?><a/>`,
      `<?a?><a/><?b ?>`,
    ];
    for (const part of around) {
      expect(await tokensOf(part), part).toEqual([
        ["open", "a", 1, {}],
        ["close", "a", 1, ""],
      ]);
    }
  });
});

describe("the text of an element", () => {
  it("is the character data directly inside an element without a child element, as stored", async () => {
    expect(await textOf("plain")).toBe("plain");
    expect(await textOf("  kept \t as it is \n")).toBe("  kept \t as it is \n");
    expect(await textOf("")).toBe("");
    expect(await textOf("Ærlig «ekte» 🎉 \u007F\u0085\uD7FF\uE000\uFFFD\u{10000}\u{10FFFF}")).toBe(
      "Ærlig «ekte» 🎉 \u007F\u0085\uD7FF\uE000\uFFFD\u{10000}\u{10FFFF}",
    );
    expect(await textOf("a > b ]] c ]> d ] ]> e")).toBe("a > b ]] c ]> d ] ]> e");
  });

  it("is empty for an element with a child element, and the child has its own", async () => {
    expect(await tokensOf("<si>before<t>in</t>between<r/>after</si>")).toEqual([
      ["open", "si", 1, {}],
      ["open", "t", 2, {}],
      ["close", "t", 2, "in"],
      ["open", "r", 2, {}],
      ["close", "r", 2, ""],
      ["close", "si", 1, ""],
    ]);
  });

  it("decodes the five entities XML defines", async () => {
    expect(await textOf("Tom &amp; Jerry &lt;b&gt; &quot;x&quot; &apos;y&apos; &amp;amp;")).toBe(
      `Tom & Jerry <b> "x" 'y' &amp;`,
    );
  });

  it("decodes a character reference, decimal or hexadecimal, to the character it names", async () => {
    expect(await textOf("&#65;&#x42;&#x63;&#x4A;&#x4a;&#0065;&#x000041;")).toBe("ABcJJAA");
    expect(await textOf("&#9;&#10;&#13;&#32;&#x1F389;&#55295;&#xE000;&#xFFFD;&#x10FFFF;")).toBe(
      "\t\n\r 🎉\uD7FF\uE000\uFFFD\u{10FFFF}",
    );
    expect(await textOf("&#60;&#38;&#62;")).toBe("<&>");
  });

  it("reads CR LF and a lone CR as LF, and a reference to a carriage return as one", async () => {
    expect(await textOf("a\r\nb\rc\nd\r\r\ne\n\rf")).toBe("a\nb\nc\nd\n\ne\n\nf");
    expect(await textOf("a&#13;b&#13;&#10;c&#xD;\n")).toBe("a\rb\r\nc\r\n");
    expect(await textOf("\r")).toBe("\n");
    const lines = Array.from({ length: 1500 }, (_, line) => line);
    expect(await textOf(lines.map((line) => `${line}\r\n&lt;\r`).join(""))).toBe(
      lines.map((line) => `${line}\n<\n`).join(""),
    );
  });

  it("takes a CDATA section as it is written, but for its line ends", async () => {
    expect(await textOf("<![CDATA[<b> & &amp; ]] ]> \r\n\r]]>")).toBe("<b> & &amp; ]] ]> \n\n");
    expect(await textOf("a<![CDATA[]]>b<![CDATA[c]]><![CDATA[]]]]><![CDATA[>]]>")).toBe("abc]]>");
    const lines = Array.from({ length: 1500 }, (_, line) => line);
    expect(await textOf(`<![CDATA[${lines.map((line) => `&${line}\r`).join("")}]]>`)).toBe(
      lines.map((line) => `&${line}\n`).join(""),
    );
  });

  it("joins the character data on both sides of a comment, an instruction and a CDATA section", async () => {
    expect(await textOf("a<!-- b -->c<?d e?>f&amp;<![CDATA[g]]>\r<!---->h")).toBe("acf&g\nh");
    expect(await textOf("<!-- only a comment -->")).toBe("");
    const runs = Array.from({ length: 2_500 }, (_, run) => `${run}`);
    expect(await textOf(runs.join("<!---->"))).toBe(runs.join(""));
    expect(await tokensOf(`<a>${runs.join("<?p?>")}<b>x</b><c/></a>`)).toEqual([
      ["open", "a", 1, {}],
      ["open", "b", 2, {}],
      ["close", "b", 2, "x"],
      ["open", "c", 2, {}],
      ["close", "c", 2, ""],
      ["close", "a", 1, ""],
    ]);
    expect(await tokensOf("<a>x<!-- y --><b/>z<![CDATA[w]]></a>")).toEqual([
      ["open", "a", 1, {}],
      ["open", "b", 2, {}],
      ["close", "b", 2, ""],
      ["close", "a", 1, ""],
    ]);
  });

  it("is asked for when its element closes: an open tag has none", async () => {
    expect(await onOpen("<a>text</a>", (xml) => xml.text())).toBe("");
  });
});

describe("the attributes of an open tag", () => {
  it("are found by the name they are written with, in either quote, with white space around the equals sign", async () => {
    const tag = `<c xmlns:x14ac="urn:x14ac" xmlns:r="urn:two" r="A1" s = '3'\tt\n=\r\n"s" x14ac:dyDescent="0.25" r:id="rId1"/>`;
    const asked = ["r", "s", "t", "x14ac:dyDescent", "r:id", "xmlns:r", "dyDescent", "id", "none"];
    expect(await tokensOf(tag, { asked })).toEqual([
      [
        "open",
        "c",
        1,
        {
          r: "A1",
          s: "3",
          t: "s",
          "x14ac:dyDescent": "0.25",
          "r:id": "rId1",
          "xmlns:r": "urn:two",
        },
      ],
      ["close", "c", 1, ""],
    ]);
  });

  it("are never the attribute of their local name when written with a prefix, however the prefix is bound", async () => {
    for (const bound of [`xmlns:x="${MAIN}"`, 'xmlns:x="urn:one"', 'xmlns:x="urn:other"']) {
      const tag = `<c ${bound} x:r="A1" x:s="3" x:t="s" x:numFmtId="14" x:date1904="1"/>`;
      expect(
        await onOpen(tag, (xml) => [
          xml.attribute("r"),
          xml.index("s"),
          xml.valueStart("t"),
          xml.attribute("numFmtId"),
          xml.attribute("date1904"),
        ]),
        bound,
      ).toEqual([undefined, -1, -1, undefined, undefined]);
    }
  });

  it("are asked for by namespace with the number the reader gives it, one of its two spellings or the other", async () => {
    const ids = (tag: string) =>
      onOpen(tag, (xml) => [xml.attribute("id", 2), xml.attribute("id", 1), xml.attribute("id")]);
    expect(await ids('<a xmlns:r="urn:two" r:id="rId1"/>')).toEqual(["rId1", undefined, undefined]);
    expect(await ids('<a xmlns:s="urn:one:strict" s:id="rId1" id="x"/>')).toEqual([
      undefined,
      "rId1",
      "x",
    ]);
    expect(await ids('<a xmlns:r="urn:three" r:id="rId1"/>')).toEqual([
      undefined,
      undefined,
      undefined,
    ]);
    expect(
      await ids('<a xmlns="urn:two" id="rId1"/>'),
      "the default holds for no attribute",
    ).toEqual([undefined, undefined, "rId1"]);
    const nested = '<a xmlns:r="urn:two"><b><c r:id="rId1"/></b></a>';
    const seen: unknown[] = [];
    await readXml(chunked(nested), {
      namespaces: NAMESPACES,
      open(xml) {
        seen.push(xml.attribute("id", 2));
      },
      close: () => undefined,
      end: () => undefined,
    });
    expect(seen).toEqual([undefined, undefined, "rId1"]);
    const twice = '<a xmlns:o="urn:one" xmlns:s="urn:one:strict" o:id="1" s:id="2"/>';
    expect(await onOpen(twice, (xml) => xml.attribute("id")), "not asked by namespace").toBe(
      undefined,
    );
    const refusal = await onOpen(twice, (xml) => xml.attribute("id", 1)).catch(
      (error: unknown) => error,
    );
    expect(refusal, "asked by namespace").toMatchObject({ code: "INVALID_XML" });
  });

  it("are not the namespace declarations beside them", async () => {
    const tag = `<a xmlns="${MAIN}" xmlns:r="urn:r" xmlns:s="urn:s" r="1"/>`;
    expect(await onOpen(tag, (xml) => [xml.attribute("r"), xml.attribute("s")])).toEqual([
      "1",
      undefined,
    ]);
  });

  it("have their references decoded and the other quote kept", async () => {
    expect(await valueOf(`"&lt;&amp;&gt;&quot;&apos;&#65;&#x42;"`)).toBe(`<&>"'AB`);
    expect(await valueOf(`"it's > here"`)).toBe("it's > here");
    expect(await valueOf(`'say "hi" ]]>'`)).toBe('say "hi" ]]>');
    expect(await valueOf(`""`)).toBe("");
  });

  it("have a tab and a line end read as a space when written as themselves, and kept when written as a reference", async () => {
    expect(await valueOf(`"a\tb\nc\rd\r\ne\r\r\nf"`)).toBe("a b c d e  f");
    expect(await valueOf(`"a&#9;b&#10;c&#13;d&#13;&#10;e"`)).toBe("a\tb\nc\rd\r\ne");
    expect(await valueOf(`" two  spaces "`)).toBe(" two  spaces ");
    const words = Array.from({ length: 1500 }, (_, word) => word);
    expect(await valueOf(`"${words.map((word) => `${word}\t&amp;\r\n`).join("")}"`)).toBe(
      words.map((word) => `${word} & `).join(""),
    );
  });

  it("are given as an index without a string being made", async () => {
    const tag = `<row r="12" a="0012" b="999999999999999" c="1234567890123456" d="" e="1.5" f="-1" g="+1" h=" 1" i="1e3" j="0x1F" k="&#49;" l="١"/>`;
    const names = ["r", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "none"];
    expect(await onOpen(tag, (xml) => names.map((name) => xml.index(name)))).toEqual([
      12,
      12,
      999_999_999_999_999,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      Number.NaN,
      -1,
    ]);
  });

  it("are given as written, as a span of the buffer", async () => {
    const written = (name: string) => (xml: XmlCursor) => {
      const start = xml.valueStart(name);
      return start < 0 ? start : xml.buffer.slice(start, xml.valueEnd);
    };
    expect(await onOpen(`<c r="A1" t="s"/>`, written("r"))).toBe("A1");
    expect(await onOpen(`<c r="A1" t="s"/>`, written("t"))).toBe("s");
    expect(await onOpen(`<c r="A&#49;" t=""/>`, written("r"))).toBe("A&#49;");
    expect(await onOpen(`<c r="A1" t=""/>`, written("t"))).toBe("");
    expect(await onOpen(`<c r="A1"/>`, written("t"))).toBe(-1);
  });

  it("are gone when the element closes", async () => {
    const seen: unknown[] = [];
    const part = `<a r="1" xmlns:p="urn:two" p:r="1"><b r="2" p:r="2"/><c r="3" p:r="3">x</c></a>`;
    await readXml(chunked(part), {
      namespaces: NAMESPACES,
      open() {},
      close(xml) {
        seen.push([
          xml.name,
          xml.attribute("r"),
          xml.index("r"),
          xml.valueStart("r"),
          xml.attribute("r", 2),
        ]);
        return undefined;
      },
      end: () => undefined,
    });
    expect(seen).toEqual([
      ["b", undefined, -1, -1, undefined],
      ["c", undefined, -1, -1, undefined],
      ["a", undefined, -1, -1, undefined],
    ]);
  });

  it("may be many: a tag with more attributes than are compared in pairs is read", async () => {
    const names = Array.from({ length: 40 }, (_, index) => `n${index}`);
    const tag = `<a ${names.map((name, index) => `${name}="${index}"`).join(" ")}/>`;
    expect(await onOpen(tag, (xml) => names.map((name) => xml.index(name)))).toEqual(
      names.map((_, index) => index),
    );
  });
});

describe("the namespace of an element", () => {
  it("is the number the reader gives the namespace its prefix or the default namespace is bound to where it stands", async () => {
    const part =
      '<a xmlns="urn:one"><b/><x:c xmlns:x="urn:two"><x:d/><e/></x:c>' +
      '<f xmlns="urn:one:strict"><g xmlns=""><h/></g></f><x:i xmlns:x="urn:three"/><j/></a>';
    const expected = [
      "<a 1",
      "<b 1",
      "/b",
      "<c 2",
      "<d 2",
      "/d",
      "<e 1",
      "/e",
      "/c",
      "<f 1",
      "<g 0",
      "<h 0",
      "/h",
      "/g",
      "/f",
      "<i 0",
      "/i",
      "<j 1",
      "/j",
      "/a",
    ];
    expect(await namespacesOf(part)).toEqual(expected);
    for (let cut = 0; cut <= part.length; cut++) {
      expect(await namespacesOf(part, [cut]), `cut after ${cut}`).toEqual(expected);
    }
  });

  it("is none without a declaration, and none for a prefix bound by the standard alone", async () => {
    expect(await namespacesOf("<a><b/></a>")).toEqual(["<a 0", "<b 0", "/b", "/a"]);
    expect(await namespacesOf('<xml:a xml:space="preserve"/>')).toEqual(["<a 0", "/a"]);
    expect(
      await namespacesOf(
        '<a xmlns:xml="http://www.w3.org/XML/1998/namespace" xml:lang="en"><xml:b/></a>',
      ),
    ).toEqual(["<a 0", "<b 0", "/b", "/a"]);
  });

  it("comes from a declaration anywhere in the tag, written as XML allows a value", async () => {
    expect(await namespacesOf('<x:a x:b="1" xmlns:x="urn:two"/>')).toEqual(["<a 2", "/a"]);
    expect(await namespacesOf('<a xmlns = "urn:&#111;ne"/>')).toEqual(["<a 1", "/a"]);
    expect(await namespacesOf("<a xmlns='urn:two'/>")).toEqual(["<a 2", "/a"]);
    expect(await namespacesOf('<a xmlns="urn:one "/>'), "a space is part of the name").toEqual([
      "<a 0",
      "/a",
    ]);
    expect(await namespacesOf('<a xmlns="URN:ONE"/>'), "and so is the case").toEqual([
      "<a 0",
      "/a",
    ]);
    expect(await namespacesOf('<a xmlnsx="urn:one" xmlns:xmlnsx="urn:two"/>')).toEqual([
      "<a 0",
      "/a",
    ]);
  });

  it("returns to the outer binding once the element that hid it closes, written closed or empty", async () => {
    const part =
      '<x:a xmlns:x="urn:one" xmlns="urn:two"><x:b xmlns:x="urn:three" xmlns="urn:one"><x:c/><c/></x:b>' +
      '<x:d xmlns:x="urn:two" xmlns=""/><x:e/><e/></x:a>';
    expect(await namespacesOf(part)).toEqual([
      "<a 1",
      "<b 0",
      "<c 0",
      "/c",
      "<c 1",
      "/c",
      "/b",
      "<d 2",
      "/d",
      "<e 1",
      "/e",
      "<e 2",
      "/e",
      "/a",
    ]);
    expect(
      await namespacesOf('<r><a xmlns="urn:one"><b xmlns="urn:two"/></a><c/></r>'),
      "an outer declaration ends after an inner one",
    ).toEqual(["<r 0", "<a 1", "<b 2", "/b", "/a", "<c 0", "/c", "/r"]);
  });

  it("comes from the binding of the prefix itself, never one whose prefix starts or ends alike", async () => {
    const part =
      '<a xmlns:p="urn:one" xmlns:pq="urn:two" xmlns:qp="urn:three"><pq:b/><p:c/><qp:d/><p:e/><pq:f/></a>';
    expect(await namespacesOf(part)).toEqual([
      "<a 0",
      "<b 2",
      "/b",
      "<c 1",
      "/c",
      "<d 0",
      "/d",
      "<e 1",
      "/e",
      "<f 2",
      "/f",
      "/a",
    ]);
  });

  it("refuses a prefix that no declaration binds where it stands, on an element or an attribute", async () => {
    for (const part of [
      "<x:a/>",
      '<a x:b="1"/>',
      '<a><b xmlns:x="urn:one"/><x:c/></a>',
      '<a><b xmlns:x="urn:one"/><c x:d="1"/></a>',
      '<a xmlns="urn:one"><x:b/></a>',
      "<xmlns:a/>",
      '<a xmlns:x="urn:one"><X:b/></a>',
      '<a XMLNS:x="urn:one"><x:b/></a>',
      '<a xmlns:x="urn:one" xml:space="preserve" y:z=""/>',
    ]) {
      await expectRefused(part);
    }
    await expectRead('<a xmlns:x="urn:one" xmlns:xmlx="urn:two" x:b="1" xmlx:c="2"><x:d/></a>');
  });

  it("refuses what XML Namespaces reserves, and a prefix bound to no namespace", async () => {
    for (const part of [
      '<a xmlns:xmlns="urn:one"/>',
      '<a xmlns:xml="urn:one"/>',
      '<a xmlns:x="http://www.w3.org/XML/1998/namespace"/>',
      '<a xmlns="http://www.w3.org/XML/1998/namespace"/>',
      '<a xmlns:x=""/>',
      '<a xmlns:x="urn:one"><b xmlns:x=""/></a>',
    ]) {
      await expectRefused(part);
    }
    await expectRead('<a xmlns:xml="http://www.w3.org/XML/1998/namespace" xmlns=""/>');
  });

  it("refuses two attributes written with one name, and reads two whose prefixes are bound alike, which only a lookup by namespace could find, and it refuses them", async () => {
    const many = Array.from({ length: 12 }, (_, index) => `n${index}="${index}"`).join(" ");
    for (const part of [
      '<a xmlns:p="urn:x" p:b="1" p:b="2"/>',
      `<a xmlns:p="urn:x" ${many} p:n9="" p:n9=""/>`,
      '<a xmlns:p="urn:x" xmlns:p="urn:y"/>',
      '<a xmlns="urn:x" xmlns="urn:y"/>',
    ]) {
      await expectRefused(part);
    }
    for (const part of [
      '<a xmlns:p="urn:x" xmlns:q="urn:x" p:b="1" q:b="2"/>',
      '<a xmlns:p="urn:x"><c xmlns:q="urn:x" q:b="2" p:b="1"/></a>',
      `<a xmlns:p="urn:x" xmlns:q="urn:x" ${many} p:n9="" q:n9=""/>`,
      '<a xmlns:p="urn:x" xmlns:q="urn:y" p:b="1" q:b="2" b="3"/>',
      `<a xmlns:p="urn:x" ${many} p:n9="" xmlns:n9="urn:y"/>`,
    ]) {
      await expectRead(part);
    }
    const bound = ['xmlns:p="urn:two"', 'xmlns:q="urn:two"'];
    for (const tag of [
      `<a ${bound.join(" ")} p:id="1" q:id="2"/>`,
      `<a ${bound[0]!}><c ${bound[1]!} q:id="2" p:id="1"/></a>`,
    ]) {
      const refusal = await readXml(chunked(tag), {
        namespaces: NAMESPACES,
        open(xml) {
          xml.attribute("id", 2);
        },
        close: () => undefined,
        end: () => undefined,
      }).catch((error: unknown) => error);
      expect(refusal, tag).toMatchObject({ code: "INVALID_XML" });
    }
  });

  it("keeps fewer than 1 024 declarations in scope, each counting 4 096 characters with the names of the open elements, and none out of scope", async () => {
    const declarations = (count: number, from = 0) =>
      Array.from({ length: count }, (_, index) => ` xmlns:p${from + index}="u"`).join("");
    const name = 1;
    const declared = (prefix: number, uri: number) => prefix + uri + 4096;
    const counted = (count: number, from = 0) =>
      Array.from({ length: count }, (_, index) => declared(`p${from + index}`.length, 1)).reduce(
        (sum, characters) => sum + characters,
        0,
      );
    let most = 0;
    while (name + counted(most + 1) <= 4_194_304) most++;
    expect(most).toBeGreaterThan(1000);
    expect(most).toBeLessThan(1024);
    await expectRead(`<a${declarations(most)}/>`);
    await expectRefused(`<a${declarations(most + 1)}/>`);
    await expectRefused(
      `<r><a/><a${declarations(most + 1)}/></r>`,
      "beside a name counted already",
    );
    let inner = 0;
    while (2 * name + counted(1000) + counted(inner + 1, 1000) <= 4_194_304) inner++;
    const nested = (count: number) =>
      `<a${declarations(1000)}><b${declarations(count, 1000)}/></a>`;
    await expectRead(nested(inner));
    await expectRefused(nested(inner + 1));
    await expectRead(
      `<a${declarations(1000)}>${`<b${declarations(inner, 1000)}/>`.repeat(50)}</a>`,
      "a declaration out of scope counts no more",
    );
  });

  it("counts the prefixes and namespace names in scope with the names of the open elements, to 4 194 304 characters", async () => {
    const names = "a".length + "b".length;
    const declared = (prefix: number, uri: number) => prefix + uri + 4096;
    const part = (rest: number) =>
      `<a xmlns:p="${"u".repeat(3_000_000)}"><b xmlns:q="${"v".repeat(rest)}"/></a>`;
    const rest = 4_194_304 - names - declared(1, 3_000_000) - declared(1, 0);
    await expectRead(part(rest));
    await expectRefused(part(rest + 1));
    const sibling = `<b xmlns:q="${"v".repeat(rest)}"/>`;
    await expectRead(
      `<a xmlns:p="${"u".repeat(3_000_000)}">${sibling}${sibling}</a>`,
      "a declaration out of scope counts no more",
    );
  }, 60_000);

  it("costs time linear in the part, however many declarations and prefixes it holds", async () => {
    // Every prefix shares a long stem, and the elements take them in turn: a lookup that walked
    // the declarations in scope would compare 450 stems on average for each name.
    const stem = "p".repeat(120);
    const prefixes = Array.from({ length: 900 }, (_, index) => `${stem}${index}`);
    const root = prefixes.map((prefix) => ` xmlns:${prefix}="urn:${prefix}"`).join("");
    const elements = Array.from({ length: 30_000 }, (_, index) => {
      const own = prefixes[index % 900]!;
      const other = prefixes[(index * 7) % 900]!;
      return `<${own}:e xmlns:q="urn:q${index % 3}" q:a="1" ${other}:b="2"/>`;
    }).join("");
    const part = `<r${root}>${elements}</r>`;
    const started = performance.now();
    await expectRead(part);
    expect(performance.now() - started).toBeLessThan(2_000);
  }, 60_000);
});

describe("parseIndex", () => {
  it("reads one to fifteen digits and nothing else", () => {
    const text = `r="41" s="007" t="x1" u="" v="123456789012345" w="1234567890123456"`;
    const between = (name: string): [number, number] => {
      const start = text.indexOf(`${name}="`) + name.length + 2;
      return [start, text.indexOf('"', start)];
    };
    expect(parseIndex(text, ...between("r"))).toBe(41);
    expect(parseIndex(text, ...between("s"))).toBe(7);
    expect(parseIndex(text, ...between("t"))).toBeNaN();
    expect(parseIndex(text, ...between("u"))).toBeNaN();
    expect(parseIndex(text, ...between("v"))).toBe(123_456_789_012_345);
    expect(parseIndex(text, ...between("w"))).toBeNaN();
    expect(parseIndex("12:34/5", 0, 2)).toBe(12);
    expect(parseIndex("12:34/5", 0, 3)).toBeNaN();
    expect(parseIndex("12:34/5", 5, 7)).toBeNaN();
    expect(parseIndex("7", 1, 0)).toBeNaN();
  });
});

describe("refused with INVALID_XML", () => {
  it("a DOCTYPE, with or without entities, and every other declaration", async () => {
    await expectRefused(`<!DOCTYPE a><a/>`);
    await expectRefused(`<?xml version="1.0"?>\n<!DOCTYPE a SYSTEM "a.dtd">\n<a/>`);
    await expectRefused(
      `<!DOCTYPE a [<!ENTITY b "bb"><!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;">]><a>&c;</a>`,
    );
    await expectRefused(`<a><!DOCTYPE a></a>`);
    await expectRefused(`<a/><!DOCTYPE a>`);
    await expectRefused(`<!ENTITY b "bb"><a/>`);
    await expectRefused(`<a><![IGNORE[x]]></a>`);
    await expectRefused(`<a><![cdata[x]]></a>`);
    await expectRefused(`<a><!></a>`);
    await expectRefused(`<a><!-</a>`);
  });

  it("bytes that are not valid UTF-8, and never a replacement character for them", async () => {
    const around = (...bytes: number[]) =>
      new Uint8Array([...utf8("<a>"), ...bytes, ...utf8("</a>")]);
    await expectRead(around(0xc3, 0xa9), "é");
    await expectRead(around(0xf0, 0x9f, 0x8e, 0x89), "🎉");
    await expectRefused(around(0xff), "a byte that starts nothing");
    await expectRefused(around(0x80), "a continuation byte alone");
    await expectRefused(around(0xc3), "a sequence cut short");
    await expectRefused(around(0xc0, 0x80), "an overlong form");
    await expectRefused(around(0xe0, 0x80, 0x80), "an overlong form of three bytes");
    await expectRefused(around(0xed, 0xa0, 0x80), "a surrogate");
    await expectRefused(around(0xf4, 0x90, 0x80, 0x80), "past U+10FFFF");
    await expectRefused(
      new Uint8Array([...utf8("<a/>"), 0xe2, 0x82]),
      "cut at the end of the part",
    );
    await expectRefused(new Uint8Array([0xff, 0xfe, 0x3c, 0, 0x61, 0, 0x2f, 0, 0x3e, 0]), "UTF-16");
    await expectRefused(new Uint8Array([0xfe, 0xff, 0, 0x3c, 0, 0x61, 0, 0x2f, 0, 0x3e]), "UTF-16");
    await expectRefused(new Uint8Array([0x3c, 0, 0x61, 0, 0x2f, 0, 0x3e, 0]), "UTF-16 unmarked");
  });

  it("tags that do not match", async () => {
    await expectRead(`<a><b></b></a>`);
    await expectRefused(`<a></b>`);
    await expectRefused(`<a><b></a></b>`);
    await expectRefused(`<a><b></a>`);
    await expectRefused(`<a></a></a>`);
    await expectRefused(`</a>`);
    await expectRefused(`<a></A>`);
    await expectRefused(`<a></ab>`);
    await expectRefused(`<ab></a>`);
    await expectRefused(`<x:a xmlns:x="u"></a>`);
    await expectRefused(`<a xmlns:x="u"></x:a>`);
    await expectRefused(`<a></ a>`);
    await expectRead(`<a></a \r\n\t>`);
    await expectRefused(`<a></a b="1">`);
  });

  it("the worksheet of a mutated workbook that a reader without tag matching read as empty", async () => {
    const mutated =
      `${XML_DECLARATION}<worksheet xmlns="http:o/schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      '<sheetData><row r="1"><c`r="A1" t="s"><v>0</v></c><b r="B1"><v>42</v></c></row></sheetData></worksh%et>';
    await expectRefused(mutated);
    await expectRefused(mutated.replace("<c`r", "<c r"), "the cell that closes as another");
    await expectRefused(
      mutated.replace("<c`r", "<c r").replace("<b r", "<c r"),
      "the root that closes as another",
    );
    await expectRead(
      mutated.replace("<c`r", "<c r").replace("<b r", "<c r").replace("worksh%et", "worksheet"),
    );
  });

  it("an entity XML does not define, and a character reference to what XML forbids", async () => {
    const references = [
      "&nbsp;",
      "&foo;",
      "&AMP;",
      "&Lt;",
      "&;",
      "&",
      "&amp",
      "&amp ;",
      "& amp;",
      "&#;",
      "&#x;",
      "&#X41;",
      "&#x4G;",
      "&#6_5;",
      "&#-1;",
      "&# 65;",
      "&#1e2;",
      "&#0;",
      "&#1;",
      "&#8;",
      "&#11;",
      "&#12;",
      "&#31;",
      "&#xD800;",
      "&#xDFFF;",
      "&#xFFFE;",
      "&#xFFFF;",
      "&#x110000;",
      "&#1114112;",
      "&#99999999999999999999;",
      "&#x100000000000041;",
      "&a.b;",
      "&#١;",
    ];
    for (const reference of references) {
      await expectRefused(`<a>${reference}</a>`);
      await expectRefused(`<a b="${reference}"/>`);
    }
    await expectRead(`<a b="&amp;&#x10FFFF;">&apos;&#1114111;</a>`);
  });

  it("nesting deeper than 256", async () => {
    const nested = (depth: number) => "<a>".repeat(depth) + "</a>".repeat(depth);
    const deepest = await tokensOf(nested(256));
    expect(deepest).toHaveLength(512);
    expect(deepest[255]).toEqual(["open", "a", 256, {}]);
    await expectRefused(nested(257));
    await expectRefused(`${"<a>".repeat(256)}<b/>`);
  });

  it("names of nested elements longer than 4 194 304 characters together", async () => {
    const quarter = 1_048_576;
    const names = ["a", "b", "c", "d"].map((letter) => letter.repeat(quarter));
    const nested = (last: string, declared = "") =>
      `${names
        .slice(0, 3)
        .map((name) => `<${name}>`)
        .join("")}<${last}${declared}></${last}>` +
      names
        .slice(0, 3)
        .reverse()
        .map((name) => `</${name}>`)
        .join("");
    await expectRead(nested(names[3]!), "4 194 304 characters of names");
    await expectRefused(nested(`${names[3]!}d`), "one character more");
    const bound = ' xmlns:x="u"';
    await expectRefused(nested(`x:${names[3]!}`, bound), "the prefix counts, and the declaration");
    await expectRead(nested(`x:${names[3]!.slice(4_100)}`, bound), "4 096 more for a declaration");
    const siblings = names.map((name) => `<${name}/>`).join("");
    await expectRead(
      `<r>${siblings}${siblings}<e${names[0]!}/></r>`,
      "siblings, one after another",
    );
  }, 60_000);

  it("a token longer than 4 194 304 characters, of any kind", async () => {
    const most = 4_194_304;
    const tokens: [string, (length: number) => string][] = [
      ["a run of text", (length) => `<a>${"x".repeat(length)}</a>`],
      ["a run of text after a child element", (length) => `<a><b/>${"x".repeat(length)}</a>`],
      ["a run of text before a child element", (length) => `<a>${"x".repeat(length)}<b/></a>`],
      ["an open tag", (length) => `<a b="${"x".repeat(length - 9)}"/>`],
      ["white space in a tag", (length) => `<a${" ".repeat(length - 3)}>`.concat("</a>")],
      ["a close tag", (length) => `<a></a${" ".repeat(length - 4)}>`],
      ["a comment", (length) => `<a><!--${"x".repeat(length - 7)}--></a>`],
      ["a CDATA section", (length) => `<a><![CDATA[${"x".repeat(length - 12)}]]></a>`],
      ["an instruction", (length) => `<a><?p ${"x".repeat(length - 6)}?></a>`],
      ["a name", (length) => `<${"a".repeat(length - 3)}/>`.concat("")],
      ["an entity", (length) => `<a>&${"a".repeat(length - 1)}</a>`],
    ];
    for (const [kind, token] of tokens) {
      const refused = token(most + 1);
      await expectRefused(refused, `${kind} of ${most + 1}`);
      expect(await outcomeOf(refused, { cuts: [5, 65_536, 2_000_000] }), `${kind}, in chunks`).toBe(
        "INVALID_XML",
      );
      const longer = token(most + 64);
      expect(
        await outcomeOf(longer, { cuts: [longer.length - 8] }),
        `${kind}, refused before its end has come`,
      ).toBe("INVALID_XML");
      if (kind !== "an entity") await expectRead(token(most), `${kind} of ${most}`);
    }
  }, 60_000);

  it("the text of one element longer than 4 194 304 characters, however it is split", async () => {
    const half = "x".repeat(2_097_152);
    expect(await textOf(`${half}<!---->${half}`)).toHaveLength(4_194_304);
    await expectRefused(`<t>${half}<!---->${half}x</t>`);
    await expectRefused(`<t>${half}<![CDATA[${half}]]>x</t>`);
    await expectRefused(`<t>${half}<![CDATA[${half}x]]></t>`);
    await expectRefused(`<t>${half}<?p?>${half}<?p?>x</t>`);
    await expectRefused(`<t>${half}<?p?>${half}x<?p?></t>`);
    await expectRead(`<t>${half}<?p?>${half}<a/>${half}</t>`, "text that a child element ends");
    const runs = `${"x".repeat(2_048)}<!---->`.repeat(2_048);
    expect(await textOf(runs)).toHaveLength(4_194_304);
    await expectRefused(`<t>${runs}x</t>`);
  }, 60_000);

  it("content after the root", async () => {
    await expectRead(`<a/> \n<!-- c --><?p?>\r\n`);
    await expectRefused(`<a/><b/>`);
    await expectRefused(`<a></a><a></a>`);
    await expectRefused(`<a/>x`);
    await expectRefused(`<a/> x`);
    await expectRefused(`<a/>&#32;`);
    await expectRefused(`<a/><![CDATA[x]]>`);
    await expectRefused(`<a/>\uFEFF`);
    await expectRefused(`<a/>\u0000`);
  });

  it("content before the root, and a part without a root", async () => {
    await expectRefused(`x<a/>`);
    await expectRefused(`&#32;<a/>`);
    await expectRefused(`<![CDATA[x]]><a/>`);
    await expectRefused(`\uFEFF\uFEFF<a/>`, "a second byte order mark is a character");
    await expectRefused(``);
    await expectRefused(` \n`);
    await expectRefused(`<!-- no root -->`);
    await expectRefused(XML_DECLARATION);
    await expectRefused(`\uFEFF`);
  });

  it("a part that stops early, wherever it stops", async () => {
    const part = utf8(
      `<?xml version="1.0"?><!-- c --><a b="1" c='&amp;'>t&#65;<![CDATA[d]]><?p i?><b/>é</a>`,
    );
    await expectRead(part);
    for (let length = 0; length < part.length; length++) {
      await expectRefused(part.subarray(0, length), `the first ${length} bytes`);
    }
  });

  it("a declaration that is not the first thing in the part, or that XML does not allow", async () => {
    const declarations = [
      ` <?xml version="1.0"?>`,
      `\n<?xml version="1.0"?>`,
      `<!-- c --><?xml version="1.0"?>`,
      `<?xml version="1.0"?><?xml version="1.0"?>`,
      `<?XML version="1.0"?>`,
      `<?Xml version="1.0"?>`,
      `<?xml?>`,
      `<?xml ?>`,
      `<?xml encoding="UTF-8"?>`,
      `<?xml version="1.1"?>`,
      `<?xml version="1.00"?>`,
      `<?xml version="1"?>`,
      `<?xml version="1.&#48;"?>`,
      `<?xml version="1.0" encoding="UTF-16"?>`,
      `<?xml version="1.0" encoding="ISO-8859-1"?>`,
      `<?xml version="1.0" encoding="US-ASCII"?>`,
      `<?xml version="1.0" encoding="utf8"?>`,
      `<?xml version="1.0" encoding=""?>`,
      `<?xml version="1.0" standalone="maybe"?>`,
      `<?xml version="1.0" standalone="YES"?>`,
      `<?xml version="1.0" standalone="yes" encoding="UTF-8"?>`,
      `<?xml standalone="yes" version="1.0"?>`,
      `<?xml version="1.0" encoding="UTF-8" other="1" standalone="yes"?>`,
      `<?xml version="1.0" other="1"?>`,
      `<?xml version="1.0" version="1.0"?>`,
      `<?xml version="1.0" standalone="yes" standalone="yes"?>`,
      `<?xml x:version="1.0"?>`,
      `<?xml version="1.0" x:encoding="UTF-8"?>`,
      `<?xml version="1.0" encoding="UTF-8" x:standalone="yes"?>`,
      `<?xml version="1.0" xmlns:x="u"?>`,
      `<?xml version="1.0"encoding="UTF-8"?>`,
      `<?xml version=1.0?>`,
      `<?xml version="1.0">`,
      `<?xml version="1.0"/>`,
      `<?xml version="1.0"? >`,
    ];
    for (const declaration of declarations) await expectRefused(`${declaration}<a/>`);
    await expectRefused(`<a><?xml version="1.0"?></a>`);
    await expectRefused(`<a/><?xml version="1.0"?>`);
    await expectRefused(`<a><?xMl?></a>`);
    await expectRead(`<a><?xmlx?><?xm?><?xml-model a?></a>`);
  });

  it("an attribute that is repeated, has no value, or is not written as XML writes one", async () => {
    await expectRead(`<a b="1" c="2"/>`);
    const tags = [
      `<a b="1" b="2"/>`,
      `<a b="1" c="2" b="3"/>`,
      `<a xmlns:r="u" r:id="1" r:id="2"/>`,
      `<a xmlns:r="u" xmlns:r="v"/>`,
      `<a xmlns="u" xmlns="v"/>`,
      `<a b/>`,
      `<a b c="1"/>`,
      `<a b=/>`,
      `<a b=1/>`,
      `<a b="1/>`,
      `<a b='1"/>`,
      `<a b="1"c="2"/>`,
      `<a b="1" "2"/>`,
      `<a ="1"/>`,
      `<a b="<"/>`,
      `<a b="a<b"/>`,
      `<a b="1"/ >`,
      `<a b="1"//>`,
      `<a/b>`,
      `<a b="1">`,
    ];
    for (const tag of tags) await expectRefused(tag);
  });

  it("an attribute repeated among more attributes than are compared in pairs", async () => {
    const attributes = Array.from({ length: 12 }, (_, index) => `n${index}="${index}"`);
    await expectRead(`<a ${attributes.join(" ")}/>`);
    for (const repeated of [0, 7, 8, 11]) {
      await expectRefused(`<a ${attributes.join(" ")} n${repeated}=""/>`, `n${repeated} again`);
    }
    await expectRefused(
      `<a xmlns:x="u" ${attributes.join(" ")} x:n9="" x:n9=""/>`,
      "x:n9 twice, under one prefix",
    );
    await expectRead(`<a xmlns:x="u" ${attributes.join(" ")} x:n9=""/>`, "n9 and x:n9 are two");
  });

  it("a name that is not one, that holds a character outside ASCII, or whose colon is out of place", async () => {
    await expectRead(`<A-b.c_1 x-y.z_2="1" _n="2"/>`);
    const tags = [
      `<1a/>`,
      `<-a/>`,
      `<.a/>`,
      `< a/>`,
      `<a$/>`,
      `<a\`b/>`,
      `<>`,
      `</>`,
      `<é/>`,
      `<aé/>`,
      `<中/>`,
      `<a:b:c/>`,
      `<:a/>`,
      `<a:/>`,
      `<a 1b="1"/>`,
      `<a é="1"/>`,
      `<a bé="1"/>`,
      `<a b:c:d="1"/>`,
      `<a :b="1"/>`,
      `<a b:="1"/>`,
      `<a xmlns:="u"/>`,
      `<a><?1p?></a>`,
      `<a><? p?></a>`,
      `<a><?p:q?></a>`,
      `<a><?é?></a>`,
    ];
    for (const tag of tags) await expectRefused(tag);
  });

  it("a character XML forbids, wherever it stands", async () => {
    const forbidden = [
      "\u0000",
      "\u0001",
      "\u0008",
      "\u000B",
      "\u000C",
      "\u001F",
      "\uFFFE",
      "\uFFFF",
    ];
    const places = [
      (c: string) => `<a>${c}</a>`,
      (c: string) => `<a>text${c}</a>`,
      (c: string) => `<a><b/>${c}</a>`,
      (c: string) => `<a b="${c}"/>`,
      (c: string) => `<a b='x${c}'/>`,
      (c: string) => `<a${c}/>`,
      (c: string) => `<a ${c}b="1"/>`,
      (c: string) => `<a><!--${c}--></a>`,
      (c: string) => `<a><![CDATA[${c}]]></a>`,
      (c: string) => `<a><?p ${c}?></a>`,
      (c: string) => `${c}<a/>`,
      (c: string) => `<a></a${c}>`,
    ];
    for (const place of places) {
      await expectRead(place(" "));
      for (const character of forbidden) await expectRefused(place(character));
    }
  });

  it("the end of a CDATA section in character data", async () => {
    await expectRefused(`<a>]]></a>`);
    await expectRefused(`<a>x]]>y</a>`);
    await expectRefused(`<a>]]]></a>`);
    await expectRefused(`<a><b/>]]></a>`);
    await expectRefused(`<a><![CDATA[x]]>]]></a>`);
    await expectRead(`<a>]]&gt;]<!---->]></a>`);
  });

  it("a comment that holds two hyphens, an instruction without a target, and markup that never ends", async () => {
    await expectRead(`<a><!----><!-- - - --><!-->--><!---> --><?p?><?p ?><?p ?x? ?><?p >?></a>`);
    const markup = [
      `<!-- a -- b -->`,
      `<!-- a --->`,
      `<!--->`,
      `<!--`,
      `<!-- a --`,
      `<!-- a -->x`,
      `<?p`,
      `<?p x`,
      `<?p?x?>`,
      `<??>`,
      `<? ?>`,
      `<![CDATA[`,
      `<![CDATA[x]]`,
      `<![CDATA[x]>`,
      `<`,
      `<b`,
      `<b c`,
      `<b c=`,
      `<b c="`,
      `<b c="1"`,
      `</`,
      `</a`,
      `&`,
      `&amp`,
    ];
    for (const cut of markup) await expectRefused(`<a>${cut}`);
    await expectRefused(`<a><!-- a -- b --></a>`);
    await expectRefused(`<a><!-- a ---></a>`);
    await expectRefused(`<a><?p?x?></a>`);
  });
});

describe("scanXml", () => {
  function values(): XmlVisitor<string, number> {
    let count = 0;
    return {
      namespaces: NAMESPACES,
      open() {},
      close(xml) {
        if (xml.name !== "v") return undefined;
        count++;
        return xml.text();
      },
      end: () => count,
    };
  }

  const row = "<row><c><v>1</v></c><c><v>two</v></c><c><v></v></c><c><v>4</v></c></row>";

  it("hands over each item a visitor returns, in order, then the visitor's result", async () => {
    const bytes = chunked(row, [7, 20, 21, 40]);
    const scan = scanXml(bytes, values());
    const items: string[] = [];
    let step = await scan.next();
    while (step.done !== true) {
      items.push(step.value);
      step = await scan.next();
    }
    expect(items).toEqual(["1", "two", "", "4"]);
    expect(step.value).toBe(4);
    expect(bytes.closes).toBe(1);
  });

  it("reads no further than the item it hands over", async () => {
    const bytes = chunked(row, [16, 30, 45, 60]);
    const scan = scanXml(bytes, values());
    expect(await scan.next()).toEqual({ done: false, value: "1" });
    expect(bytes.reads).toBe(1);
    expect(bytes.closes).toBe(0);
  });

  it("closes the part when the consumer stops early", async () => {
    const bytes = chunked(row, [16, 30, 45, 60]);
    for await (const value of scanXml(bytes, values())) {
      expect(value).toBe("1");
      break;
    }
    expect(bytes.reads).toBe(1);
    expect(bytes.closes).toBe(1);
  });

  it("asks a visitor for its result only once the part has been read to its last chunk", async () => {
    const bytes = chunked("<a/>\n", [4]);
    const readsAtEnd = await readXml(bytes, {
      namespaces: NAMESPACES,
      open() {},
      close: () => undefined,
      end: () => bytes.reads,
    });
    expect(readsAtEnd).toBe(3);
    expect(bytes.closes).toBe(1);
  });

  it("closes the part when it is refused, by XML or by the visitor, and passes the refusal on", async () => {
    const refused = chunked("<a><b></a>");
    await expect(readXml(refused, collector())).rejects.toMatchObject({ code: "INVALID_XML" });
    expect(refused.closes).toBe(1);

    const failure = new XlsxReadError("INVALID_WORKBOOK", "not a workbook");
    for (const fails of ["open", "close", "end"] as const) {
      const bytes = chunked("<a/>");
      const visitor: XmlVisitor<never, undefined> = {
        namespaces: NAMESPACES,
        open() {
          if (fails === "open") throw failure;
        },
        close() {
          if (fails === "close") throw failure;
          return undefined;
        },
        end() {
          if (fails === "end") throw failure;
          return undefined;
        },
      };
      await expect(readXml(bytes, visitor), fails).rejects.toBe(failure);
      expect(bytes.closes, fails).toBe(1);
    }
  });

  it("passes on what reading the part's bytes fails with, and closes the part", async () => {
    const failure = new XlsxReadError("INVALID_PACKAGE", "a part has another CRC-32 than stated");
    let closes = 0;
    const bytes = {
      read: () => Promise.reject(failure),
      close: () => {
        closes++;
      },
    };
    await expect(readXml(bytes, collector())).rejects.toBe(failure);
    expect(closes).toBe(1);
  });
});
