import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { NEUTRAL_PATH, anonymize, replaceEntries } from "./corpus/anonymize.js";
import { type ZipEntry, unzip } from "./helpers/unzip.js";

const corpus = join(dirname(fileURLToPath(import.meta.url)), "corpus");

const files = readdirSync(corpus, { recursive: true, encoding: "utf8" })
  .map((name) => name.replaceAll("\\", "/"))
  .filter((name) => /\.[a-z]+$/.test(name))
  .sort();

const FIELD = /<(?:\w+:)?(creator|lastModifiedBy|Company|Manager|author)\b[^>]*>([^<]*)</g;
const PERSON = /<(?:\w+:)?person\b/;
const HOME = /\/Users\/|\/home\/|\\Users\\/i;
const ADDRESS = /[\w.%+-]+@((?:[\w-]+\.)+[a-z]{2,})/gi;

/** A finding names the place and the kind, never the value: the logs of a failed check are public. */
function findingsIn(text: string, where: string, markup: boolean): string[] {
  const findings: string[] = [];
  if (markup) {
    for (const [, field, value = ""] of text.matchAll(FIELD)) {
      if (value.trim() !== "") findings.push(`${where}: ${field} holds ${value.length} characters`);
    }
    if (PERSON.test(text)) findings.push(`${where}: a person`);
  }
  if (HOME.test(text)) findings.push(`${where}: a home folder`);
  for (const [, domain = ""] of text.matchAll(ADDRESS)) {
    if (domain.toLowerCase() !== "example.com") findings.push(`${where}: an address`);
  }
  return findings;
}

const latin1 = new TextDecoder("latin1");

function findingsInFile(name: string, bytes: Uint8Array): string[] {
  if (!name.endsWith(".xlsx")) return findingsIn(latin1.decode(bytes), name, false);
  return unzip(bytes).flatMap((entry) =>
    findingsIn(latin1.decode(entry.data), `${name} ${entry.name}`, entry.name.endsWith(".xml")),
  );
}

const read = (name: string) => new Uint8Array(readFileSync(join(corpus, name)));

describe("the corpus names no one", () => {
  it("holds files to check", () => {
    expect(files.filter((name) => name.endsWith(".xlsx")).length).toBeGreaterThan(80);
  });

  it("holds no name, person, home folder or address but at example.com in any file", () => {
    expect(files.flatMap((name) => findingsInFile(name, read(name)))).toEqual([]);
  });

  it.each([
    ["<dc:creator>A</dc:creator>", "creator holds 1 characters"],
    ["<cp:lastModifiedBy>A B</cp:lastModifiedBy>", "lastModifiedBy holds 3 characters"],
    ["<Company>C</Company>", "Company holds 1 characters"],
    ["<Manager>M</Manager>", "Manager holds 1 characters"],
    ["<authors><author>A</author></authors>", "author holds 1 characters"],
    ['<personList><person displayName="A" id="{1}"/></personList>', "a person"],
    ['<x15ac:absPath url="/Users/a/"/>', "a home folder"],
    ["/home/a/book.xlsx", "a home folder"],
    ["C:\\users\\a\\book.xlsx", "a home folder"],
    ["mailto:a@example.org", "an address"],
  ])("finds %s", (text, finding) => {
    expect(findingsIn(text, "part", true)).toEqual([`part: ${finding}`]);
  });

  it.each([
    "<dc:creator></dc:creator><cp:lastModifiedBy/><Company> </Company>",
    "<personList/>",
    "mailto:hello@example.com",
  ])("finds nothing in %s", (text) => {
    expect(findingsIn(text, "part", true)).toEqual([]);
  });
});

function localRecord(bytes: Uint8Array, entry: ZipEntry): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const at = entry.localHeaderOffset;
  const end = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true);
  return bytes.subarray(at, end + entry.compressedSize);
}

function expectLocalHeadersAgree(bytes: Uint8Array): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (const { name, localHeaderOffset: at, crc32, compressedSize, uncompressedSize } of unzip(
    bytes,
  )) {
    const local = [14, 18, 22].map((field) => view.getUint32(at + field, true));
    expect(local, name).toEqual([crc32, compressedSize, uncompressedSize]);
  }
}

const text = (entries: readonly ZipEntry[], name: string) =>
  new TextDecoder().decode(entries.find((entry) => entry.name === name)!.data);

describe("anonymize", () => {
  it("finds nothing left to blank in any workbook of the corpus, with or without the two parts", () => {
    const workbooks = files.filter((name) => name.endsWith(".xlsx"));
    expect(workbooks.filter((name) => anonymize(read(name)) !== undefined)).toEqual([]);
    expect(unzip(read("google-sheets/minimal.xlsx")).map(({ name }) => name)).not.toContain(
      "docProps/core.xml",
    );
  });

  it("blanks both fields and copies every other entry byte for byte", () => {
    const original = read("excel/minimal.xlsx");
    const entries = unzip(original);
    const core = text(entries, "docProps/core.xml");
    const workbook = text(entries, "xl/workbook.xml");
    expect(core).toContain("<cp:lastModifiedBy></cp:lastModifiedBy>");
    expect(workbook).toContain(`absPath url="${NEUTRAL_PATH}"`);
    const named = replaceEntries(
      original,
      new Map([
        [
          "docProps/core.xml",
          new TextEncoder().encode(
            core.replace("<cp:lastModifiedBy>", "<cp:lastModifiedBy>Some One"),
          ),
        ],
        [
          "xl/workbook.xml",
          new TextEncoder().encode(workbook.replace(NEUTRAL_PATH, "/Users/someone/Documents/")),
        ],
      ]),
    );
    expect(findingsInFile("named.xlsx", named)).toHaveLength(2);

    const blanked = anonymize(named)!;
    const after = unzip(blanked);
    expect(findingsInFile("blanked.xlsx", blanked)).toEqual([]);
    expect(after.map((entry) => entry.name)).toEqual(entries.map((entry) => entry.name));
    expect(text(after, "docProps/core.xml")).toBe(core);
    expect(text(after, "xl/workbook.xml")).toBe(workbook);
    for (const entry of entries.filter(({ name }) => !/core\.xml$|workbook\.xml$/.test(name))) {
      const twin = after.find(({ name }) => name === entry.name)!;
      expect(localRecord(blanked, twin), entry.name).toEqual(localRecord(original, entry));
    }
    expectLocalHeadersAgree(named);
    expectLocalHeadersAgree(blanked);
    expect(anonymize(blanked)).toBeUndefined();
  });

  it("refuses a part that states a field twice", () => {
    const original = read("excel/minimal.xlsx");
    const core = text(unzip(original), "docProps/core.xml");
    const twice = core.replace(
      "<cp:lastModifiedBy></cp:lastModifiedBy>",
      "<cp:lastModifiedBy>A</cp:lastModifiedBy><cp:lastModifiedBy>B</cp:lastModifiedBy>",
    );
    const doctored = replaceEntries(
      original,
      new Map([["docProps/core.xml", new TextEncoder().encode(twice)]]),
    );
    expect(() => anonymize(doctored)).toThrow("states the field 2 times");
  });

  function records(bytes: Uint8Array): { end: number; central: number; local: number } {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const end = bytes.length - 22;
    const local = unzip(bytes).find(({ name }) => name === "docProps/core.xml")!.localHeaderOffset;
    let central = view.getUint32(end + 16, true);
    while (view.getUint32(central + 42, true) !== local) {
      const name = view.getUint16(central + 28, true);
      central +=
        46 + name + view.getUint16(central + 30, true) + view.getUint16(central + 32, true);
    }
    return { end, central, local };
  }

  it.each<[string, (view: DataView, at: ReturnType<typeof records>) => void, string]>([
    ["a zip64 count", (view, { end }) => view.setUint16(end + 10, 0xffff, true), "zip64"],
    ["a zip64 offset", (view, { end }) => view.setUint32(end + 16, 0xffffffff, true), "zip64"],
    ["a zip64 locator", (view, { end }) => view.setUint32(end - 20, 0x07064b50, true), "zip64"],
    ["no end record", (view, { end }) => view.setUint32(end, 0, true), "no end of central"],
    ["a bad central header", (view, { central }) => view.setUint32(central, 0, true), "central"],
    ["a bad local header", (view, { local }) => view.setUint32(local, 0, true), "local header"],
    [
      "a data descriptor",
      (view, { central }) => view.setUint16(central + 8, 8, true),
      "descriptor",
    ],
    [
      "an encrypted entry",
      (view, { central }) => view.setUint16(central + 8, 1, true),
      "encrypted",
    ],
    ["another method", (view, { central }) => view.setUint16(central + 10, 12, true), "method 12"],
  ])("refuses a package with %s", (_, doctor, message) => {
    const bytes = read("excel/minimal.xlsx").slice();
    doctor(new DataView(bytes.buffer), records(bytes));
    expect(() => replaceEntries(bytes, new Map([["docProps/core.xml", new Uint8Array()]]))).toThrow(
      message,
    );
  });
});
