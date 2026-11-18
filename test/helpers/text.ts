export const textsToReadBack: readonly string[] = [
  "_x0041_",
  "part _x00e6_ no",
  "_x005F_",
  "_x000D_",
  "_xZZZZ_",
  "_x41_",
  "_X0041_",
  "_x0041_x0042_",
  "_x0041\r",
  "a\rb",
  "a\r\nb",
  "a\nb",
];

const ENTITIES: Readonly<Record<string, string>> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

function decodeXml(xml: string): string {
  return xml
    .replace(/\r\n?/g, "\n")
    .replace(/&(#\d+|\w+);/g, (reference, name: string) =>
      name.startsWith("#")
        ? String.fromCodePoint(Number(name.slice(1)))
        : (ENTITIES[name] ?? reference),
    );
}

/** Element text as Excel and LibreOffice read it, decoded independently of the writer. */
export function readText(xml: string): string {
  return decodeXml(xml).replace(/_x([\dA-Fa-f]{4})_/g, (_, code: string) =>
    String.fromCharCode(Number.parseInt(code, 16)),
  );
}

/**
 * A shared string as Numbers 14.4 reads it, as measured: `_x005F_` with the `xHHHH_` after it
 * is the text `_xHHHH_`, so the underscore that closes it never starts another escape. It reads
 * `_x005F_x0041_x005F_x0042_` as `_x0041_x005FB`. Numbers decodes no escape in an inline string.
 */
export function readTextAsNumbers(xml: string): string {
  return decodeXml(xml).replace(
    /_x005F_(x[\dA-Fa-f]{4}_)|_x[\dA-Fa-f]{4}_/g,
    (escape, literal?: string) => (literal === undefined ? readText(escape) : `_${literal}`),
  );
}

export function keptByXml(text: string): string {
  let kept = "";
  for (const character of text) {
    const point = character.codePointAt(0) ?? 0;
    if (
      point === 0x9 ||
      point === 0xa ||
      point === 0xd ||
      (point >= 0x20 && point <= 0xd7ff) ||
      (point >= 0xe000 && point <= 0xfffd) ||
      point >= 0x10000
    ) {
      kept += character;
    }
  }
  return kept;
}
