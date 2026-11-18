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
  "_x_x0041_",
  "box_xyz_x0041_",
  "a_x b_x0041_",
  "_xZ_x0041_",
  "_x\r_x0041_",
  "id_x12_x0041_",
  "_x_x0078_",
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

export function readTextAsExcel(xml: string): string {
  return decodeXml(xml).replace(/_x([\dA-Fa-f]{4})_/g, (_, code: string) =>
    String.fromCharCode(Number.parseInt(code, 16)),
  );
}

/**
 * A shared string as Numbers 14.4 reads it, as measured on 1 014 texts: an `_x` opens a piece and
 * the next underscore closes it. A piece of `_xHHHH_` is that character, and `_x005F_` with the
 * `xHHHH_` after it is the text `_xHHHH_`. Any other piece is text up to and with its underscore,
 * so that underscore starts no escape: `_x_x005F_x0041_` reads `_x_x005FA`. Numbers decodes no
 * escape in an inline string.
 */
export function readTextAsNumbers(xml: string): string {
  const text = decodeXml(xml);
  let read = "";
  let index = 0;
  for (let open = text.indexOf("_x"); open !== -1; open = text.indexOf("_x", index)) {
    const close = text.indexOf("_", open + 2);
    if (close === -1) break;
    read += text.slice(index, open);
    index = close + 1;
    const piece = text.slice(open, index);
    if (!/^_x[\dA-Fa-f]{4}_$/.test(piece)) {
      read += piece;
      continue;
    }
    const code = Number.parseInt(piece.slice(2, 6), 16);
    const escaped = /^x[\dA-Fa-f]{4}_/.exec(text.slice(index))?.[0];
    if (code === 0x5f && escaped !== undefined) {
      read += `_${escaped}`;
      index += escaped.length;
    } else {
      read += String.fromCharCode(code);
    }
  }
  return read + text.slice(index);
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
