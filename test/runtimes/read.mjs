// Run by Node, Bun and Deno: reads the sheet of each package it is given in base64 with the
// package reader in dist/, and prints one line of JSON, a `ReadReport`.
import { resolveLimits } from "../../dist/errors/read.js";
import { openPackage, partName } from "../../dist/opc/read.js";

const sheet = "xl/worksheets/sheet1.xml";

async function outcome(base64) {
  const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
  const limits = resolveLimits(undefined);
  try {
    const pkg = await openPackage(bytes, limits.book.parts, limits.book.structure);
    const part = pkg.open(partName(sheet, "worksheet"), limits.rows().inflate);
    let binary = "";
    for (let chunk = await part.read(); chunk !== undefined; chunk = await part.read()) {
      for (const byte of chunk) binary += String.fromCharCode(byte);
    }
    return { sheet: btoa(binary) };
  } catch (error) {
    return { refused: error.code ?? String(error) };
  }
}

const packages = globalThis.Deno?.args ?? process.argv.slice(2);
const outcomes = [];
for (const base64 of packages) outcomes.push(await outcome(base64));
console.log(JSON.stringify({ outcomes }));
