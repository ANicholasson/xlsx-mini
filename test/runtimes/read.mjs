import { resolveLimits } from "../../dist/errors/read.js";
import { openPackage, partName } from "../../dist/opc/read.js";
import { openXlsx } from "../../dist/read.js";

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
    return { sheet: btoa(binary), sheets: (await openXlsx(bytes)).sheets };
  } catch (error) {
    return { refused: error.code ?? String(error) };
  }
}

const decoder = new TextDecoder();
let input = "";
for await (const chunk of globalThis.Deno?.stdin.readable ?? process.stdin) {
  input += decoder.decode(chunk, { stream: true });
}
const packages = JSON.parse(input + decoder.decode());
const outcomes = [];
for (const base64 of packages) outcomes.push(await outcome(base64));
console.log(JSON.stringify({ outcomes }));
