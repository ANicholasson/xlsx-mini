// Run by Node (ESM), Bun and Deno: the package resolves by name through its own exports map.
import * as browser from "xlsx-mini/browser";
import * as read from "xlsx-mini/read";
import * as lib from "xlsx-mini";
import { run } from "./scenario.mjs";

console.log(JSON.stringify(await run(lib, browser, read)));
