// Bundled by the matrix test as a classic script and evaluated inside the Vercel edge VM.
import * as browser from "../../dist/browser.js";
import * as lib from "../../dist/index.js";
import { run } from "./scenario.mjs";

globalThis.xlsxMiniReport = run(lib, browser);
