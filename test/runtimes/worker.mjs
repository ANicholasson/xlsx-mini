// Bundled by the matrix test and run inside workerd through Miniflare.
import * as browser from "../../dist/browser.js";
import * as lib from "../../dist/index.js";
import { run } from "./scenario.mjs";

export default {
  async fetch() {
    return Response.json(await run(lib, browser));
  },
};
