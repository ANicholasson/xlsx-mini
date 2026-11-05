// Run by Node as CommonJS: `require` of the ESM-only package and of the ESM scenario.
"use strict";

const browser = require("xlsx-mini/browser");
const lib = require("xlsx-mini");
const { run } = require("./scenario.mjs");

run(lib, browser).then(
  (report) => console.log(JSON.stringify(report)),
  (error) => {
    console.error(error);
    process.exitCode = 1;
  },
);
