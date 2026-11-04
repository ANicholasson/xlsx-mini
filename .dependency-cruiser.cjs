// @ts-check
// Deep-module enforcement for dependency-cruiser. Adapted from the setup-ts-deep-modules
// convention: every feature folder directly under src/ is a DEEP MODULE whose PUBLIC
// SURFACE is its root files (entry points). Anything in a subfolder (lib/, tests/, any
// other) is private. Root files of src/ itself (index.ts, write-xlsx.ts) are app code:
// they may import entry points only.
//
// The only things to edit here are PACKAGES_ROOT and the optional layering rules.

/** Where features live. One immediate child dir per feature (flat, no nesting). */
const PACKAGES_ROOT = "src";

// --- derived patterns (no need to edit) -------------------------------------
const R = PACKAGES_ROOT;
/** A feature's private internals: anything nested inside one of its subfolders. */
const PACKAGE_INTERNALS = `^${R}/[^/]+/[^/]+/`;
/** Any file that lives inside a feature folder. */
const INSIDE_A_PACKAGE = `^${R}/[^/]+/`;

/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    {
      name: "entrypoint-boundary-from-app",
      comment:
        "Code outside the features (src root files, test/, config) may import a feature's entry points (its root files), but nothing inside its subfolders.",
      severity: "error",
      from: { pathNot: INSIDE_A_PACKAGE },
      to: { path: PACKAGE_INTERNALS },
    },
    {
      name: "entrypoint-boundary-across-packages",
      comment:
        "A feature's own files import each other freely, but may reach OTHER features only through their entry points, never their internals.",
      severity: "error",
      from: { path: `^${R}/([^/]+)/`, pathNot: `^${R}/[^/]+/tests/` },
      to: { path: PACKAGE_INTERNALS, pathNot: `^${R}/$1/` },
    },
    {
      name: "tests-through-entrypoints",
      comment:
        "A feature's tests exercise it through its entry points: they may import any feature's entry points and their own tests/ fixtures, never any feature's internals, not even their own.",
      severity: "error",
      from: { path: `^${R}/([^/]+)/tests/` },
      to: { path: PACKAGE_INTERNALS, pathNot: `^${R}/$1/tests/` },
    },
    {
      name: "tests-folder-is-private",
      comment:
        "A feature's tests/ folder is reachable only from tests: nothing else may import fixtures.",
      severity: "error",
      from: { pathNot: `^${R}/[^/]+/tests/` },
      to: { path: `^${R}/[^/]+/tests/` },
    },
    {
      name: "no-circular",
      comment: "No dependency cycles.",
      severity: "error",
      from: {},
      to: { circular: true },
    },

    // --- Layering ---------------------------------------------------------------
    // The kernels (xml, opc) know nothing about spreadsheets: they may not import features.
    {
      name: "kernels-stay-ignorant",
      comment: "xml/ and opc/ are kernels; they may not depend on any spreadsheet feature.",
      severity: "error",
      from: { path: `^${R}/(xml|opc)/` },
      to: { path: `^${R}/(workbook|worksheet|styles|shared-strings|drawing)/` },
    },
  ],
  options: {
    // dependency-cruiser 18 cannot parse TypeScript 7 through the typescript package; swc can.
    parser: "swc",
    doNotFollow: { path: "node_modules" },
    tsPreCompilationDeps: true,
    tsConfig: { fileName: "tsconfig.json" },
    enhancedResolveOptions: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json"],
    },
  },
};
