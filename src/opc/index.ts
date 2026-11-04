/**
 * Open Packaging Conventions container: the zip archive an .xlsx file is.
 * Knows nothing about spreadsheets.
 */
export { crc32 } from "./lib/crc32.js";
export { createPackage, type PackagePart } from "./lib/zip.js";
