/**
 * Open Packaging Conventions container: the zip archive an .xlsx file is.
 * Knows nothing about spreadsheets.
 */
export {
  type CompressionMode,
  type StreamCompression,
  selectCompression,
  selectStreamCompression,
  store,
} from "./lib/compression.js";
export { CRC32_INITIAL, crc32, crc32Finish, crc32Update } from "./lib/crc32.js";
export { MAX_STORED_BLOCK, storedBlock } from "./lib/stored-blocks.js";
export { createPackage, type PackagePart } from "./lib/zip.js";
export { type StreamedPart, streamPackage } from "./lib/zip-stream.js";
