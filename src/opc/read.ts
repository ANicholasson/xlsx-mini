/**
 * The read side of the package: finds the parts of a zip container and hands out one part's bytes
 * at a time. Knows nothing about spreadsheets and nothing about XML.
 *
 * Everything that makes a hostile package safe to open is behind `openPackage` and `open`: the
 * directory is the only source of offsets and sizes, a part is charged before it inflates, its
 * size and CRC-32 are enforced, and the platform's inflate is the only one.
 */
import type { Meter, PartMeter, PartUse } from "../errors/read.js";
import { readDirectory } from "./read/directory.js";
import { foldName } from "./read/layout.js";
import { type PartBytes, framingOnce, openPart } from "./read/part.js";
import { byteSource } from "./read/source.js";

export type { PartBytes } from "./read/part.js";

const partUse = Symbol();

/**
 * The name of a part together with what its bytes will be read for. The use is decided once,
 * where the kind of the part is known, and `open` then takes only a meter of that use. A name
 * whose use is one of several opens under no meter: a meter has one use.
 */
export interface PartName<Use extends PartUse> {
  readonly name: string;
  readonly [partUse]: Use;
}

/** Says what a part will be read for. The only way to make a `PartName`. */
export function partName<Use extends PartUse>(name: string, use: Use): PartName<Use> {
  return { name, [partUse]: use };
}

/** An opened package: its directory has been read and checked, no part has. */
export interface PackageReader {
  /**
   * Opens a part by its name, compared without regard to the case of the letters A to Z, or
   * returns `undefined` when the package has no such part: the caller decides what a missing part
   * means.
   *
   * The part's declared size is charged to `meter` here, before anything of the part is read, so a
   * part that does not fit is refused with `LIMIT_EXCEEDED` and not a byte of it is inflated. There
   * is no way to a part's bytes without a meter of the use its name carries: the use is taken from
   * the name alone, so a meter of another use does not compile.
   */
  open<Use extends PartUse>(
    part: PartName<Use>,
    meter: PartMeter<NoInfer<Use>>,
  ): PartBytes | undefined;
}

/**
 * Reads the package's directory and nothing else: the end of the file, then the directory. A
 * `Blob` is read in slices and never whole. Bytes are viewed, so the caller must not change them
 * while the package is in use; bytes over shared memory are copied.
 *
 * The directory is bounded before it is read: its size as the end of the file states it is
 * charged to `structure`, and every entry to `parts` before it is kept. A directory that does not
 * fit is refused with `LIMIT_EXCEEDED` after one read of the end of the file.
 *
 * Throws `NOT_XLSX` when the file is no package, naming a compound file, and `INVALID_PACKAGE`
 * when the package is cut short, has bytes in front of it or after it, has an end that can be read
 * in two ways, a damaged directory or a part outside it, or names a part twice. A zip64 package is
 * accepted, and so is a comment after the end record. An error of the `Blob` itself is passed on
 * as it is: it says nothing about the file.
 */
export async function openPackage(
  file: Uint8Array | ArrayBuffer | Blob,
  parts: Meter<"parts">,
  structure: PartMeter<"structure">,
): Promise<PackageReader> {
  const source = byteSource(file);
  const directory = await readDirectory(source, parts, structure);
  const framing = framingOnce();
  return {
    open(part, meter) {
      const entry = directory.entries.get(foldName(part.name));
      if (entry === undefined) return undefined;
      meter.charge(entry.size);
      return openPart(source, directory.start, entry, framing);
    },
  };
}
