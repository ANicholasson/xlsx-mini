import { resolveLimits } from "../../errors/read.js";
import { type StreamedPart, streamPackage } from "../index.js";
import { type PackageReader, type PartBytes, openPackage, partName } from "../read.js";

export type Limits = ReturnType<typeof resolveLimits>;

export const encode = (text: string): Uint8Array<ArrayBuffer> => new TextEncoder().encode(text);

/** Bytes that no deflate shortens. */
export function noise(length: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(length);
  let state = 0x2545f491;
  for (let at = 0; at < length; at++) {
    state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
    bytes[at] = state >>> 24;
  }
  return bytes;
}

export function join(chunks: readonly Uint8Array[]): Uint8Array<ArrayBuffer> {
  const joined = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    joined.set(chunk, offset);
    offset += chunk.length;
  }
  return joined;
}

export function open(
  file: Uint8Array | ArrayBuffer | Blob,
  limits: Limits = resolveLimits(undefined),
): Promise<PackageReader> {
  return openPackage(file, limits.book.parts, limits.book.structure);
}

/** Every chunk of a part, joined. */
export async function gather(part: PartBytes): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for (let chunk = await part.read(); chunk !== undefined; chunk = await part.read()) {
    chunks.push(chunk.slice());
  }
  return join(chunks);
}

/** One part read whole the way a worksheet is: under the meter of one rows iteration. */
export async function read(
  pkg: PackageReader,
  name: string,
  limits: Limits = resolveLimits(undefined),
): Promise<Uint8Array | undefined> {
  const part = pkg.open(partName(name, "worksheet"), limits.rows().inflate);
  return part === undefined ? undefined : gather(part);
}

/** What `streamPackage` writes for parts whose bodies arrive in one chunk each. */
export async function streamed(
  parts: readonly { readonly path: string; readonly data: Uint8Array }[],
  compression: Parameters<typeof streamPackage>[1],
): Promise<Uint8Array<ArrayBuffer>> {
  async function* bodies(): AsyncGenerator<StreamedPart> {
    for (const { path, data } of parts) {
      yield { path, body: new Blob([new Uint8Array(data)]).stream() };
    }
  }
  return new Uint8Array(await new Response(streamPackage(bodies(), compression)).arrayBuffer());
}

/** A `Blob` that records which ranges were read from it. */
export class WatchedBlob extends Blob {
  readonly reads: (readonly [start: number, end: number])[] = [];

  override slice(start = 0, end = this.size): Blob {
    this.reads.push([start, end]);
    return super.slice(start, end);
  }
}

/** Where the fields this suite changes sit in a record. */
export const CENTRAL = {
  flags: 8,
  method: 10,
  compressedSize: 20,
  size: 24,
  nameLength: 28,
  extraLength: 30,
  offset: 42,
  name: 46,
} as const;
export const LOCAL = {
  signature: 0,
  flags: 6,
  method: 8,
  crc32: 14,
  compressedSize: 18,
  size: 22,
  nameLength: 26,
  extraLength: 28,
  name: 30,
} as const;
export const END = { entries: 10, size: 12, offset: 16, commentLength: 20, bytes: 22 } as const;

/** Where the records of a package without zip64 and without a comment start. */
export function records(bytes: Uint8Array): {
  readonly end: number;
  readonly central: readonly number[];
  readonly local: readonly number[];
} {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = bytes.length - END.bytes;
  const central: number[] = [];
  const local: number[] = [];
  const directoryEnd =
    view.getUint32(end + END.offset, true) + view.getUint32(end + END.size, true);
  for (let at = view.getUint32(end + END.offset, true); at < directoryEnd;) {
    central.push(at);
    local.push(view.getUint32(at + CENTRAL.offset, true));
    at +=
      CENTRAL.name +
      view.getUint16(at + CENTRAL.nameLength, true) +
      view.getUint16(at + CENTRAL.extraLength, true);
  }
  return { end, central, local };
}

/** A package that has no comment yet, with this one after its end record. */
export function commented(bytes: Uint8Array, comment: Uint8Array): Uint8Array<ArrayBuffer> {
  return changed(join([bytes, comment]), (view) =>
    view.setUint16(bytes.length - END.bytes + END.commentLength, comment.length, true),
  );
}

/** A copy of a package with the changes `change` makes to it. */
export function changed(
  bytes: Uint8Array,
  change: (view: DataView, bytes: Uint8Array) => void,
): Uint8Array<ArrayBuffer> {
  const copy = new Uint8Array(bytes);
  change(new DataView(copy.buffer), copy);
  return copy;
}
