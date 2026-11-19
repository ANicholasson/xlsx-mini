import { XlsxReadError } from "../../errors/read.js";

/**
 * Random access to the file the caller gave, so that bytes and a `Blob` share every code path
 * after this one: a `Blob` is read a slice at a time and never whole, and bytes are viewed.
 */
export interface ByteSource {
  readonly size: number;
  /**
   * The bytes of `[start, end)`, or those of them the file has. `start` is never negative and
   * never past `end`: both a view and a slice would count a negative one from the end.
   */
  read(start: number, end: number): Promise<Uint8Array<ArrayBuffer>>;
}

export function byteSource(file: Uint8Array | ArrayBuffer | Blob): ByteSource {
  if (file instanceof Blob) {
    return {
      size: file.size,
      read: async (start, end) => new Uint8Array(await file.slice(start, end).arrayBuffer()),
    };
  }
  const bytes = bytesOf(file);
  return {
    size: bytes.length,
    read: (start, end) => Promise.resolve(bytes.subarray(start, end)),
  };
}

function bytesOf(file: Uint8Array | ArrayBuffer): Uint8Array<ArrayBuffer> {
  if (file instanceof ArrayBuffer) return new Uint8Array(file);
  if (!(file instanceof Uint8Array)) {
    throw new XlsxReadError("NOT_XLSX", "expected a Uint8Array, an ArrayBuffer or a Blob");
  }
  const { buffer } = file;
  // The inflate of Node 24, Bun and Deno refuses bytes over shared memory, so those are copied, once.
  return buffer instanceof ArrayBuffer
    ? new Uint8Array(buffer, file.byteOffset, file.byteLength)
    : new Uint8Array(file);
}
