import { crc32 as nodeCrc32 } from "node:zlib";
import { describe, expect, it } from "vite-plus/test";
import { crc32 } from "../index.js";

const bytes = (text: string) => new TextEncoder().encode(text);

describe("crc32", () => {
  it("matches the published check values", () => {
    expect(crc32(bytes(""))).toBe(0);
    expect(crc32(bytes("123456789"))).toBe(0xcbf43926);
    expect(crc32(bytes("The quick brown fox jumps over the lazy dog"))).toBe(0x414fa339);
  });

  it("agrees with Node's implementation on binary data", () => {
    const data = new Uint8Array(4096);
    for (let i = 0; i < data.length; i++) data[i] = (i * 7919) & 0xff;
    expect(crc32(data)).toBe(nodeCrc32(data));
    expect(crc32(data.subarray(13, 2999))).toBe(nodeCrc32(data.subarray(13, 2999)));
  });
});
