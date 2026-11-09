import { crc32 as nodeCrc32 } from "node:zlib";
import { describe, expect, it } from "vite-plus/test";
import { CRC32_INITIAL, crc32, crc32Finish, crc32Update } from "../index.js";

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

  it("gives the same checksum fed in chunks as in one piece", () => {
    const data = new Uint8Array(1000);
    for (let i = 0; i < data.length; i++) data[i] = (i * 31) & 0xff;
    let state = CRC32_INITIAL;
    for (let offset = 0; offset < data.length; offset += 37) {
      state = crc32Update(state, data.subarray(offset, offset + 37));
    }
    expect(crc32Finish(state)).toBe(crc32(data));
    expect(crc32Finish(CRC32_INITIAL)).toBe(0);
  });
});
