// SPEC-087 AC-46: hand-built byte cases for the raw Snappy block decoder —
// literals (short and length-byte forms), all three copy tag forms,
// overlapping copies, and truncation/corruption raising the typed error.
import { describe, expect, it } from "vitest";
import { snappyDecode } from "./snappy";
import { KeynoteOpenError } from "./errors";

const encoder = new TextEncoder();

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

function concat(...parts: (Uint8Array | number[])[]): Uint8Array {
  const arrays = parts.map((part) => (Array.isArray(part) ? Uint8Array.from(part) : part));
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const arr of arrays) {
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

describe("snappyDecode", () => {
  it("decodes a short literal", () => {
    // length 5, literal tag (5-1)<<2 = 16, then the bytes.
    const block = concat([5, 16], encoder.encode("hello"));
    expect(new TextDecoder().decode(snappyDecode(block))).toBe("hello");
  });

  it("decodes a literal with a one-byte length (len > 60)", () => {
    const payload = new Uint8Array(100).fill(0x41);
    // tag 60<<2 = 240 means "length in next 1 byte", stored as len-1.
    const block = concat([100, 240, 99], payload);
    expect(snappyDecode(block)).toEqual(payload);
  });

  it("decodes a literal with a two-byte length", () => {
    const payload = new Uint8Array(300).fill(0x42);
    // varint 300 = 0xac 0x02; tag 61<<2 = 244 = two length bytes, LE 299.
    const block = concat([0xac, 0x02, 244, 299 & 0xff, 299 >> 8], payload);
    expect(snappyDecode(block)).toEqual(payload);
  });

  it("decodes a near copy (tag 1: 3-bit length, 11-bit offset)", () => {
    // "abcd" then copy 4 bytes from offset 4 → "abcdabcd".
    // copy-1 tag: kind 1, length 4 → ((4-4)&7)<<2 = 0, offset 4 → high 3
    // bits 0, low byte 4: tag = 0b000_000_01 = 1, then offset byte 4.
    const block = concat([8, 12], encoder.encode("abcd"), [1, 4]);
    expect(new TextDecoder().decode(snappyDecode(block))).toBe("abcdabcd");
  });

  it("decodes an overlapping copy (RLE-style repetition)", () => {
    // "ab" then copy 6 bytes from offset 2 → "abababab".
    // copy-1: length 6 → ((6-4)&7)<<2 = 8, kind 1 → tag 9, offset byte 2.
    const block = concat([8, 4], encoder.encode("ab"), [9, 2]);
    expect(new TextDecoder().decode(snappyDecode(block))).toBe("abababab");
  });

  it("decodes a copy with a 16-bit offset (tag 2)", () => {
    // 70 literal bytes, then copy 5 from offset 70 (start of output).
    const payload = concat(encoder.encode("start"), new Uint8Array(65).fill(0x2e));
    // copy-2 tag: (5-1)<<2 | 2 = 18, offset LE 70,0.
    const block = concat([75, 240, 69], payload, [18, 70, 0]);
    expect(new TextDecoder().decode(snappyDecode(block).subarray(70))).toBe("start");
  });

  it("decodes a copy with a 32-bit offset (tag 3)", () => {
    const block = concat([8, 12], encoder.encode("wxyz"), [(4 - 1) * 4 + 3, 4, 0, 0, 0]);
    expect(new TextDecoder().decode(snappyDecode(block))).toBe("wxyzwxyz");
  });

  it("decodes an empty block", () => {
    expect(snappyDecode(bytes(0))).toEqual(new Uint8Array(0));
  });

  it("throws the typed error on a literal running past the end", () => {
    expect(() => snappyDecode(bytes(5, 16, 0x61))).toThrow(KeynoteOpenError);
  });

  it("throws the typed error when output falls short of the declared length", () => {
    expect(() => snappyDecode(concat([9, 16], encoder.encode("hello")))).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on a copy offset before the start of output", () => {
    expect(() => snappyDecode(concat([8, 4], encoder.encode("ab"), [9, 200]))).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on a truncated length varint", () => {
    expect(() => snappyDecode(bytes(0x80))).toThrow(KeynoteOpenError);
  });
});
