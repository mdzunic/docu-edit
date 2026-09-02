// SPEC-087 AC-47/AC-48: varint edge values (0, 300, 2^32, 2^53-ish) and each
// supported wire type, plus unknown-field and error behaviour — all
// hand-built bytes, no encoder shared with the code under test.
import { describe, expect, it } from "vitest";
import { fields, float32Of, readVarint, WIRE_FIXED32, WIRE_FIXED64, WIRE_LENGTH_DELIMITED, WIRE_VARINT } from "./protobuf";
import { KeynoteOpenError } from "./errors";

function varintOf(bytes: number[]): number {
  return readVarint({ bytes: Uint8Array.from(bytes), pos: 0 });
}

describe("readVarint", () => {
  it("reads 0", () => {
    expect(varintOf([0x00])).toBe(0);
  });

  it("reads 300", () => {
    expect(varintOf([0xac, 0x02])).toBe(300);
  });

  it("reads 2^32", () => {
    // 2^32 = 0x100000000 → groups of 7 bits LE: 0,0,0,0,16.
    expect(varintOf([0x80, 0x80, 0x80, 0x80, 0x10])).toBe(2 ** 32);
  });

  it("reads 2^53 - 1 (the exactness ceiling)", () => {
    // 53 set bits → seven 0xff continuation bytes then 0x0f.
    expect(varintOf([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x0f])).toBe(2 ** 53 - 1);
  });

  it("throws the typed error at 2^53 (no longer a safe integer)", () => {
    expect(() => varintOf([0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x10])).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on an unterminated varint", () => {
    expect(() => varintOf([0x80, 0x80])).toThrow(KeynoteOpenError);
  });
});

describe("fields", () => {
  it("decodes a varint field (wire 0)", () => {
    // field 1, wire 0, value 150: key 0x08, varint 0x96 0x01.
    const decoded = fields(Uint8Array.from([0x08, 0x96, 0x01]));
    expect(decoded).toEqual([{ num: 1, wire: WIRE_VARINT, value: 150 }]);
  });

  it("decodes a fixed64 field (wire 1)", () => {
    // field 2, wire 1: key 0x11, then 8 payload bytes.
    const decoded = fields(Uint8Array.from([0x11, 1, 2, 3, 4, 5, 6, 7, 8]));
    expect(decoded).toHaveLength(1);
    expect(decoded[0].num).toBe(2);
    expect(decoded[0].wire).toBe(WIRE_FIXED64);
    expect(Array.from(decoded[0].bytes ?? [])).toEqual([1, 2, 3, 4, 5, 6, 7, 8]);
  });

  it("decodes a length-delimited field (wire 2)", () => {
    // field 3, wire 2: key 0x1a, length 3, "abc".
    const decoded = fields(Uint8Array.from([0x1a, 3, 0x61, 0x62, 0x63]));
    expect(decoded).toHaveLength(1);
    expect(decoded[0].num).toBe(3);
    expect(decoded[0].wire).toBe(WIRE_LENGTH_DELIMITED);
    expect(new TextDecoder().decode(decoded[0].bytes)).toBe("abc");
  });

  it("decodes a fixed32 field (wire 5) and its float32 value", () => {
    // field 4, wire 5: key 0x25, float32 LE 1920.
    const view = new DataView(new ArrayBuffer(4));
    view.setFloat32(0, 1920, true);
    const decoded = fields(Uint8Array.from([0x25, ...new Uint8Array(view.buffer)]));
    expect(decoded[0].num).toBe(4);
    expect(decoded[0].wire).toBe(WIRE_FIXED32);
    expect(float32Of(decoded[0])).toBe(1920);
  });

  it("walks past unknown field numbers by wire type (AC-40)", () => {
    // field 999 varint (key 999·8 = 7992 → 0xb8 0x3e), then field 1 varint 7
    // — the reader sees both; a consumer wanting field 1 ignores the other.
    const decoded = fields(Uint8Array.from([0xb8, 0x3e, 0x05, 0x08, 0x07]));
    expect(decoded.map((field) => field.num)).toEqual([999, 1]);
  });

  it("throws the typed error on a group wire type (unskippable)", () => {
    // field 1, wire 3 (group start): key 0x0b.
    expect(() => fields(Uint8Array.from([0x0b]))).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on a truncated length-delimited field", () => {
    expect(() => fields(Uint8Array.from([0x1a, 10, 0x61]))).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on field number 0", () => {
    expect(() => fields(Uint8Array.from([0x00]))).toThrow(KeynoteOpenError);
  });
});
