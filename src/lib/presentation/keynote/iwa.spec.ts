// SPEC-087 AC-46/AC-49: IWA chunk framing (including a message split across
// two chunks, which only works because decoded chunks concatenate into one
// stream) and TSP.ArchiveInfo parsing — hand-built bytes throughout.
import { describe, expect, it } from "vitest";
import { decodeIwa, parseArchives } from "./iwa";
import { KeynoteOpenError } from "./errors";

/** A raw Snappy block that is one big literal of `payload`. */
function literalBlock(payload: Uint8Array): Uint8Array {
  if (payload.length === 0 || payload.length > 60) {
    throw new Error("test helper handles 1..60-byte payloads");
  }
  return Uint8Array.from([payload.length, (payload.length - 1) << 2, ...payload]);
}

/** Wraps a Snappy block as one IWA chunk: 0x00 + 3-byte LE length. */
function chunk(block: Uint8Array): Uint8Array {
  return Uint8Array.from([0x00, block.length & 0xff, (block.length >> 8) & 0xff, (block.length >> 16) & 0xff, ...block]);
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

describe("decodeIwa", () => {
  it("decodes a single chunk", () => {
    const payload = Uint8Array.from([1, 2, 3, 4]);
    expect(decodeIwa(chunk(literalBlock(payload)))).toEqual(payload);
  });

  it("concatenates multiple chunks into one stream", () => {
    const a = Uint8Array.from([10, 20]);
    const b = Uint8Array.from([30, 40, 50]);
    expect(decodeIwa(concat(chunk(literalBlock(a)), chunk(literalBlock(b))))).toEqual(
      Uint8Array.from([10, 20, 30, 40, 50]),
    );
  });

  it("throws the typed error on an unknown chunk type byte", () => {
    const bad = concat(Uint8Array.from([0x01]), literalBlock(Uint8Array.from([1])));
    expect(() => decodeIwa(bad)).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on a chunk length past the end", () => {
    expect(() => decodeIwa(Uint8Array.from([0x00, 0xff, 0x00, 0x00, 0x01]))).toThrow(KeynoteOpenError);
  });
});

describe("parseArchives", () => {
  // ArchiveInfo for object 7, one MessageInfo {type 5, length 2}:
  //   field 1 (varint 7): 0x08 0x07
  //   field 2 (len-delim MessageInfo): 0x12 len [0x08 type, 0x18 length]
  const info = Uint8Array.from([0x08, 0x07, 0x12, 0x04, 0x08, 0x05, 0x18, 0x02]);
  const payload = Uint8Array.from([0x08, 0x2a]); // any 2 bytes
  const record = Uint8Array.from([info.length, ...info, ...payload]);

  it("parses an object's identifier, type and payload (AC-49)", () => {
    const archives = parseArchives(record);
    expect(archives.size).toBe(1);
    const messages = archives.get(7n);
    expect(messages).toHaveLength(1);
    expect(messages?.[0].type).toBe(5);
    expect(Array.from(messages?.[0].payload ?? [])).toEqual([0x08, 0x2a]);
  });

  it("parses consecutive payloads for multiple MessageInfos", () => {
    // Object 9 with two MessageInfos: {type 1, len 1} and {type 2, len 3}.
    const twoInfo = Uint8Array.from([
      0x08, 0x09,
      0x12, 0x04, 0x08, 0x01, 0x18, 0x01,
      0x12, 0x04, 0x08, 0x02, 0x18, 0x03,
    ]);
    const stream = Uint8Array.from([twoInfo.length, ...twoInfo, 0xaa, 0xbb, 0xcc, 0xdd]);
    const messages = parseArchives(stream).get(9n);
    expect(messages).toHaveLength(2);
    expect(Array.from(messages?.[0].payload ?? [])).toEqual([0xaa]);
    expect(messages?.[1].type).toBe(2);
    expect(Array.from(messages?.[1].payload ?? [])).toEqual([0xbb, 0xcc, 0xdd]);
  });

  it("parses a stream reassembled from two IWA chunks (record split across the boundary)", () => {
    // The whole record's bytes split mid-record between two chunks — legal
    // because chunk payloads concatenate before the stream is parsed.
    const first = literalBlock(record.subarray(0, 5));
    const second = literalBlock(record.subarray(5));
    const stream = decodeIwa(concat(chunk(first), chunk(second)));
    expect(parseArchives(stream).get(7n)?.[0].type).toBe(5);
  });

  it("throws the typed error on a payload running past the end (AC-42)", () => {
    const truncated = record.subarray(0, record.length - 1);
    expect(() => parseArchives(truncated)).toThrow(KeynoteOpenError);
  });

  it("throws the typed error on an ArchiveInfo without an identifier", () => {
    const noId = Uint8Array.from([0x02, 0x18, 0x01]); // just some field 3
    expect(() => parseArchives(noId)).toThrow(KeynoteOpenError);
  });
});
