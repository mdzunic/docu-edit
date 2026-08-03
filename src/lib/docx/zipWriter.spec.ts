import { describe, it, expect } from "vitest";
import { crc32, zip } from "./zipWriter";
import { unzip } from "./zip";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

describe("crc32", () => {
  it("matches the standard check value for '123456789'", () => {
    expect(crc32(encoder.encode("123456789"))).toBe(0xcbf43926);
  });

  it("is 0 for empty input", () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });
});

describe("zip", () => {
  const entries = [
    { name: "[Content_Types].xml", data: encoder.encode("<Types/>") },
    { name: "word/document.xml", data: encoder.encode("<w:document>Hello, docx!</w:document>") },
    { name: "docProps/core.xml", data: encoder.encode("x".repeat(10_000)) },
  ];

  it("produces an archive our OPC reader unzips back to identical parts", async () => {
    const archive = await zip(entries);
    const parts = await unzip(archive);
    expect(Array.from(parts.keys())).toEqual(entries.map((entry) => entry.name));
    for (const entry of entries) {
      expect(decoder.decode(parts.get(entry.name))).toBe(decoder.decode(entry.data));
    }
  });

  it("starts with a local file header signature (AC-3)", async () => {
    const archive = await zip(entries);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
  });

  it("writes an end-of-central-directory record with the right entry count", async () => {
    const archive = await zip(entries);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    const eocdOffset = archive.byteLength - 22; // no archive comment
    expect(view.getUint32(eocdOffset, true)).toBe(0x06054b50);
    expect(view.getUint16(eocdOffset + 10, true)).toBe(entries.length);
  });

  it("records CRC-32 and sizes that match each entry's actual content", async () => {
    const archive = await zip(entries);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    const eocdOffset = archive.byteLength - 22;
    let offset = view.getUint32(eocdOffset + 16, true);
    for (const entry of entries) {
      expect(view.getUint32(offset, true)).toBe(0x02014b50);
      expect(view.getUint32(offset + 16, true)).toBe(crc32(entry.data));
      expect(view.getUint32(offset + 24, true)).toBe(entry.data.length);
      const nameLength = view.getUint16(offset + 28, true);
      expect(decoder.decode(archive.subarray(offset + 46, offset + 46 + nameLength))).toBe(
        entry.name,
      );
      offset += 46 + nameLength;
    }
  });

  it("compresses: the archive is smaller than highly repetitive input", async () => {
    const archive = await zip([
      { name: "big.xml", data: encoder.encode("repeat ".repeat(5_000)) },
    ]);
    expect(archive.byteLength).toBeLessThan(5_000);
  });
});
