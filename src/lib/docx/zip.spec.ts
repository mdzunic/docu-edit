import { describe, it, expect } from "vitest";
import { deflateRawSync } from "node:zlib";
import { unzip, ZipError } from "./zip";

// Test-only ZIP writer. Uses Node's zlib (only available under Vitest, never
// bundled into the app) so DEFLATE entries are byte-for-byte what a real
// ZIP tool would produce, exercising the same DecompressionStream path the
// browser build relies on.
interface FixtureEntry {
  name: string;
  content: string;
  method?: 0 | 8;
}

function writeUint32LE(value: number): Uint8Array {
  const buf = new Uint8Array(4);
  new DataView(buf.buffer).setUint32(0, value, true);
  return buf;
}

function writeUint16LE(value: number): Uint8Array {
  const buf = new Uint8Array(2);
  new DataView(buf.buffer).setUint16(0, value, true);
  return buf;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, p) => sum + p.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const p of parts) {
    out.set(p, offset);
    offset += p.length;
  }
  return out;
}

function buildZip(entries: FixtureEntry[]): Uint8Array {
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const method = entry.method ?? 8;
    const nameBytes = encoder.encode(entry.name);
    const rawBytes = encoder.encode(entry.content);
    const dataBytes = method === 8 ? deflateRawSync(rawBytes) : rawBytes;

    const localHeader = concat([
      writeUint32LE(0x04034b50),
      writeUint16LE(20),
      writeUint16LE(0),
      writeUint16LE(method),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint32LE(0),
      writeUint32LE(dataBytes.length),
      writeUint32LE(rawBytes.length),
      writeUint16LE(nameBytes.length),
      writeUint16LE(0),
      nameBytes,
    ]);
    localParts.push(localHeader, dataBytes);

    const centralHeader = concat([
      writeUint32LE(0x02014b50),
      writeUint16LE(20),
      writeUint16LE(20),
      writeUint16LE(0),
      writeUint16LE(method),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint32LE(0),
      writeUint32LE(dataBytes.length),
      writeUint32LE(rawBytes.length),
      writeUint16LE(nameBytes.length),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint32LE(0),
      writeUint32LE(offset),
      nameBytes,
    ]);
    centralParts.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  }

  const centralDirectory = concat(centralParts);
  const localData = concat(localParts);
  const eocd = concat([
    writeUint32LE(0x06054b50),
    writeUint16LE(0),
    writeUint16LE(0),
    writeUint16LE(entries.length),
    writeUint16LE(entries.length),
    writeUint32LE(centralDirectory.length),
    writeUint32LE(localData.length),
    writeUint16LE(0),
  ]);

  return concat([localData, centralDirectory, eocd]);
}

describe("unzip", () => {
  it("extracts a stored (uncompressed) entry", async () => {
    const zip = buildZip([{ name: "hello.txt", content: "hello world", method: 0 }]);
    const parts = await unzip(zip);
    expect(new TextDecoder().decode(parts.get("hello.txt"))).toBe("hello world");
  });

  it("extracts a deflate-compressed entry", async () => {
    const content = "the quick brown fox jumps over the lazy dog ".repeat(50);
    const zip = buildZip([{ name: "big.txt", content, method: 8 }]);
    const parts = await unzip(zip);
    expect(new TextDecoder().decode(parts.get("big.txt"))).toBe(content);
  });

  it("extracts multiple entries, mixing stored and deflated", async () => {
    const zip = buildZip([
      { name: "a.xml", content: "<a/>", method: 0 },
      { name: "dir/b.xml", content: "<b/>", method: 8 },
    ]);
    const parts = await unzip(zip);
    expect(parts.size).toBe(2);
    expect(new TextDecoder().decode(parts.get("a.xml"))).toBe("<a/>");
    expect(new TextDecoder().decode(parts.get("dir/b.xml"))).toBe("<b/>");
  });

  it("throws ZipError for a file that is not a ZIP", async () => {
    const notZip = new TextEncoder().encode("this is definitely not a zip file");
    await expect(unzip(notZip)).rejects.toThrow(ZipError);
  });

  it("throws ZipError for a truncated/too-small buffer", async () => {
    await expect(unzip(new Uint8Array([1, 2, 3]))).rejects.toThrow(ZipError);
  });
});
