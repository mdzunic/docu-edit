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
  /** Overrides the uncompressed size written into both headers, so a
   * fixture can lie about how far it expands the way a bomb does. */
  declaredSize?: number;
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
    const declaredSize = entry.declaredSize ?? rawBytes.length;

    const localHeader = concat([
      writeUint32LE(0x04034b50),
      writeUint16LE(20),
      writeUint16LE(0),
      writeUint16LE(method),
      writeUint16LE(0),
      writeUint16LE(0),
      writeUint32LE(0),
      writeUint32LE(dataBytes.length),
      writeUint32LE(declaredSize),
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
      writeUint32LE(declaredSize),
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

  // The decompression budget. DEFLATE reaches ~1000:1 on repetitive input,
  // so the caller's limit on the compressed file bounds nothing here — the
  // budget is what stops a small archive from inflating into an OOM.
  describe("decompression budget", () => {
    const FLOOR = 64 * 1024 * 1024;
    // Under the floor, so a legitimate small package is never touched.
    const UNDER_BUDGET = "x".repeat(1024 * 1024);

    it("still opens a package that inflates well inside the budget", async () => {
      const zip = buildZip([{ name: "word/document.xml", content: UNDER_BUDGET }]);
      const parts = await unzip(zip);
      expect(parts.get("word/document.xml")?.length).toBe(UNDER_BUDGET.length);
    });

    it("rejects an archive whose central directory declares more than the budget", async () => {
      // What a bomb built by an ordinary ZIP tool looks like: the declared
      // size is honest, and enormous. Rejected before anything inflates.
      const zip = buildZip([
        { name: "bomb.bin", content: "x", declaredSize: FLOOR + 1 },
      ]);
      await expect(unzip(zip)).rejects.toThrow(/declares more decompressed content/);
    });

    it("sums declared sizes across entries rather than checking each alone", async () => {
      const half = Math.ceil(FLOOR / 2) + 1;
      const zip = buildZip([
        { name: "a.bin", content: "x", declaredSize: half },
        { name: "b.bin", content: "x", declaredSize: half },
      ]);
      await expect(unzip(zip)).rejects.toThrow(/declares more decompressed content/);
    });

    it("stops inflating an entry that under-declares its size", async () => {
      // The bomb that gets past the cheap check: the central directory says
      // one byte, the entry really expands past the floor. Only the running
      // total over what DecompressionStream actually produces catches this.
      // Deflating the fixture costs ~100ms and ~110MB peak — the price of
      // testing the path that matters rather than a proxy for it.
      const zip = buildZip([
        { name: "liar.bin", content: " ".repeat(FLOOR + 1024), declaredSize: 1 },
      ]);
      await expect(unzip(zip)).rejects.toThrow(/expands well beyond its declared size/);
    });

    it("names the offending entry so a bad package can be diagnosed", async () => {
      const zip = buildZip([
        { name: "word/media/image1.png", content: " ".repeat(FLOOR + 1024), declaredSize: 1 },
      ]);
      await expect(unzip(zip)).rejects.toThrow(/word\/media\/image1\.png/);
    });
  });
});
