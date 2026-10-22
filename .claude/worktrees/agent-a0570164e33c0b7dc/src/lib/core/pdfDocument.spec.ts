// The structural PDF reader's spec (SPEC-107): both xref flavours, /Prev
// chains, hybrid /XRefStm, object streams, page-tree inheritance and
// rotation, generations, stream filters and /Length fallbacks, the
// reconstruction path, encryption refusal, and one test per parsing bound
// that genuinely trips it. Fixtures are hand-written bytes, plus the
// checked-in e2e fixtures (spec107-pages.pdf is the app's own buildPdf
// output), so unit and e2e assertions cannot drift.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { deflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import {
  MAX_INFLATED_STREAM_BYTES,
  MAX_PDF_OBJECTS,
  MAX_RESOLVE_DEPTH,
  MAX_XREF_CHAIN,
  parsePdf,
  parsePdfDate,
  PdfEncryptedError,
  pdfTextOf,
  PdfParseError,
  resolve,
  streamData,
  type PdfValue,
} from "./pdfDocument";
import { buildEncryptedPdfBytes, buildXrefStreamPdfBytes } from "../media/fixtures/build";

const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../../e2e/fixtures/spec107-pages.pdf")));

function ascii(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    out[i] = text.charCodeAt(i) & 0xff;
  }
  return out;
}

/** Bytes as a string, one char per byte — exact, so `ascii(latin1(bytes))`
 * round-trips binary stream data (TextDecoder's "latin1" is windows-1252,
 * which would not). */
function latin1(bytes: Uint8Array): string {
  let text = "";
  for (let i = 0; i < bytes.length; i += 8192) {
    text += String.fromCharCode(...bytes.subarray(i, i + 8192));
  }
  return text;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let cursor = 0;
  for (const part of parts) {
    out.set(part, cursor);
    cursor += part.length;
  }
  return out;
}

/** A little incremental file writer that records byte offsets. */
class FileBuilder {
  private parts: Uint8Array[] = [];
  length = 0;

  push(data: string | Uint8Array): number {
    const bytes = typeof data === "string" ? ascii(data) : data;
    const at = this.length;
    this.parts.push(bytes);
    this.length += bytes.length;
    return at;
  }

  /** Writes `N G obj … endobj`, returning the header's offset. */
  object(num: number, body: string, gen = 0): number {
    return this.push(`${num} ${gen} obj\n${body}\nendobj\n`);
  }

  get bytes(): Uint8Array {
    return concat(this.parts);
  }
}

/** A classic xref section + trailer listing `entries` ([num, offset, gen]),
 * one subsection per entry. Returns the section's own start offset. */
function pushClassicXref(
  file: FileBuilder,
  entries: Array<[number, number, number?]>,
  size: number,
  trailerExtra: string,
): number {
  const at = file.length;
  let xref = "xref\n0 1\n0000000000 65535 f \n";
  for (const [num, offset, gen = 0] of entries.toSorted((a, b) => a[0] - b[0])) {
    xref += `${num} 1\n${String(offset).padStart(10, "0")} ${String(gen).padStart(5, "0")} n \n`;
  }
  xref += `trailer\n<< /Size ${size} ${trailerExtra} >>\n`;
  file.push(xref);
  return at;
}

/** The three objects of a one-page document (catalog 1, pages 2, page 3). */
function pushOnePage(file: FileBuilder, mediaBox = "[0 0 300 500]"): Array<[number, number]> {
  return [
    [1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")],
    [2, file.object(2, `<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox ${mediaBox} >>`)],
    [3, file.object(3, "<< /Type /Page /Parent 2 0 R >>")],
  ];
}

/** A one-page classic PDF with extra objects and trailer entries. */
function simplePdf(extraObjects: Array<[number, string]> = [], trailerExtra = ""): Uint8Array {
  const file = new FileBuilder();
  file.push("%PDF-1.4\n");
  const entries: Array<[number, number]> = pushOnePage(file);
  for (const [num, body] of extraObjects) {
    entries.push([num, file.object(num, body)]);
  }
  const xrefAt = pushClassicXref(file, entries, 100, `/Root 1 0 R ${trailerExtra}`);
  file.push(`startxref\n${xrefAt}\n%%EOF\n`);
  return file.bytes;
}

/** A `stream` object body with the given dictionary entries and data. */
function streamObject(file: FileBuilder, num: number, dict: string, data: Uint8Array, length = String(data.length)): number {
  const at = file.push(`${num} 0 obj\n<< ${dict} /Length ${length} >>\nstream\n`);
  file.push(data);
  file.push("\nendstream\nendobj\n");
  return at;
}

/** Replaces the startxref offset with a nonsense one, so only
 * reconstruction can read the file. */
function breakStartxref(bytes: Uint8Array): Uint8Array {
  return ascii(latin1(bytes).replace(/startxref\s+\d+/, "startxref\n2"));
}

/** PNG-predicts `rows` (1 byte per pixel) with the given per-row filter
 * types — the encoder side of what the reader must undo. */
function pngPredict(rows: number[][], types: number[]): Uint8Array {
  const out: number[] = [];
  let prior = rows[0].map(() => 0);
  rows.forEach((row, r) => {
    const type = types[r];
    out.push(type);
    row.forEach((raw, i) => {
      const left = i > 0 ? row[i - 1] : 0;
      const up = prior[i];
      const upLeft = i > 0 ? prior[i - 1] : 0;
      const p = left + up - upLeft;
      const paeth =
        Math.abs(p - left) <= Math.abs(p - up) && Math.abs(p - left) <= Math.abs(p - upLeft)
          ? left
          : Math.abs(p - up) <= Math.abs(p - upLeft) ? up : upLeft;
      const predicted = [0, left, up, Math.floor((left + up) / 2), paeth][type];
      out.push((raw - predicted) & 0xff);
    });
    prior = row;
  });
  return new Uint8Array(out);
}

async function rejection(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => null,
    (thrown: unknown) => thrown,
  );
}

describe("the checked-in fixtures", () => {
  it("parses spec107-pages.pdf — buildPdf's output: 3 pages, mixed sizes, a rotated page, /Info", async () => {
    const doc = await parsePdf(PAGES_PDF);
    expect(doc.version).toBe("1.7");
    expect(doc.reconstructed).toBe(false);
    expect(doc.linearised).toBe(false);
    expect(doc.pages.map((page) => [page.widthPt, page.heightPt, page.rotate])).toEqual([
      [612, 792, 0],
      [595.28, 841.89, 90],
      [720, 540, 0],
    ]);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 pages");
    expect(pdfTextOf(doc.info?.get("Author"))).toBe("docu-edit factory");
    expect(pdfTextOf(doc.info?.get("Producer"))).toBe("docu-edit");
    expect(parsePdfDate(pdfTextOf(doc.info?.get("CreationDate")))).toBe("2026-08-30 12:00");
    expect(doc.bytes).toBe(PAGES_PDF);
  });

  it("parses spec107-xrefstream.pdf: an xref STREAM with a PNG predictor, catalog in an object stream", async () => {
    const doc = await parsePdf(await buildXrefStreamPdfBytes());
    expect(doc.version).toBe("1.5");
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(1);
    expect(doc.pages[0].widthPt).toBe(400);
    expect(doc.pages[0].rotate).toBe(180);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 xref stream");
    // The catalog (object 1) came out of object stream 4.
    expect(doc.objects.get(1)?.t).toBe("dict");
  });

  it("refuses spec107-encrypted.pdf with PdfEncryptedError", async () => {
    const error = await rejection(parsePdf(buildEncryptedPdfBytes()));
    expect(error).toBeInstanceOf(PdfEncryptedError);
    expect(error).toBeInstanceOf(PdfParseError);
    expect((error as Error).message).toBe("This PDF is password-protected");
  });
});

describe("classic xref and the page tree", () => {
  it("walks the page tree with attribute inheritance, CropBox clipping and rotation", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = [
      [1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")],
      // The root carries MediaBox, Rotate and Resources for everything below.
      [
        2,
        file.object(
          2,
          "<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 3 /MediaBox [0 0 612 792] /Rotate 90 /Resources << /ProcSet [/PDF] >> >>",
        ),
      ],
      // An intermediate node adds a CropBox.
      [3, file.object(3, "<< /Type /Pages /Parent 2 0 R /Kids [4 0 R 5 0 R] /Count 2 /CropBox [36 36 576 756] >>")],
      // Inherits everything: the effective box is the inherited CropBox.
      [4, file.object(4, "<< /Type /Page /Parent 3 0 R >>")],
      // Overrides: its own MediaBox and Rotate; a CropBox larger than the
      // page is clipped to the MediaBox.
      [5, file.object(5, "<< /Type /Page /Parent 3 0 R /MediaBox [0 0 595 842] /CropBox [0 0 1000 1000] /Rotate 0 >>")],
      // A reversed, offset MediaBox normalises; -90 is 270.
      [6, file.object(6, "<< /Type /Page /Parent 2 0 R /MediaBox [110 220 10 20] /Rotate -90 >>")],
    ];
    const xrefAt = pushClassicXref(file, entries, 7, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages.map((page) => page.ref.num)).toEqual([4, 5, 6]);

    const [inherits, overrides, reversed] = doc.pages;
    expect([inherits.widthPt, inherits.heightPt, inherits.rotate]).toEqual([540, 720, 90]);
    expect(inherits.mediaBoxPt).toEqual([0, 0, 612, 792]);
    expect(inherits.cropBoxPt).toEqual([36, 36, 576, 756]);
    expect(inherits.inherited.get("Resources")?.t).toBe("dict");
    expect(inherits.inherited.get("Rotate")).toEqual({ t: "num", v: 90 });

    expect([overrides.widthPt, overrides.heightPt, overrides.rotate]).toEqual([595, 842, 0]);
    expect(overrides.cropBoxPt).toEqual([0, 0, 595, 842]);

    expect([reversed.widthPt, reversed.heightPt, reversed.rotate]).toEqual([100, 200, 270]);
    expect(reversed.mediaBoxPt).toEqual([10, 20, 110, 220]);
    expect(reversed.cropBoxPt).toBeUndefined();
  });

  it("follows a /Prev chain, newest section winning", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = pushOnePage(file);
    entries.push([4, file.object(4, "<< /Title (old) >>")]);
    const firstXref = pushClassicXref(file, entries, 5, "/Root 1 0 R /Info 4 0 R");
    file.push(`startxref\n${firstXref}\n%%EOF\n`);
    // Incremental update: a NEW object 4 and an xref with /Prev.
    const secondXref = pushClassicXref(file, [[4, file.object(4, "<< /Title (new) >>")]], 5, `/Root 1 0 R /Info 4 0 R /Prev ${firstXref}`);
    file.push(`startxref\n${secondXref}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("new");
    expect(doc.pages).toHaveLength(1);
  });

  it("reads a hybrid file: a classic table plus /XRefStm for object-stream members", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.5\n");
    const offsets = new Map<number, number>();
    offsets.set(2, file.object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 111 222] >>"));
    offsets.set(3, file.object(3, "<< /Type /Page /Parent 2 0 R >>"));
    // The catalog lives in object stream 4.
    const header = "1 0 ";
    const objStm = new Uint8Array(deflateSync(ascii(`${header}<< /Type /Catalog /Pages 2 0 R >>\n`)));
    offsets.set(4, streamObject(file, 4, `/Type /ObjStm /N 1 /First ${header.length} /Filter /FlateDecode`, objStm));
    // The /XRefStm stream maps object 1 into stream 4 (type 2, index 0).
    const xrefStm = new Uint8Array(deflateSync(new Uint8Array([2, 4, 0])));
    const xrefStmAt = streamObject(file, 5, "/Type /XRef /Size 6 /W [1 1 1] /Index [1 1] /Filter /FlateDecode", xrefStm);
    // Classic table: object 1 marked FREE (the hybrid convention).
    const at = file.length;
    let xref = "xref\n0 2\n0000000000 65535 f \n0000000000 00000 f \n";
    for (const num of [2, 3, 4]) {
      xref += `${num} 1\n${String(offsets.get(num)).padStart(10, "0")} 00000 n \n`;
    }
    file.push(`${xref}trailer\n<< /Size 6 /Root 1 0 R /XRefStm ${xrefStmAt} >>\nstartxref\n${at}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(1);
    expect([doc.pages[0].widthPt, doc.pages[0].heightPt]).toEqual([111, 222]);
  });

  it("parses a file with bytes before its %PDF- header (offset recorded)", async () => {
    const junk = "x".repeat(100) + "\n";
    const file = new FileBuilder();
    file.push(junk);
    file.push("%PDF-1.4\n");
    const xrefAt = pushClassicXref(file, pushOnePage(file), 4, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.version).toBe("1.4");
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(1);
  });

  it("never interprets a content stream: garbage page content still parses", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = [
      [1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")],
      [2, file.object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 10 10] >>")],
      [3, file.object(3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>")],
    ];
    entries.push([4, streamObject(file, 4, "", ascii("((( not [ operators BT Tj <zz"))]);
    const xrefAt = pushClassicXref(file, entries, 5, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.pages).toHaveLength(1);
    expect(doc.objects.get(4)?.t).toBe("stream");
  });
});

describe("generations", () => {
  it("honours a non-zero generation when reading, keyed by object number alone", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number, number?]> = pushOnePage(file);
    entries.push([4, file.object(4, "<< /Title (gen two) >>", 2), 2]);
    const xrefAt = pushClassicXref(file, entries, 5, "/Root 1 0 R /Info 4 2 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("gen two");
    expect([...doc.objects.keys()].toSorted((a, b) => a - b)).toEqual([1, 2, 3, 4]);
  });

  it("takes the newest xref section's generation for a reused object number", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number, number?]> = pushOnePage(file);
    entries.push([4, file.object(4, "<< /Title (generation 0) >>")]);
    const firstXref = pushClassicXref(file, entries, 5, "/Root 1 0 R /Info 4 0 R");
    file.push(`startxref\n${firstXref}\n%%EOF\n`);
    // The update deletes 4/0 and reuses the number as 4/1.
    const reused = file.object(4, "<< /Title (generation 1) >>", 1);
    const secondXref = pushClassicXref(file, [[4, reused, 1]], 5, `/Root 1 0 R /Info 4 1 R /Prev ${firstXref}`);
    file.push(`startxref\n${secondXref}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("generation 1");
    expect(doc.objects.size).toBe(4);
  });

  it("does not read an object whose header disagrees with its xref entry's generation", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number, number?]> = pushOnePage(file);
    // The xref says 4/0; the header at that offset says 4/3.
    entries.push([4, file.object(4, "<< /Title (stale) >>", 3), 0]);
    const xrefAt = pushClassicXref(file, entries, 5, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(doc.objects.has(4)).toBe(false);
  });

  it("refuses a file that needs generation-keyed lookup, rather than half-reading it", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number, number?]> = pushOnePage(file);
    entries.push([4, file.object(4, "<< /Title (live) >>", 1), 1]);
    // Object 5 still points at generation 0 of object 4 — a deleted object
    // a number-keyed map would silently answer with generation 1.
    const referrer = file.object(5, "<< /Stale 4 0 R >>");
    entries.push([5, referrer]);
    const xrefAt = pushClassicXref(file, entries, 6, "/Root 1 0 R /Info 4 1 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const error = await rejection(parsePdf(file.bytes));
    expect(error).toBeInstanceOf(PdfParseError);
    expect((error as Error).message).toMatch(/generation 0.*generation 1.*generation-keyed lookup/);
    expect((error as PdfParseError).offset).toBe(referrer);
  });
});

describe("streams and filters", () => {
  /** Parses a one-page file whose object 4 is a stream with `dict`/`data`
   * (and `length`, when the /Length entry should lie), returning the doc
   * and that stream. */
  async function withStream(dict: string, data: Uint8Array, length?: string, extra: Array<[number, string]> = []) {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = pushOnePage(file);
    entries.push([4, streamObject(file, 4, dict, data, length)]);
    for (const [num, body] of extra) {
      entries.push([num, file.object(num, body)]);
    }
    const xrefAt = pushClassicXref(file, entries, 10, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    const stream = doc.objects.get(4);
    if (stream?.t !== "stream") {
      throw new Error("no stream");
    }
    return { doc, stream };
  }

  it("resolves an indirect /Length after the xref is built", async () => {
    const compressed = new Uint8Array(deflateSync(ascii("hello stream payload")));
    const { doc, stream } = await withStream("/Filter /FlateDecode", compressed, "5 0 R", [[5, String(compressed.length)]]);
    expect(stream.raw.length).toBe(compressed.length);
    expect(new TextDecoder().decode(await streamData(doc, stream))).toBe("hello stream payload");
  });

  it("bounds a stream by scanning to endstream when /Length is wrong", async () => {
    const { stream } = await withStream("", ascii("twenty-one bytes long"), "3");
    expect(latin1(stream.raw)).toBe("twenty-one bytes long");
  });

  it("bounds a stream by scanning to endstream when an indirect /Length is unresolvable", async () => {
    const { stream } = await withStream("", ascii("dangling length"), "99 0 R");
    expect(latin1(stream.raw)).toBe("dangling length");
  });

  it("bounds a stream by scanning when /Length points past the end of the file", async () => {
    const { stream } = await withStream("", ascii("short"), "999999");
    expect(latin1(stream.raw)).toBe("short");
  });

  it("keeps the image-compression filters raw — never decoded", async () => {
    const doc = await parsePdf(simplePdf());
    const raw = new Uint8Array([1, 2, 3, 0xff]);
    for (const filter of ["DCTDecode", "JPXDecode", "CCITTFaxDecode", "JBIG2Decode", "RunLengthDecode"]) {
      const stream = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: filter }]]), raw };
      expect(await streamData(doc, stream)).toEqual(raw);
    }
    // Decoding stops at the image filter in a chain: the hex layer comes
    // off, the JPEG bytes stay as they are.
    const chained = {
      dict: new Map<string, PdfValue>([
        ["Filter", { t: "array", v: [{ t: "name", v: "ASCIIHexDecode" }, { t: "name", v: "DCTDecode" }] }],
      ]),
      raw: ascii("FFD8FF>"),
    };
    expect([...(await streamData(doc, chained))]).toEqual([0xff, 0xd8, 0xff]);
  });

  it("decodes ASCIIHexDecode and ASCII85Decode", async () => {
    const doc = await parsePdf(simplePdf());
    const hex = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "ASCIIHexDecode" }]]), raw: ascii("48 65 6C 6C 6F>") };
    expect(latin1(await streamData(doc, hex))).toBe("Hello");
    const a85 = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "ASCII85Decode" }]]), raw: ascii("87cURD_*#4DfTZ)+T~>") };
    expect(latin1(await streamData(doc, a85))).toBe("Hello, World!");
    // `z` is four zero bytes; a short final group decodes to fewer bytes.
    const zeros = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "ASCII85Decode" }]]), raw: ascii("z@:E^~>") };
    expect([...(await streamData(doc, zeros))]).toEqual([0, 0, 0, 0, 0x61, 0x62, 0x63]);
  });

  it("decodes FlateDecode with every PNG predictor (10–15): None, Sub, Up, Average, Paeth", async () => {
    const doc = await parsePdf(simplePdf());
    const rows = [
      [10, 200, 30, 250],
      [11, 190, 35, 5],
      [100, 101, 102, 103],
      [7, 77, 177, 255],
      [0, 128, 64, 32],
    ];
    for (const predictor of [10, 11, 12, 13, 14, 15]) {
      const stream = {
        dict: new Map<string, PdfValue>([
          ["Filter", { t: "name", v: "FlateDecode" }],
          ["DecodeParms", { t: "dict", v: new Map<string, PdfValue>([["Predictor", { t: "num", v: predictor }], ["Columns", { t: "num", v: 4 }]]) }],
        ]),
        raw: new Uint8Array(deflateSync(pngPredict(rows, [0, 1, 2, 3, 4]))),
      };
      expect([...(await streamData(doc, stream))]).toEqual(rows.flat());
    }
  });

  it("decodes FlateDecode with TIFF Predictor 2 at 8, 16 and 4 bits per component", async () => {
    const doc = await parsePdf(simplePdf());
    const decode = async (data: number[], parms: Array<[string, number]>) =>
      [
        ...(await streamData(doc, {
          dict: new Map<string, PdfValue>([
            ["Filter", { t: "name", v: "FlateDecode" }],
            ["DecodeParms", { t: "dict", v: new Map(parms.map(([key, v]) => [key, { t: "num", v } as PdfValue])) }],
          ]),
          raw: new Uint8Array(deflateSync(new Uint8Array(data))),
        })),
      ];
    // 8-bit RGB, 2 pixels per row: each component adds its left neighbour's.
    expect(
      await decode([10, 20, 30, 1, 2, 3, 200, 0, 0, 100, 0, 0], [["Predictor", 2], ["Columns", 2], ["Colors", 3]]),
    ).toEqual([10, 20, 30, 11, 22, 33, 200, 0, 0, 44, 0, 0]);
    // 16-bit gray, 3 samples: 0x0102, +0x00FF, +0x0001 with carries.
    expect(
      await decode([0x01, 0x02, 0x00, 0xff, 0x00, 0x01], [["Predictor", 2], ["Columns", 3], ["BitsPerComponent", 16]]),
    ).toEqual([0x01, 0x02, 0x02, 0x01, 0x02, 0x02]);
    // 4-bit gray, 4 samples in two bytes: 1, +2, +3, +15 (mod 16).
    expect(await decode([0x12, 0x3f], [["Predictor", 2], ["Columns", 4], ["BitsPerComponent", 4]])).toEqual([0x13, 0x65]);
  });

  it("refuses an unknown filter and impossible predictor parameters", async () => {
    const doc = await parsePdf(simplePdf());
    const raw = new Uint8Array([1, 2, 3]);
    const unknown = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "Bogus" }]]), raw };
    await expect(streamData(doc, unknown)).rejects.toBeInstanceOf(PdfParseError);
    const badParms = {
      dict: new Map<string, PdfValue>([
        ["Filter", { t: "name", v: "FlateDecode" }],
        ["DecodeParms", { t: "dict", v: new Map<string, PdfValue>([["Predictor", { t: "num", v: 12 }], ["Columns", { t: "num", v: 1e12 }]]) }],
      ]),
      raw: new Uint8Array(deflateSync(raw)),
    };
    await expect(streamData(doc, badParms)).rejects.toThrow(/DecodeParms/);
  });

  it("reports a stream's own byte offset when its data cannot be decoded", async () => {
    const { doc, stream } = await withStream("/Filter /Bogus", ascii("data"));
    const error = (await rejection(streamData(doc, stream))) as PdfParseError;
    expect(error).toBeInstanceOf(PdfParseError);
    expect(latin1(doc.bytes.subarray(error.offset, error.offset + 4))).toBe("data");
  });
});

describe("encryption", () => {
  it("refuses an encrypted file even when its xref is broken (reconstruction path)", async () => {
    await expect(parsePdf(breakStartxref(buildEncryptedPdfBytes()))).rejects.toBeInstanceOf(PdfEncryptedError);
  });

  it("refuses an xref-stream file whose stream dictionary names /Encrypt, on either path", async () => {
    const encrypted = ascii(latin1(await buildXrefStreamPdfBytes()).replace("/Type /XRef", "/Type /XRef /Encrypt 9 0 R"));
    await expect(parsePdf(encrypted)).rejects.toBeInstanceOf(PdfEncryptedError);
    await expect(parsePdf(breakStartxref(encrypted))).rejects.toBeInstanceOf(PdfEncryptedError);
  });
});

describe("reconstruction", () => {
  it("rebuilds a file whose startxref points at garbage (reconstructed flag)", async () => {
    const doc = await parsePdf(breakStartxref(PAGES_PDF));
    expect(doc.reconstructed).toBe(true);
    expect(doc.pages).toHaveLength(3);
    expect(doc.pages[1].rotate).toBe(90);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 pages");
  });

  it("rebuilds an xref-stream file from the XRef stream's own dictionary", async () => {
    const doc = await parsePdf(breakStartxref(await buildXrefStreamPdfBytes()));
    expect(doc.reconstructed).toBe(true);
    expect(doc.pages).toHaveLength(1);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 xref stream");
  });

  it("finds the /Type /Catalog when there is no xref or trailer at all", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    pushOnePage(file, "[0 0 50 60]");
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(true);
    expect([doc.pages[0].widthPt, doc.pages[0].heightPt]).toEqual([50, 60]);
  });

  it("takes the LAST definition of each object number", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    file.object(1, "<< /Type /Catalog /Pages 2 0 R >>");
    file.object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>");
    file.object(3, "<< /Type /Page /Parent 2 0 R >>");
    file.object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 222 333] >>");
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(true);
    expect([doc.pages[0].widthPt, doc.pages[0].heightPt]).toEqual([222, 333]);
  });

  it("throws PdfParseError with the byte offset when even reconstruction cannot read the file", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    file.object(1, "<< /Type /Catalog /Pages 2 0 R >>");
    const broken = file.object(2, "42");
    const error = (await rejection(parsePdf(file.bytes))) as PdfParseError;
    expect(error).toBeInstanceOf(PdfParseError);
    expect(error.message).toMatch(/page tree node 2/);
    expect(error.offset).toBe(broken);
  });

  it("throws PdfParseError with an offset for a file that is not a PDF at all", async () => {
    const error = (await rejection(parsePdf(ascii("this is not a pdf at all")))) as PdfParseError;
    expect(error).toBeInstanceOf(PdfParseError);
    expect(error.offset).toBe(0);
  });
});

describe("bounds — each one tripped", () => {
  it("MAX_PDF_OBJECTS: an xref section claiming more entries is unusable (exactly the bound still reads)", async () => {
    expect(MAX_PDF_OBJECTS).toBe(500_000);
    const fixture = latin1(await buildXrefStreamPdfBytes());
    const withIndex = (count: number) => ascii(fixture.replace("/Index [0 7]", `/Index [0 ${count}]`));
    // The stream's rows run out after 7 entries either way; only the bound
    // makes the claim itself fatal to the xref.
    expect((await parsePdf(withIndex(MAX_PDF_OBJECTS))).reconstructed).toBe(false);
    expect((await parsePdf(withIndex(MAX_PDF_OBJECTS + 1))).reconstructed).toBe(true);
  });

  it("MAX_PDF_OBJECTS: reconstruction refuses a buffer of more object headers", async () => {
    const headers: string[] = ["%PDF-1.4\n"];
    for (let num = 1; num <= MAX_PDF_OBJECTS + 1; num += 1) {
      headers.push(`${num} 0 obj `);
    }
    await expect(parsePdf(ascii(headers.join("")))).rejects.toThrow(/too many objects/);
  });

  it("MAX_XREF_CHAIN: a /Prev chain one section longer is unusable (exactly the bound still reads)", async () => {
    expect(MAX_XREF_CHAIN).toBe(64);
    const chained = (sections: number) => {
      const file = new FileBuilder();
      file.push("%PDF-1.4\n");
      const entries = pushOnePage(file);
      let prev = -1;
      for (let i = 0; i < sections; i += 1) {
        prev = pushClassicXref(file, entries, 4, `/Root 1 0 R${prev >= 0 ? ` /Prev ${prev}` : ""}`);
      }
      file.push(`startxref\n${prev}\n%%EOF\n`);
      return file.bytes;
    };
    expect((await parsePdf(chained(MAX_XREF_CHAIN))).reconstructed).toBe(false);
    expect((await parsePdf(chained(MAX_XREF_CHAIN + 1))).reconstructed).toBe(true);
  });

  it("a cyclic /Prev chain terminates (the xref is unusable; reconstruction reads the file)", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries = pushOnePage(file, "[0 0 100 100]");
    const firstAt = file.length;
    const section = (prev: number) =>
      "xref\n0 1\n0000000000 65535 f \n" +
      entries.map(([num, offset]) => `${num} 1\n${String(offset).padStart(10, "0")} 00000 n \n`).join("") +
      `trailer\n<< /Size 4 /Root 1 0 R /Prev ${String(prev).padStart(10, "0")} >>\n`;
    const secondAt = firstAt + section(0).length;
    file.push(section(secondAt));
    file.push(section(firstAt));
    file.push(`startxref\n${firstAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(true);
    expect(doc.pages).toHaveLength(1);
  });

  it("MAX_RESOLVE_DEPTH: a reference chain longer than the cap fails instead of resolving", async () => {
    expect(MAX_RESOLVE_DEPTH).toBe(64);
    const chain: Array<[number, string]> = [];
    for (let i = 0; i < MAX_RESOLVE_DEPTH + 4; i += 1) {
      chain.push([10 + i, `${11 + i} 0 R`]);
    }
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = [
      [1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")],
      [2, file.object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>")],
      [3, file.object(3, "<< /Type /Page /Parent 2 0 R /Rotate 10 0 R >>")],
    ];
    for (const [num, body] of chain) {
      entries.push([num, file.object(num, body)]);
    }
    const xrefAt = pushClassicXref(file, entries, 200, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    await expect(parsePdf(file.bytes)).rejects.toThrow(/reference chain too deep/);
  });

  it("MAX_RESOLVE_DEPTH: resolve() terminates on a reference cycle", async () => {
    const doc = await parsePdf(simplePdf([[10, "11 0 R"], [11, "10 0 R"]]));
    const error = (() => {
      try {
        resolve(doc, { t: "ref", num: 10, gen: 0 });
        return null;
      } catch (thrown) {
        return thrown;
      }
    })();
    expect(error).toBeInstanceOf(PdfParseError);
    expect((error as Error).message).toMatch(/too deep/);
  });

  it("the page-tree visited set: a self-referential page tree fails instead of hanging", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = [
      [1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")],
      [2, file.object(2, "<< /Type /Pages /Kids [2 0 R] /Count 1 >>")],
    ];
    const xrefAt = pushClassicXref(file, entries, 3, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    await expect(parsePdf(file.bytes)).rejects.toThrow(/cyclic page tree/);
  });

  it("the page-tree walk is depth-capped too", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const entries: Array<[number, number]> = [[1, file.object(1, "<< /Type /Catalog /Pages 2 0 R >>")]];
    const depth = MAX_RESOLVE_DEPTH + 2;
    for (let i = 0; i < depth; i += 1) {
      const num = 2 + i;
      entries.push([num, file.object(num, `<< /Type /Pages /Kids [${num + 1} 0 R] /Count 1 /MediaBox [0 0 9 9] >>`)]);
    }
    entries.push([2 + depth, file.object(2 + depth, "<< /Type /Page >>")]);
    const xrefAt = pushClassicXref(file, entries, depth + 4, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    await expect(parsePdf(file.bytes)).rejects.toThrow(/page tree too deep/);
  });

  it("MAX_INFLATED_STREAM_BYTES: a decompression bomb stops at the cap", async () => {
    expect(MAX_INFLATED_STREAM_BYTES).toBe(256 * 1024 * 1024);
    const doc = await parsePdf(simplePdf());
    const bomb = {
      dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "FlateDecode" }]]),
      raw: new Uint8Array(deflateSync(new Uint8Array(64 * 1024))),
    };
    // The real cap needs a 256 MB inflate; the cap is a parameter so the
    // same code path trips at test scale.
    await expect(streamData(doc, bomb, 1024)).rejects.toThrow(/size limit/);
    expect((await streamData(doc, bomb, 64 * 1024)).length).toBe(64 * 1024);
  });
});

describe("values and helpers", () => {
  it("parses names with #-escapes, nested strings, hex strings and refs", async () => {
    const doc = await parsePdf(
      simplePdf([[4, "<< /A#42 (par(en)s) /Hex <48690> /Ref 3 0 R /Neg -1.5 /Arr [1 2 3] >>"]], "/Info 4 0 R"),
    );
    const info = doc.info;
    expect(info?.has("AB")).toBe(true);
    expect(pdfTextOf(info?.get("AB"))).toBe("par(en)s");
    const hex = info?.get("Hex");
    expect(hex?.t === "str" && [...hex.v]).toEqual([0x48, 0x69, 0x00]);
    const ref = info?.get("Ref");
    expect(ref?.t).toBe("ref");
    if (ref) {
      expect(resolve(doc, ref).t).toBe("dict");
    }
    const neg = info?.get("Neg");
    expect(neg?.t === "num" && neg.v).toBe(-1.5);
  });

  it("decodes UTF-16BE strings and PDF dates", () => {
    const utf16: PdfValue = { t: "str", v: new Uint8Array([0xfe, 0xff, 0x00, 0x48, 0x00, 0x69]) };
    expect(pdfTextOf(utf16)).toBe("Hi");
    expect(parsePdfDate("D:20260830120500Z")).toBe("2026-08-30 12:05");
    expect(parsePdfDate("D:2026")).toBe("2026");
    expect(parsePdfDate("yesterday")).toBeNull();
  });

  it("reports the linearised flag from the leading dictionary", async () => {
    const linearised = ascii(latin1(PAGES_PDF).replace("%PDF-1.7\n", "%PDF-1.7\n7 0 obj\n<< /Linearized 1 >>\nendobj\n"));
    // Offsets shifted: the structured path fails and reconstruction (which
    // ignores offsets) takes over — the flag must still be read.
    expect((await parsePdf(linearised)).linearised).toBe(true);
    expect((await parsePdf(PAGES_PDF)).linearised).toBe(false);
  });
});
