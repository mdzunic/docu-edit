// The structural PDF reader's spec (SPEC-107): both xref flavours, /Prev
// chains, hybrid /XRefStm, object streams, page-tree inheritance and
// rotation, the reconstruction path, encryption refusal, and one test per
// parsing bound. Fixtures are hand-written bytes (built by the same helpers
// the e2e fixtures use, so unit and e2e assertions cannot drift).
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
import {
  buildEncryptedPdfBytes,
  buildPagesPdfBytes,
  buildXrefStreamPdfBytes,
} from "../media/fixtures/build";

function ascii(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    out[i] = text.charCodeAt(i) & 0xff;
  }
  return out;
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

  get bytes(): Uint8Array {
    return concat(this.parts);
  }
}

/** A classic xref section + trailer for the given object offsets. Returns
 * the section's own start offset. */
function pushClassicXref(
  file: FileBuilder,
  offsets: Map<number, number>,
  size: number,
  trailerExtra: string,
): number {
  const at = file.length;
  let xref = "xref\n";
  const nums = [...offsets.keys()].toSorted((a, b) => a - b);
  xref += "0 1\n0000000000 65535 f \n";
  for (const num of nums) {
    xref += `${num} 1\n${String(offsets.get(num)).padStart(10, "0")} 00000 n \n`;
  }
  xref += `trailer\n<< /Size ${size} ${trailerExtra} >>\n`;
  file.push(xref);
  return at;
}

/** A one-page classic PDF whose trailer/extra objects the tests vary. */
function simplePdf(extraObjects: Array<[number, string]> = [], trailerExtra = ""): Uint8Array {
  const file = new FileBuilder();
  file.push("%PDF-1.4\n");
  const offsets = new Map<number, number>();
  offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
  offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 300 500] >>\nendobj\n"));
  offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n"));
  for (const [num, body] of extraObjects) {
    offsets.set(num, file.push(`${num} 0 obj\n${body}\nendobj\n`));
  }
  const xrefAt = pushClassicXref(file, offsets, 10, `/Root 1 0 R ${trailerExtra}`);
  file.push(`startxref\n${xrefAt}\n%%EOF\n`);
  return file.bytes;
}

describe("classic xref", () => {
  it("parses the 3-page mixed-size fixture with inheritance and rotation", async () => {
    const doc = await parsePdf(buildPagesPdfBytes());
    expect(doc.version).toBe("1.7");
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(3);
    // Page 1 inherits US Letter from the page-tree root.
    expect(doc.pages[0].widthPt).toBe(612);
    expect(doc.pages[0].heightPt).toBe(792);
    expect(doc.pages[0].rotate).toBe(0);
    // Page 2 overrides with A4 and rotates.
    expect(doc.pages[1].widthPt).toBe(595);
    expect(doc.pages[1].heightPt).toBe(842);
    expect(doc.pages[1].rotate).toBe(90);
    // Page 3's effective box is its CropBox.
    expect(doc.pages[2].widthPt).toBe(700);
    expect(doc.pages[2].heightPt).toBe(520);
    expect(doc.pages[2].mediaBoxPt).toEqual([0, 0, 720, 540]);
    expect(doc.pages[2].cropBoxPt).toEqual([10, 10, 710, 530]);
    // /Info comes through.
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 pages");
    expect(pdfTextOf(doc.info?.get("Producer"))).toBe("docu-edit");
  });

  it("resolves inherited attributes onto each page", async () => {
    const doc = await parsePdf(buildPagesPdfBytes());
    const inherited = doc.pages[0].inherited.get("MediaBox");
    expect(inherited?.t).toBe("array");
  });

  it("follows a /Prev chain, newest section winning", async () => {
    // Original file: catalog + pages + page + info(A). Update: info(B).
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 200 200] >>\nendobj\n"));
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n"));
    offsets.set(4, file.push("4 0 obj\n<< /Title (old) >>\nendobj\n"));
    const firstXref = pushClassicXref(file, offsets, 5, "/Root 1 0 R /Info 4 0 R");
    file.push(`startxref\n${firstXref}\n%%EOF\n`);
    // Incremental update: a NEW object 4 and an xref with /Prev.
    const updated = new Map<number, number>();
    updated.set(4, file.push("4 0 obj\n<< /Title (new) >>\nendobj\n"));
    const secondXref = pushClassicXref(
      file,
      updated,
      5,
      `/Root 1 0 R /Info 4 0 R /Prev ${firstXref}`,
    );
    file.push(`startxref\n${secondXref}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("new");
    expect(doc.pages).toHaveLength(1);
  });
});

describe("xref streams and object streams", () => {
  it("parses the xref-stream fixture (PNG predictor, catalog in an ObjStm)", async () => {
    const doc = await parsePdf(await buildXrefStreamPdfBytes());
    expect(doc.version).toBe("1.5");
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(1);
    expect(doc.pages[0].widthPt).toBe(400);
    expect(doc.pages[0].rotate).toBe(180);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 xref stream");
  });

  it("reads a hybrid file: classic table plus /XRefStm for object-stream members", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.5\n");
    const offsets = new Map<number, number>();
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 111 222] >>\nendobj\n"));
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n"));
    // The catalog lives in object stream 4.
    const header = "1 0 ";
    const body = "<< /Type /Catalog /Pages 2 0 R >>\n";
    const objStmData = deflateSync(ascii(header + body));
    offsets.set(
      4,
      file.push(
        `4 0 obj\n<< /Type /ObjStm /N 1 /First ${header.length} /Filter /FlateDecode /Length ${objStmData.length} >>\nstream\n`,
      ),
    );
    file.push(new Uint8Array(objStmData));
    file.push("\nendstream\nendobj\n");
    // The /XRefStm stream mapping object 1 into stream 4.
    const rows = new Uint8Array([2, 4, 0]); // type 2, container 4, index 0 (W [1 1 1])
    const xrefStmData = deflateSync(rows);
    const xrefStmAt = file.push(
      `5 0 obj\n<< /Type /XRef /Size 6 /W [1 1 1] /Index [1 1] /Root 1 0 R /Filter /FlateDecode /Length ${xrefStmData.length} >>\nstream\n`,
    );
    file.push(new Uint8Array(xrefStmData));
    file.push("\nendstream\nendobj\n");
    // Classic table: object 1 marked FREE (the hybrid convention), the rest live.
    const at = file.length;
    let xref = "xref\n0 2\n0000000000 65535 f \n0000000000 00000 f \n";
    for (const num of [2, 3, 4]) {
      xref += `${num} 1\n${String(offsets.get(num)).padStart(10, "0")} 00000 n \n`;
    }
    xref += `trailer\n<< /Size 6 /Root 1 0 R /XRefStm ${xrefStmAt} >>\n`;
    file.push(xref);
    file.push(`startxref\n${at}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(false);
    expect(doc.pages).toHaveLength(1);
    expect(doc.pages[0].widthPt).toBe(111);
    expect(doc.pages[0].heightPt).toBe(222);
  });
});

describe("streams", () => {
  it("resolves an indirect /Length and decodes FlateDecode data", async () => {
    const payload = ascii("hello stream payload");
    const compressed = deflateSync(payload);
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>\nendobj\n"));
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>\nendobj\n"));
    offsets.set(4, file.push(`4 0 obj\n<< /Length 5 0 R /Filter /FlateDecode >>\nstream\n`));
    file.push(new Uint8Array(compressed));
    file.push("\nendstream\nendobj\n");
    offsets.set(5, file.push(`5 0 obj\n${compressed.length}\nendobj\n`));
    const xrefAt = pushClassicXref(file, offsets, 6, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    const contents = doc.objects.get(4);
    expect(contents?.t).toBe("stream");
    if (contents?.t === "stream") {
      expect(contents.raw.length).toBe(compressed.length);
      const data = await streamData(doc, contents);
      expect(new TextDecoder().decode(data)).toBe("hello stream payload");
    }
  });

  it("keeps image-compression filters raw and rejects unknown filters", async () => {
    const doc = await parsePdf(buildPagesPdfBytes());
    const raw = new Uint8Array([1, 2, 3]);
    const dct = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "DCTDecode" }]]), raw };
    expect(await streamData(doc, dct)).toEqual(raw);
    const unknown = { dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "Bogus" }]]), raw };
    await expect(streamData(doc, unknown)).rejects.toBeInstanceOf(PdfParseError);
  });

  it("decodes ASCIIHexDecode", async () => {
    const doc = await parsePdf(buildPagesPdfBytes());
    const stream = {
      dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "ASCIIHexDecode" }]]),
      raw: ascii("48 65 6C 6C 6F>"),
    };
    expect(new TextDecoder().decode(await streamData(doc, stream))).toBe("Hello");
  });

  it("bounds the inflated size of a stream (decompression bomb)", async () => {
    const doc = await parsePdf(buildPagesPdfBytes());
    const bomb = deflateSync(new Uint8Array(64 * 1024));
    const stream = {
      dict: new Map<string, PdfValue>([["Filter", { t: "name", v: "FlateDecode" }]]),
      raw: new Uint8Array(bomb),
    };
    // The real cap would need a 256 MB fixture; the cap is a parameter so
    // the same code path trips at test scale.
    await expect(streamData(doc, stream, 1024)).rejects.toBeInstanceOf(PdfParseError);
    expect(MAX_INFLATED_STREAM_BYTES).toBe(256 * 1024 * 1024);
  });
});

describe("refusal and bounds", () => {
  it("throws PdfEncryptedError for an encrypted file", async () => {
    await expect(parsePdf(buildEncryptedPdfBytes())).rejects.toBeInstanceOf(PdfEncryptedError);
    await expect(parsePdf(buildEncryptedPdfBytes())).rejects.toThrow("password-protected");
  });

  it("throws PdfEncryptedError even when the xref is broken (reconstruction path)", async () => {
    const bytes = buildEncryptedPdfBytes();
    // Corrupt the startxref offset so the structured path fails first.
    const text = new TextDecoder("latin1").decode(bytes);
    const corrupted = ascii(text.replace(/startxref\n\d+/, "startxref\n999999"));
    await expect(parsePdf(corrupted)).rejects.toBeInstanceOf(PdfEncryptedError);
  });

  it("throws PdfParseError with a byte offset for a non-PDF file", async () => {
    const error = await parsePdf(ascii("this is not a pdf at all")).then(
      () => null,
      (thrown: unknown) => thrown,
    );
    expect(error).toBeInstanceOf(PdfParseError);
    expect(typeof (error as PdfParseError).offset).toBe("number");
  });

  it("rebuilds a file whose startxref points at garbage (reconstructed flag)", async () => {
    const bytes = buildPagesPdfBytes();
    const text = new TextDecoder("latin1").decode(bytes);
    const corrupted = ascii(text.replace(/startxref\n\d+/, "startxref\n2"));
    const doc = await parsePdf(corrupted);
    expect(doc.reconstructed).toBe(true);
    expect(doc.pages).toHaveLength(3);
    expect(doc.pages[1].rotate).toBe(90);
    expect(pdfTextOf(doc.info?.get("Title"))).toBe("SPEC-107 pages");
  });

  it("terminates on a cyclic /Prev chain (via reconstruction, not a hang)", async () => {
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>\nendobj\n"));
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n"));
    // Two xref sections whose /Prev entries point at each other.
    const firstAt = file.length;
    const xrefBody = (prev: number): string =>
      `xref\n0 1\n0000000000 65535 f \n1 3\n` +
      `${String(offsets.get(1)).padStart(10, "0")} 00000 n \n` +
      `${String(offsets.get(2)).padStart(10, "0")} 00000 n \n` +
      `${String(offsets.get(3)).padStart(10, "0")} 00000 n \n` +
      `trailer\n<< /Size 4 /Root 1 0 R /Prev ${prev} >>\n`;
    // First section's /Prev points at the second, which points back.
    const firstText = xrefBody(0);
    const secondAt = firstAt + firstText.length;
    file.push(xrefBody(secondAt));
    file.push(xrefBody(firstAt));
    file.push(`startxref\n${firstAt}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(true);
    expect(doc.pages).toHaveLength(1);
    expect(MAX_XREF_CHAIN).toBe(64);
  });

  it("caps the number of xref entries at MAX_PDF_OBJECTS", async () => {
    // A subsection header claiming an absurd count trips the bound in the
    // structured path; reconstruction then salvages the real objects.
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>\nendobj\n"));
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n"));
    const at = file.length;
    file.push(`xref\n0 ${MAX_PDF_OBJECTS + 1}\n0000000000 65535 f \n`);
    file.push(`trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`);
    const doc = await parsePdf(file.bytes);
    expect(doc.reconstructed).toBe(true);
    expect(MAX_PDF_OBJECTS).toBe(500_000);
  });

  it("caps reference-resolution depth", async () => {
    // /Rotate is a reference chain longer than MAX_RESOLVE_DEPTH: both the
    // structured and reconstruction paths hit the cap, so the parse fails
    // rather than spinning.
    const chain: Array<[number, string]> = [];
    for (let i = 0; i < MAX_RESOLVE_DEPTH + 4; i += 1) {
      chain.push([10 + i, `${11 + i} 0 R`]);
    }
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(
      2,
      file.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>\nendobj\n"),
    );
    offsets.set(3, file.push("3 0 obj\n<< /Type /Page /Parent 2 0 R /Rotate 10 0 R >>\nendobj\n"));
    for (const [num, body] of chain) {
      offsets.set(num, file.push(`${num} 0 obj\n${body}\nendobj\n`));
    }
    const xrefAt = pushClassicXref(file, offsets, 200, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    await expect(parsePdf(file.bytes)).rejects.toBeInstanceOf(PdfParseError);
  });

  it("detects a self-referential page tree (visited set)", async () => {
    // The pages node lists ITSELF as a kid.
    const file = new FileBuilder();
    file.push("%PDF-1.4\n");
    const offsets = new Map<number, number>();
    offsets.set(1, file.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"));
    offsets.set(2, file.push("2 0 obj\n<< /Type /Pages /Kids [2 0 R] /Count 1 >>\nendobj\n"));
    const xrefAt = pushClassicXref(file, offsets, 3, "/Root 1 0 R");
    file.push(`startxref\n${xrefAt}\n%%EOF\n`);
    await expect(parsePdf(file.bytes)).rejects.toBeInstanceOf(PdfParseError);
  });
});

describe("values and helpers", () => {
  it("parses names with #-escapes, nested strings, hex strings and refs", async () => {
    const bytes = simplePdf(
      [[4, "<< /A#42 (par(en)s) /Hex <48690> /Ref 3 0 R /Neg -1.5 /Arr [1 2 3] >>"]],
      "/Info 4 0 R",
    );
    const doc = await parsePdf(bytes);
    const info = doc.info;
    expect(info?.has("AB")).toBe(true);
    expect(pdfTextOf(info?.get("AB"))).toBe("par(en)s");
    const hex = info?.get("Hex");
    expect(hex?.t === "str" && [...hex.v]).toEqual([0x48, 0x69, 0x00]);
    const ref = info?.get("Ref");
    expect(ref?.t).toBe("ref");
    if (ref) {
      const target = resolve(doc, ref);
      expect(target.t).toBe("dict");
    }
    const neg = info?.get("Neg");
    expect(neg?.t === "num" && neg.v).toBe(-1.5);
  });

  it("decodes UTF-16BE strings and PDF dates", () => {
    const utf16: PdfValue = {
      t: "str",
      v: new Uint8Array([0xfe, 0xff, 0x00, 0x48, 0x00, 0x69]),
    };
    expect(pdfTextOf(utf16)).toBe("Hi");
    expect(parsePdfDate("D:20260830120500Z")).toBe("2026-08-30 12:05");
    expect(parsePdfDate("D:2026")).toBe("2026");
    expect(parsePdfDate("yesterday")).toBeNull();
  });

  it("reports the linearised flag from the leading dictionary", async () => {
    // A linearisation dict at the head of the file (values are stubs; the
    // flag only reports its presence).
    const raw = buildPagesPdfBytes();
    const text = new TextDecoder("latin1").decode(raw);
    const linearised = ascii(text.replace("%PDF-1.7\n", "%PDF-1.7\n7 0 obj\n<< /Linearized 1 >>\nendobj\n"));
    // Offsets shifted: the structured path fails and reconstruction (which
    // ignores offsets) takes over — the flag must still be read.
    const doc = await parsePdf(linearised);
    expect(doc.linearised).toBe(true);
    const plain = await parsePdf(raw);
    expect(plain.linearised).toBe(false);
  });
});
