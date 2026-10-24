// SPEC-111: assemblePdf, the writer half of core/pdfDocument.ts. The anchor
// is single-source identity (parse → assemble every page → re-parse: same
// page count, same geometry, equalBytes content streams); around it, object
// renumbering across sources, inherited-attribute materialisation, rotation
// arithmetic, what survives by reachability (resources, annotations with
// their appearances, boxes, optional content), link remapping and inert-link
// removal, named-destination merging, outline re-parenting, the drop-record
// table with one case per PdfDropKind, the encrypted and unreadable
// refusals, every bound, yielding/progress/cancellation, the twice-identical
// property, frozen inputs, and the re-parse self-check. Fixtures are the
// checked-in files scripts/build-media-fixtures.ts writes (spec107-pages,
// spec107-xrefstream, spec107-encrypted, spec111-forms, spec111-outlines),
// plus a few hand-assembled files for single constructs.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import {
  assemblePdf,
  checkAssembledPdf,
  MAX_MERGE_PAGES,
  MAX_MERGE_SOURCE_BYTES,
  mergeLimitRefusal,
  MERGE_YIELD_PAGES,
  normaliseRotation,
  openPdfSources,
  parsePdf,
  PDF_DROP_KINDS,
  PdfAssemblyCancelled,
  PdfAssemblyRefusal,
  PdfEncryptedError,
  plannedPdfPages,
  pdfTextOf,
  resolve,
  streamData,
  type PdfAssemblyPlan,
  type PdfDocument,
  type PdfDropKind,
  type PdfDropRecord,
  type PdfPageSelection,
  type PdfValue,
} from "./pdfDocument";
import { equalBytes } from "./equality";
import { assembleClassicPdf, buildJpegFixtureBytes } from "../media/fixtures/build";

const FIXTURES = join(__dirname, "../../../e2e/fixtures");
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(FIXTURES, name)));

const PAGES = fixture("spec107-pages.pdf");
const XREF_STREAM = fixture("spec107-xrefstream.pdf");
const ENCRYPTED = fixture("spec107-encrypted.pdf");
const FORMS = fixture("spec111-forms.pdf");
const OUTLINES = fixture("spec111-outlines.pdf");

const noYield = async (): Promise<void> => {};

/** Every page of `doc`, in order, under one key. */
function allPages(source: string, doc: PdfDocument): PdfPageSelection[] {
  return doc.pages.map((_, pageIndex) => ({ source, doc, pageIndex }));
}

function plan(pages: PdfPageSelection[], extra: Partial<PdfAssemblyPlan> = {}): PdfAssemblyPlan {
  return { pages, title: "Merged", outline: true, ...extra };
}

async function assemble(pages: PdfPageSelection[], extra: Partial<PdfAssemblyPlan> = {}) {
  const result = await assemblePdf(plan(pages, extra), { between: noYield });
  return { ...result, out: await parsePdf(result.bytes) };
}

function dictOf(doc: PdfDocument, value: PdfValue | undefined): Map<string, PdfValue> {
  const resolved = value === undefined ? ({ t: "null" } as PdfValue) : resolve(doc, value);
  if (resolved.t !== "dict") {
    throw new Error(`expected a dictionary, got ${resolved.t}`);
  }
  return resolved.v;
}

function streamOf(doc: PdfDocument, value: PdfValue | undefined): { dict: Map<string, PdfValue>; raw: Uint8Array } {
  const resolved = value === undefined ? ({ t: "null" } as PdfValue) : resolve(doc, value);
  if (resolved.t !== "stream") {
    throw new Error(`expected a stream, got ${resolved.t}`);
  }
  return resolved;
}

/** A page's own dictionary, as written (not through its parents). */
function pageDict(doc: PdfDocument, index: number): Map<string, PdfValue> {
  const ref = doc.pages[index].ref;
  return dictOf(doc, { t: "ref", num: ref.num, gen: ref.gen });
}

function plain(value: PdfValue | undefined): unknown {
  if (value === undefined) {
    return undefined;
  }
  switch (value.t) {
    case "null":
      return null;
    case "str":
      return pdfTextOf(value);
    case "array":
      return value.v.map(plain);
    case "dict":
      return Object.fromEntries(Array.from(value.v, ([key, item]) => [key, plain(item)]));
    case "stream":
      return { stream: Object.fromEntries(Array.from(value.dict, ([key, item]) => [key, plain(item)])) };
    case "ref":
      return `${value.num} ${value.gen} R`;
    default:
      return value.v;
  }
}

function refsIn(value: PdfValue, into: Array<{ num: number; gen: number }> = []): Array<{ num: number; gen: number }> {
  if (value.t === "ref") {
    into.push(value);
  } else if (value.t === "array") {
    value.v.forEach((item) => refsIn(item, into));
  } else if (value.t === "dict" || value.t === "stream") {
    (value.t === "dict" ? value.v : value.dict).forEach((item) => refsIn(item, into));
  }
  return into;
}

const latin1 = (bytes: Uint8Array): string => Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");

const ascii = (text: string): Uint8Array => Uint8Array.from(text, (char) => char.charCodeAt(0));

function refuseMutation(): never {
  throw new TypeError("a frozen input Map was mutated");
}

/** Deep-freezes a value: objects and arrays frozen, a Map's mutators made to
 * throw (Object.freeze does not reach a Map's entries). Typed arrays cannot
 * be frozen; callers compare their bytes instead. */
function deepFreeze<T>(value: T, seen = new Set<unknown>()): T {
  if (value === null || typeof value !== "object" || seen.has(value)) {
    return value;
  }
  seen.add(value);
  if (ArrayBuffer.isView(value)) {
    return value;
  }
  if (value instanceof Map) {
    for (const [key, item] of value) {
      deepFreeze(key, seen);
      deepFreeze(item, seen);
    }
    Object.defineProperties(value, { set: { value: refuseMutation }, delete: { value: refuseMutation }, clear: { value: refuseMutation } });
  } else {
    for (const item of Object.values(value)) {
      deepFreeze(item, seen);
    }
  }
  return Object.freeze(value);
}

/** A two-page document (pages 3 and 4) from object bodies 1–2 fixed and the
 * rest supplied: 1 is the catalog (with `catalogExtra`), 2 the page tree. */
function twoPagePdf(page1: string, page2: string, extra: string[] = [], catalogExtra = ""): Uint8Array {
  return assembleClassicPdf(
    "1.7",
    [
      `<< /Type /Catalog /Pages 2 0 R ${catalogExtra} >>`,
      "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /MediaBox [0 0 300 400] >>",
      `<< /Type /Page /Parent 2 0 R ${page1} >>`,
      `<< /Type /Page /Parent 2 0 R ${page2} >>`,
      ...extra,
    ],
    "/Root 1 0 R",
  );
}

describe("assemblePdf — single-source identity", () => {
  it("keeps the page count, the geometry and every content stream byte for byte", async () => {
    const source = await parsePdf(PAGES);
    const { out } = await assemble(allPages("pages", source));
    expect(out.pages).toHaveLength(source.pages.length);
    out.pages.forEach((page, index) => {
      const original = source.pages[index];
      expect(page.widthPt).toBe(original.widthPt);
      expect(page.heightPt).toBe(original.heightPt);
      expect(page.rotate).toBe(original.rotate);
      expect(page.mediaBoxPt).toEqual(original.mediaBoxPt);
      expect(page.cropBoxPt).toEqual(original.cropBoxPt);
      const copied = streamOf(out, pageDict(out, index).get("Contents"));
      const sourceContents = streamOf(source, pageDict(source, index).get("Contents"));
      expect(equalBytes(copied.raw, sourceContents.raw)).toBe(true);
      expect(plain(copied.dict.get("Filter"))).toBe("FlateDecode");
    });
  });

  it("copies streams verbatim: the filters, the parameters and the raw bytes, never re-encoded", async () => {
    const source = await parsePdf(FORMS);
    const { out } = await assemble(allPages("forms", source));
    for (const index of [0, 1]) {
      const copied = streamOf(out, pageDict(out, index).get("Contents"));
      const original = streamOf(source, pageDict(source, index).get("Contents"));
      expect(equalBytes(copied.raw, original.raw)).toBe(true);
      expect(plain(copied.dict.get("Filter"))).toEqual(plain(original.dict.get("Filter")));
      expect(equalBytes(await streamData(out, copied), await streamData(source, original))).toBe(true);
    }
    // The DCT image is the JPEG fixture's own bytes, still raw.
    const resources = dictOf(out, pageDict(out, 0).get("Resources"));
    const image = streamOf(out, dictOf(out, resources.get("XObject")).get("Im1"));
    expect(plain(image.dict.get("Filter"))).toBe("DCTDecode");
    expect(equalBytes(image.raw, buildJpegFixtureBytes())).toBe(true);
  });

  it("keeps /DecodeParms beside its filter", async () => {
    const data = ascii("BT ET");
    const bytes = twoPagePdf(
      "/Contents 5 0 R",
      "",
      [`<< /Filter [/ASCIIHexDecode] /DecodeParms [null] /Length ${data.length * 2 + 1} >>\nstream\n${Array.from(data, (b) => b.toString(16).padStart(2, "0")).join("")}>\nendstream`],
    );
    const source = await parsePdf(bytes);
    const { out } = await assemble([{ source: "hex", doc: source, pageIndex: 0 }]);
    const copied = streamOf(out, pageDict(out, 0).get("Contents"));
    expect(plain(copied.dict.get("Filter"))).toEqual(["ASCIIHexDecode"]);
    expect(plain(copied.dict.get("DecodeParms"))).toEqual([null]);
    expect(latin1(await streamData(out, copied))).toBe("BT ET");
  });

  it("writes a PDF 1.7 file with a classic xref table, whatever the source used", async () => {
    const source = await parsePdf(XREF_STREAM);
    const { bytes, out } = await assemble(allPages("xref", source));
    const text = latin1(bytes);
    expect(text.startsWith("%PDF-1.7\n")).toBe(true);
    expect(text).toMatch(/\nxref\n0 \d+\n0000000000 65535 f \n/);
    expect(text).toMatch(/\ntrailer\n<< \/Size \d+ \/Root 1 0 R \/Info \d+ 0 R >>\nstartxref\n\d+\n%%EOF\n$/);
    expect(text).not.toContain("/ObjStm");
    expect(text).not.toContain("/XRef");
    expect(text).not.toContain("/Linearized");
    expect(out.version).toBe("1.7");
    expect(out.reconstructed).toBe(false);
    expect(out.pages[0]).toMatchObject({ widthPt: 400, heightPt: 400, rotate: 180 });
  });
});

describe("assemblePdf — renumbering", () => {
  it("numbers objects from 1 with no gaps or collisions, and every reference resolves", async () => {
    const pages = await parsePdf(PAGES);
    const outlines = await parsePdf(OUTLINES);
    const forms = await parsePdf(FORMS);
    const { bytes, out } = await assemble([
      ...allPages("pages", pages),
      ...allPages("outlines", outlines),
      ...allPages("forms", forms),
    ]);
    const text = latin1(bytes);
    const size = Number(/\/Size (\d+)/.exec(text)?.[1]);
    const headers = Array.from(text.matchAll(/(?:^|\n)(\d+) 0 obj\n/g), (match) => Number(match[1]));
    expect(headers).toEqual(Array.from({ length: size - 1 }, (_, i) => i + 1));
    expect(out.objects.size).toBe(size - 1);
    for (const [num, value] of out.objects) {
      for (const ref of refsIn(value)) {
        expect(out.objects.has(ref.num), `object ${num} refers to missing ${ref.num}`).toBe(true);
        expect(ref.gen).toBe(0);
      }
    }
    expect(out.pages).toHaveLength(7);
  });

  it("copies an object shared by pages of one file once, and the same object in two files twice", async () => {
    const first = await parsePdf(PAGES);
    const second = await parsePdf(PAGES);
    const { out } = await assemble([...allPages("first", first), ...allPages("second", second)]);
    const resources = out.pages.map((_, index) => plain(pageDict(out, index).get("Resources")));
    expect(new Set(resources.slice(0, 3)).size).toBe(1);
    expect(new Set(resources.slice(3)).size).toBe(1);
    expect(resources[0]).not.toBe(resources[3]);
  });
});

describe("assemblePdf — inherited attributes", () => {
  it("writes /Resources, /MediaBox, /CropBox and /Rotate onto each copied page", async () => {
    const source = await parsePdf(FORMS);
    expect(pageDict(source, 0).has("Rotate")).toBe(false);
    const { out } = await assemble(allPages("forms", source));
    const first = pageDict(out, 0);
    expect(first.get("Resources")?.t).toBe("ref");
    expect(Array.from(dictOf(out, first.get("Resources")).keys())).toEqual(
      expect.arrayContaining(["Font", "XObject", "Properties", "ColorSpace", "Shading", "Pattern", "ExtGState"]),
    );
    expect(plain(first.get("MediaBox"))).toEqual([0, 0, 612, 792]);
    expect(plain(first.get("CropBox"))).toEqual([10, 10, 602, 782]);
    expect(plain(first.get("Rotate"))).toBe(90);
    const second = pageDict(out, 1);
    expect(plain(second.get("MediaBox"))).toEqual([0, 0, 500, 500]);
    expect(plain(second.get("CropBox"))).toEqual([10, 10, 602, 782]);
    expect(plain(second.get("Rotate"))).toBe(0);
    // The new tree carries nothing to inherit.
    const tree = dictOf(out, pageDict(out, 0).get("Parent"));
    expect(Array.from(tree.keys())).toEqual(["Type", "Kids", "Count"]);
    expect(out.pages.map((page) => [page.widthPt, page.heightPt, page.rotate])).toEqual(
      source.pages.map((page) => [page.widthPt, page.heightPt, page.rotate]),
    );
  });

  it("keeps /BleedBox, /TrimBox and /ArtBox on the page", async () => {
    const { out } = await assemble([{ source: "forms", doc: await parsePdf(FORMS), pageIndex: 0 }]);
    const page = pageDict(out, 0);
    expect(plain(page.get("BleedBox"))).toEqual([5, 5, 607, 787]);
    expect(plain(page.get("TrimBox"))).toEqual([15, 15, 597, 777]);
    expect(plain(page.get("ArtBox"))).toEqual([20, 20, 592, 772]);
  });

  it("gives a page with no resources anywhere an empty dictionary, and inherits a parent's box", async () => {
    const source = await parsePdf(twoPagePdf("", "/MediaBox [0 0 100 100]"));
    const { out } = await assemble(allPages("bare", source));
    expect(plain(pageDict(out, 0).get("Resources"))).toEqual({});
    expect(plain(pageDict(out, 0).get("MediaBox"))).toEqual([0, 0, 300, 400]);
    expect(plain(pageDict(out, 1).get("MediaBox"))).toEqual([0, 0, 100, 100]);
  });
});

describe("assemblePdf — rotation arithmetic", () => {
  it.each([
    [0, 0],
    [90, 90],
    [180, 180],
    [270, 270],
    [360, 0],
    [450, 90],
    [-90, 270],
    [-180, 180],
    [-450, 270],
    [Number.NaN, 0],
  ])("normaliseRotation(%s) is %s", (degrees, expected) => {
    expect(normaliseRotation(degrees)).toBe(expected);
  });

  it("adds rotateBy to each page's own /Rotate, modulo 360", async () => {
    const pages = await parsePdf(PAGES); // page 2 is /Rotate 90
    const xref = await parsePdf(XREF_STREAM); // its page is /Rotate 180
    const selections: PdfPageSelection[] = [
      { source: "pages", doc: pages, pageIndex: 0, rotateBy: 90 },
      { source: "pages", doc: pages, pageIndex: 1, rotateBy: 90 },
      { source: "pages", doc: pages, pageIndex: 1, rotateBy: 270 },
      { source: "pages", doc: pages, pageIndex: 2, rotateBy: 180 },
      { source: "xref", doc: xref, pageIndex: 0, rotateBy: 270 },
      { source: "xref", doc: xref, pageIndex: 0 },
    ];
    const { out } = await assemble(selections);
    expect(out.pages.map((page) => page.rotate)).toEqual([90, 180, 0, 180, 90, 180]);
    expect(plannedPdfPages(plan(selections)).map((page) => page.rotate)).toEqual([90, 180, 0, 180, 90, 180]);
  });
});

describe("assemblePdf — what survives by reachability", () => {
  it("carries content, fonts, XObjects, colour spaces, shadings, patterns and optional content groups", async () => {
    const { out } = await assemble([{ source: "forms", doc: await parsePdf(FORMS), pageIndex: 0 }]);
    const resources = dictOf(out, pageDict(out, 0).get("Resources"));
    expect(plain(dictOf(out, dictOf(out, resources.get("Font")).get("F1")).get("BaseFont"))).toBe("Helvetica");
    const form = streamOf(out, dictOf(out, resources.get("XObject")).get("Fm1"));
    expect(latin1(form.raw)).toBe("BT /F1 12 Tf 4 14 Td (Layer two) Tj ET");
    expect(plain(dictOf(out, form.dict.get("OC")).get("Name"))).toBe("Layer two");
    const colourSpace = resolve(out, dictOf(out, resources.get("ColorSpace")).get("CS1") ?? { t: "null" });
    expect(plain(colourSpace)).toEqual(["Indexed", "DeviceRGB", 1, expect.any(String)]);
    const shading = dictOf(out, dictOf(out, resources.get("Shading")).get("Sh1"));
    expect(plain(dictOf(out, shading.get("Function")).get("FunctionType"))).toBe(2);
    const pattern = dictOf(out, dictOf(out, resources.get("Pattern")).get("P1"));
    expect(plain(dictOf(out, pattern.get("Shading")).get("ShadingType"))).toBe(2);
    expect(plain(dictOf(out, resources.get("ExtGState")).get("GS1"))).toEqual({ CA: 0.5, ca: 0.5 });
    expect(plain(dictOf(out, dictOf(out, resources.get("Properties")).get("oc1")).get("Name"))).toBe("Layer one");
  });

  it("keeps widgets with their appearance streams — a filled form still looks filled — and leaves out the signature's bytes", async () => {
    const { bytes, out } = await assemble([{ source: "forms", doc: await parsePdf(FORMS), pageIndex: 0 }]);
    const page = pageDict(out, 0);
    const annots = resolve(out, page.get("Annots") ?? { t: "null" });
    expect(annots.t).toBe("array");
    const widgets = (annots as { v: PdfValue[] }).v.map((entry) => dictOf(out, entry));
    expect(widgets.map((widget) => plain(widget.get("Subtype")))).toEqual(["Widget", "Widget", "Widget"]);
    const [text, checkbox, signature] = widgets;
    expect(latin1(streamOf(out, dictOf(out, text.get("AP")).get("N")).raw)).toContain("(Ada Lovelace) Tj");
    expect(plain(text.get("V"))).toBe("Ada Lovelace");
    expect(plain(checkbox.get("AS"))).toBe("Yes");
    const states = dictOf(out, dictOf(out, checkbox.get("AP")).get("N"));
    expect(latin1(streamOf(out, states.get("Yes")).raw)).toBe("0 0 1 rg 3 3 12 12 re f");
    // The field hierarchy belongs to the dropped /AcroForm.
    expect(checkbox.has("Parent")).toBe(false);
    expect(latin1(streamOf(out, dictOf(out, signature.get("AP")).get("N")).raw)).toContain("Signed: A. Lovelace");
    expect(signature.has("V")).toBe(false);
    expect(latin1(bytes)).not.toContain("0123456789ABCDEF");
    // Each annotation names the page it sits on.
    for (const widget of widgets) {
      expect(plain(widget.get("P"))).toBe(`${out.pages[0].ref.num} 0 R`);
    }
  });

  it("copies an external /URI link unchanged", async () => {
    const source = await parsePdf(OUTLINES);
    const { out } = await assemble(allPages("outlines", source));
    const annots = resolve(out, pageDict(out, 0).get("Annots") ?? { t: "null" }) as { v: PdfValue[] };
    const web = annots.v.map((entry) => dictOf(out, entry)).find((annot) => annot.get("A") !== undefined && dictOf(out, annot.get("A")).has("URI"));
    const original = (resolve(source, pageDict(source, 0).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v
      .map((entry) => dictOf(source, entry))
      .find((annot) => annot.get("A") !== undefined && dictOf(source, annot.get("A")).has("URI"));
    expect(web && plain(web.get("A"))).toEqual({ S: "URI", URI: "https://example.com/spec111" });
    expect(web && plain(web.get("Rect"))).toEqual(original && plain(original.get("Rect")));
  });

  it("gives a page selected twice its own copy of each annotation", async () => {
    const source = await parsePdf(FORMS);
    const { out } = await assemble([
      { source: "forms", doc: source, pageIndex: 0 },
      { source: "forms", doc: source, pageIndex: 0 },
    ]);
    const annotsOf = (index: number) => (resolve(out, pageDict(out, index).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v;
    const first = annotsOf(0);
    const second = annotsOf(1);
    expect(first).toHaveLength(3);
    expect(second).toHaveLength(3);
    const numbers = [...first, ...second].map((entry) => (entry as { num: number }).num);
    expect(new Set(numbers).size).toBe(6);
    first.forEach((entry) => expect(plain(dictOf(out, entry).get("P"))).toBe(`${out.pages[0].ref.num} 0 R`));
    second.forEach((entry) => expect(plain(dictOf(out, entry).get("P"))).toBe(`${out.pages[1].ref.num} 0 R`));
    // What they reach is shared: one appearance stream for both copies.
    expect(plain(dictOf(out, first[0]).get("AP"))).toEqual(plain(dictOf(out, second[0]).get("AP")));
  });
});

describe("assemblePdf — links", () => {
  it("remaps an internal /GoTo destination to the copied page's new object", async () => {
    const source = await parsePdf(OUTLINES);
    // Reversed: source page 2 is output page 1.
    const { out, dropped } = await assemble([
      { source: "outlines", doc: source, pageIndex: 1 },
      { source: "outlines", doc: source, pageIndex: 0 },
    ]);
    expect(dropped.filter((record) => record.kind === "annotation")).toEqual([]);
    const links = (resolve(out, pageDict(out, 1).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v.map((entry) =>
      dictOf(out, entry),
    );
    const explicit = links.find((link) => link.has("Dest"));
    expect(plain(explicit?.get("Dest"))).toEqual([`${out.pages[0].ref.num} 0 R`, "XYZ", 0, 420, null]);
    // Named destinations stay named, and the names point at the new objects.
    const names = plain(dictOf(out, dictOf(out, dictOf(out, out.trailer.get("Root")).get("Names")).get("Dests")).get("Names"));
    expect(names).toEqual([
      "chapter1",
      [`${out.pages[1].ref.num} 0 R`, "XYZ", 0, 595, null],
      "chapter2",
      [`${out.pages[0].ref.num} 0 R`, "XYZ", 0, 420, null],
    ]);
  });

  it("removes a link to a page not in the output, reporting it, and keeps the rest", async () => {
    const source = await parsePdf(OUTLINES);
    const { out, dropped } = await assemble([{ source: "outlines", doc: source, pageIndex: 0 }]);
    const annots = plain(pageDict(out, 0).get("Annots")) as string[];
    expect(annots).toHaveLength(1);
    expect(plain(dictOf(out, pageDict(out, 0).get("Annots") && { t: "ref", num: Number(annots[0].split(" ")[0]), gen: 0 }).get("A"))).toEqual({
      S: "URI",
      URI: "https://example.com/spec111",
    });
    const removed = dropped.filter((record) => record.kind === "annotation");
    expect(removed).toEqual([
      { kind: "annotation", source: "outlines", page: 1, reason: "A link on page 1 pointed at page 2, which is not in the merged PDF, so it was removed." },
      { kind: "annotation", source: "outlines", page: 1, reason: "A link on page 1 pointed at page 2, which is not in the merged PDF, so it was removed." },
    ]);
  });

  it("makes a non-link annotation whose action points at a dropped page inert, keeping the annotation", async () => {
    const source = await parsePdf(
      twoPagePdf("/Annots [5 0 R]", "", [
        "<< /Type /Annot /Subtype /Widget /FT /Btn /Rect [0 0 50 20] /A << /S /GoTo /D [4 0 R /Fit] >> /AP << /N 6 0 R >> >>",
        "<< /Type /XObject /Subtype /Form /BBox [0 0 50 20] /Length 0 >>\nstream\n\nendstream",
      ]),
    );
    const { out, dropped } = await assemble([{ source: "button", doc: source, pageIndex: 0 }]);
    const [button] = (resolve(out, pageDict(out, 0).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v.map((entry) => dictOf(out, entry));
    expect(button.has("A")).toBe(false);
    expect(button.has("AP")).toBe(true);
    expect(dropped).toEqual([
      {
        kind: "annotation",
        source: "button",
        page: 1,
        reason: "A Widget annotation on page 1 pointed at page 2, which is not in the merged PDF, so its action was removed.",
      },
    ]);
  });
});

describe("assemblePdf — named destinations", () => {
  it("merges /Dests with a source-qualified prefix on collision, and links follow the new names", async () => {
    const first = await parsePdf(OUTLINES);
    const second = await parsePdf(OUTLINES);
    const { out } = await assemble([...allPages("first", first), ...allPages("second", second)]);
    const tree = plain(dictOf(out, dictOf(out, dictOf(out, out.trailer.get("Root")).get("Names")).get("Dests")).get("Names")) as unknown[];
    expect(tree.filter((_, index) => index % 2 === 0)).toEqual(["chapter1", "chapter2", "second:chapter1", "second:chapter2"]);
    expect(tree[5]).toEqual([`${out.pages[2].ref.num} 0 R`, "XYZ", 0, 595, null]);
    const linksOf = (index: number) =>
      (resolve(out, pageDict(out, index).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v.map((entry) => plain(resolve(out, entry)));
    expect(linksOf(1)[0]).toMatchObject({ Dest: "chapter1" });
    expect(linksOf(3)[0]).toMatchObject({ Dest: "second:chapter1" });
    expect(linksOf(0)[1]).toMatchObject({ A: { S: "GoTo", D: "chapter2" } });
    expect(linksOf(2)[1]).toMatchObject({ A: { S: "GoTo", D: "second:chapter2" } });
  });

  it("reads a PDF 1.1 /Dests dictionary, turning name destinations into strings in the merged name tree", async () => {
    const source = await parsePdf(
      twoPagePdf("/Annots [5 0 R]", "", ["<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] /Dest /intro >>"], "/Dests << /intro [4 0 R /Fit] >>"),
    );
    const { out, dropped } = await assemble(allPages("old", source));
    expect(dropped).toEqual([]);
    const [link] = (resolve(out, pageDict(out, 0).get("Annots") ?? { t: "null" }) as { v: PdfValue[] }).v.map((entry) => dictOf(out, entry));
    expect(link.get("Dest")).toEqual({ t: "str", v: ascii("intro") });
    const tree = plain(dictOf(out, dictOf(out, dictOf(out, out.trailer.get("Root")).get("Names")).get("Dests")).get("Names"));
    expect(tree).toEqual(["intro", [`${out.pages[1].ref.num} 0 R`, "Fit"]]);
  });

  it("drops a named destination whose page is not in the output, with a namedDestination record", async () => {
    const { out, dropped } = await assemble([{ source: "outlines", doc: await parsePdf(OUTLINES), pageIndex: 0 }]);
    expect(dropped.filter((record) => record.kind === "namedDestination")).toEqual([
      {
        kind: "namedDestination",
        source: "outlines",
        page: 2,
        reason: "The named destination “chapter2” pointed at page 2, which is not in the merged PDF, so it was removed.",
      },
    ]);
    const tree = plain(dictOf(out, dictOf(out, dictOf(out, out.trailer.get("Root")).get("Names")).get("Dests")).get("Names")) as unknown[];
    expect(tree[0]).toBe("chapter1");
    expect(tree).toHaveLength(2);
  });
});

describe("assemblePdf — optional content", () => {
  it("concatenates /OCGs and unions /D's ON and OFF across sources", async () => {
    const first = await parsePdf(FORMS);
    const second = await parsePdf(FORMS);
    const { out } = await assemble([
      { source: "first", doc: first, pageIndex: 0 },
      { source: "second", doc: second, pageIndex: 0 },
    ]);
    const props = dictOf(out, dictOf(out, out.trailer.get("Root")).get("OCProperties"));
    const groups = (resolve(out, props.get("OCGs") ?? { t: "null" }) as { v: PdfValue[] }).v;
    expect(groups.map((group) => plain(dictOf(out, group).get("Name")))).toEqual(["Layer one", "Layer two", "Layer one", "Layer two"]);
    const defaults = dictOf(out, props.get("D"));
    const names = (key: string) =>
      (resolve(out, defaults.get(key) ?? { t: "null" }) as { v: PdfValue[] }).v.map((group) => plain(dictOf(out, group).get("Name")));
    expect(names("ON")).toEqual(["Layer one", "Layer one"]);
    expect(names("OFF")).toEqual(["Layer two", "Layer two"]);
    expect(names("Order")).toEqual(["Layer one", "Layer two", "Layer one", "Layer two"]);
    // The page's own group references are the merged ones.
    const resources = dictOf(out, pageDict(out, 1).get("Resources"));
    expect(plain(dictOf(out, resources.get("Properties")).get("oc1"))).toBe(plain(groups[2]));
  });

  it("turns a /BaseState /OFF source's groups off unless it turns them on", async () => {
    const source = await parsePdf(
      twoPagePdf("", "", ["<< /Type /OCG /Name (A) >>", "<< /Type /OCG /Name (B) >>"], "/OCProperties << /OCGs [5 0 R 6 0 R] /D << /BaseState /OFF /ON [5 0 R] >> >>"),
    );
    const { out } = await assemble(allPages("base", source));
    const defaults = dictOf(out, dictOf(out, dictOf(out, out.trailer.get("Root")).get("OCProperties")).get("D"));
    const names = (key: string) =>
      (resolve(out, defaults.get(key) ?? { t: "null" }) as { v: PdfValue[] }).v.map((group) => plain(dictOf(out, group).get("Name")));
    expect(names("ON")).toEqual(["A"]);
    expect(names("OFF")).toEqual(["B"]);
  });
});

describe("assemblePdf — outlines", () => {
  it("re-parents each source's outline under one top-level entry per source, in first-appearance order", async () => {
    const outlines = await parsePdf(OUTLINES);
    const pages = await parsePdf(PAGES);
    const { out, dropped } = await assemble([
      { source: "outlines", doc: outlines, pageIndex: 0 },
      { source: "pages", doc: pages, pageIndex: 0 },
      { source: "outlines", doc: outlines, pageIndex: 1 },
    ]);
    expect(dropped.filter((record) => record.kind === "outlineEntry")).toEqual([]);
    const root = dictOf(out, dictOf(out, out.trailer.get("Root")).get("Outlines"));
    const pageRef = (index: number) => `${out.pages[index].ref.num} 0 R`;
    type Item = { title: string; dest: unknown; count?: number; children: Item[] };
    const read = (first: PdfValue | undefined, parent: Map<string, PdfValue>): Item[] => {
      const items: Item[] = [];
      for (let cursor = first; cursor; ) {
        const item = dictOf(out, cursor);
        expect(dictOf(out, item.get("Parent"))).toBe(parent);
        const action = item.get("A") ? dictOf(out, item.get("A")) : null;
        items.push({
          title: pdfTextOf(item.get("Title")),
          dest: plain(item.get("Dest") ?? action?.get("D")),
          count: plain(item.get("Count")) as number | undefined,
          children: read(item.get("First"), item),
        });
        cursor = item.get("Next");
      }
      return items;
    };
    expect(read(root.get("First"), root)).toEqual([
      {
        title: "outlines",
        dest: [pageRef(0), "Fit"],
        count: 3,
        children: [
          { title: "Chapter 1", dest: [pageRef(0), "Fit"], count: 1, children: [{ title: "Section 1.1", dest: "chapter1", children: [] }] },
          { title: "Chapter 2", dest: [pageRef(2), "Fit"], children: [] },
        ],
      },
      { title: "pages", dest: [pageRef(1), "Fit"], children: [] },
    ]);
    expect(plain(root.get("Count"))).toBe(5);
  });

  it("removes an entry pointing at a dropped page and promotes its surviving children", async () => {
    const source = await parsePdf(
      twoPagePdf(
        "",
        "",
        [
          "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 2 >>",
          "<< /Title (Part one) /Parent 5 0 R /First 7 0 R /Last 7 0 R /Count 1 /Dest [3 0 R /Fit] >>",
          "<< /Title (Later) /Parent 6 0 R /Dest [4 0 R /Fit] >>",
        ],
        "/Outlines 5 0 R",
      ),
    );
    const { out, dropped } = await assemble([{ source: "parts", doc: source, pageIndex: 1 }]);
    expect(dropped).toEqual([
      { kind: "outlineEntry", source: "parts", page: 1, reason: "The bookmark “Part one” pointed at page 1, which is not in the merged PDF, so it was removed." },
    ]);
    const root = dictOf(out, dictOf(out, out.trailer.get("Root")).get("Outlines"));
    const top = dictOf(out, root.get("First"));
    expect(pdfTextOf(top.get("Title"))).toBe("parts");
    const child = dictOf(out, top.get("First"));
    expect(pdfTextOf(child.get("Title"))).toBe("Later");
    expect(plain(child.get("Dest"))).toEqual([`${out.pages[0].ref.num} 0 R`, "Fit"]);
    expect(child.has("Next")).toBe(false);
  });

  it("writes no outline when the plan asks for none, reporting the bookmarks left out", async () => {
    const { out, dropped } = await assemble(allPages("outlines", await parsePdf(OUTLINES)), { outline: false });
    expect(dictOf(out, out.trailer.get("Root")).has("Outlines")).toBe(false);
    expect(dropped.filter((record) => record.kind === "outlineEntry")).toEqual([
      { kind: "outlineEntry", source: "outlines", reason: "Its bookmarks were left out: this merge has no outline." },
    ]);
  });
});

/** The records a refused set of sources carries. */
async function refusalOf(inputs: Array<{ source: string; bytes: Uint8Array }>): Promise<PdfDropRecord[]> {
  try {
    await openPdfSources(inputs);
  } catch (error) {
    if (error instanceof PdfAssemblyRefusal) {
      return error.dropped;
    }
    throw error;
  }
  throw new Error("expected a refusal");
}

describe("assemblePdf — the drop-record table", () => {
  /** What one merge (or refusal) reports. */
  async function dropsOf(selections: () => Promise<PdfPageSelection[]>): Promise<PdfDropRecord[]> {
    return (await assemblePdf(plan(await selections()), { between: noYield })).dropped;
  }
  const forms = async () => allPages("forms", await parsePdf(FORMS));
  const outlinesPage1 = async () => [{ source: "outlines", doc: await parsePdf(OUTLINES), pageIndex: 0 }];

  const cases: Record<PdfDropKind, { run: () => Promise<PdfDropRecord[]>; expected: PdfDropRecord }> = {
    acroForm: {
      run: () => dropsOf(forms),
      expected: {
        kind: "acroForm",
        source: "forms",
        reason:
          "Its form fields are no longer fillable: what was filled in still shows, but the form itself is not carried into a merged PDF. " +
          "Its digital signature does not survive a merge and was not carried.",
      },
    },
    structureTree: {
      run: () => dropsOf(async () => allPages("outlines", await parsePdf(OUTLINES))),
      expected: {
        kind: "structureTree",
        source: "outlines",
        reason: "Its structure tree (the tags that make it accessible) was not carried, so the merged PDF is not a tagged PDF.",
      },
    },
    javascript: {
      run: () => dropsOf(forms),
      expected: { kind: "javascript", source: "forms", reason: "Its document scripts (JavaScript) were removed." },
    },
    openAction: {
      run: () => dropsOf(forms),
      expected: { kind: "openAction", source: "forms", reason: "What it did on opening (its open action) was removed." },
    },
    embeddedFile: {
      run: () => dropsOf(forms),
      expected: { kind: "embeddedFile", source: "forms", reason: "The attached file “data.csv” was not carried." },
    },
    namedDestination: {
      run: () => dropsOf(outlinesPage1),
      expected: {
        kind: "namedDestination",
        source: "outlines",
        page: 2,
        reason: "The named destination “chapter2” pointed at page 2, which is not in the merged PDF, so it was removed.",
      },
    },
    outlineEntry: {
      run: () => dropsOf(outlinesPage1),
      expected: {
        kind: "outlineEntry",
        source: "outlines",
        page: 2,
        reason: "The bookmark “Chapter 2” pointed at page 2, which is not in the merged PDF, so it was removed.",
      },
    },
    annotation: {
      run: () => dropsOf(outlinesPage1),
      expected: {
        kind: "annotation",
        source: "outlines",
        page: 1,
        reason: "A link on page 1 pointed at page 2, which is not in the merged PDF, so it was removed.",
      },
    },
    pageLabel: {
      run: () => dropsOf(forms),
      expected: { kind: "pageLabel", source: "forms", reason: "Its page labels were removed; the merged PDF numbers its pages 1 to 2." },
    },
    encryptedSource: {
      run: () => refusalOf([{ source: "locked", bytes: ENCRYPTED }]),
      expected: { kind: "encryptedSource", source: "locked", reason: "“locked” is password-protected, so it cannot be merged." },
    },
    unreadableSource: {
      run: () => refusalOf([{ source: "junk", bytes: ascii("%PDF-1.7\nnot a pdf at all") }]),
      expected: {
        kind: "unreadableSource",
        source: "junk",
        reason: expect.stringMatching(/^“junk” could not be read as a PDF \(.+\), so it cannot be merged\.$/) as unknown as string,
      },
    },
  };

  it("has exactly one case per PdfDropKind", () => {
    expect(Object.keys(cases).toSorted()).toEqual([...PDF_DROP_KINDS].toSorted());
  });

  it.each(PDF_DROP_KINDS.map((kind) => [kind]))("reports %s", async (kind) => {
    const { run, expected } = cases[kind];
    expect(await run()).toContainEqual(expected);
  });

  it("records the form once per source", async () => {
    const first = await parsePdf(FORMS);
    const second = await parsePdf(FORMS);
    const dropped = await dropsOf(async () => [...allPages("first", first), ...allPages("second", second)]);
    expect(dropped.filter((record) => record.kind === "acroForm").map((record) => record.source)).toEqual(["first", "second"]);
  });

  it("drops nothing from a plain file", async () => {
    expect(await dropsOf(async () => allPages("pages", await parsePdf(PAGES)))).toEqual([]);
  });
});

describe("assemblePdf — refusals", () => {
  it("refuses the whole plan when any source is encrypted or unreadable, naming every file", async () => {
    const parse = vi.fn(parsePdf);
    let refusal: unknown;
    try {
      await openPdfSources(
        [
          { source: "pages", bytes: PAGES },
          { source: "locked", bytes: ENCRYPTED },
          { source: "junk", bytes: ascii("not a PDF") },
        ],
        parse,
      );
    } catch (error) {
      refusal = error;
    }
    expect(refusal).toBeInstanceOf(PdfAssemblyRefusal);
    const error = refusal as PdfAssemblyRefusal;
    expect(error.message).toContain("“locked” is password-protected");
    expect(error.message).toContain("“junk” could not be read as a PDF");
    expect(error.message).toContain("Remove them from the plan or replace them, then try again.");
    expect(error.dropped.map((record) => [record.kind, record.source])).toEqual([
      ["encryptedSource", "locked"],
      ["unreadableSource", "junk"],
    ]);
    // SPEC-107's own refusal is the cause.
    expect(error.cause).toBeInstanceOf(PdfEncryptedError);
    expect(parse).toHaveBeenCalledTimes(3);
  });

  it("returns every document when no source is refused", async () => {
    const docs = await openPdfSources([
      { source: "pages", bytes: PAGES },
      { source: "forms", bytes: FORMS },
    ]);
    expect(docs.map((doc) => doc.pages.length)).toEqual([3, 2]);
  });

  it("refuses an encrypted source before anything else runs", async () => {
    const source = await parsePdf(PAGES);
    const locked: PdfDocument = { ...source, trailer: new Map([...source.trailer, ["Encrypt", { t: "dict", v: new Map() }]]) };
    const onProgress = vi.fn();
    await expect(
      assemblePdf(plan([...allPages("pages", source), { source: "locked", doc: locked, pageIndex: 0 }]), { onProgress }),
    ).rejects.toMatchObject({
      name: "Error",
      message: "“locked” is password-protected, so it cannot be merged. Remove it from the plan or replace it, then try again.",
      dropped: [{ kind: "encryptedSource", source: "locked", reason: "“locked” is password-protected, so it cannot be merged." }],
    });
    expect(onProgress).not.toHaveBeenCalled();
  });

  it("refuses a selection of a page its source does not have", async () => {
    const source = await parsePdf(PAGES);
    await expect(assemblePdf(plan([{ source: "pages", doc: source, pageIndex: 3 }]))).rejects.toMatchObject({
      dropped: [{ kind: "unreadableSource", source: "pages", reason: "“pages” could not be read as a PDF (it has no page 4), so it cannot be merged." }],
    });
  });
});

describe("assemblePdf — bounds", () => {
  it("states each bound, naming the limit", () => {
    expect(mergeLimitRefusal(MAX_MERGE_PAGES, MAX_MERGE_SOURCE_BYTES)).toBeNull();
    expect(mergeLimitRefusal(MAX_MERGE_PAGES + 1, 0)).toBe(
      "A merged PDF can have at most 2000 pages; this one would have 2001. Remove 1 and try again.",
    );
    expect(mergeLimitRefusal(1, MAX_MERGE_SOURCE_BYTES + 1)).toBe(
      "The PDFs in one merge can add up to at most 500 MB; these add up to 501 MB. Remove a source and try again.",
    );
  });

  it("refuses a plan over MAX_MERGE_PAGES before copying a page", async () => {
    const source = await parsePdf(PAGES);
    const pages = Array.from({ length: MAX_MERGE_PAGES + 1 }, () => ({ source: "pages", doc: source, pageIndex: 0 }));
    const onProgress = vi.fn();
    await expect(assemblePdf(plan(pages), { onProgress })).rejects.toThrow(/at most 2000 pages/);
    expect(onProgress).not.toHaveBeenCalled();
  });

  it("merges a plan of exactly MAX_MERGE_PAGES", async () => {
    const source = await parsePdf(PAGES);
    const pages = Array.from({ length: MAX_MERGE_PAGES }, (_, i) => ({ source: "pages", doc: source, pageIndex: i % 3 }));
    const { bytes } = await assemblePdf(plan(pages, { outline: false }), { between: noYield });
    expect((await parsePdf(bytes)).pages).toHaveLength(MAX_MERGE_PAGES);
  });

  it("refuses sources over MAX_MERGE_SOURCE_BYTES before copying a page", async () => {
    const source = await parsePdf(PAGES);
    const huge = new Uint8Array(source.bytes);
    Object.defineProperty(huge, "byteLength", { value: MAX_MERGE_SOURCE_BYTES - 1000 });
    const big: PdfDocument = { ...source, bytes: huge };
    const onProgress = vi.fn();
    await expect(
      assemblePdf(plan([{ source: "big", doc: big, pageIndex: 0 }, { source: "pages", doc: source, pageIndex: 0 }]), { onProgress }),
    ).rejects.toThrow(/at most 500 MB/);
    expect(onProgress).not.toHaveBeenCalled();
  });

  it("refuses an empty plan", async () => {
    await expect(assemblePdf(plan([]))).rejects.toThrow("There are no pages to merge.");
  });
});

describe("assemblePdf — yielding, progress and cancellation", () => {
  it(`yields every ${MERGE_YIELD_PAGES} pages and reports each one`, async () => {
    const source = await parsePdf(PAGES);
    const between = vi.fn(noYield);
    const progress: Array<[number, number]> = [];
    await assemblePdf(plan(Array.from({ length: 20 }, () => ({ source: "pages", doc: source, pageIndex: 0 }))), {
      between,
      onProgress: (done, total) => progress.push([done, total]),
    });
    expect(between).toHaveBeenCalledTimes(2);
    expect(progress).toEqual(Array.from({ length: 21 }, (_, done) => [done, 20]));
  });

  it("stops at the next page boundary when cancelled, producing nothing", async () => {
    const source = await parsePdf(PAGES);
    let done = 0;
    await expect(
      assemblePdf(plan(Array.from({ length: 10 }, () => ({ source: "pages", doc: source, pageIndex: 0 }))), {
        between: noYield,
        onProgress: (count) => {
          done = count;
        },
        cancelled: () => done >= 3,
      }),
    ).rejects.toBeInstanceOf(PdfAssemblyCancelled);
    expect(done).toBe(3);
  });
});

describe("assemblePdf — determinism and purity", () => {
  it("writes equalBytes-identical output when run twice on the same plan", async () => {
    const pages = await parsePdf(PAGES);
    const outlines = await parsePdf(OUTLINES);
    const forms = await parsePdf(FORMS);
    const twice = plan(
      [
        { source: "outlines", doc: outlines, pageIndex: 1, rotateBy: 90 },
        ...allPages("pages", pages),
        ...allPages("forms", forms),
        { source: "outlines", doc: outlines, pageIndex: 0 },
      ],
      { modified: new Date(Date.UTC(2026, 9, 24, 12, 0, 0)), author: "Ada", subject: "Merging" },
    );
    const a = await assemblePdf(twice, { between: noYield });
    const b = await assemblePdf(twice, { between: noYield });
    expect(equalBytes(a.bytes, b.bytes)).toBe(true);
    expect(a.dropped).toEqual(b.dropped);
    const info = dictOf(await parsePdf(a.bytes), (await parsePdf(a.bytes)).trailer.get("Info"));
    expect(plain(info.get("Title"))).toBe("Merged");
    expect(plain(info.get("Author"))).toBe("Ada");
    expect(plain(info.get("Subject"))).toBe("Merging");
    expect(plain(info.get("Producer"))).toBe("docu-edit");
    expect(plain(info.get("ModDate"))).toBe("D:20261024120000Z");
  });

  it("writes no date without plan.modified, so even an undated plan is twice-identical", async () => {
    const source = await parsePdf(PAGES);
    const a = await assemblePdf(plan(allPages("pages", source)), { between: noYield });
    const b = await assemblePdf(plan(allPages("pages", source)), { between: noYield });
    expect(equalBytes(a.bytes, b.bytes)).toBe(true);
    expect(latin1(a.bytes)).not.toContain("/ModDate");
  });

  it("is given deep-frozen inputs and mutates none of them", async () => {
    const docs = await Promise.all([PAGES, OUTLINES, FORMS].map((bytes) => parsePdf(bytes)));
    const before = docs.map((doc) => ({
      bytes: doc.bytes.slice(),
      objects: JSON.stringify(Array.from(doc.objects, ([num, value]) => [num, plain(value)])),
    }));
    const frozen = deepFreeze(
      plan([
        ...allPages("outlines", docs[1]),
        { source: "pages", doc: docs[0], pageIndex: 1, rotateBy: 270 },
        ...allPages("forms", docs[2]),
      ]),
    );
    const { bytes } = await assemblePdf(frozen, { between: noYield });
    expect((await parsePdf(bytes)).pages).toHaveLength(5);
    docs.forEach((doc, index) => {
      expect(equalBytes(doc.bytes, before[index].bytes)).toBe(true);
      expect(JSON.stringify(Array.from(doc.objects, ([num, value]) => [num, plain(value)]))).toBe(before[index].objects);
    });
  });
});

describe("checkAssembledPdf — the re-parse self-check", () => {
  it("returns the re-parsed document when it matches the plan", async () => {
    const source = await parsePdf(PAGES);
    const merged = plan([{ source: "pages", doc: source, pageIndex: 1, rotateBy: 90 }, { source: "pages", doc: source, pageIndex: 0 }]);
    const { bytes } = await assemblePdf(merged, { between: noYield });
    const doc = await checkAssembledPdf(bytes, plannedPdfPages(merged));
    expect(doc.pages.map((page) => page.rotate)).toEqual([180, 0]);
  });

  it("fails loudly on output that does not read back, or reads back as something else", async () => {
    const source = await parsePdf(PAGES);
    const merged = plan(allPages("pages", source));
    const expected = plannedPdfPages(merged);
    await expect(checkAssembledPdf(ascii("junk"), expected)).rejects.toThrow(/^the merged PDF does not read back \(.+\), so nothing was kept\.$/);
    await expect(checkAssembledPdf(PAGES, expected.slice(0, 2))).rejects.toThrow(
      "the merged PDF has 3 pages instead of 2, so nothing was kept.",
    );
    const turned = expected.map((page, index) => (index === 2 ? { ...page, rotate: 90 as const } : page));
    await expect(checkAssembledPdf(PAGES, turned)).rejects.toThrow(
      "page 3 of the merged PDF is 720 × 540 pt, rotated 0° instead of 720 × 540 pt, rotated 90°, so nothing was kept.",
    );
  });
});
