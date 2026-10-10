// A from-scratch PDF 1.7 writer for client-side export (SPEC-034). Builds a
// searchable, selectable document: every text item becomes a real Tj
// operator against an embedded Liberation face (CIDFontType2 via Identity-H,
// with /W widths and a /ToUnicode CMap so extraction and search work),
// hyperlinks become /Link annotations, Heading 1–3 become a nested
// /Outlines tree, and the document title lands in the Info dictionary.
// package.json is frozen for feature work, so no PDF library — the writer
// emits objects, xref and trailer itself. Framework-free per src/lib.
import { deflate } from "../core/deflate";
import type { FaceKey } from "./fontMap";
import type { SfntFont } from "./sfnt";

/** One positioned text fragment (a single rendered line of one styled run).
 * Coordinates are in PDF points measured from the page's TOP-left — the
 * builder flips into PDF's bottom-up space. */
export interface PdfTextItem {
  text: string;
  xPt: number;
  /** Baseline y from the page top. */
  yPt: number;
  sizePt: number;
  face: FaceKey;
  /** rgb, each 0..1. */
  color: [number, number, number];
}

/** One clickable rectangle: an external URI or a jump to another page. */
export interface PdfLinkItem {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
  href?: string;
  /** 0-based target page index for internal links. */
  destPage?: number;
}

/** One outline entry (a Heading 1–3), in document order. */
export interface PdfHeading {
  /** 1, 2 or 3. */
  level: number;
  title: string;
  yPt: number;
}

/** One placed image. Coordinates match PdfLinkItem — the box's top-left, in
 * points from the page's TOP-left. */
export interface PdfImageItem {
  /** Key into PdfBuildInput.images; the same source placed twice is
   * embedded once. */
  src: string;
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
  /** Painted over the text rather than under it — an image whose wrap puts
   * it in front of text (SPEC-024). */
  onTop?: boolean;
}

/** A decoded image, in exactly the form its PDF stream carries. */
export interface PdfImageData {
  widthPx: number;
  heightPx: number;
  /** "DCTDecode" for JPEG bytes passed through, "FlateDecode" for
   * zlib-compressed 8-bit samples. */
  filter: "DCTDecode" | "FlateDecode";
  colorSpace: "DeviceRGB" | "DeviceGray";
  data: Uint8Array;
  /** zlib-compressed 8-bit alpha, present only when the image is not fully
   * opaque; becomes the XObject's /SMask. */
  alpha?: Uint8Array;
}

/** One drawn sub-path of a vector item, in points from the page's TOP-left
 * (the builder flips into PDF's bottom-up space, exactly as for text and
 * images). A segment with `c1`/`c2` is a cubic, otherwise a straight line. */
export interface PdfPathSubpath {
  start: { xPt: number; yPt: number };
  segments: Array<{
    toXPt: number;
    toYPt: number;
    c1XPt?: number;
    c1YPt?: number;
    c2XPt?: number;
    c2YPt?: number;
  }>;
  closed?: boolean;
}

/** One vector path — how a shape the page paints as SVG rather than as a
 * bitmap reaches the PDF (SPEC-070's connectors): real stroked and filled
 * geometry, so it stays crisp at any zoom instead of being rasterized.
 * A path with neither `stroke` nor `fill` paints nothing and is skipped. */
export interface PdfPathItem {
  subpaths: PdfPathSubpath[];
  /** rgb 0..1 plus the line width in points, an optional dash pattern (also
   * in points, PDF's `d` operator array), and the cap/join styles as PDF's
   * own codes (cap 0 butt / 1 round / 2 square; join 0 miter / 1 round /
   * 2 bevel) — the extractor reads them off the painted element so an elbow's
   * corners and a dash's ends look the same in the file as on the page.
   * Absent means PDF's defaults, which are SVG's too (butt, miter). */
  stroke?: {
    color: [number, number, number];
    widthPt: number;
    dashPt?: number[];
    cap?: 0 | 1 | 2;
    join?: 0 | 1 | 2;
  };
  /** rgb 0..1; non-zero winding, matching SVG's default fill-rule for the
   * convex heads this draws. */
  fill?: [number, number, number];
  /** Painted over the text rather than under it, like PdfImageItem.onTop. */
  onTop?: boolean;
}

/** A filled rectangle — SPEC-108's sheet background. Coordinates match
 * PdfImageItem: the box's top-left, in points from the page's TOP-left.
 * Rects paint first, beneath every image, path and text item. */
export interface PdfRectItem {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
  /** "#rrggbb"; anything else paints nothing. */
  fill: string;
}

export interface PdfPageInput {
  widthPt: number;
  heightPt: number;
  items: PdfTextItem[];
  links: PdfLinkItem[];
  headings: PdfHeading[];
  images: PdfImageItem[];
  /** Vector paths (SPEC-070 connectors). Optional: a page input built before
   * this existed draws none. */
  paths?: PdfPathItem[];
  /** Filled rectangles under everything else (SPEC-108). Optional: a page
   * input built before this existed draws none. */
  rects?: PdfRectItem[];
  /** The page's /Rotate — how a viewer turns it for display, clockwise.
   * Absent or 0 writes no entry. SPEC-107's fixture needs a rotated page. */
  rotate?: 0 | 90 | 180 | 270;
}

/** Info entries beyond the title. Absent entries are not written. */
export interface PdfInfoInput {
  author?: string;
  creationDate?: Date;
  modDate?: Date;
}

export interface PdfBuildInput {
  /** Info /Title — the empty string when the document has no title. */
  title: string;
  /** Further Info entries (SPEC-107's fixture records an author and dates). */
  info?: PdfInfoInput;
  pages: PdfPageInput[];
  /** Parsed faces available for embedding; only faces the items actually
   * use are written into the file. */
  fonts: Map<FaceKey, SfntFont>;
  /** Decoded images by src. A placed image missing from this map (one that
   * could not be decoded) is skipped, so the rest of the page still
   * exports. */
  images?: Map<string, PdfImageData>;
}

export const PDF_MIME_TYPE = "application/pdf";

/** Mirrors docxFileName/odtFileName: the trimmed title plus the extension.
 * An untitled document still downloads under the model's default title —
 * only the PDF *metadata* title is allowed to be empty (AC-32). */
export function pdfFileName(title: string, fallback: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : fallback}.pdf`;
}

const ASCII_PRINTABLE = /^[\x20-\x7e]*$/;

function fmt(value: number): string {
  const rounded = Math.round(value * 100) / 100;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

/** A text string operand: literal when printable ASCII, UTF-16BE hex (with
 * BOM) otherwise — parentheses and backslashes escaped either way. */
function pdfTextString(text: string): string {
  if (ASCII_PRINTABLE.test(text)) {
    return `(${text.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)")})`;
  }
  let hex = "FEFF";
  for (const char of text) {
    const code = char.charCodeAt(0);
    hex += code.toString(16).padStart(4, "0").toUpperCase();
    if (char.length > 1) {
      hex += char.charCodeAt(1).toString(16).padStart(4, "0").toUpperCase();
    }
  }
  return `<${hex}>`;
}

function pad(value: number, width = 2): string {
  return String(value).padStart(width, "0");
}

/** A date operand (ISO 32000 7.9.4), in UTC: (D:YYYYMMDDHHmmSSZ). */
function pdfDateString(date: Date): string {
  return (
    `(D:${pad(date.getUTCFullYear(), 4)}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}` +
    `${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z)`
  );
}

/** A URI operand: control characters stripped, delimiters escaped, so a
 * crafted href can never break out of the string and inject objects. */
function pdfUriString(uri: string): string {
  const cleaned = Array.from(uri)
    .filter((char) => char.charCodeAt(0) >= 32 && char.charCodeAt(0) !== 127)
    .join("");
  return `(${cleaned.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)")})`;
}

type ObjectBody = Array<string | Uint8Array>;

class PdfObjects {
  private bodies: Array<ObjectBody | null> = [];

  alloc(): number {
    this.bodies.push(null);
    return this.bodies.length;
  }

  set(id: number, ...body: ObjectBody): void {
    this.bodies[id - 1] = body;
  }

  add(...body: ObjectBody): number {
    const id = this.alloc();
    this.set(id, ...body);
    return id;
  }

  stream(dict: string, bytes: Uint8Array | string): number {
    const data = typeof bytes === "string" ? asciiBytes(bytes) : bytes;
    return this.add(
      `<< ${dict} /Length ${data.length} >>\nstream\n`,
      data,
      "\nendstream",
    );
  }

  /** Serializes header, every object, xref table and trailer. */
  assemble(rootId: number, infoId: number): Uint8Array {
    const chunks: Uint8Array[] = [];
    let position = 0;
    const push = (chunk: Uint8Array) => {
      chunks.push(chunk);
      position += chunk.length;
    };
    // The comment line's high bytes mark the file as binary for transports.
    push(asciiBytes("%PDF-1.7\n"));
    push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
    const offsets: number[] = [];
    this.bodies.forEach((body, index) => {
      offsets.push(position);
      push(asciiBytes(`${index + 1} 0 obj\n`));
      for (const part of body ?? ["null"]) {
        push(typeof part === "string" ? asciiBytes(part) : part);
      }
      push(asciiBytes("\nendobj\n"));
    });
    const xrefStart = position;
    let xref = `xref\n0 ${this.bodies.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets) {
      xref += `${String(offset).padStart(10, "0")} 00000 n \n`;
    }
    xref += `trailer\n<< /Size ${this.bodies.length + 1} /Root ${rootId} 0 R /Info ${infoId} 0 R >>\n`;
    xref += `startxref\n${xrefStart}\n%%EOF\n`;
    push(asciiBytes(xref));
    const out = new Uint8Array(position);
    let cursor = 0;
    for (const chunk of chunks) {
      out.set(chunk, cursor);
      cursor += chunk.length;
    }
    return out;
  }
}

function asciiBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    bytes[i] = text.charCodeAt(i) & 0xff;
  }
  return bytes;
}

/** Per-face embedding state: which glyphs the document used, and what
 * Unicode text they stand for (the /ToUnicode CMap's input). */
interface FaceUsage {
  font: SfntFont;
  resourceName: string;
  /** glyph id → the code point it renders (first one seen wins). */
  usedGlyphs: Map<number, number>;
  spaceGlyph: number;
}

function toUnicodeCMap(usage: FaceUsage): string {
  const entries = Array.from(usage.usedGlyphs.entries()).toSorted((a, b) => a[0] - b[0]);
  let body = "";
  for (let start = 0; start < entries.length; start += 100) {
    const batch = entries.slice(start, start + 100);
    body += `${batch.length} beginbfchar\n`;
    for (const [glyph, code] of batch) {
      const hex =
        code > 0xffff
          ? // Surrogate pair for astral code points.
            (0xd800 + ((code - 0x10000) >> 10)).toString(16).padStart(4, "0") +
            (0xdc00 + ((code - 0x10000) & 0x3ff)).toString(16).padStart(4, "0")
          : code.toString(16).padStart(4, "0");
      body += `<${glyph.toString(16).padStart(4, "0")}> <${hex}>\n`;
    }
    body += "endbfchar\n";
  }
  return (
    "/CIDInit /ProcSet findresource begin\n" +
    "12 dict begin\n" +
    "begincmap\n" +
    "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n" +
    "/CMapName /Adobe-Identity-UCS def\n" +
    "/CMapType 2 def\n" +
    "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n" +
    body +
    "endcmap\n" +
    "CMapName currentdict /CMap defineresource pop\n" +
    "end\nend"
  );
}

function descriptorFlags(face: FaceKey): number {
  let flags = 32; // Nonsymbolic
  // Cambria's twin (Caladea) is a serif; Calibri's twin (Carlito) is sans.
  if (face.startsWith("serif") || face.startsWith("cambria")) {
    flags |= 2;
  }
  if (face.startsWith("mono")) {
    flags |= 1;
  }
  if (face.endsWith("italic")) {
    flags |= 64;
  }
  return flags;
}

/** Writes one face's font object graph; returns the Type0 font's id.
 *
 * The whole font program is embedded (nothing here subsets glyphs), so it is
 * by far the largest thing in a short document — four raw Liberation faces are
 * 1.6 MB, and the Carlito set most .docx files resolve to is 2.7 MB, dwarfing
 * a one-page PDF's actual content. /FontFile2 is a normal stream and takes a
 * filter like any other, so deflate it: /Length1 stays the UNCOMPRESSED
 * length, which is what a reader uses to reconstruct the sfnt. */
async function writeFont(objects: PdfObjects, face: FaceKey, usage: FaceUsage): Promise<number> {
  const { font } = usage;
  const scale = 1000 / font.unitsPerEm;
  const fileId = objects.stream(
    `/Length1 ${font.bytes.length} /Filter /FlateDecode`,
    await deflate(font.bytes),
  );
  const descriptorId = objects.add(
    `<< /Type /FontDescriptor /FontName /${font.postScriptName}` +
      ` /Flags ${descriptorFlags(face)}` +
      ` /FontBBox [${font.bbox.map((v) => fmt(v * scale)).join(" ")}]` +
      ` /ItalicAngle ${fmt(font.italicAngle)}` +
      ` /Ascent ${fmt(font.ascent * scale)}` +
      ` /Descent ${fmt(font.descent * scale)}` +
      ` /CapHeight ${fmt(font.capHeight * scale)}` +
      ` /StemV ${face.includes("bold") ? 120 : 80}` +
      ` /FontFile2 ${fileId} 0 R >>`,
  );
  const widths = Array.from(usage.usedGlyphs.keys())
    .toSorted((a, b) => a - b)
    .map((glyph) => `${glyph} [${fmt(font.advanceOf(glyph) * scale)}]`)
    .join(" ");
  const cidFontId = objects.add(
    `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${font.postScriptName}` +
      ` /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>` +
      ` /FontDescriptor ${descriptorId} 0 R /DW 1000 /W [ ${widths} ] /CIDToGIDMap /Identity >>`,
  );
  const toUnicodeId = objects.stream("", toUnicodeCMap(usage));
  return objects.add(
    `<< /Type /Font /Subtype /Type0 /BaseFont /${font.postScriptName}` +
      ` /Encoding /Identity-H /DescendantFonts [${cidFontId} 0 R] /ToUnicode ${toUnicodeId} 0 R >>`,
  );
}

/** Encodes a text item's characters as Identity-H glyph ids, recording the
 * glyphs used. Unmapped characters fall back to the space glyph so layout
 * survives even when the face lacks a script. */
function encodeText(usage: FaceUsage, text: string): string {
  let hex = "";
  for (const char of text) {
    const code = char.codePointAt(0) ?? 32;
    let glyph = usage.font.glyphFor(code);
    if (glyph === 0) {
      glyph = usage.spaceGlyph;
    } else if (!usage.usedGlyphs.has(glyph)) {
      usage.usedGlyphs.set(glyph, code);
    }
    hex += glyph.toString(16).padStart(4, "0");
  }
  return hex;
}

interface OutlineEntry {
  heading: PdfHeading;
  pageIndex: number;
}

interface OutlineNode {
  id: number;
  level: number;
  parentId: number;
  children: OutlineNode[];
  entry: OutlineEntry;
}

const countDescendants = (node: OutlineNode): number =>
  node.children.reduce((sum, child) => sum + 1 + countDescendants(child), 0);

/** Writes the /Outlines tree from the flat heading list, nesting each entry
 * under the closest preceding entry of a shallower level. Returns the root
 * outline dictionary's id, or 0 when there are no headings. */
function writeOutlines(
  objects: PdfObjects,
  entries: OutlineEntry[],
  pageIds: number[],
  pageHeights: number[],
): number {
  if (entries.length === 0) {
    return 0;
  }
  const rootId = objects.alloc();
  const roots: OutlineNode[] = [];
  const stack: OutlineNode[] = [];
  for (const entry of entries) {
    const level = Math.min(Math.max(entry.heading.level, 1), 3);
    while (stack.length > 0 && stack[stack.length - 1].level >= level) {
      stack.pop();
    }
    const parent = stack[stack.length - 1];
    const node: OutlineNode = {
      id: objects.alloc(),
      level,
      parentId: parent ? parent.id : rootId,
      children: [],
      entry,
    };
    (parent ? parent.children : roots).push(node);
    stack.push(node);
  }
  let total = 0;
  const writeNode = (node: OutlineNode, prev: OutlineNode | undefined, next: OutlineNode | undefined) => {
    total += 1;
    const page = node.entry.pageIndex;
    const y = fmt(Math.max(0, (pageHeights[page] ?? 0) - node.entry.heading.yPt));
    let dict =
      `<< /Title ${pdfTextString(node.entry.heading.title)} /Parent ${node.parentId} 0 R` +
      ` /Dest [${pageIds[page]} 0 R /XYZ 0 ${y} 0]`;
    if (prev) {
      dict += ` /Prev ${prev.id} 0 R`;
    }
    if (next) {
      dict += ` /Next ${next.id} 0 R`;
    }
    if (node.children.length > 0) {
      dict +=
        ` /First ${node.children[0].id} 0 R /Last ${node.children[node.children.length - 1].id} 0 R` +
        ` /Count ${countDescendants(node)}`;
    }
    objects.set(node.id, dict + " >>");
    node.children.forEach((child, index) =>
      writeNode(child, node.children[index - 1], node.children[index + 1]),
    );
  };
  roots.forEach((node, index) => writeNode(node, roots[index - 1], roots[index + 1]));
  objects.set(
    rootId,
    `<< /Type /Outlines /First ${roots[0].id} 0 R /Last ${roots[roots.length - 1].id} 0 R /Count ${total} >>`,
  );
  return rootId;
}

/** Builds the complete PDF file. Throws when text exists but no font was
 * provided to embed — the exporter turns that into its failure toast. */
export async function buildPdf(input: PdfBuildInput): Promise<Uint8Array> {
  const objects = new PdfObjects();
  const catalogId = objects.alloc();
  const pagesId = objects.alloc();
  const infoId = objects.alloc();
  const pageIds = input.pages.map(() => objects.alloc());

  // Face usage is discovered while encoding page content, so content
  // streams are written first and font objects after.
  const usages = new Map<FaceKey, FaceUsage>();
  const fallbackFace = input.fonts.keys().next();
  const usageFor = (face: FaceKey): FaceUsage => {
    let key = face;
    if (!input.fonts.has(key)) {
      if (fallbackFace.done) {
        throw new Error("no fonts available for embedding");
      }
      key = fallbackFace.value;
    }
    let usage = usages.get(key);
    if (!usage) {
      const font = input.fonts.get(key) as SfntFont;
      const spaceGlyph = font.glyphFor(32);
      usage = {
        font,
        resourceName: `F${usages.size + 1}`,
        usedGlyphs: new Map(spaceGlyph > 0 ? [[spaceGlyph, 32]] : []),
        spaceGlyph,
      };
      usages.set(key, usage);
    }
    return usage;
  };

  // Image XObjects are written before the content streams that reference
  // them, so every draw already knows its resource name. One object per
  // distinct source, however many pages place it.
  const imageNames = new Map<string, string>();
  const imageIds = new Map<string, number>();
  for (const page of input.pages) {
    for (const item of page.images) {
      const image = input.images?.get(item.src);
      if (!image || imageNames.has(item.src)) {
        continue;
      }
      const size = `/Width ${image.widthPx} /Height ${image.heightPx}`;
      let dict =
        `/Type /XObject /Subtype /Image ${size} /ColorSpace /${image.colorSpace}` +
        ` /BitsPerComponent 8 /Filter /${image.filter}`;
      if (image.alpha) {
        const smaskId = objects.stream(
          `/Type /XObject /Subtype /Image ${size} /ColorSpace /DeviceGray` +
            ` /BitsPerComponent 8 /Filter /FlateDecode`,
          image.alpha,
        );
        dict += ` /SMask ${smaskId} 0 R`;
      }
      const name = `Im${imageNames.size + 1}`;
      imageNames.set(item.src, name);
      imageIds.set(name, objects.stream(dict, image.data));
    }
  }

  /** The draw operators for one page's images, in the requested layer: the
   * unit image square is scaled to the placed box and moved into PDF's
   * bottom-up space. */
  const drawImages = (page: PdfPageInput, onTop: boolean): string => {
    let ops = "";
    for (const item of page.images) {
      const name = imageNames.get(item.src);
      if (!name || Boolean(item.onTop) !== onTop) {
        continue;
      }
      const y = page.heightPt - item.yPt - item.heightPt;
      ops +=
        `q ${fmt(item.widthPt)} 0 0 ${fmt(item.heightPt)} ${fmt(item.xPt)} ${fmt(y)} cm` +
        ` /${name} Do Q\n`;
    }
    return ops;
  };

  /** The draw operators for one page's vector paths, in the requested
   * layer. Coordinates arrive measured from the page top, so every y is
   * flipped into PDF space here — the one place that conversion happens for
   * paths. */
  const drawPaths = (page: PdfPageInput, onTop: boolean): string => {
    let ops = "";
    for (const item of page.paths ?? []) {
      if (Boolean(item.onTop) !== onTop || (!item.stroke && !item.fill)) {
        continue;
      }
      const flip = (yPt: number) => page.heightPt - yPt;
      let geometry = "";
      for (const subpath of item.subpaths) {
        if (subpath.segments.length === 0 && !subpath.closed) {
          continue;
        }
        geometry += `${fmt(subpath.start.xPt)} ${fmt(flip(subpath.start.yPt))} m\n`;
        for (const segment of subpath.segments) {
          if (segment.c1XPt !== undefined && segment.c2XPt !== undefined) {
            geometry +=
              `${fmt(segment.c1XPt)} ${fmt(flip(segment.c1YPt ?? 0))} ` +
              `${fmt(segment.c2XPt)} ${fmt(flip(segment.c2YPt ?? 0))} ` +
              `${fmt(segment.toXPt)} ${fmt(flip(segment.toYPt))} c\n`;
          } else {
            geometry += `${fmt(segment.toXPt)} ${fmt(flip(segment.toYPt))} l\n`;
          }
        }
        if (subpath.closed) {
          geometry += "h\n";
        }
      }
      if (geometry.length === 0) {
        continue;
      }
      let state = "";
      if (item.fill) {
        state += `${item.fill.map(fmt).join(" ")} rg `;
      }
      if (item.stroke) {
        state +=
          `${item.stroke.color.map(fmt).join(" ")} RG ` +
          `${fmt(Math.max(0, item.stroke.widthPt))} w ` +
          `[${(item.stroke.dashPt ?? []).map(fmt).join(" ")}] 0 d ` +
          `${item.stroke.cap ?? 0} J ${item.stroke.join ?? 0} j `;
      }
      const paint = item.fill && item.stroke ? "B" : item.fill ? "f" : "S";
      ops += `q ${state}${geometry}${paint} Q\n`;
    }
    return ops;
  };

  /** The fill operators for one page's rects, flipped into PDF space. */
  const drawRects = (page: PdfPageInput): string => {
    let ops = "";
    for (const rect of page.rects ?? []) {
      const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(rect.fill);
      if (!hex) {
        continue;
      }
      const rgb = hex.slice(1).map((channel) => fmt(parseInt(channel, 16) / 255));
      const y = page.heightPt - rect.yPt - rect.heightPt;
      ops +=
        `q ${rgb.join(" ")} rg ${fmt(rect.xPt)} ${fmt(y)} ${fmt(rect.widthPt)} ${fmt(rect.heightPt)} re f Q\n`;
    }
    return ops;
  };

  const outlineEntries: OutlineEntry[] = [];
  const pageAnnotIds: number[][] = [];
  const contentIds: number[] = [];

  // A plain loop, not forEach: the content stream is deflated per page, and
  // the pages' content objects must be allocated in page order.
  for (const [pageIndex, page] of input.pages.entries()) {
    // Behind-the-text images first, then the text, then the images that sit
    // in front of it — PDF paints in content-stream order.
    let content = drawRects(page) + drawImages(page, false) + drawPaths(page, false);
    for (const item of page.items) {
      if (item.text.length === 0) {
        continue;
      }
      const usage = usageFor(item.face);
      const [r, g, b] = item.color;
      content +=
        `BT /${usage.resourceName} ${fmt(item.sizePt)} Tf ` +
        `${fmt(r)} ${fmt(g)} ${fmt(b)} rg ` +
        `1 0 0 1 ${fmt(item.xPt)} ${fmt(page.heightPt - item.yPt)} Tm ` +
        `<${encodeText(usage, item.text)}> Tj ET\n`;
    }
    content += drawImages(page, true) + drawPaths(page, true);
    // Page content is operator text — highly compressible, and the bulk of
    // any document long enough for the fonts not to dominate.
    contentIds.push(objects.stream("/Filter /FlateDecode", await deflate(asciiBytes(content))));
    const annotIds: number[] = [];
    for (const link of page.links) {
      const x1 = fmt(link.xPt);
      const y2 = fmt(page.heightPt - link.yPt);
      const x2 = fmt(link.xPt + link.widthPt);
      const y1 = fmt(page.heightPt - link.yPt - link.heightPt);
      let action = "";
      if (link.href !== undefined) {
        action = `/A << /S /URI /URI ${pdfUriString(link.href)} >>`;
      } else if (
        link.destPage !== undefined &&
        link.destPage >= 0 &&
        link.destPage < pageIds.length
      ) {
        const targetHeight = input.pages[link.destPage].heightPt;
        action = `/Dest [${pageIds[link.destPage]} 0 R /XYZ 0 ${fmt(targetHeight)} 0]`;
      } else {
        continue;
      }
      annotIds.push(
        objects.add(
          `<< /Type /Annot /Subtype /Link /Rect [${x1} ${y1} ${x2} ${y2}] /Border [0 0 0] ${action} >>`,
        ),
      );
    }
    pageAnnotIds.push(annotIds);
    for (const heading of page.headings) {
      const title = heading.title.trim();
      if (title.length > 0) {
        outlineEntries.push({ heading: { ...heading, title }, pageIndex });
      }
    }
  }

  const fontIds = new Map<string, number>();
  for (const [face, usage] of usages) {
    // Sequential, not Promise.all: each writeFont allocates object ids as it
    // goes, and the descriptor/CIDFont/Type0 graph has to stay contiguous.
    fontIds.set(usage.resourceName, await writeFont(objects, face, usage));
  }
  const fontResources =
    fontIds.size > 0
      ? `/Font << ${Array.from(fontIds.entries())
          .map(([name, id]) => `/${name} ${id} 0 R`)
          .join(" ")} >>`
      : "";
  const imageResources =
    imageIds.size > 0
      ? `/XObject << ${Array.from(imageIds.entries())
          .map(([name, id]) => `/${name} ${id} 0 R`)
          .join(" ")} >>`
      : "";
  const resourcesId = objects.add(
    `<< ${[fontResources, imageResources].filter((part) => part.length > 0).join(" ")} >>`,
  );

  input.pages.forEach((page, pageIndex) => {
    const annots =
      pageAnnotIds[pageIndex].length > 0
        ? ` /Annots [${pageAnnotIds[pageIndex].map((id) => `${id} 0 R`).join(" ")}]`
        : "";
    const rotate = page.rotate ? ` /Rotate ${page.rotate}` : "";
    objects.set(
      pageIds[pageIndex],
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${fmt(page.widthPt)} ${fmt(page.heightPt)}]` +
        `${rotate} /Resources ${resourcesId} 0 R /Contents ${contentIds[pageIndex]} 0 R${annots} >>`,
    );
  });

  const outlinesId = writeOutlines(
    objects,
    outlineEntries,
    pageIds,
    input.pages.map((page) => page.heightPt),
  );

  objects.set(
    pagesId,
    `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`,
  );
  objects.set(
    catalogId,
    `<< /Type /Catalog /Pages ${pagesId} 0 R` +
      (outlinesId > 0 ? ` /Outlines ${outlinesId} 0 R /PageMode /UseOutlines` : "") +
      " >>",
  );
  const info = input.info ?? {};
  objects.set(
    infoId,
    `<< /Title ${pdfTextString(input.title)}` +
      (info.author !== undefined ? ` /Author ${pdfTextString(info.author)}` : "") +
      " /Producer (docu-edit)" +
      (info.creationDate ? ` /CreationDate ${pdfDateString(info.creationDate)}` : "") +
      (info.modDate ? ` /ModDate ${pdfDateString(info.modDate)}` : "") +
      " >>",
  );

  return objects.assemble(catalogId, infoId);
}
