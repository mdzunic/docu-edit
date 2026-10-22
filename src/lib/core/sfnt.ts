// A minimal TrueType (sfnt) parser for PDF font embedding (SPEC-034).
// Reads exactly the tables the PDF builder needs to embed a face as a
// CIDFontType2 font: character→glyph mapping (cmap format 4), advance
// widths (hhea/hmtx), design metrics (head, OS/2, post) and the PostScript
// name. The raw bytes are kept verbatim — the whole file is embedded as
// the /FontFile2 stream, so nothing here rewrites or subsets the font.
// Framework-free per the src/lib convention.

/** The parsed face: everything buildPdf needs to embed and use the font. */
export interface SfntFont {
  /** The complete font file, embedded verbatim as /FontFile2. */
  bytes: Uint8Array;
  /** Design units per em (glyph-space values scale by 1000/unitsPerEm). */
  unitsPerEm: number;
  /** hhea ascender/descender in font units (descender is negative). */
  ascent: number;
  descent: number;
  /** head glyph bounding box [xMin, yMin, xMax, yMax] in font units. */
  bbox: [number, number, number, number];
  /** post italicAngle in degrees (0 for upright faces). */
  italicAngle: number;
  /** OS/2 sCapHeight in font units (approximated when absent). */
  capHeight: number;
  /** name table PostScript name (nameID 6), for /BaseFont. */
  postScriptName: string;
  /** The glyph id for a Unicode code point, or 0 (.notdef) when unmapped. */
  glyphFor(codePoint: number): number;
  /** A glyph's advance width in font units. */
  advanceOf(glyphId: number): number;
}

interface TableDirectory {
  [tag: string]: { offset: number; length: number };
}

function readTables(view: DataView): TableDirectory {
  const version = view.getUint32(0);
  // 0x00010000 = TrueType outlines, 0x74727565 = 'true' (Apple TrueType).
  if (version !== 0x00010000 && version !== 0x74727565) {
    throw new Error("not a TrueType font file");
  }
  const numTables = view.getUint16(4);
  const tables: TableDirectory = {};
  for (let i = 0; i < numTables; i += 1) {
    const base = 12 + i * 16;
    const tag = String.fromCharCode(
      view.getUint8(base),
      view.getUint8(base + 1),
      view.getUint8(base + 2),
      view.getUint8(base + 3),
    );
    tables[tag] = { offset: view.getUint32(base + 8), length: view.getUint32(base + 12) };
  }
  return tables;
}

function requireTable(tables: TableDirectory, tag: string): { offset: number; length: number } {
  const table = tables[tag];
  if (!table) {
    throw new Error(`font is missing its ${tag} table`);
  }
  return table;
}

/** A cmap format 4 subtable compiled into a segment lookup. */
interface CmapSegments {
  endCodes: Uint16Array;
  startCodes: Uint16Array;
  idDeltas: Int16Array;
  idRangeOffsets: Uint16Array;
  /** Byte offset of the idRangeOffset array (glyphIdArray reads need it). */
  rangeOffsetBase: number;
  view: DataView;
}

function parseCmap(view: DataView, cmapOffset: number): CmapSegments {
  const numSubtables = view.getUint16(cmapOffset + 2);
  let best = -1;
  for (let i = 0; i < numSubtables; i += 1) {
    const base = cmapOffset + 4 + i * 8;
    const platform = view.getUint16(base);
    const encoding = view.getUint16(base + 2);
    const offset = view.getUint32(base + 4);
    // Windows Unicode BMP (3,1) is what Liberation ships; Unicode platform
    // (0, any) is an equivalent fallback.
    if ((platform === 3 && encoding === 1) || platform === 0) {
      best = cmapOffset + offset;
      if (platform === 3) {
        break;
      }
    }
  }
  if (best < 0) {
    throw new Error("font has no Unicode cmap subtable");
  }
  if (view.getUint16(best) !== 4) {
    throw new Error("font cmap subtable is not format 4");
  }
  const segCount = view.getUint16(best + 6) / 2;
  const endCodes = new Uint16Array(segCount);
  const startCodes = new Uint16Array(segCount);
  const idDeltas = new Int16Array(segCount);
  const idRangeOffsets = new Uint16Array(segCount);
  const endBase = best + 14;
  const startBase = endBase + segCount * 2 + 2;
  const deltaBase = startBase + segCount * 2;
  const rangeOffsetBase = deltaBase + segCount * 2;
  for (let i = 0; i < segCount; i += 1) {
    endCodes[i] = view.getUint16(endBase + i * 2);
    startCodes[i] = view.getUint16(startBase + i * 2);
    idDeltas[i] = view.getInt16(deltaBase + i * 2);
    idRangeOffsets[i] = view.getUint16(rangeOffsetBase + i * 2);
  }
  return { endCodes, startCodes, idDeltas, idRangeOffsets, rangeOffsetBase, view };
}

function glyphFromCmap(cmap: CmapSegments, codePoint: number): number {
  if (codePoint > 0xffff) {
    return 0; // format 4 covers the BMP only
  }
  const { endCodes, startCodes, idDeltas, idRangeOffsets, rangeOffsetBase, view } = cmap;
  // Binary search the first segment whose endCode >= codePoint.
  let lo = 0;
  let hi = endCodes.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (endCodes[mid] < codePoint) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  if (startCodes[lo] > codePoint) {
    return 0;
  }
  const rangeOffset = idRangeOffsets[lo];
  if (rangeOffset === 0) {
    return (codePoint + idDeltas[lo]) & 0xffff;
  }
  const glyphAddress = rangeOffsetBase + lo * 2 + rangeOffset + (codePoint - startCodes[lo]) * 2;
  const glyph = view.getUint16(glyphAddress);
  return glyph === 0 ? 0 : (glyph + idDeltas[lo]) & 0xffff;
}

function parsePostScriptName(view: DataView, nameOffset: number): string {
  const count = view.getUint16(nameOffset + 2);
  const stringsBase = nameOffset + view.getUint16(nameOffset + 4);
  for (let i = 0; i < count; i += 1) {
    const base = nameOffset + 6 + i * 12;
    if (view.getUint16(base + 6) !== 6) {
      continue; // nameID 6 = PostScript name
    }
    const platform = view.getUint16(base);
    const length = view.getUint16(base + 8);
    const offset = stringsBase + view.getUint16(base + 10);
    let name = "";
    if (platform === 3) {
      // UTF-16BE; PostScript names are ASCII, so the low bytes suffice.
      for (let b = 0; b < length; b += 2) {
        name += String.fromCharCode(view.getUint16(offset + b));
      }
    } else {
      for (let b = 0; b < length; b += 1) {
        name += String.fromCharCode(view.getUint8(offset + b));
      }
    }
    // /BaseFont is a PDF name: keep it to the safe ASCII subset.
    const cleaned = name.replace(/[^!-~]/g, "").replace(/[()<>[\]{}/%#]/g, "");
    if (cleaned.length > 0) {
      return cleaned;
    }
  }
  return "EmbeddedFont";
}

/** Parses a TrueType font file. Throws on anything that is not a usable
 * TrueType font — the PDF exporter surfaces that as its failure reason. */
export function parseSfnt(bytes: Uint8Array): SfntFont {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tables = readTables(view);
  const head = requireTable(tables, "head");
  const hhea = requireTable(tables, "hhea");
  const maxp = requireTable(tables, "maxp");
  const hmtx = requireTable(tables, "hmtx");
  const cmapTable = requireTable(tables, "cmap");

  const unitsPerEm = view.getUint16(head.offset + 18);
  const bbox: [number, number, number, number] = [
    view.getInt16(head.offset + 36),
    view.getInt16(head.offset + 38),
    view.getInt16(head.offset + 40),
    view.getInt16(head.offset + 42),
  ];
  const ascent = view.getInt16(hhea.offset + 4);
  const descent = view.getInt16(hhea.offset + 6);
  const numberOfHMetrics = view.getUint16(hhea.offset + 34);
  const numGlyphs = view.getUint16(maxp.offset + 4);

  const post = tables.post;
  // post italicAngle is a 16.16 fixed number.
  const italicAngle = post ? view.getInt32(post.offset + 4) / 65536 : 0;

  const os2 = tables["OS/2"];
  let capHeight = Math.round(ascent * 0.7);
  if (os2 && view.getUint16(os2.offset) >= 2 && os2.length >= 90) {
    const sCapHeight = view.getInt16(os2.offset + 88);
    if (sCapHeight > 0) {
      capHeight = sCapHeight;
    }
  }

  const name = tables.name;
  const postScriptName = name ? parsePostScriptName(view, name.offset) : "EmbeddedFont";

  const cmap = parseCmap(view, cmapTable.offset);

  function advanceOf(glyphId: number): number {
    if (numberOfHMetrics === 0) {
      return unitsPerEm;
    }
    const index = Math.min(Math.max(glyphId, 0), Math.min(numberOfHMetrics, numGlyphs) - 1);
    return view.getUint16(hmtx.offset + index * 4);
  }

  return {
    bytes,
    unitsPerEm,
    ascent,
    descent,
    bbox,
    italicAngle,
    capHeight,
    postScriptName,
    glyphFor: (codePoint) => glyphFromCmap(cmap, codePoint),
    advanceOf,
  };
}
