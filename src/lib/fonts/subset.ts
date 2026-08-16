// SPEC-048 — a conservative TrueType subsetter for embed-on-save (AC-15).
// Rather than renumber glyphs (which forces a cmap/hmtx/post rebuild and is
// easy to get subtly wrong), it keeps the glyph count, cmap and metrics
// intact and empties the outlines of glyphs the document never uses — the
// bulk of a font's bytes are glyph outlines, so this still shrinks the file
// substantially while staying trivially valid. Composite-glyph components of
// kept glyphs are kept too. Framework-free per the src/lib convention.

interface Table {
  tag: string;
  data: Uint8Array;
}

function readTableDirectory(bytes: Uint8Array): { version: number; tables: Map<string, Table> } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const version = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Table>();
  for (let i = 0; i < numTables; i += 1) {
    const base = 12 + i * 16;
    const tag = String.fromCharCode(
      view.getUint8(base),
      view.getUint8(base + 1),
      view.getUint8(base + 2),
      view.getUint8(base + 3),
    );
    const offset = view.getUint32(base + 8);
    const length = view.getUint32(base + 12);
    tables.set(tag, { tag, data: bytes.subarray(offset, offset + length) });
  }
  return { version, tables };
}

/** Maps requested code points to glyph ids via the font's best Unicode cmap
 * (format 12 preferred, then 4). Unmapped points are skipped. */
function glyphsForCodepoints(cmap: Uint8Array, codepoints: Iterable<number>): Set<number> {
  const view = new DataView(cmap.buffer, cmap.byteOffset, cmap.byteLength);
  const numSubtables = view.getUint16(2);
  let best4 = -1;
  let best12 = -1;
  for (let i = 0; i < numSubtables; i += 1) {
    const platform = view.getUint16(4 + i * 8);
    const encoding = view.getUint16(4 + i * 8 + 2);
    const offset = view.getUint32(4 + i * 8 + 4);
    const format = view.getUint16(offset);
    const unicode = platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10));
    if (!unicode) {
      continue;
    }
    if (format === 12) {
      best12 = offset;
    } else if (format === 4) {
      best4 = offset;
    }
  }
  const result = new Set<number>();
  const points = Array.from(codepoints);
  if (best12 >= 0) {
    const nGroups = view.getUint32(best12 + 12);
    for (const cp of points) {
      for (let g = 0; g < nGroups; g += 1) {
        const base = best12 + 16 + g * 12;
        const start = view.getUint32(base);
        const end = view.getUint32(base + 4);
        if (cp >= start && cp <= end) {
          result.add(view.getUint32(base + 8) + (cp - start));
          break;
        }
      }
    }
    return result;
  }
  if (best4 >= 0) {
    const segX2 = view.getUint16(best4 + 6);
    const segCount = segX2 / 2;
    const endBase = best4 + 14;
    const startBase = endBase + segX2 + 2;
    const deltaBase = startBase + segX2;
    const rangeBase = deltaBase + segX2;
    for (const cp of points) {
      for (let s = 0; s < segCount; s += 1) {
        const end = view.getUint16(endBase + s * 2);
        if (cp > end) {
          continue;
        }
        const start = view.getUint16(startBase + s * 2);
        if (cp < start) {
          break;
        }
        const delta = view.getUint16(deltaBase + s * 2);
        const rangeOffset = view.getUint16(rangeBase + s * 2);
        let gid: number;
        if (rangeOffset === 0) {
          gid = (cp + delta) & 0xffff;
        } else {
          const gIndex = rangeBase + s * 2 + rangeOffset + (cp - start) * 2;
          gid = view.getUint16(gIndex);
          if (gid !== 0) {
            gid = (gid + delta) & 0xffff;
          }
        }
        if (gid !== 0) {
          result.add(gid);
        }
        break;
      }
    }
  }
  return result;
}

/** Reads loca into absolute glyph byte ranges. */
function readLoca(loca: Uint8Array, longFormat: boolean, numGlyphs: number): number[] {
  const view = new DataView(loca.buffer, loca.byteOffset, loca.byteLength);
  const offsets: number[] = [];
  for (let i = 0; i <= numGlyphs; i += 1) {
    offsets.push(longFormat ? view.getUint32(i * 4) : view.getUint16(i * 2) * 2);
  }
  return offsets;
}

/** Adds the component glyphs a composite glyph references (transitively). */
function addComposites(glyf: Uint8Array, offsets: number[], used: Set<number>): void {
  const view = new DataView(glyf.buffer, glyf.byteOffset, glyf.byteLength);
  const queue = Array.from(used);
  while (queue.length > 0) {
    const gid = queue.pop() as number;
    const start = offsets[gid];
    const end = offsets[gid + 1];
    if (end - start < 10) {
      continue;
    }
    const numberOfContours = view.getInt16(start);
    if (numberOfContours >= 0) {
      continue; // simple glyph, no components
    }
    let p = start + 10;
    // Walk the component records.
    for (;;) {
      const flags = view.getUint16(p);
      const componentGid = view.getUint16(p + 2);
      if (!used.has(componentGid)) {
        used.add(componentGid);
        queue.push(componentGid);
      }
      p += 4;
      p += flags & 0x0001 ? 4 : 2; // ARG_1_AND_2_ARE_WORDS
      if (flags & 0x0008) {
        p += 2; // WE_HAVE_A_SCALE
      } else if (flags & 0x0040) {
        p += 4; // WE_HAVE_AN_X_AND_Y_SCALE
      } else if (flags & 0x0080) {
        p += 8; // WE_HAVE_A_TWO_BY_TWO
      }
      if (!(flags & 0x0020)) {
        break; // no MORE_COMPONENTS
      }
    }
  }
}

function pad4(n: number): number {
  return (n + 3) & ~3;
}

function tableChecksum(data: Uint8Array): number {
  let sum = 0;
  const padded = pad4(data.length);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  for (let i = 0; i < padded; i += 4) {
    const b0 = i < data.length ? view.getUint8(i) : 0;
    const b1 = i + 1 < data.length ? view.getUint8(i + 1) : 0;
    const b2 = i + 2 < data.length ? view.getUint8(i + 2) : 0;
    const b3 = i + 3 < data.length ? view.getUint8(i + 3) : 0;
    sum = (sum + ((b0 << 24) | (b1 << 16) | (b2 << 8) | b3)) >>> 0;
  }
  return sum >>> 0;
}

/** Reassembles a font from its tables, recomputing offsets, per-table
 * checksums and the head checkSumAdjustment. */
function assemble(version: number, tables: Table[]): Uint8Array {
  const numTables = tables.length;
  const sorted = [...tables].sort((a, b) => (a.tag < b.tag ? -1 : 1));
  const headerSize = 12 + numTables * 16;
  let offset = headerSize;
  const layout = sorted.map((table) => {
    const at = offset;
    offset = pad4(offset + table.data.length);
    return { table, at };
  });
  const total = offset;
  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  // Offset table.
  view.setUint32(0, version);
  view.setUint16(4, numTables);
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= numTables) {
    searchRange *= 2;
    entrySelector += 1;
  }
  view.setUint16(6, searchRange * 16);
  view.setUint16(8, entrySelector);
  view.setUint16(10, (numTables - searchRange) * 16);
  // Records + table bodies.
  layout.forEach(({ table, at }, i) => {
    const base = 12 + i * 16;
    for (let c = 0; c < 4; c += 1) {
      view.setUint8(base + c, table.tag.charCodeAt(c));
    }
    view.setUint32(base + 4, tableChecksum(table.data));
    view.setUint32(base + 8, at);
    view.setUint32(base + 12, table.data.length);
    out.set(table.data, at);
  });
  // checkSumAdjustment: 0xB1B0AFBA minus the whole-file checksum.
  const headEntry = layout.find((l) => l.table.tag === "head");
  if (headEntry) {
    const headOffset = headEntry.at;
    view.setUint32(headOffset + 8, 0);
    const fileChecksum = tableChecksum(out);
    view.setUint32(headOffset + 8, (0xb1b0afba - fileChecksum) >>> 0);
  }
  return out;
}

/** Returns a subset of `bytes` keeping only the glyphs needed to render the
 * given code points (plus .notdef and composite components), or the original
 * bytes when the font cannot be parsed as glyf-based TrueType. */
export function subsetFont(bytes: Uint8Array, codepoints: Iterable<number>): Uint8Array {
  let version: number;
  let tables: Map<string, Table>;
  try {
    ({ version, tables } = readTableDirectory(bytes));
  } catch {
    return bytes;
  }
  const head = tables.get("head");
  const maxp = tables.get("maxp");
  const loca = tables.get("loca");
  const glyf = tables.get("glyf");
  if (!head || !maxp || !loca || !glyf) {
    // Not glyf-based (e.g. CFF/OTTO): return as-is rather than corrupt it.
    return bytes;
  }
  const headView = new DataView(head.data.buffer, head.data.byteOffset, head.data.byteLength);
  const longLoca = headView.getInt16(50) === 1;
  const numGlyphs = new DataView(maxp.data.buffer, maxp.data.byteOffset, maxp.data.byteLength).getUint16(4);
  const offsets = readLoca(loca.data, longLoca, numGlyphs);

  const used = new Set<number>([0]);
  const cmap = tables.get("cmap");
  if (cmap) {
    for (const gid of glyphsForCodepoints(cmap.data, codepoints)) {
      used.add(gid);
    }
  }
  addComposites(glyf.data, offsets, used);

  // Rebuild glyf (kept outlines, empty otherwise) and a long-format loca.
  const chunks: Uint8Array[] = [];
  const newOffsets: number[] = [0];
  let running = 0;
  for (let gid = 0; gid < numGlyphs; gid += 1) {
    if (used.has(gid)) {
      const slice = glyf.data.subarray(offsets[gid], offsets[gid + 1]);
      const padded = pad4(slice.length);
      const buf = new Uint8Array(padded);
      buf.set(slice);
      chunks.push(buf);
      running += padded;
    }
    newOffsets.push(running);
  }
  const newGlyf = new Uint8Array(running);
  let at = 0;
  for (const chunk of chunks) {
    newGlyf.set(chunk, at);
    at += chunk.length;
  }
  const newLoca = new Uint8Array((numGlyphs + 1) * 4);
  const locaView = new DataView(newLoca.buffer);
  for (let i = 0; i <= numGlyphs; i += 1) {
    locaView.setUint32(i * 4, newOffsets[i]);
  }
  // Force long loca format in head.
  const newHead = head.data.slice();
  new DataView(newHead.buffer).setInt16(50, 1);

  const rebuilt: Table[] = [];
  for (const [tag, table] of tables) {
    if (tag === "glyf") {
      rebuilt.push({ tag, data: newGlyf });
    } else if (tag === "loca") {
      rebuilt.push({ tag, data: newLoca });
    } else if (tag === "head") {
      rebuilt.push({ tag, data: newHead });
    } else {
      rebuilt.push({ tag, data: table.data });
    }
  }
  return assemble(version, rebuilt);
}
