// SPEC-048 — extracting fonts embedded inside an opened .docx or .odt so the
// editor can render with the author's own faces (AC-11, AC-13). The font
// binaries and their declarations are already preserved verbatim as opaque
// parts by the readers (SPEC-008), so this reads them back out at the point
// of use rather than changing the readers/writers — which keeps a plain
// read→write round-trip byte-preserving (AC-12, AC-14) for free.
//
// Framework-free (no Vue imports) per the src/lib convention.
import type { DocumentModel } from "../documentModel";
import { opaquePartBytes } from "../opaque";
import { childElements, parseXml } from "../core/xml";
import { WORD_NAMESPACE } from "../docx/xml";
import { STYLE_NAMESPACE, SVG_NAMESPACE, XLINK_NAMESPACE } from "../odf/xml";

/** The OOXML `r:id` relationship-reference namespace (distinct from the
 * package relationships namespace of a `.rels` part itself). */
const OFFICE_REL_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** LibreOffice's loext extension namespace, where an ODF `svg:font-face-uri`
 * carries the per-URI weight/style of the face it points at. */
const LOEXT_NAMESPACE = "urn:oasis:names:tc:openoffice:xmlns:loext:1.0";

export interface EmbeddedFontFace {
  /** The family name the face declares itself under — matched against the
   * document's run families to decide when to render with it. */
  family: string;
  bold: boolean;
  italic: boolean;
  /** The de-obfuscated, ready-to-register font program bytes. */
  bytes: Uint8Array;
  /** Whether the font's OS/2 fsType permits embedding. A face whose licence
   * bits forbid embedding is neither rendered nor re-embedded (AC-11/13's
   * "subject to their embedding permission bits", AC-17). */
  embeddable: boolean;
  /** The package part the bytes came from (diagnostics / de-duplication). */
  partName: string;
}

// --- OS/2 fsType embedding permission -----------------------------------

/** Reads the OS/2 `fsType` embedding-permission bitfield, or null when the
 * font has no OS/2 table (permission then treated as unrestricted). Accepts
 * TrueType (0x00010000 / 'true') and OpenType-CFF ('OTTO') alike. */
export function readFsType(bytes: Uint8Array): number | null {
  if (bytes.length < 12) {
    return null;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = view.getUint16(4);
  for (let i = 0; i < numTables; i += 1) {
    const base = 12 + i * 16;
    if (base + 16 > bytes.length) {
      break;
    }
    const tag = String.fromCharCode(
      view.getUint8(base),
      view.getUint8(base + 1),
      view.getUint8(base + 2),
      view.getUint8(base + 3),
    );
    if (tag === "OS/2") {
      const offset = view.getUint32(base + 8);
      if (offset + 10 <= bytes.length) {
        return view.getUint16(offset + 8);
      }
      return null;
    }
  }
  return null;
}

/** Whether an fsType value permits embedding. Bit 1 (0x0002, "Restricted
 * License embedding") forbids it outright; the "no embedding" bit set with
 * others is still restricted. An absent OS/2 table (null) is treated as
 * installable/unrestricted, matching how Word treats such fonts. */
export function fsTypeAllowsEmbedding(fsType: number | null): boolean {
  if (fsType === null) {
    return true;
  }
  // Mask to the defined licence bits; bit 1 set → restricted, no embedding.
  return (fsType & 0x0002) === 0;
}

// --- .docx obfuscated embedded fonts ------------------------------------

/** Reverses the OOXML font obfuscation (ECMA-376 §17.8.1): the first 32 bytes
 * of an `.odttf` part are XORed with the 16-byte GUID key (used twice, in
 * reverse byte order). Returns a fresh de-obfuscated copy. */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[^0-9a-fA-F]/g, "");
  if (hex.length !== 32) {
    return bytes.slice();
  }
  const guid = new Uint8Array(16);
  for (let i = 0; i < 16; i += 1) {
    guid[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  const key = guid.slice().reverse();
  const out = bytes.slice();
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i += 1) {
    out[i] ^= key[i % 16];
  }
  return out;
}

interface EmbedRef {
  bold: boolean;
  italic: boolean;
  relId: string;
  fontKey: string;
}

const EMBED_ELEMENTS: Array<{ local: string; bold: boolean; italic: boolean }> = [
  { local: "embedRegular", bold: false, italic: false },
  { local: "embedBold", bold: true, italic: false },
  { local: "embedItalic", bold: false, italic: true },
  { local: "embedBoldItalic", bold: true, italic: true },
];

/** Resolves a fontTable rels target ("fonts/font1.odttf") to a full package
 * part path under `word/`. */
function resolveWordPart(target: string): string {
  const clean = target.replace(/^(\.\/|\/)/, "");
  return clean.startsWith("word/") ? clean : `word/${clean}`;
}

export function extractDocxEmbeddedFonts(parts: Map<string, Uint8Array>): EmbeddedFontFace[] {
  const tableBytes = parts.get("word/fontTable.xml");
  const relsBytes = parts.get("word/_rels/fontTable.xml.rels");
  if (!tableBytes || !relsBytes) {
    return [];
  }
  let rels: Map<string, string>;
  let table: Document;
  try {
    rels = parseFontTableRels(new TextDecoder("utf-8").decode(relsBytes));
    table = parseXml(new TextDecoder("utf-8").decode(tableBytes));
  } catch {
    return [];
  }
  const faces: EmbeddedFontFace[] = [];
  for (const font of childElements(table.documentElement, WORD_NAMESPACE, "font")) {
    const family = font.getAttributeNS(WORD_NAMESPACE, "name") ?? "";
    for (const { local, bold, italic } of EMBED_ELEMENTS) {
      const embed = childElements(font, WORD_NAMESPACE, local)[0];
      if (!embed) {
        continue;
      }
      const ref: EmbedRef = {
        bold,
        italic,
        relId: embed.getAttributeNS(OFFICE_REL_NAMESPACE, "id") ?? "",
        fontKey: embed.getAttributeNS(WORD_NAMESPACE, "fontKey") ?? "",
      };
      const target = rels.get(ref.relId);
      if (!target) {
        continue;
      }
      const raw = parts.get(resolveWordPart(target));
      if (!raw) {
        continue;
      }
      const bytes = deobfuscateOdttf(raw, ref.fontKey);
      faces.push({
        family,
        bold: ref.bold,
        italic: ref.italic,
        bytes,
        embeddable: fsTypeAllowsEmbedding(readFsType(bytes)),
        partName: resolveWordPart(target),
      });
    }
  }
  return faces;
}

/** Parses a fontTable `.rels` part into id → target. Kept local (rather than
 * core/xml's `parseRelationships`) to avoid pulling in its Relationship type
 * for a two-attribute read. */
function parseFontTableRels(xmlText: string): Map<string, string> {
  const doc = parseXml(xmlText);
  const map = new Map<string, string>();
  const root = doc.documentElement;
  for (const rel of Array.from(root.getElementsByTagName("*"))) {
    if (rel.localName === "Relationship") {
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      if (id && target) {
        map.set(id, target);
      }
    }
  }
  return map;
}

// --- .odt embedded fonts (Fonts/*.ttf, non-obfuscated) ------------------

export function extractOdtEmbeddedFonts(
  fontFaceDeclsXml: string | undefined,
  parts: Map<string, Uint8Array>,
): EmbeddedFontFace[] {
  if (!fontFaceDeclsXml) {
    return [];
  }
  let decls: Document;
  try {
    decls = parseXml(fontFaceDeclsXml);
  } catch {
    return [];
  }
  const faces: EmbeddedFontFace[] = [];
  for (const face of childElements(decls.documentElement, STYLE_NAMESPACE, "font-face")) {
    const family = face.getAttributeNS(SVG_NAMESPACE, "font-family") ?? "";
    const src = childElements(face, SVG_NAMESPACE, "font-face-src")[0];
    if (!src) {
      continue;
    }
    for (const uri of childElements(src, SVG_NAMESPACE, "font-face-uri")) {
      const href = uri.getAttributeNS(XLINK_NAMESPACE, "href");
      if (!href) {
        continue;
      }
      const raw = parts.get(href.replace(/^\.?\//, ""));
      if (!raw) {
        continue;
      }
      // LibreOffice tags each URI with the weight/style of the face it backs.
      const weight = uri.getAttributeNS(LOEXT_NAMESPACE, "font-weight") ?? uri.getAttributeNS(SVG_NAMESPACE, "font-weight");
      const style = uri.getAttributeNS(LOEXT_NAMESPACE, "font-style") ?? uri.getAttributeNS(SVG_NAMESPACE, "font-style");
      const bytes = raw.slice();
      faces.push({
        family,
        bold: weight === "bold" || weight === "700",
        italic: style === "italic" || style === "oblique",
        bytes,
        embeddable: fsTypeAllowsEmbedding(readFsType(bytes)),
        partName: href.replace(/^\.?\//, ""),
      });
    }
  }
  return faces;
}

/** Every font embedded in the open document, de-obfuscated and ready to
 * register. Empty for a document with no embedded fonts. */
export function extractEmbeddedFonts(model: DocumentModel): EmbeddedFontFace[] {
  const parts = opaquePartBytes(model);
  if (model.format === "odt") {
    return extractOdtEmbeddedFonts(model.sourceXml?.odtFontFaceDecls, parts);
  }
  return extractDocxEmbeddedFonts(parts);
}
