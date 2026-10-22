// SPEC-048 — injects subset embedded fonts into already-serialized .docx
// bytes (AC-15). Post-processing the package (rather than threading a font
// plan through the writer) keeps writeDocx untouched and a plain save
// byte-preserving; the font parts land as new OPC parts plus a rebuilt
// word/fontTable.xml and the Content-Types / document-rels declarations that
// reference them. Framework-free per the src/lib convention.
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { deobfuscateOdttf } from "./embeddedFonts";
import type { EmbedFace } from "./embedPlan";

const OBFUSCATED_FONT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const FONT_TABLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** A stable {GUID} for a font part's obfuscation key. Deterministic in the
 * face index so output is reproducible (and Math.random-free). */
function fontKeyFor(index: number): string {
  const hex = (index + 1).toString(16).padStart(8, "0");
  return `{${hex}-1000-4000-8000-00000000${hex.slice(0, 4)}}`;
}

interface PlannedFace {
  face: EmbedFace;
  relId: string;
  partName: string;
  fileName: string;
  fontKey: string;
  tag: "embedRegular" | "embedBold" | "embedItalic" | "embedBoldItalic";
}

function embedTag(face: EmbedFace): PlannedFace["tag"] {
  if (face.bold && face.italic) {
    return "embedBoldItalic";
  }
  return face.bold ? "embedBold" : face.italic ? "embedItalic" : "embedRegular";
}

/** The next free numeric rId past those a rels part already uses. */
function nextRelId(relsXml: string): number {
  let max = 0;
  for (const match of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

function buildFontTableXml(byFamily: Map<string, PlannedFace[]>): string {
  let body = "";
  for (const [family, planned] of byFamily) {
    body += `<w:font w:name="${xmlEscape(family)}">`;
    for (const p of planned) {
      body += `<w:${p.tag} r:id="${p.relId}" w:fontKey="${p.fontKey}" w:subsetted="true"/>`;
    }
    body += `</w:font>`;
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</w:fonts>`
  );
}

function buildFontTableRels(planned: PlannedFace[]): string {
  let body = "";
  for (const p of planned) {
    body += `<Relationship Id="${p.relId}" Type="${FONT_TABLE_REL_TYPE.replace("/fontTable", "/font")}" Target="fonts/${p.fileName}"/>`;
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
  );
}

function patchContentTypes(xml: string): string {
  if (xml.includes('Extension="odttf"')) {
    return xml;
  }
  const decl = `<Default Extension="odttf" ContentType="${OBFUSCATED_FONT_CONTENT_TYPE}"/>`;
  return xml.replace("</Types>", `${decl}</Types>`);
}

function patchDocumentRels(xml: string): string {
  if (xml.includes(FONT_TABLE_REL_TYPE)) {
    return xml;
  }
  const id = `rId${nextRelId(xml)}`;
  const rel = `<Relationship Id="${id}" Type="${FONT_TABLE_REL_TYPE}" Target="fontTable.xml"/>`;
  return xml.replace("</Relationships>", `${rel}</Relationships>`);
}

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

/** Rewrites `.docx` bytes to embed the planned (already subset) faces. */
export async function embedFontsIntoDocx(bytes: Uint8Array, faces: EmbedFace[]): Promise<Uint8Array> {
  if (faces.length === 0) {
    return bytes;
  }
  const parts = await unzip(bytes);

  const planned: PlannedFace[] = faces.map((face, i) => {
    const fileName = `font${i + 1}.odttf`;
    return {
      face,
      relId: `rId${i + 1}`,
      partName: `word/fonts/${fileName}`,
      fileName,
      fontKey: fontKeyFor(i),
      tag: embedTag(face),
    };
  });

  // Group into one <w:font> per family, preserving first-seen order.
  const byFamily = new Map<string, PlannedFace[]>();
  for (const p of planned) {
    const list = byFamily.get(p.face.family) ?? [];
    list.push(p);
    byFamily.set(p.face.family, list);
  }

  const out = new Map<string, Uint8Array>(parts);
  out.set("word/fontTable.xml", encoder.encode(buildFontTableXml(byFamily)));
  out.set("word/_rels/fontTable.xml.rels", encoder.encode(buildFontTableRels(planned)));
  for (const p of planned) {
    out.set(p.partName, deobfuscateOdttf(p.face.bytes, p.fontKey));
  }
  const contentTypes = parts.get("[Content_Types].xml");
  if (contentTypes) {
    out.set("[Content_Types].xml", encoder.encode(patchContentTypes(decoder.decode(contentTypes))));
  }
  const docRels = parts.get("word/_rels/document.xml.rels");
  if (docRels) {
    out.set("word/_rels/document.xml.rels", encoder.encode(patchDocumentRels(decoder.decode(docRels))));
  }

  return zip(Array.from(out, ([name, data]) => ({ name, data })));
}
