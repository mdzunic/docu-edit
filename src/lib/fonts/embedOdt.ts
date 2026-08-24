// SPEC-048 — injects subset embedded fonts into already-serialized .odt bytes
// (AC-15). Like the .docx path, this post-processes the package so writeOdt
// stays untouched: each face becomes a Fonts/<name>.ttf part (ODF stores
// embedded fonts un-obfuscated), declared via a <style:font-face> in
// content.xml's office:font-face-decls and a manifest entry. Framework-free.
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { SVG_NAMESPACE, XLINK_NAMESPACE } from "../core/odfXml";
import type { EmbedFace } from "./embedPlan";

const FONT_MEDIA_TYPE = "application/x-font-ttf";

// writeOdt declares only office/style/text/fo on the content.xml root, so the
// svg: and xlink: prefixes these font-face declarations use are undeclared
// there. Bind them locally on each <style:font-face> (the TABLE_XMLNS/LINK_XMLNS
// pattern from the writer) so the child svg:/xlink: nodes inherit them and the
// package is namespace-well-formed — a strict browser DOMParser (the reader's)
// rejects the file otherwise.
const FACE_XMLNS = ` xmlns:svg="${SVG_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}"`;

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

interface PlannedOdtFace {
  face: EmbedFace;
  styleName: string;
  path: string;
}

function fontFaceXml(planned: PlannedOdtFace[]): string {
  let out = "";
  for (const p of planned) {
    const weight = p.face.bold ? ' svg:font-weight="bold"' : "";
    const style = p.face.italic ? ' svg:font-style="italic"' : "";
    out +=
      `<style:font-face${FACE_XMLNS} style:name="${xmlEscape(p.styleName)}" svg:font-family="${xmlEscape(p.face.family)}">` +
      `<svg:font-face-src><svg:font-face-uri xlink:href="${p.path}"${weight}${style}>` +
      `<svg:font-face-format svg:string="truetype"/></svg:font-face-uri></svg:font-face-src>` +
      `</style:font-face>`;
  }
  return out;
}

/** Splices the new font-face declarations into content.xml's
 * office:font-face-decls, creating the element if the document has none. */
function patchContentXml(xml: string, faceXml: string): string {
  if (xml.includes("</office:font-face-decls>")) {
    return xml.replace("</office:font-face-decls>", `${faceXml}</office:font-face-decls>`);
  }
  const selfClosing = xml.match(/<office:font-face-decls\s*\/>/);
  if (selfClosing) {
    return xml.replace(selfClosing[0], `<office:font-face-decls>${faceXml}</office:font-face-decls>`);
  }
  // No decls element: insert one right after the document-content open tag.
  return xml.replace(
    /(<office:document-content\b[^>]*>)/,
    `$1<office:font-face-decls>${faceXml}</office:font-face-decls>`,
  );
}

function patchManifest(xml: string, planned: PlannedOdtFace[]): string {
  let entries = "";
  for (const p of planned) {
    if (!xml.includes(`manifest:full-path="${p.path}"`)) {
      entries += `<manifest:file-entry manifest:full-path="${p.path}" manifest:media-type="${FONT_MEDIA_TYPE}"/>`;
    }
  }
  return xml.replace("</manifest:manifest>", `${entries}</manifest:manifest>`);
}

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

/** Rewrites `.odt` bytes to embed the planned (already subset) faces. */
export async function embedFontsIntoOdt(bytes: Uint8Array, faces: EmbedFace[]): Promise<Uint8Array> {
  if (faces.length === 0) {
    return bytes;
  }
  const parts = await unzip(bytes);
  const content = parts.get("content.xml");
  const manifest = parts.get("META-INF/manifest.xml");
  if (!content || !manifest) {
    return bytes;
  }

  const planned: PlannedOdtFace[] = faces.map((face, i) => ({
    face,
    styleName: `EmbeddedFont${i + 1}`,
    path: `Fonts/${face.family.replace(/[^A-Za-z0-9]+/g, "")}${face.bold ? "b" : ""}${face.italic ? "i" : ""}_${i + 1}.ttf`,
  }));

  const out = new Map<string, Uint8Array>(parts);
  out.set("content.xml", encoder.encode(patchContentXml(decoder.decode(content), fontFaceXml(planned))));
  out.set("META-INF/manifest.xml", encoder.encode(patchManifest(decoder.decode(manifest), planned)));
  for (const p of planned) {
    out.set(p.path, p.face.bytes);
  }

  // mimetype must stay first and STOREd for ODF sniffers.
  const entries = Array.from(out, ([name, data]) => ({ name, data, store: name === "mimetype" }));
  entries.sort((a, b) => (a.name === "mimetype" ? -1 : b.name === "mimetype" ? 1 : 0));
  return zip(entries);
}
