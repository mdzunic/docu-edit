// The `.odp` writer (SPEC-080). Like the pptx writer, three paths chosen by
// what the model holds:
//
//  1. No retained package (from-scratch / `stripRetention`ed deck): build the
//     minimal ODF package — `mimetype` (first, STOREd), META-INF/manifest.xml,
//     meta.xml, styles.xml (one `Default` master page carrying the page size),
//     content.xml with one `draw:page` per model slide (empty placeholder
//     frames; SPEC-081 owns real frame serialisation).
//  2. Retained package, nothing structural changed: re-emit every part
//     verbatim, `mimetype` first + STOREd — the preparation-milestone
//     behaviour, so a no-edit download stays byte-identical (SPEC-008 / AC-51).
//  3. Retained package + structural changes: re-emit every part verbatim
//     except content.xml, rebuilt as `envelope.prefix + pages + envelope.suffix`
//     — each untouched page contributes its `sourceXml` verbatim
//     (splice-equality), duplicated pages the source `sourceXml` with
//     `draw:name` uniqued, new pages generated markup. META-INF/manifest.xml
//     is spliced only when parts were added/removed (never, this spec — pages
//     live inside content.xml).
import { decodedPartBytes } from "../../core/packageParts";
import { buildMetaXml } from "../../core/odfMeta";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import {
  DEFAULT_PRESENTATION_TITLE,
  type PlaceholderRole,
  type PresentationModel,
  type Slide,
  type TextFrame,
} from "../model";
import { ODP_MIMETYPE } from "./reader";

export { ODP_MIMETYPE };

const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

const encoder = new TextEncoder();
const encode = (text: string): Uint8Array => encoder.encode(text);

export function odpFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.odp`;
}

export interface WriteOdpOptions {
  /** Injectable so tests assert an exact meta.xml (the clock-coupled-test
   * rule): the timestamp stamped into a from-scratch meta.xml. */
  modified?: Date;
}

/** Serialises a presentation to `.odp` bytes (see the module header). */
export async function writeOdp(model: PresentationModel, options: WriteOdpOptions = {}): Promise<{ bytes: Uint8Array }> {
  const hasPackage = model.opaque !== undefined && model.opaque.length > 0;
  if (!hasPackage) {
    return { bytes: await zip(buildFromScratch(model, options.modified ?? new Date())) };
  }
  const parts = decodedPartBytes(model.opaque);
  const structural = model.structureDirty === true || model.slides.some((slide) => slide.dirty === true);
  const mimetype = parts.get("mimetype") ?? encode(ODP_MIMETYPE);
  const entries: ZipEntryInput[] = [{ name: "mimetype", data: mimetype, store: true }];

  const rebuiltContent =
    structural && model.odpContentEnvelope ? encode(rebuildContentXml(model, model.odpContentEnvelope)) : undefined;

  for (const [name, data] of parts) {
    if (name === "mimetype") {
      continue;
    }
    if (name === "content.xml" && rebuiltContent) {
      entries.push({ name, data: rebuiltContent });
    } else {
      entries.push({ name, data });
    }
  }
  return { bytes: await zip(entries) };
}

// --- substitution: rebuild content.xml from the page envelope --------------

function rebuildContentXml(model: PresentationModel, envelope: { prefix: string; suffix: string }): string {
  const usedNames = new Set<string>();
  let lastMaster = "Default";
  const pages = model.slides.map((slide) => {
    if (slide.sourceXml) {
      const master = drawNameAttr(slide.sourceXml, "draw:master-page-name");
      if (master) {
        lastMaster = master;
      }
      return uniqueDrawName(slide.sourceXml, usedNames);
    }
    const page = generatePage(slide, uniqueName("page", usedNames), lastMaster);
    return page;
  });
  return envelope.prefix + pages.join("") + envelope.suffix;
}

/** The value of an attribute like `draw:name="…"` in a page's markup. */
function drawNameAttr(xml: string, attr: string): string | undefined {
  const match = new RegExp(`${attr}="([^"]*)"`).exec(xml);
  return match ? match[1] : undefined;
}

/** A page's `sourceXml` with its `draw:name` made unique against `used` (and
 * recorded there). A duplicated page shares its source's name until here. */
function uniqueDrawName(sourceXml: string, used: Set<string>): string {
  const current = drawNameAttr(sourceXml, "draw:name");
  if (current === undefined) {
    return sourceXml;
  }
  if (!used.has(current)) {
    used.add(current);
    return sourceXml;
  }
  const fresh = uniqueName(current, used);
  return sourceXml.replace(/draw:name="[^"]*"/, `draw:name="${fresh}"`);
}

function uniqueName(base: string, used: Set<string>): string {
  let candidate = base;
  let n = 2;
  while (used.has(candidate)) {
    candidate = `${base}_${n++}`;
  }
  used.add(candidate);
  return candidate;
}

// --- shared page serialisation ---------------------------------------------

const ROLE_TO_CLASS: Record<PlaceholderRole, string> = {
  title: "title",
  subtitle: "subtitle",
  body: "body",
  other: "outline",
};

function esc(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function frameXml(frame: TextFrame): string {
  const cls = frame.role ? ROLE_TO_CLASS[frame.role] : "outline";
  const geom =
    ` svg:x="${frame.xCm}cm" svg:y="${frame.yCm}cm" svg:width="${frame.widthCm}cm" svg:height="${frame.heightCm}cm"`;
  return `<draw:frame presentation:class="${esc(cls)}"${geom}><draw:text-box/></draw:frame>`;
}

/** A fresh `draw:page` for a placeholder-only slide (from-scratch and new
 * pages). Uses the given master page name (AC-28). */
function generatePage(slide: Slide, name: string, master: string): string {
  const frames = slide.frames
    .filter((frame): frame is TextFrame => frame.kind === "text")
    .map(frameXml)
    .join("");
  return `<draw:page draw:name="${esc(name)}" draw:master-page-name="${esc(master)}">${frames}</draw:page>`;
}

// --- from-scratch package --------------------------------------------------

function buildFromScratch(model: PresentationModel, modified: Date): ZipEntryInput[] {
  const used = new Set<string>();
  const pages = model.slides.map((slide) => generatePage(slide, uniqueName("page", used), "Default")).join("");

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:text="${TEXT_NS}"` +
    ` xmlns:draw="${DRAW_NS}" xmlns:fo="${FO_NS}" xmlns:svg="${SVG_NS}" xmlns:presentation="${PRESENTATION_NS}"` +
    ` office:version="1.2">` +
    `<office:automatic-styles/>` +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    `</office:document-content>`;

  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:draw="${DRAW_NS}"` +
    ` xmlns:fo="${FO_NS}" xmlns:svg="${SVG_NS}" xmlns:presentation="${PRESENTATION_NS}" office:version="1.2">` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="${model.slideWidthCm}cm"` +
    ` fo:page-height="${model.slideHeightCm}cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles>` +
    `<style:master-page style:name="Default" style:page-layout-name="PL1"/>` +
    `</office:master-styles>` +
    `</office:document-styles>`;

  const meta = buildMetaXml({ title: model.title }, modified);

  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="${ODP_MIMETYPE}"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;

  return [
    { name: "mimetype", data: encode(ODP_MIMETYPE), store: true },
    { name: "content.xml", data: encode(content) },
    { name: "styles.xml", data: encode(styles) },
    { name: "meta.xml", data: encode(meta) },
    { name: "META-INF/manifest.xml", data: encode(manifest) },
  ];
}
