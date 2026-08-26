// The `.pptx` writer (SPEC-080). Three paths, chosen by what the model holds:
//
//  1. No retained package (`opaque` absent/empty — a from-scratch or
//     `stripRetention`ed deck): build the full minimal package from the model
//     — Content_Types, root rels, docProps/core.xml, presentation.xml + rels,
//     one slideMaster + theme + slideLayout, and one slide part per model
//     slide (empty placeholders; SPEC-081 owns real frame serialisation).
//  2. Retained package, nothing structural changed (no `structureDirty`, no
//     dirty slide): re-emit every part verbatim, in package order — the
//     preparation-milestone behaviour, so a no-edit download stays
//     byte-identical (SPEC-008 / AC-51).
//  3. Retained package + structural changes (the substitution path): start
//     from `opaque`, rebuild ONLY the structural parts (presentation.xml, its
//     rels, [Content_Types].xml), copy every untouched slide part and every
//     unmodelled part byte-verbatim, skip deleted slides' parts, and append
//     new/duplicated slide parts. The normative rules live in SPEC-080.
import { base64ToBytes, decodedPartBytes } from "../../core/packageParts";
import { cmToEmu } from "../../core/units";
import {
  escapeXmlAttribute,
  parseRelationships,
  relsPathFor,
  resolvePartPath,
  RELATIONSHIPS_NAMESPACE,
  type Relationship,
} from "../../core/xml";
import { buildCorePropertiesXml } from "../../core/opcProperties";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import type { MediaPart } from "../../core/media";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel, type Slide, type SlideFrame } from "../model";
import { PPTX_MAIN_CONTENT_TYPE } from "./reader";
import { serializeSlideXml } from "./serialize";

const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export { PPTX_MIME_TYPE } from "./reader";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CORE_PROPERTIES_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const NOTES_SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");
const encode = (text: string): Uint8Array => encoder.encode(text);
const decode = (bytes: Uint8Array): string => decoder.decode(bytes);

export function pptxFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.pptx`;
}

export interface WritePptxOptions {
  /** Injectable so tests assert an exact core.xml (the clock-coupled-test
   * rule): the timestamp stamped into a from-scratch docProps/core.xml. */
  modified?: Date;
}

/** Serialises a presentation to `.pptx` bytes (see the module header for the
 * three paths). */
export async function writePptx(model: PresentationModel, options: WritePptxOptions = {}): Promise<{ bytes: Uint8Array }> {
  const hasPackage = model.opaque !== undefined && model.opaque.length > 0;
  if (!hasPackage) {
    return { bytes: await zip(buildFromScratch(model, options.modified ?? new Date())) };
  }
  const parts = decodedPartBytes(model.opaque);
  const anyDirty = model.slides.some((slide) => slide.dirty === true);
  if (!model.structureDirty && !anyDirty) {
    // Nothing changed — re-emit verbatim (AC-51).
    const bytes = await zip([...parts].map(([name, data]) => ({ name, data })));
    return { bytes };
  }
  if (!model.structureDirty) {
    // Text-only edits: substitute each dirty slide's own part in place —
    // presentation.xml, its rels and [Content_Types].xml stay byte-verbatim,
    // and every untouched slide/unmodelled part is copied through (SPEC-081
    // Dirtying: a text edit does not rebuild the structural parts).
    return { bytes: await zip(buildTextEditSubstitution(model, parts)) };
  }
  return { bytes: await zip(buildSubstitution(model, parts)) };
}

/** Rebuilds only the dirty slides' own parts, in place; every other part
 * (structural parts, untouched slides) is copied byte-verbatim. A dirty slide
 * that gained an inserted image also has its `.rels` re-emitted (with the new
 * image relationship spliced in), the media part added under `ppt/media/`, and
 * the Content-Types extension default added if missing (SPEC-082). */
function buildTextEditSubstitution(model: PresentationModel, parts: Map<string, Uint8Array>): ZipEntryInput[] {
  const contentTypesXml = decode(parts.get("[Content_Types].xml") ?? encode(""));
  const media = planMedia(model, parts, contentTypesXml);
  const rebuilt = new Map<string, Uint8Array>();
  for (const slide of model.slides) {
    if (slide.dirty && slide.partName && parts.has(slide.partName)) {
      const relsPath = relsPathFor(slide.partName);
      const relsXml = decode(parts.get(relsPath) ?? encode(""));
      const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();
      const newRels = linkImageRels(slide, rels, slide.partName);
      rebuilt.set(slide.partName, encode(serializeSlideXml(slide, rels, slide.partName)));
      if (newRels.length > 0) {
        rebuilt.set(relsPath, encode(appendRelationships(relsXml, newRels)));
      }
    }
  }
  const entries: ZipEntryInput[] = [];
  for (const [name, data] of parts) {
    if (rebuilt.has(name)) {
      entries.push({ name, data: rebuilt.get(name)! });
    } else if (media.override.has(name)) {
      entries.push({ name, data: media.override.get(name)! });
    } else if (name === "[Content_Types].xml" && media.defaults) {
      entries.push({ name, data: encode(spliceContentDefaults(decode(data), media.defaults)) });
    } else {
      entries.push({ name, data });
    }
  }
  entries.push(...media.appended);
  return entries;
}

// --- inserted media (SPEC-082) ---------------------------------------------

/** Every media part referenced by an ImageFrame anywhere in the deck (groups
 * descended), from `model.media` — the writers' source of truth for inserted
 * image parts. */
function referencedMedia(model: PresentationModel): MediaPart[] {
  const used = new Set<string>();
  const walk = (frames: SlideFrame[]): void => {
    for (const frame of frames) {
      if (frame.kind === "image") {
        used.add(frame.mediaId);
      } else if (frame.kind === "group") {
        walk(frame.children);
      }
    }
  };
  for (const slide of model.slides) {
    walk(slide.frames);
  }
  return (model.media ?? []).filter((part) => used.has(part.mediaId));
}

interface MediaPlan {
  /** Package parts already present, overridden with the model's media bytes
   * (a media part takes precedence over an `opaque` twin). */
  override: Map<string, Uint8Array>;
  /** Media parts not yet in the package, to append (`ppt/media/…`). */
  appended: ZipEntryInput[];
  /** Content-Types `<Default>` entries for image extensions not yet declared,
   * to splice before the first `<Override>`; "" when none are missing. */
  defaults: string;
}

/** Plans how the deck's referenced media rides in the package: which parts to
 * override in place, which to append, and which Content-Types defaults are
 * missing (AC-47). `contentTypesXml` is the package's current
 * `[Content_Types].xml` (""/from-scratch when it does not exist yet). */
function planMedia(model: PresentationModel, parts: Map<string, Uint8Array>, contentTypesXml: string): MediaPlan {
  const override = new Map<string, Uint8Array>();
  const appended: ZipEntryInput[] = [];
  const extToType = new Map<string, string>();
  for (const part of referencedMedia(model)) {
    const bytes = base64ToBytes(part.data);
    if (parts.has(part.mediaId)) {
      override.set(part.mediaId, bytes);
    } else {
      appended.push({ name: part.mediaId, data: bytes });
    }
    const ext = part.mediaId.slice(part.mediaId.lastIndexOf(".") + 1).toLowerCase();
    if (ext) {
      extToType.set(ext, part.contentType);
    }
  }
  let defaults = "";
  for (const [ext, contentType] of extToType) {
    if (!new RegExp(`<Default\\s+Extension="${ext}"`, "i").test(contentTypesXml)) {
      defaults += `<Default Extension="${ext}" ContentType="${contentType}"/>`;
    }
  }
  return { override, appended, defaults };
}

/** Adds image relationships to `rels` for any ImageFrame whose media is not
 * yet resolvable in it (a freshly inserted image), minting ids after the
 * highest existing one. Mutates `rels`; returns the new relationships (empty
 * when the slide gained none). */
function linkImageRels(slide: Slide, rels: Map<string, Relationship>, slidePartName: string): Relationship[] {
  const resolved = new Set<string>();
  let maxId = 0;
  for (const rel of rels.values()) {
    maxId = Math.max(maxId, relIdNumber(rel.id));
    if (rel.type === IMAGE_REL) {
      resolved.add(resolvePartPath(slidePartName, rel.target));
    }
  }
  const added: Relationship[] = [];
  const seen = new Set<string>();
  const walk = (frames: SlideFrame[]): void => {
    for (const frame of frames) {
      if (frame.kind === "image") {
        if (!resolved.has(frame.mediaId) && !seen.has(frame.mediaId)) {
          seen.add(frame.mediaId);
          maxId += 1;
          const rel: Relationship = { id: `rId${maxId}`, type: IMAGE_REL, target: relativeTarget(slidePartName, frame.mediaId) };
          rels.set(rel.id, rel);
          added.push(rel);
        }
      } else if (frame.kind === "group") {
        walk(frame.children);
      }
    }
  };
  walk(slide.frames);
  return added;
}

/** The `.rels` XML of a slide with `newRels` appended before `</Relationships>`
 * — splicing keeps every existing relationship byte-verbatim (TargetMode and
 * all), unlike re-serialising the parsed map. */
function appendRelationships(relsXml: string, newRels: Relationship[]): string {
  if (newRels.length === 0) {
    return relsXml;
  }
  const additions = newRels
    .map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXmlAttribute(rel.target)}"/>`)
    .join("");
  if (/<\/Relationships>/.test(relsXml)) {
    return relsXml.replace(/<\/Relationships>/, `${additions}</Relationships>`);
  }
  return `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${additions}</Relationships>`;
}

/** Splices Content-Types `<Default>` entries before the first `<Override>` (or
 * `</Types>` when there is none) — schema-correct order. */
function spliceContentDefaults(contentTypesXml: string, defaults: string): string {
  if (defaults.length === 0) {
    return contentTypesXml;
  }
  if (/<Override/.test(contentTypesXml)) {
    return contentTypesXml.replace(/<Override/, `${defaults}<Override`);
  }
  return contentTypesXml.replace(/<\/Types>/, `${defaults}</Types>`);
}

// --- shared slide serialisation --------------------------------------------

/** A slide part's XML from its frames (SPEC-081's serialiser). `rels`/`path`
 * resolve an ImageFrame's media back to its `r:embed`; a from-scratch / new
 * slide passes an empty rels map (it bears no images yet). */
function slideXml(slide: Slide, rels: Map<string, Relationship> = new Map(), path = ""): string {
  return serializeSlideXml(slide, rels, path);
}

function slideRelsXml(layoutTarget: string): string {
  return (
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${SLIDE_LAYOUT_REL}" Target="${escapeXmlAttribute(layoutTarget)}"/>` +
    `</Relationships>`
  );
}

// --- from-scratch package --------------------------------------------------

function buildFromScratch(model: PresentationModel, modified: Date): ZipEntryInput[] {
  const slidePaths = model.slides.map((_, index) => `ppt/slides/slide${index + 1}.xml`);

  const mediaDefaults = mediaContentDefaults(referencedMedia(model));
  const contentTypes =
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    mediaDefaults +
    `<Override PartName="/ppt/presentation.xml" ContentType="${PPTX_MAIN_CONTENT_TYPE}"/>` +
    slidePaths.map((path) => `<Override PartName="/${path}" ContentType="${SLIDE_CONTENT_TYPE}"/>`).join("") +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `</Types>`;

  const rootRels =
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="ppt/presentation.xml"/>` +
    `<Relationship Id="rId2" Type="${CORE_PROPERTIES_REL}" Target="docProps/core.xml"/>` +
    `</Relationships>`;

  // presentation.xml.rels: master (rId1), then one slide rel each, then theme.
  const slideRelIds = model.slides.map((_, index) => `rId${index + 2}`);
  const themeRelId = `rId${model.slides.length + 2}`;
  const presentationRels =
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${SLIDE_MASTER_REL}" Target="slideMasters/slideMaster1.xml"/>` +
    model.slides
      .map((_, index) => `<Relationship Id="${slideRelIds[index]}" Type="${SLIDE_REL}" Target="slides/slide${index + 1}.xml"/>`)
      .join("") +
    `<Relationship Id="${themeRelId}" Type="${THEME_REL}" Target="theme/theme1.xml"/>` +
    `</Relationships>`;

  const sldIdLst = model.slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="${slideRelIds[index]}"/>`).join("");
  const presentation =
    `${XML_DECL}<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${sldIdLst}</p:sldIdLst>` +
    `<p:sldSz cx="${cmToEmu(model.slideWidthCm)}" cy="${cmToEmu(model.slideHeightCm)}"/>` +
    `<p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;

  const core = buildCorePropertiesXml({ title: model.title }, modified);

  const entries: ZipEntryInput[] = [
    { name: "[Content_Types].xml", data: encode(contentTypes) },
    { name: "_rels/.rels", data: encode(rootRels) },
    { name: "docProps/core.xml", data: encode(core) },
    { name: "ppt/presentation.xml", data: encode(presentation) },
    { name: "ppt/_rels/presentation.xml.rels", data: encode(presentationRels) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: encode(SLIDE_MASTER_XML) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: encode(SLIDE_MASTER_RELS) },
    { name: "ppt/slideLayouts/slideLayout1.xml", data: encode(SLIDE_LAYOUT_XML) },
    { name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: encode(SLIDE_LAYOUT_RELS) },
    { name: "ppt/theme/theme1.xml", data: encode(THEME_XML) },
  ];
  model.slides.forEach((slide, index) => {
    const slidePath = slidePaths[index];
    const rels = new Map<string, Relationship>();
    rels.set("rId1", { id: "rId1", type: SLIDE_LAYOUT_REL, target: "../slideLayouts/slideLayout1.xml" });
    const imageRels = linkImageRels(slide, rels, slidePath);
    entries.push({ name: slidePath, data: encode(slideXml(slide, rels, slidePath)) });
    entries.push({
      name: `ppt/slides/_rels/slide${index + 1}.xml.rels`,
      data: encode(appendRelationships(slideRelsXml("../slideLayouts/slideLayout1.xml"), imageRels)),
    });
  });
  for (const part of referencedMedia(model)) {
    entries.push({ name: part.mediaId, data: base64ToBytes(part.data) });
  }
  return entries;
}

/** Content-Types `<Default>` entries for the extensions a media list uses,
 * de-duplicated by extension. */
function mediaContentDefaults(media: MediaPart[]): string {
  const byExt = new Map<string, string>();
  for (const part of media) {
    const ext = part.mediaId.slice(part.mediaId.lastIndexOf(".") + 1).toLowerCase();
    if (ext) {
      byExt.set(ext, part.contentType);
    }
  }
  return [...byExt].map(([ext, contentType]) => `<Default Extension="${ext}" ContentType="${contentType}"/>`).join("");
}

// --- substitution path -----------------------------------------------------

interface Emitted {
  slide: Slide;
  emitName: string;
  kind: "verbatim" | "duplicate" | "new";
  sourceName?: string;
  relId?: string;
}

function slideNumberOf(path: string): number | null {
  const match = /\/slide(\d+)\.xml$/.exec(path);
  return match ? Number(match[1]) : null;
}

function relativeTarget(fromPart: string, toPart: string): string {
  const fromDir = fromPart.split("/").slice(0, -1);
  const toSegments = toPart.split("/");
  let i = 0;
  while (i < fromDir.length && i < toSegments.length - 1 && fromDir[i] === toSegments[i]) {
    i++;
  }
  const up = fromDir.slice(i).map(() => "..");
  return [...up, ...toSegments.slice(i)].join("/");
}

function buildSubstitution(model: PresentationModel, parts: Map<string, Uint8Array>): ZipEntryInput[] {
  // Locate the presentation part through the root relationships.
  const rootRels = parseRelationships(decode(parts.get("_rels/.rels") ?? encode("")));
  const officeRel = [...rootRels.values()].find((rel) => rel.type === OFFICE_DOCUMENT_REL);
  const presPath = officeRel ? resolvePartPath("", officeRel.target) : "ppt/presentation.xml";
  const presRelsPath = relsPathFor(presPath);
  const origPresXml = decode(parts.get(presPath) ?? encode(""));
  const origPresRels = parseRelationships(decode(parts.get(presRelsPath) ?? encode("")));

  // The original slide parts (by their relationship targets).
  const origSlidePaths = new Set<string>();
  for (const rel of origPresRels.values()) {
    if (rel.type === SLIDE_REL) {
      origSlidePaths.add(resolvePartPath(presPath, rel.target));
    }
  }
  const slidesDir = [...origSlidePaths][0]?.split("/").slice(0, -1).join("/") ?? `${presPath.split("/").slice(0, -1).join("/")}/slides`;

  // Fresh slide-part naming: past the highest existing slideN, skipping taken.
  const takenNumbers = new Set<number>();
  for (const name of parts.keys()) {
    const number = slideNumberOf(name);
    if (number !== null && name.startsWith(`${slidesDir}/`)) {
      takenNumbers.add(number);
    }
  }
  let nextNumber = takenNumbers.size > 0 ? Math.max(...takenNumbers) : 0;
  const takenNames = new Set(parts.keys());
  const nextSlideName = (): string => {
    let name: string;
    do {
      nextNumber += 1;
      name = `${slidesDir}/slide${nextNumber}.xml`;
    } while (takenNames.has(name));
    takenNames.add(name);
    return name;
  };

  // Assign each model slide an emitted part name and kind, in model order.
  const usedSource = new Set<string>();
  const emitted: Emitted[] = model.slides.map((slide) => {
    if (slide.partName && parts.has(slide.partName) && !usedSource.has(slide.partName)) {
      usedSource.add(slide.partName);
      return { slide, emitName: slide.partName, kind: "verbatim" };
    }
    if (slide.partName && parts.has(slide.partName)) {
      return { slide, emitName: nextSlideName(), kind: "duplicate", sourceName: slide.partName };
    }
    return { slide, emitName: nextSlideName(), kind: "new" };
  });

  // Deleted slides: their part, their .rels, and the notesSlide they alone
  // referenced (shared targets like layouts/media stay).
  const removed = new Set<string>();
  for (const path of origSlidePaths) {
    if (usedSource.has(path)) {
      continue;
    }
    removed.add(path);
    removed.add(relsPathFor(path));
    const rels = parseRelationships(decode(parts.get(relsPathFor(path)) ?? encode("")));
    for (const rel of rels.values()) {
      if (rel.type === NOTES_SLIDE_REL) {
        const notesPath = resolvePartPath(path, rel.target);
        removed.add(notesPath);
        removed.add(relsPathFor(notesPath));
      }
    }
  }

  // Rebuild presentation.xml.rels: non-slide rels kept with original ids;
  // slide rels renumbered after the highest surviving id, in model order.
  const keptRels = [...origPresRels.values()].filter((rel) => rel.type !== SLIDE_REL);
  const maxSurviving = keptRels.reduce((max, rel) => Math.max(max, relIdNumber(rel.id)), 0);
  let nextRelId = maxSurviving + 1;
  const slideRelXml = emitted
    .map((entry) => {
      entry.relId = `rId${nextRelId++}`;
      return `<Relationship Id="${entry.relId}" Type="${SLIDE_REL}" Target="${escapeXmlAttribute(relativeTarget(presPath, entry.emitName))}"/>`;
    })
    .join("");
  const keptRelXml = keptRels
    .map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXmlAttribute(rel.target)}"/>`)
    .join("");
  const rebuiltPresRels = `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${keptRelXml}${slideRelXml}</Relationships>`;

  // Rebuild presentation.xml: splice a fresh p:sldIdLst (ids re-minted 256…).
  const sldIdLstInner = emitted.map((entry, index) => `<p:sldId id="${256 + index}" r:id="${entry.relId}"/>`).join("");
  const rebuiltPresXml = spliceSldIdLst(origPresXml, `<p:sldIdLst>${sldIdLstInner}</p:sldIdLst>`);

  // Rebuild [Content_Types].xml: add overrides for added slides, drop overrides
  // for removed parts; everything else byte-verbatim. A dirty slide re-emitted
  // under its own name (kind "verbatim") already has an override — only genuinely
  // new part names are added.
  const contentTypesXml = decode(parts.get("[Content_Types].xml") ?? encode(""));
  const media = planMedia(model, parts, contentTypesXml);
  const addedSlidePaths = emitted.filter((entry) => entry.kind !== "verbatim").map((entry) => entry.emitName);
  const rebuiltContentTypes = spliceContentDefaults(
    rebuildContentTypes(contentTypesXml, addedSlidePaths, [...removed]),
    media.defaults,
  );

  // A dirty slide kept under its own part name is rebuilt in place; one that
  // gained an inserted image also has its `.rels` re-emitted with the new image
  // relationship spliced in. An untouched slide re-emits byte-verbatim.
  const rebuiltOwn = new Map<string, Uint8Array>();
  for (const entry of emitted) {
    if (entry.kind === "verbatim" && entry.slide.dirty) {
      const relsPath = relsPathFor(entry.emitName);
      const relsXml = decode(parts.get(relsPath) ?? encode(""));
      const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();
      const newRels = linkImageRels(entry.slide, rels, entry.emitName);
      rebuiltOwn.set(entry.emitName, encode(slideXml(entry.slide, rels, entry.emitName)));
      if (newRels.length > 0) {
        rebuiltOwn.set(relsPath, encode(appendRelationships(relsXml, newRels)));
      }
    }
  }

  // Assemble: iterate opaque parts in order, substituting the three structural
  // parts, rebuilding dirty own-name slides, and skipping removed parts; then
  // append new/duplicated slide parts.
  const entries: ZipEntryInput[] = [];
  for (const [name, data] of parts) {
    if (removed.has(name)) {
      continue;
    }
    if (name === presPath) {
      entries.push({ name, data: encode(rebuiltPresXml) });
    } else if (name === presRelsPath) {
      entries.push({ name, data: encode(rebuiltPresRels) });
    } else if (name === "[Content_Types].xml") {
      entries.push({ name, data: encode(rebuiltContentTypes) });
    } else {
      entries.push({ name, data: rebuiltOwn.get(name) ?? media.override.get(name) ?? data });
    }
  }
  for (const entry of emitted) {
    if (entry.kind === "verbatim") {
      continue;
    }
    const relsName = relsPathFor(entry.emitName);
    if (entry.kind === "duplicate" && entry.sourceName) {
      // A duplicated-then-edited slide is rebuilt from the model; its rels are
      // still the source's (carried, ids intact) plus any inserted image. An
      // unedited duplicate copies the source part's bytes verbatim (AC-33).
      const sourceRelsBytes = parts.get(relsPathFor(entry.sourceName));
      const sourceRelsXml = sourceRelsBytes ? decode(sourceRelsBytes) : "";
      const rels = sourceRelsXml ? parseRelationships(sourceRelsXml) : new Map<string, Relationship>();
      const newRels = entry.slide.dirty ? linkImageRels(entry.slide, rels, entry.emitName) : [];
      const xml = entry.slide.dirty
        ? encode(slideXml(entry.slide, rels, entry.emitName))
        : parts.get(entry.sourceName)!;
      entries.push({ name: entry.emitName, data: xml });
      if (sourceRelsBytes) {
        entries.push({ name: relsName, data: encode(appendRelationships(sourceRelsXml, newRels)) });
      } else if (newRels.length > 0) {
        entries.push({ name: relsName, data: encode(appendRelationships("", newRels)) });
      }
    } else {
      const layoutTarget = relativeTarget(entry.emitName, resolveLayoutPath(model, entry, parts, slidesDir, presPath));
      const rels = new Map<string, Relationship>();
      rels.set("rId1", { id: "rId1", type: SLIDE_LAYOUT_REL, target: layoutTarget });
      const imageRels = linkImageRels(entry.slide, rels, entry.emitName);
      entries.push({ name: entry.emitName, data: encode(slideXml(entry.slide, rels, entry.emitName)) });
      entries.push({ name: relsName, data: encode(appendRelationships(slideRelsXml(layoutTarget), imageRels)) });
    }
  }
  entries.push(...media.appended);
  return entries;
}

function relIdNumber(id: string): number {
  const match = /(\d+)$/.exec(id);
  return match ? Number(match[1]) : 0;
}

/** The layout part a new slide references (AC-27): the current slide's layout
 * — the nearest preceding model slide that maps to an existing part — else
 * the deck's first layout. Returns an absolute part path. */
function resolveLayoutPath(
  model: PresentationModel,
  entry: Emitted,
  parts: Map<string, Uint8Array>,
  slidesDir: string,
  presPath: string,
): string {
  const index = model.slides.indexOf(entry.slide);
  for (let i = index - 1; i >= 0; i--) {
    const neighbour = model.slides[i];
    if (neighbour.partName && parts.has(relsPathFor(neighbour.partName))) {
      const rels = parseRelationships(decode(parts.get(relsPathFor(neighbour.partName))!));
      const layoutRel = [...rels.values()].find((rel) => rel.type === SLIDE_LAYOUT_REL);
      if (layoutRel) {
        return resolvePartPath(neighbour.partName, layoutRel.target);
      }
    }
  }
  const layouts = [...parts.keys()].filter((name) => /\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).sort();
  if (layouts[0]) {
    return layouts[0];
  }
  void slidesDir;
  return `${presPath.split("/").slice(0, -1).join("/")}/slideLayouts/slideLayout1.xml`;
}

function spliceSldIdLst(xml: string, replacement: string): string {
  const re = /<p:sldIdLst(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/p:sldIdLst>)/;
  if (re.test(xml)) {
    return xml.replace(re, replacement);
  }
  return xml.replace(/(<p:sldSz)/, `${replacement}$1`);
}

function rebuildContentTypes(xml: string, addedSlidePaths: string[], removedPaths: string[]): string {
  let result = xml;
  for (const path of removedPaths) {
    const escaped = `/${path}`.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    result = result.replace(new RegExp(`<Override PartName="${escaped}"[^>]*/>`), "");
  }
  const additions = addedSlidePaths
    .map((path) => `<Override PartName="/${path}" ContentType="${SLIDE_CONTENT_TYPE}"/>`)
    .join("");
  return result.replace(/(<\/Types>\s*)$/, `${additions}$1`);
}

// --- from-scratch master / layout / theme (fixed minimal Office look) -------

const SLIDE_MASTER_XML =
  `${XML_DECL}<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
  `<p:txStyles>` +
  `<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
  `<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>` +
  `</p:txStyles></p:sldMaster>`;

const SLIDE_MASTER_RELS =
  `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
  `<Relationship Id="rId1" Type="${SLIDE_LAYOUT_REL}" Target="../slideLayouts/slideLayout1.xml"/>` +
  `<Relationship Id="rId2" Type="${THEME_REL}" Target="../theme/theme1.xml"/>` +
  `</Relationships>`;

const SLIDE_LAYOUT_XML =
  `${XML_DECL}<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const SLIDE_LAYOUT_RELS =
  `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
  `<Relationship Id="rId1" Type="${SLIDE_MASTER_REL}" Target="../slideMasters/slideMaster1.xml"/>` +
  `</Relationships>`;

const THEME_XML =
  `${XML_DECL}<a:theme xmlns:a="${A_NS}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
  `<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>` +
  `<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
  `</a:themeElements></a:theme>`;
