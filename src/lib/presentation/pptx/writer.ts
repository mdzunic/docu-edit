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
import { buildCorePropertiesXml, type CorePropertiesInput } from "../../core/opcProperties";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import type { MediaPart } from "../../core/media";
import {
  DEFAULT_PRESENTATION_TITLE,
  type LayoutSlot,
  type PresentationModel,
  type Slide,
  type SlideFrame,
  type SlideLayout,
} from "../model";
import { classifyFromSlots, STANDARD_LAYOUTS, standardLayoutKey, synthesiseStandardLayouts } from "../layouts";
import { PPTX_MAIN_CONTENT_TYPE } from "./reader";
import { serializeSlideXml } from "./serialize";
import { applyNotesContentTypes, planNotes } from "./notesWriter";

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
  const anyNotes = model.slides.some((slide) => slide.notesDirty === true);
  const deckDirty = model.deckDirty === true;
  if (!model.structureDirty && !deckDirty && !anyDirty && !anyNotes) {
    // Nothing changed — re-emit verbatim (AC-51).
    const bytes = await zip([...parts].map(([name, data]) => ({ name, data })));
    return { bytes };
  }
  if (!model.structureDirty && !deckDirty) {
    // Text-only edits: substitute each dirty slide's own part in place —
    // presentation.xml, its rels and [Content_Types].xml stay byte-verbatim,
    // and every untouched slide/unmodelled part is copied through (SPEC-081
    // Dirtying: a text edit does not rebuild the structural parts).
    return { bytes: await zip(buildTextEditSubstitution(model, parts)) };
  }
  // A deck-level change (size / first slide number / metadata, SPEC-090) rides
  // the substitution path too: it rebuilds presentation.xml (p:sldSz +
  // firstSlideNum) and docProps/core.xml.
  return { bytes: await zip(buildSubstitution(model, parts, options.modified ?? new Date())) };
}

/** Rebuilds only the dirty slides' own parts, in place; every other part
 * (structural parts, untouched slides) is copied byte-verbatim. A dirty slide
 * that gained an inserted image also has its `.rels` re-emitted (with the new
 * image relationship spliced in), the media part added under `ppt/media/`, and
 * the Content-Types extension default added if missing (SPEC-082). */
function buildTextEditSubstitution(model: PresentationModel, parts: Map<string, Uint8Array>): ZipEntryInput[] {
  const contentTypesXml = decode(parts.get("[Content_Types].xml") ?? encode(""));
  const media = planMedia(model, parts, contentTypesXml);
  const rootRels = parseRelationships(decode(parts.get("_rels/.rels") ?? encode("")));
  const officeRel = [...rootRels.values()].find((rel) => rel.type === OFFICE_DOCUMENT_REL);
  const presPath = officeRel ? resolvePartPath("", officeRel.target) : "ppt/presentation.xml";
  // Layout changes (SPEC-089): a slide whose layoutKey differs from its source
  // .rels has only that .rels repointed; a synthesised layout is authored once.
  const plan = planLayouts(model, parts, presPath);
  const rebuilt = new Map<string, Uint8Array>();
  for (const slide of model.slides) {
    if (slide.dirty && slide.partName && parts.has(slide.partName)) {
      const relsPath = relsPathFor(slide.partName);
      const relsXml = decode(parts.get(relsPath) ?? encode(""));
      const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();
      const newRels = linkImageRels(slide, rels, slide.partName);
      rebuilt.set(slide.partName, encode(serializeSlideXml(slide, rels, slide.partName)));
      const outRels = relsForLayoutChange(slide, slide.partName, relsXml, newRels, plan);
      if (outRels !== undefined) {
        rebuilt.set(relsPath, encode(outRels));
      }
    }
  }

  // Notes surgery (SPEC-083): the slide part itself stays byte-verbatim when
  // only notes changed — only the notes-slide parts (and their machinery) move.
  const notes = planNotes({
    model,
    parts,
    slidePartFor: (slide) => (slide.partName && parts.has(slide.partName) ? slide.partName : undefined),
    slideRelsXml: (slide) => {
      const relsPath = relsPathFor(slide.partName!);
      const fromRebuilt = rebuilt.get(relsPath);
      if (fromRebuilt) {
        return decode(fromRebuilt);
      }
      const orig = parts.get(relsPath);
      return orig ? decode(orig) : "";
    },
    presPath,
    presXml: decode(parts.get(presPath) ?? encode("")),
    presRelsXml: decode(parts.get(relsPathFor(presPath)) ?? encode("")),
    contentTypesXml,
  });

  let combinedContentTypes = applyNotesContentTypes(contentTypesXml, notes);
  if (media.defaults) {
    combinedContentTypes = spliceContentDefaults(combinedContentTypes, media.defaults);
  }
  combinedContentTypes = spliceLayoutOverrides(combinedContentTypes, plan.overrides);

  const entries: ZipEntryInput[] = [];
  for (const [name, data] of parts) {
    if (notes.removed.has(name)) {
      continue;
    }
    if (name === "[Content_Types].xml") {
      entries.push({ name, data: encode(combinedContentTypes) });
    } else if (notes.parts.has(name)) {
      entries.push({ name, data: notes.parts.get(name)! });
    } else if (rebuilt.has(name)) {
      entries.push({ name, data: rebuilt.get(name)! });
    } else if (plan.masterXml.has(name)) {
      entries.push({ name, data: encode(plan.masterXml.get(name)!) });
    } else if (plan.masterRelsXml.has(name)) {
      entries.push({ name, data: encode(plan.masterRelsXml.get(name)!) });
    } else if (media.override.has(name)) {
      entries.push({ name, data: media.override.get(name)! });
    } else {
      entries.push({ name, data });
    }
  }
  // New notes parts (notes-slide, its rels, a created notes master + rels).
  for (const [name, data] of notes.parts) {
    if (!parts.has(name)) {
      entries.push({ name, data });
    }
  }
  // Authored layout parts + their .rels (SPEC-089 AC-14).
  for (const [name, data] of plan.newParts) {
    if (!parts.has(name)) {
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

/** `relsXml` with any notesSlide relationship removed. A duplicated slide must
 * not inherit the SOURCE's notesSlide target (that would make two model slides
 * share one on-disk notes part); the notes planner mints the duplicate its own. */
function stripNotesSlideRel(relsXml: string): string {
  return relsXml.replace(new RegExp(`<Relationship\\b[^>]*\\bType="${NOTES_SLIDE_REL}"[^>]*/>`, "g"), "");
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

/** The standard-set index a from-scratch slide references: its `layoutKey` when
 * it names a standard layout, else the kind inferred from its placeholder roles
 * (so a title+body slide references the Title-and-Content layout). */
function standardLayoutIndexForSlide(slide: Slide): number {
  if (slide.layoutKey) {
    const named = STANDARD_LAYOUTS.findIndex((d) => standardLayoutKey(d.kind) === slide.layoutKey);
    if (named >= 0) {
      return named;
    }
  }
  const slots: LayoutSlot[] = slide.frames
    .filter((f): f is Slide["frames"][number] & { kind: "text" } => f.kind === "text" && f.role !== undefined)
    .map((f) => ({ role: f.role!, xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 }));
  const kind = classifyFromSlots(slots);
  const inferred = STANDARD_LAYOUTS.findIndex((d) => d.kind === kind);
  return inferred >= 0 ? inferred : 0;
}

function buildFromScratch(model: PresentationModel, modified: Date): ZipEntryInput[] {
  const slidePaths = model.slides.map((_, index) => `ppt/slides/slide${index + 1}.xml`);
  // The minimal standard set (SPEC-089 AC-17): slideLayout1..6, in the Reference
  // table order (title/obj/secHead/twoObj/titleOnly/blank), at the deck size.
  const standardLayouts = synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm);
  const layoutPaths = standardLayouts.map((_, i) => `ppt/slideLayouts/slideLayout${i + 1}.xml`);

  const mediaDefaults = mediaContentDefaults(referencedMedia(model));
  const contentTypes =
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    mediaDefaults +
    `<Override PartName="/ppt/presentation.xml" ContentType="${PPTX_MAIN_CONTENT_TYPE}"/>` +
    slidePaths.map((path) => `<Override PartName="/${path}" ContentType="${SLIDE_CONTENT_TYPE}"/>`).join("") +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    layoutPaths.map((path) => `<Override PartName="/${path}" ContentType="${SLIDE_LAYOUT_CONTENT_TYPE}"/>`).join("") +
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
  const firstSlideNumAttr = model.firstSlideNumber !== undefined ? ` firstSlideNum="${model.firstSlideNumber}"` : "";
  const presentation =
    `${XML_DECL}<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"${firstSlideNumAttr}>` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${sldIdLst}</p:sldIdLst>` +
    `<p:sldSz cx="${cmToEmu(model.slideWidthCm)}" cy="${cmToEmu(model.slideHeightCm)}"/>` +
    `<p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;

  const core = buildCorePropertiesXml(coreInputFor(model), modified);

  // The master lists six layouts (rId1..rId6), theme at rId7.
  const masterLayoutIdLst =
    `<p:sldLayoutIdLst>` +
    standardLayouts.map((_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join("") +
    `</p:sldLayoutIdLst>`;
  const masterXml = SLIDE_MASTER_XML.replace(
    /<p:sldLayoutIdLst>.*?<\/p:sldLayoutIdLst>/,
    masterLayoutIdLst,
  );
  const masterRels =
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    standardLayouts
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${SLIDE_LAYOUT_REL}" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`)
      .join("") +
    `<Relationship Id="rId${standardLayouts.length + 1}" Type="${THEME_REL}" Target="../theme/theme1.xml"/>` +
    `</Relationships>`;

  const entries: ZipEntryInput[] = [
    { name: "[Content_Types].xml", data: encode(contentTypes) },
    { name: "_rels/.rels", data: encode(rootRels) },
    { name: "docProps/core.xml", data: encode(core) },
    { name: "ppt/presentation.xml", data: encode(presentation) },
    { name: "ppt/_rels/presentation.xml.rels", data: encode(presentationRels) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: encode(masterXml) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: encode(masterRels) },
    { name: "ppt/theme/theme1.xml", data: encode(THEME_XML) },
  ];
  standardLayouts.forEach((layout, i) => {
    entries.push({ name: layoutPaths[i], data: encode(authorLayoutXml(layout)) });
    entries.push({
      name: relsPathFor(layoutPaths[i]),
      data: encode(authorLayoutRels(layoutPaths[i], "ppt/slideMasters/slideMaster1.xml")),
    });
  });
  model.slides.forEach((slide, index) => {
    const slidePath = slidePaths[index];
    const layoutTarget = `../slideLayouts/slideLayout${standardLayoutIndexForSlide(slide) + 1}.xml`;
    const rels = new Map<string, Relationship>();
    rels.set("rId1", { id: "rId1", type: SLIDE_LAYOUT_REL, target: layoutTarget });
    const imageRels = linkImageRels(slide, rels, slidePath);
    entries.push({ name: slidePath, data: encode(slideXml(slide, rels, slidePath)) });
    entries.push({
      name: `ppt/slides/_rels/slide${index + 1}.xml.rels`,
      data: encode(appendRelationships(slideRelsXml(layoutTarget), imageRels)),
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

function buildSubstitution(model: PresentationModel, parts: Map<string, Uint8Array>, modified: Date = new Date()): ZipEntryInput[] {
  // Locate the presentation part through the root relationships.
  const rootRels = parseRelationships(decode(parts.get("_rels/.rels") ?? encode("")));
  const officeRel = [...rootRels.values()].find((rel) => rel.type === OFFICE_DOCUMENT_REL);
  const presPath = officeRel ? resolvePartPath("", officeRel.target) : "ppt/presentation.xml";
  // A deck-level change (SPEC-090) rewrites the slide size + first-slide-number
  // in presentation.xml and rebuilds docProps/core.xml from the deck metadata.
  const deckDirty = model.deckDirty === true;
  const coreRel = [...rootRels.values()].find((rel) => rel.type === CORE_PROPERTIES_REL);
  const corePath = coreRel ? resolvePartPath("", coreRel.target) : undefined;
  const deckCoreXml = deckDirty && corePath ? buildCorePropertiesXml(coreInputFor(model), modified) : undefined;
  const presRelsPath = relsPathFor(presPath);
  const origPresXml = decode(parts.get(presPath) ?? encode(""));
  const origPresRels = parseRelationships(decode(parts.get(presRelsPath) ?? encode("")));
  // Layout changes (SPEC-089): resolve every referenced layoutKey to a part
  // (existing or authored) and repoint/author accordingly.
  const plan = planLayouts(model, parts, presPath);

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

  // Rebuild presentation.xml: splice a fresh p:sldIdLst (ids re-minted 256…),
  // then (deck-level change only) the slide size and first-slide-number origin.
  const sldIdLstInner = emitted.map((entry, index) => `<p:sldId id="${256 + index}" r:id="${entry.relId}"/>`).join("");
  let rebuiltPresXml = spliceSldIdLst(origPresXml, `<p:sldIdLst>${sldIdLstInner}</p:sldIdLst>`);
  if (deckDirty) {
    rebuiltPresXml = spliceSldSz(rebuiltPresXml, model);
    rebuiltPresXml = applyFirstSlideNum(rebuiltPresXml, model.firstSlideNumber);
  }

  // Rebuild [Content_Types].xml: add overrides for added slides, drop overrides
  // for removed parts; everything else byte-verbatim. A dirty slide re-emitted
  // under its own name (kind "verbatim") already has an override — only genuinely
  // new part names are added.
  const contentTypesXml = decode(parts.get("[Content_Types].xml") ?? encode(""));
  const media = planMedia(model, parts, contentTypesXml);
  const addedSlidePaths = emitted.filter((entry) => entry.kind !== "verbatim").map((entry) => entry.emitName);
  const rebuiltContentTypes = spliceLayoutOverrides(
    spliceContentDefaults(rebuildContentTypes(contentTypesXml, addedSlidePaths, [...removed]), media.defaults),
    plan.overrides,
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
      const outRels = relsForLayoutChange(entry.slide, entry.emitName, relsXml, newRels, plan);
      if (outRels !== undefined) {
        rebuiltOwn.set(relsPath, encode(outRels));
      }
    }
  }

  // Precompute the emitted bytes and base `.rels` XML for every slide emitted
  // under a fresh part (added & duplicated slides), so the notes planner sees
  // the exact rels the writer will emit and can splice a notesSlide rel in. A
  // duplicate inherits the source's rels; any notesSlide relationship there
  // points at the SOURCE's notes part, so it is stripped and the planner mints
  // the duplicate its own part (two model slides never share one notes part).
  const baseSlideParts = new Map<string, Uint8Array>();
  const baseSlideRels = new Map<string, string>();
  for (const entry of emitted) {
    if (entry.kind === "verbatim") {
      continue;
    }
    const relsName = relsPathFor(entry.emitName);
    if (entry.kind === "duplicate" && entry.sourceName) {
      const sourceRelsBytes = parts.get(relsPathFor(entry.sourceName));
      const sourceRelsXml = sourceRelsBytes ? stripNotesSlideRel(decode(sourceRelsBytes)) : "";
      const rels = sourceRelsXml ? parseRelationships(sourceRelsXml) : new Map<string, Relationship>();
      const newRels = entry.slide.dirty ? linkImageRels(entry.slide, rels, entry.emitName) : [];
      const xml = entry.slide.dirty
        ? encode(slideXml(entry.slide, rels, entry.emitName))
        : parts.get(entry.sourceName)!;
      baseSlideParts.set(entry.emitName, xml);
      if (sourceRelsBytes) {
        baseSlideRels.set(relsName, appendRelationships(sourceRelsXml, newRels));
      } else if (newRels.length > 0) {
        baseSlideRels.set(relsName, appendRelationships("", newRels));
      }
    } else {
      const layoutTarget = relativeTarget(entry.emitName, resolveLayoutPath(model, entry, parts, slidesDir, presPath, plan));
      const rels = new Map<string, Relationship>();
      rels.set("rId1", { id: "rId1", type: SLIDE_LAYOUT_REL, target: layoutTarget });
      const imageRels = linkImageRels(entry.slide, rels, entry.emitName);
      baseSlideParts.set(entry.emitName, encode(slideXml(entry.slide, rels, entry.emitName)));
      baseSlideRels.set(relsName, appendRelationships(slideRelsXml(layoutTarget), imageRels));
    }
  }

  // Notes surgery (SPEC-083) for EVERY model slide whose notes changed — the
  // notes parts move while a verbatim slide part stays byte-verbatim, and a
  // freshly-emitted (added/duplicated) slide gets its own notes part too. A
  // created notes master splices into the already-rebuilt presentation.xml.
  const emitNameBySlide = new Map<Slide, string>();
  const freshSlides = new Set<Slide>();
  for (const entry of emitted) {
    emitNameBySlide.set(entry.slide, entry.emitName);
    if (entry.kind !== "verbatim") {
      freshSlides.add(entry.slide);
    }
  }
  const notes = planNotes({
    model,
    parts,
    freshSlides,
    slidePartFor: (slide) => emitNameBySlide.get(slide),
    slideRelsXml: (slide) => {
      const path = emitNameBySlide.get(slide);
      if (!path) {
        return "";
      }
      const relsPath = relsPathFor(path);
      const own = rebuiltOwn.get(relsPath);
      if (own) {
        return decode(own);
      }
      const base = baseSlideRels.get(relsPath);
      if (base !== undefined) {
        return base;
      }
      const orig = parts.get(relsPath);
      return orig ? decode(orig) : "";
    },
    presPath,
    presXml: rebuiltPresXml,
    presRelsXml: rebuiltPresRels,
    contentTypesXml: rebuiltContentTypes,
  });
  const finalPresXml = notes.parts.get(presPath) ? decode(notes.parts.get(presPath)!) : rebuiltPresXml;
  const finalPresRels = notes.parts.get(presRelsPath) ? decode(notes.parts.get(presRelsPath)!) : rebuiltPresRels;
  const finalContentTypes = applyNotesContentTypes(rebuiltContentTypes, notes);

  // Assemble: iterate opaque parts in order, substituting the three structural
  // parts, rebuilding dirty own-name slides, and skipping removed parts; then
  // append new/duplicated slide parts.
  const entries: ZipEntryInput[] = [];
  for (const [name, data] of parts) {
    if (removed.has(name) || notes.removed.has(name)) {
      continue;
    }
    if (name === presPath) {
      entries.push({ name, data: encode(finalPresXml) });
    } else if (name === presRelsPath) {
      entries.push({ name, data: encode(finalPresRels) });
    } else if (name === "[Content_Types].xml") {
      entries.push({ name, data: encode(finalContentTypes) });
    } else if (name === corePath && deckCoreXml) {
      // docProps/core.xml rebuilt from the deck metadata (SPEC-090); app.xml
      // and custom.xml stay byte-verbatim.
      entries.push({ name, data: encode(deckCoreXml) });
    } else if (notes.parts.has(name)) {
      entries.push({ name, data: notes.parts.get(name)! });
    } else if (plan.masterXml.has(name)) {
      entries.push({ name, data: encode(plan.masterXml.get(name)!) });
    } else if (plan.masterRelsXml.has(name)) {
      entries.push({ name, data: encode(plan.masterRelsXml.get(name)!) });
    } else {
      entries.push({ name, data: rebuiltOwn.get(name) ?? media.override.get(name) ?? data });
    }
  }
  const emittedFresh = new Set<string>();
  for (const entry of emitted) {
    if (entry.kind === "verbatim") {
      continue;
    }
    const relsName = relsPathFor(entry.emitName);
    // The slide part was precomputed above (bytes copied verbatim for an
    // unedited duplicate, serialised from the model otherwise).
    entries.push({ name: entry.emitName, data: baseSlideParts.get(entry.emitName)! });
    emittedFresh.add(entry.emitName);
    // Its rels: the notes planner's version (a notesSlide rel spliced in) wins
    // over the base rels; a slide with neither has no rels part.
    const relsData = notes.parts.get(relsName) ?? (baseSlideRels.has(relsName) ? encode(baseSlideRels.get(relsName)!) : undefined);
    if (relsData) {
      entries.push({ name: relsName, data: relsData });
      emittedFresh.add(relsName);
    }
  }
  // New notes parts (notes-slide, its rels, a created notes master + rels) not
  // already emitted above with the fresh slide parts.
  for (const [name, data] of notes.parts) {
    if (!parts.has(name) && !emittedFresh.has(name)) {
      entries.push({ name, data });
    }
  }
  // Authored layout parts + their .rels (SPEC-089 AC-14).
  for (const [name, data] of plan.newParts) {
    if (!parts.has(name) && !emittedFresh.has(name)) {
      entries.push({ name, data });
    }
  }
  entries.push(...media.appended);
  return entries;
}

function relIdNumber(id: string): number {
  const match = /(\d+)$/.exec(id);
  return match ? Number(match[1]) : 0;
}

// ── Layout catalogue writing (SPEC-089) ────────────────────────────────────
const SLIDE_LAYOUT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

/** The `p:sldLayout/@type` for a modelled kind when authoring a missing layout
 * part, from the standard table; unknown kinds fall back to `obj`. */
function pptxTypeForKind(kind: string): string {
  return STANDARD_LAYOUTS.find((d) => d.kind === kind)?.pptxType ?? "obj";
}

/** The absolute layout part a slide's `.rels` currently targets. */
function currentLayoutTarget(relsXml: string, slidePartName: string): string | undefined {
  if (!relsXml) {
    return undefined;
  }
  const rel = [...parseRelationships(relsXml).values()].find((r) => r.type === SLIDE_LAYOUT_REL);
  return rel ? resolvePartPath(slidePartName, rel.target) : undefined;
}

/** Rewrites ONLY the slideLayout relationship's `Target` to point at
 * `newLayoutPath`, keeping every other relationship byte-identical (AC-13). */
function repointLayoutRel(relsXml: string, slidePartName: string, newLayoutPath: string): string {
  const target = escapeXmlAttribute(relativeTarget(slidePartName, newLayoutPath));
  return relsXml.replace(/<Relationship\b[^>]*Type="[^"]*\/slideLayout"[^>]*\/>/, (el) =>
    el.replace(/Target="[^"]*"/, `Target="${target}"`),
  );
}

/** One authored placeholder `p:sp` for a layout slot. */
function authorLayoutPlaceholder(slot: LayoutSlot, id: number): string {
  const ph = `<p:ph${slot.phType ? ` type="${escapeXmlAttribute(slot.phType)}"` : ""}${slot.phIdx !== undefined ? ` idx="${slot.phIdx}"` : ""}/>`;
  const xfrm =
    `<a:xfrm><a:off x="${cmToEmu(slot.xCm)}" y="${cmToEmu(slot.yCm)}"/>` +
    `<a:ext cx="${cmToEmu(slot.widthCm)}" cy="${cmToEmu(slot.heightCm)}"/></a:xfrm>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`
  );
}

/** A whole authored `p:sldLayout` part for a catalogue entry with no counterpart. */
function authorLayoutXml(layout: SlideLayout): string {
  const shapes = layout.slots.map((slot, i) => authorLayoutPlaceholder(slot, i + 2)).join("");
  return (
    `${XML_DECL}<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="${pptxTypeForKind(layout.kind)}" preserve="1">` +
    `<p:cSld name="${escapeXmlAttribute(layout.name)}"><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

function authorLayoutRels(layoutPath: string, masterPath: string): string {
  return (
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${SLIDE_MASTER_REL}" Target="${escapeXmlAttribute(relativeTarget(layoutPath, masterPath))}"/>` +
    `</Relationships>`
  );
}

/** The first slide master's absolute part path, via presentation.xml.rels. */
function firstMasterPath(parts: Map<string, Uint8Array>, presPath: string): string | undefined {
  const rels = parseRelationships(decode(parts.get(relsPathFor(presPath)) ?? encode("")));
  const rel = [...rels.values()].find((r) => r.type === SLIDE_MASTER_REL);
  return rel ? resolvePartPath(presPath, rel.target) : undefined;
}

/** The whole-writer plan for layout changes: resolving every referenced
 * `layoutKey` to a part path (existing or freshly authored), the authored parts,
 * the Content-Types overrides they need, and the master edits (a `p:sldLayoutId`
 * spliced into `p:sldLayoutIdLst` and a matching master `.rels` relationship).
 * A deck whose layouts were untouched produces an empty-effect plan. */
interface LayoutPlan {
  pathForKey: Map<string, string>;
  newParts: Map<string, Uint8Array>;
  overrides: string[];
  masterXml: Map<string, string>;
  masterRelsXml: Map<string, string>;
}

function planLayouts(model: PresentationModel, parts: Map<string, Uint8Array>, presPath: string): LayoutPlan {
  const plan: LayoutPlan = {
    pathForKey: new Map(),
    newParts: new Map(),
    overrides: [],
    masterXml: new Map(),
    masterRelsXml: new Map(),
  };
  const anyLayout = [...parts.keys()].find((n) => /\/slideLayouts\/slideLayout\d+\.xml$/.test(n));
  const layoutDir = anyLayout
    ? anyLayout.split("/").slice(0, -1).join("/")
    : `${presPath.split("/").slice(0, -1).join("/")}/slideLayouts`;
  let maxN = 0;
  for (const name of parts.keys()) {
    const m = /\/slideLayout(\d+)\.xml$/.exec(name);
    if (m) {
      maxN = Math.max(maxN, Number(m[1]));
    }
  }
  const masterPath = firstMasterPath(parts, presPath);
  const keys = new Set<string>();
  for (const slide of model.slides) {
    if (slide.layoutKey) {
      keys.add(slide.layoutKey);
    }
  }
  for (const key of keys) {
    if (parts.has(key)) {
      plan.pathForKey.set(key, key);
      continue;
    }
    const entry = (model.layouts ?? []).find((l) => l.key === key);
    if (!entry || !masterPath) {
      continue; // unknown key or masterless deck: never repaired (AC-1).
    }
    maxN += 1;
    const path = `${layoutDir}/slideLayout${maxN}.xml`;
    plan.newParts.set(path, encode(authorLayoutXml(entry)));
    plan.newParts.set(relsPathFor(path), encode(authorLayoutRels(path, masterPath)));
    plan.overrides.push(`<Override PartName="/${path}" ContentType="${SLIDE_LAYOUT_CONTENT_TYPE}"/>`);
    plan.pathForKey.set(key, path);

    // Splice the layout into the master: a new relationship + a p:sldLayoutId.
    const masterRelsPath = relsPathFor(masterPath);
    const masterRels = plan.masterRelsXml.get(masterRelsPath) ?? decode(parts.get(masterRelsPath) ?? encode(""));
    const relIds = [...parseRelationships(masterRels).values()].map((r) => relIdNumber(r.id));
    const nextRid = `rId${(relIds.length ? Math.max(...relIds) : 0) + 1}`;
    const layoutRelXml = `<Relationship Id="${nextRid}" Type="${SLIDE_LAYOUT_REL}" Target="${escapeXmlAttribute(relativeTarget(masterPath, path))}"/>`;
    plan.masterRelsXml.set(
      masterRelsPath,
      /<\/Relationships>/.test(masterRels)
        ? masterRels.replace(/<\/Relationships>/, `${layoutRelXml}</Relationships>`)
        : `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${layoutRelXml}</Relationships>`,
    );

    const masterXml = plan.masterXml.get(masterPath) ?? decode(parts.get(masterPath) ?? encode(""));
    const existingIds = [...masterXml.matchAll(/<p:sldLayoutId id="(\d+)"/g)].map((m) => Number(m[1]));
    const nextId = (existingIds.length ? Math.max(...existingIds) : 2147483648) + 1;
    const idXml = `<p:sldLayoutId id="${nextId}" r:id="${nextRid}"/>`;
    plan.masterXml.set(
      masterPath,
      /<\/p:sldLayoutIdLst>/.test(masterXml)
        ? masterXml.replace(/<\/p:sldLayoutIdLst>/, `${idXml}</p:sldLayoutIdLst>`)
        : masterXml.replace(/<p:sldLayoutIdLst\/>/, `<p:sldLayoutIdLst>${idXml}</p:sldLayoutIdLst>`),
    );
  }
  return plan;
}

/** Splices layout Content-Types overrides before `</Types>` (authored layouts). */
function spliceLayoutOverrides(contentTypesXml: string, overrides: string[]): string {
  if (overrides.length === 0) {
    return contentTypesXml;
  }
  return contentTypesXml.replace(/<\/Types>/, `${overrides.join("")}</Types>`);
}

/** The `.rels` bytes for a dirty slide whose `layoutKey` changed: repointed at
 * the plan's target, with image relationships appended. Returns undefined when
 * nothing changed. */
function relsForLayoutChange(
  slide: Slide,
  slidePartName: string,
  relsXml: string,
  newImageRels: { id: string; type: string; target: string }[],
  plan: LayoutPlan,
): string | undefined {
  let out = relsXml;
  const desired = slide.layoutKey ? plan.pathForKey.get(slide.layoutKey) : undefined;
  if (desired && currentLayoutTarget(relsXml, slidePartName) !== desired) {
    out = repointLayoutRel(out, slidePartName, desired);
  }
  if (newImageRels.length > 0) {
    out = appendRelationships(out, newImageRels);
  }
  return out === relsXml ? undefined : out;
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
  plan: LayoutPlan,
): string {
  // SPEC-089: a slide that names a layout honours it (existing or authored),
  // ahead of the nearest-preceding-slide heuristic.
  if (entry.slide.layoutKey) {
    const path = plan.pathForKey.get(entry.slide.layoutKey);
    if (path) {
      return path;
    }
  }
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

/** The docProps/core.xml input for the deck's metadata — the deck's `properties`
 * with `title` falling back to the app-bar title (SPEC-090). Shared by the
 * from-scratch and substitution paths so both emit the same shape. */
function coreInputFor(model: PresentationModel): CorePropertiesInput {
  const props = model.properties ?? {};
  return {
    title: props.title ?? model.title,
    subject: props.subject,
    author: props.author,
    lastModifiedBy: props.lastModifiedBy,
    keywords: props.keywords,
    comments: props.comments,
    created: props.created,
    extraFields: props.extraFields,
  };
}

/** Rewrites `<p:sldSz>`'s `cx`/`cy` (whole EMU) in place, keeping its other
 * attributes and the surrounding bytes; inserts one before `<p:notesSz` (else
 * `</p:presentation>`) when the deck has none. `p:notesSz` is never touched —
 * PowerPoint sizes notes pages independently (SPEC-090). */
function spliceSldSz(presXml: string, model: PresentationModel): string {
  const cx = cmToEmu(model.slideWidthCm);
  const cy = cmToEmu(model.slideHeightCm);
  if (/<p:sldSz\b/.test(presXml)) {
    return presXml.replace(/<p:sldSz\b[^>]*\/?>/, (tag) => {
      let out = tag;
      out = /\bcx="[^"]*"/.test(out) ? out.replace(/\bcx="[^"]*"/, `cx="${cx}"`) : out.replace(/<p:sldSz/, `<p:sldSz cx="${cx}"`);
      out = /\bcy="[^"]*"/.test(out) ? out.replace(/\bcy="[^"]*"/, `cy="${cy}"`) : out.replace(/<p:sldSz/, `<p:sldSz cy="${cy}"`);
      return out;
    });
  }
  const element = `<p:sldSz cx="${cx}" cy="${cy}"/>`;
  if (presXml.includes("<p:notesSz")) {
    return presXml.replace("<p:notesSz", `${element}<p:notesSz`);
  }
  return presXml.replace("</p:presentation>", `${element}</p:presentation>`);
}

/** Sets (or removes) `firstSlideNum` on the `p:presentation` start tag — set
 * when the deck has a non-default origin, removed when absent (SPEC-090). */
function applyFirstSlideNum(presXml: string, firstSlideNumber: number | undefined): string {
  return presXml.replace(/<p:presentation\b[^>]*>/, (tag) => {
    let out = tag.replace(/\s+firstSlideNum="[^"]*"/, "");
    if (firstSlideNumber !== undefined) {
      out = out.replace(/\s*>$/, ` firstSlideNum="${firstSlideNumber}">`);
    }
    return out;
  });
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
