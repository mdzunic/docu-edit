// The `.pptx` notes-slide package surgery (SPEC-083). A notes edit sets a
// slide's `notesDirty`; this planner turns that into the OPC part changes the
// writer applies, WITHOUT rebuilding the slide part itself (that stays
// byte-verbatim when only notes changed):
//
//  - a slide whose notes part already exists → rebuild that part in place,
//    keeping its `.rels` verbatim;
//  - a slide gaining its FIRST notes → a fresh `ppt/notesSlides/notesSlideN.xml`
//    (N unique), its `.rels` (→ the slide, → the notes master), the slide's
//    `.rels` gaining the notesSlide relationship, a Content-Types override, and
//    — when the package has NO notes master — a minimal
//    `ppt/notesMasters/notesMaster1.xml` (+ rels + Content-Types +
//    `p:notesMasterIdLst` spliced into presentation.xml);
//  - a slide whose notes were deleted → drop the notes part, its `.rels`, the
//    slide's notesSlide relationship and the Content-Types override.
//
// Framework-free; imports only presentation siblings and src/lib/core.
import {
  parseRelationships,
  relsPathFor,
  resolvePartPath,
  escapeXmlAttribute,
  RELATIONSHIPS_NAMESPACE,
  type Relationship,
} from "../../core/xml";
import type { PresentationModel, Slide } from "../model";
import { preserveNotesFieldPlaceholders, serializeNotesSlideXml } from "./serialize";
import { fieldText, slideNumberFor } from "../headerFooter";
import { now as clockNow } from "../../core/clock";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const NOTES_SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

const NOTES_SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const NOTES_MASTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");
const encode = (text: string): Uint8Array => encoder.encode(text);

export interface NotesPlanInput {
  model: PresentationModel;
  /** The package's parts (post any structural rebuild the caller already did). */
  parts: Map<string, Uint8Array>;
  /** The instant a rebuilt notes part recomputes its field placeholders'
   * cached `a:t` against (SPEC-093); defaults to core `now()` when absent. */
  now?: Date;
  /** The emitted slide-part path for a model slide, or undefined to skip it
   * (a slide with no part cannot host a notes part in this path). */
  slidePartFor: (slide: Slide) => string | undefined;
  /** Slides emitted under a FRESH part this save (added or duplicated). Their
   * notes must be materialised into their own notes part even when they are not
   * `notesDirty` — a duplicate carries copied `notes` but its inherited rels
   * point at the SOURCE's notes part, so the caller strips that relationship and
   * the planner mints a new one here (two model slides never share one part). */
  freshSlides?: Set<Slide>;
  /** The slide's effective `.rels` XML the caller will emit (post image-rel
   * splicing), or "" when it has none yet. */
  slideRelsXml: (slide: Slide) => string;
  presPath: string;
  /** The presentation.xml the caller will emit — transformed when a notes
   * master is created (the one case a notes edit touches presentation.xml). */
  presXml: string;
  /** The presentation.xml.rels the caller will emit. */
  presRelsXml: string;
  /** The current `[Content_Types].xml`, to avoid re-adding an existing
   * override. */
  contentTypesXml: string;
}

export interface NotesPlan {
  /** Parts to add or override, by absolute name. */
  parts: Map<string, Uint8Array>;
  /** Parts to drop from the package. */
  removed: Set<string>;
  /** Content-Types `<Override>` lines to splice before `</Types>`. */
  contentTypeAdditions: string;
  /** Absolute part paths whose `<Override>` must be removed. */
  contentTypeRemovals: string[];
  /** True when at least one slide changed its notes (the caller must apply the
   * plan). */
  touched: boolean;
}

/** The relative target from one part to another (mirrors the writer's own). */
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

function relIdNumber(id: string): number {
  const match = /(\d+)$/.exec(id);
  return match ? Number(match[1]) : 0;
}

function relationshipsXml(rels: Relationship[]): string {
  return (
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeXmlAttribute(r.target)}"/>`).join("") +
    `</Relationships>`
  );
}

/** `relsXml` with one relationship appended before `</Relationships>` (kept a
 * splice so the existing rows stay byte-verbatim). */
function appendRel(relsXml: string, rel: Relationship): string {
  const row = `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXmlAttribute(rel.target)}"/>`;
  if (relsXml && /<\/Relationships>/.test(relsXml)) {
    return relsXml.replace(/<\/Relationships>/, `${row}</Relationships>`);
  }
  return `${XML_DECL}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${row}</Relationships>`;
}

/** `relsXml` with the relationship whose id is `id` removed. */
function removeRel(relsXml: string, id: string): string {
  return relsXml.replace(new RegExp(`<Relationship Id="${id}"[^>]*/>`), "");
}

function hasOverride(contentTypesXml: string, absoluteName: string): boolean {
  const escaped = `/${absoluteName}`.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`<Override PartName="${escaped}"`).test(contentTypesXml);
}

const NOTES_MASTER_XML =
  `${XML_DECL}<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle></p:notesMaster>`;

/** The `<p:notesMasterIdLst>` spliced into presentation.xml before
 * `<p:sldIdLst>` (schema order), else before `<p:sldSz>`. */
function spliceNotesMasterIdLst(presXml: string, relId: string): string {
  const lst = `<p:notesMasterIdLst><p:notesMasterId r:id="${relId}"/></p:notesMasterIdLst>`;
  if (/<p:notesMasterIdLst/.test(presXml)) {
    return presXml; // already present
  }
  if (/<p:sldIdLst/.test(presXml)) {
    return presXml.replace(/<p:sldIdLst/, `${lst}<p:sldIdLst`);
  }
  if (/<p:sldSz/.test(presXml)) {
    return presXml.replace(/<p:sldSz/, `${lst}<p:sldSz`);
  }
  return presXml.replace(/<\/p:presentation>/, `${lst}</p:presentation>`);
}

/** Plans the notes-part changes for every `notesDirty` slide. Returns
 * `touched: false` (and empty transforms) when no slide changed its notes. */
export function planNotes(input: NotesPlanInput): NotesPlan {
  const { model, parts } = input;
  const plan: NotesPlan = {
    parts: new Map(),
    removed: new Set(),
    contentTypeAdditions: "",
    contentTypeRemovals: [],
    touched: false,
  };
  // Work every slide whose notes changed, plus every freshly-emitted slide that
  // carries notes (an added/duplicated slide gets its own notes part even when
  // it was never edited, so its notes are not lost / not shared with a source).
  const fresh = input.freshSlides ?? new Set<Slide>();
  const dirtySlides = model.slides.filter(
    (slide) => slide.notesDirty === true || (fresh.has(slide) && slide.notes !== undefined && slide.notes.length > 0),
  );
  if (dirtySlides.length === 0) {
    return plan;
  }
  plan.touched = true;

  // Fresh notes-slide naming: past the highest existing notesSlideN.
  const takenNames = new Set(parts.keys());
  let nextNum = 0;
  for (const name of parts.keys()) {
    const match = /\/notesSlides\/notesSlide(\d+)\.xml$/.exec(name);
    if (match) {
      nextNum = Math.max(nextNum, Number(match[1]));
    }
  }
  const nextNotesName = (): string => {
    let name: string;
    do {
      nextNum += 1;
      name = `ppt/notesSlides/notesSlide${nextNum}.xml`;
    } while (takenNames.has(name) || plan.parts.has(name));
    takenNames.add(name);
    return name;
  };

  // Does the package (already) carry a notes master?
  let notesMasterPath: string | undefined;
  for (const name of parts.keys()) {
    if (/\/notesMasters\/notesMaster\d+\.xml$/.test(name)) {
      notesMasterPath = name;
      break;
    }
  }
  const presDir = input.presPath.split("/").slice(0, -1).join("/");
  let presXml = input.presXml;
  let presRelsXml = input.presRelsXml;
  let masterCreated = false;

  const ensureNotesMaster = (): string => {
    if (notesMasterPath) {
      return notesMasterPath;
    }
    // Create a minimal notes master mirroring SPEC-080's from-scratch master.
    const masterPath = `${presDir}/notesMasters/notesMaster1.xml`;
    notesMasterPath = masterPath;
    masterCreated = true;
    plan.parts.set(masterPath, encode(NOTES_MASTER_XML));
    // Its rels → the presentation's theme (so Office keeps the colour map).
    const themePath = [...parts.keys()].find((name) => /\/theme\/theme\d+\.xml$/.test(name));
    const masterRels: Relationship[] = [];
    if (themePath) {
      masterRels.push({ id: "rId1", type: THEME_REL, target: relativeTarget(masterPath, themePath) });
    }
    plan.parts.set(relsPathFor(masterPath), encode(relationshipsXml(masterRels)));
    plan.contentTypeAdditions += `<Override PartName="/${masterPath}" ContentType="${NOTES_MASTER_CONTENT_TYPE}"/>`;
    // presentation.xml.rels gains the notesMaster relationship; presentation.xml
    // gains p:notesMasterIdLst pointing at it.
    const presRels = parseRelationships(presRelsXml);
    const maxId = [...presRels.values()].reduce((max, r) => Math.max(max, relIdNumber(r.id)), 0);
    const relId = `rId${maxId + 1}`;
    presRelsXml = appendRel(presRelsXml, {
      id: relId,
      type: NOTES_MASTER_REL,
      target: relativeTarget(input.presPath, masterPath),
    });
    presXml = spliceNotesMasterIdLst(presXml, relId);
    return masterPath;
  };

  for (const slide of dirtySlides) {
    const slidePath = input.slidePartFor(slide);
    if (!slidePath) {
      continue;
    }
    const relsPath = relsPathFor(slidePath);
    const curRelsXml = input.slideRelsXml(slide);
    const rels = curRelsXml ? parseRelationships(curRelsXml) : new Map<string, Relationship>();
    const existing = [...rels.values()].find((r) => r.type === NOTES_SLIDE_REL);
    const hasNotes = slide.notes !== undefined && slide.notes.length > 0;

    if (hasNotes) {
      if (existing) {
        // Rebuild the existing notes part in place; keep its .rels verbatim.
        // Its non-body field placeholders (hdr/ftr/dt/sldNum) are preserved with
        // their cached a:t refreshed (SPEC-093 AC-92/93).
        const notesPath = resolvePartPath(slidePath, existing.target);
        const index = model.slides.indexOf(slide);
        const nonBody = preserveNotesFieldPlaceholders(
          parts.has(notesPath) ? decoder.decode(parts.get(notesPath)) : "",
          String(slideNumberFor(model, index)),
          fieldText(model, index, "date", input.now ?? clockNow()),
        );
        plan.parts.set(notesPath, encode(serializeNotesSlideXml(slide.notes!, nonBody)));
        if (!hasOverride(input.contentTypesXml, notesPath) && !plan.contentTypeAdditions.includes(`/${notesPath}"`)) {
          plan.contentTypeAdditions += `<Override PartName="/${notesPath}" ContentType="${NOTES_SLIDE_CONTENT_TYPE}"/>`;
        }
      } else {
        // First notes for this slide: mint the part, its rels, splice the slide
        // rels, add the Content-Types override, ensure a notes master.
        const notesPath = nextNotesName();
        const masterPath = ensureNotesMaster();
        plan.parts.set(notesPath, encode(serializeNotesSlideXml(slide.notes!)));
        plan.parts.set(
          relsPathFor(notesPath),
          encode(
            relationshipsXml([
              { id: "rId1", type: SLIDE_REL, target: relativeTarget(notesPath, slidePath) },
              { id: "rId2", type: NOTES_MASTER_REL, target: relativeTarget(notesPath, masterPath) },
            ]),
          ),
        );
        const maxId = [...rels.values()].reduce((max, r) => Math.max(max, relIdNumber(r.id)), 0);
        const slideRel: Relationship = {
          id: `rId${maxId + 1}`,
          type: NOTES_SLIDE_REL,
          target: relativeTarget(slidePath, notesPath),
        };
        plan.parts.set(relsPath, encode(appendRel(curRelsXml, slideRel)));
        plan.contentTypeAdditions += `<Override PartName="/${notesPath}" ContentType="${NOTES_SLIDE_CONTENT_TYPE}"/>`;
      }
    } else if (existing) {
      // Notes deleted: drop the notes part, its rels, the slide relationship and
      // the Content-Types override (no orphan referencing empty notes).
      const notesPath = resolvePartPath(slidePath, existing.target);
      plan.removed.add(notesPath);
      plan.removed.add(relsPathFor(notesPath));
      plan.contentTypeRemovals.push(notesPath);
      plan.parts.set(relsPath, encode(removeRel(curRelsXml, existing.id)));
    }
  }

  if (masterCreated) {
    plan.parts.set(input.presPath, encode(presXml));
    plan.parts.set(relsPathFor(input.presPath), encode(presRelsXml));
  }
  return plan;
}

/** `[Content_Types].xml` with the plan's overrides added (before `</Types>`)
 * and its removals dropped. */
export function applyNotesContentTypes(contentTypesXml: string, plan: NotesPlan): string {
  let result = contentTypesXml;
  for (const path of plan.contentTypeRemovals) {
    const escaped = `/${path}`.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    result = result.replace(new RegExp(`<Override PartName="${escaped}"[^>]*/>`), "");
  }
  if (plan.contentTypeAdditions.length > 0) {
    result = result.replace(/(<\/Types>\s*)$/, `${plan.contentTypeAdditions}$1`);
  }
  return result;
}

export { decoder as notesDecoder };
