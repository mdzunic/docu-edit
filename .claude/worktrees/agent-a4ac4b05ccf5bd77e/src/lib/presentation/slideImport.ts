// Import repair (SPEC-099, normative): every slide entering a deck from
// elsewhere — a paste from the slide clipboard (this deck or another one in
// the library) or the "Reuse slides" picker — goes through `importSlides`, so
// the rules a cross-deck copy obeys live in exactly one place. Per slide:
//  1. structuredClone the source slide; a fresh `Slide.id`; fresh
//     `SlideFrame.id`s for every frame and group child (`remintFrameIds`).
//  2. Drop frames this editor cannot re-home — a `ForeignFrame` whose verbatim
//     xml carries a package reference (`r:embed`/`r:id`/`r:link`, an odp
//     `xlink:href` into the package) or that belongs to the other format, a
//     `GroupFrame` whose verbatim xml carries a reference other than a plain
//     image — recording their accessible names; never a dangling reference.
//  3. Retarget the animations through the same id map (a dropped frame takes
//     its animations with it); a modelled `Slide.transition` survives.
//  4. `stripSlideParts` — the slide rebuilds from the model in both formats,
//     so this path sets `dirty` and `notesDirty`.
//  5. Re-home media into the target (de-duplicated), rescale on a slide-size
//     mismatch, apply the theme/layout option.
//  6. Splice into `target.slides` taking the insertion point's section, then
//     `normaliseSections(target)` and `target.structureDirty = true`.
// Layering rule (presentation/boundary.spec.ts): presentation siblings and
// src/lib/core only — so the frame-naming function the view uses is passed in.
import { resolveColor, resolveFont, DEFAULT_THEME, type DocumentTheme, type ModelFont, type ThemeColorRef } from "../core/theme";
import { flushPendingMedia, type MediaPart } from "../core/media";
import type {
  DeckTheme,
  PresentationFormat,
  PresentationModel,
  Slide,
  SlideFrame,
  SlideLayout,
  SlideParagraph,
} from "./model";
import { retargetAnimations } from "./animations";
import { scaleSlideGeometry } from "./deckSetup";
import { mintMediaId } from "./frames";
import { normaliseSections, sectionKeyForInsertion } from "./sections";
import { remintFrameIds, stripSlideParts } from "./slides";

/** Where imported slides come from: a deck, or a slide-clipboard payload —
 * anything carrying slides plus the deck facts the repair needs. */
export interface SlideImportSource {
  format: PresentationFormat;
  slides: Slide[];
  slideWidthCm: number;
  slideHeightCm: number;
  media?: MediaPart[];
  theme?: DeckTheme;
  layouts?: SlideLayout[];
}

export interface SlideImportOptions {
  /** The index the first imported slide lands at (clamped; default: the end). */
  at?: number;
  /** "destination" (default) keeps every theme reference so it resolves
   * against the target's theme and maps `layoutKey` to the target layout of the
   * same kind; "source" resolves every reference to a literal through the
   * source's theme and drops `layoutKey`. */
  formatting?: "destination" | "source";
  /** The accessible name a dropped frame is reported under — the view passes
   * `frameAccessibleName`; the default mirrors it for the kinds dropped here. */
  frameName?: (frame: SlideFrame) => string;
}

/** What an import did, honestly: what arrived, what could not come along. */
export interface SlideImportReport {
  slides: number;
  droppedFrames: number;
  droppedFrameNames: string[];
  /** Transitions that existed only as preserved source XML (out of subset). */
  lostTransitions: number;
  /** Timings that existed only as preserved source XML, plus animations whose
   * target frame was dropped. */
  lostAnimations: number;
  /** Media references that resolved to a part the target already held. */
  mediaReused: number;
  rescaled: boolean;
}

function defaultFrameName(frame: SlideFrame): string {
  switch (frame.kind) {
    case "foreign":
      return frame.label || "Object";
    case "group":
      return "Group";
    case "chart":
      return "Chart";
    case "image":
      return "Image";
    case "media":
      return frame.media === "audio" ? "Audio" : "Video";
    default:
      return "Object";
  }
}

const OOXML_REFERENCE = /\br:(?:embed|id|link|pict)="[^"]*"/;
const ODF_PACKAGE_HREF = /\bxlink:href="(?!(?:[a-z][a-z0-9+.-]*:|#))[^"]*"/i;

/** Whether verbatim xml points into its source package (a relationship id or
 * an odp package path). With `allowPlainImages`, a plain picture's reference
 * (`a:blip r:embed`, `draw:image xlink:href`) does not count — the group
 * rebuild re-homes those through the modelled image children. */
export function carriesPackageReference(xml: string, allowPlainImages = false): boolean {
  const scanned = allowPlainImages
    ? xml.replace(/<a:blip\b[^>]*\br:embed="[^"]*"[^>]*>/g, "").replace(/<draw:image\b[^>]*>/g, "")
    : xml;
  return OOXML_REFERENCE.test(scanned) || ODF_PACKAGE_HREF.test(scanned);
}

/** A frame's own paragraphs — text, shape text and table cells — where a
 * picture bullet's `bulletImageId` lives. */
function paragraphsOf(frame: SlideFrame): SlideParagraph[] {
  if (frame.kind === "text") {
    return frame.paragraphs;
  }
  if (frame.kind === "shape") {
    return frame.paragraphs ?? [];
  }
  if (frame.kind === "table") {
    return frame.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs));
  }
  return [];
}

interface DropContext {
  crossFormat: boolean;
  sourceMedia: Map<string, MediaPart>;
  name: (frame: SlideFrame) => string;
  dropped: string[];
  droppedIds: Set<string>;
}

/** Filters a frame list down to what can be re-homed (step 2), recording the
 * dropped frames' names and ids, and readying the kept verbatim containers
 * for a rebuild from the model. */
function keepRehomable(frames: SlideFrame[], ctx: DropContext): SlideFrame[] {
  const kept: SlideFrame[] = [];
  for (const frame of frames) {
    if (rehome(frame, ctx)) {
      kept.push(frame);
    } else {
      ctx.dropped.push(ctx.name(frame));
      collectIds(frame, ctx.droppedIds);
    }
  }
  return kept;
}

function collectIds(frame: SlideFrame, into: Set<string>): void {
  into.add(frame.id);
  if (frame.kind === "group") {
    for (const child of frame.children) {
      collectIds(child, into);
    }
  }
}

function rehome(frame: SlideFrame, ctx: DropContext): boolean {
  switch (frame.kind) {
    case "foreign":
      // Emitted verbatim: any package reference would dangle, and another
      // format's markup is not this format's.
      return !ctx.crossFormat && !carriesPackageReference(frame.xml);
    case "group":
      if (frame.xml !== undefined && carriesPackageReference(frame.xml, true)) {
        return false;
      }
      // Rebuilt from its (re-homed) children in the target's format.
      delete frame.xml;
      frame.children = keepRehomable(frame.children, ctx);
      return frame.children.length > 0;
    case "chart":
      if (frame.unsupportedChartXml !== undefined) {
        return false;
      }
      // Regenerated from the model into a fresh chart part of the target.
      delete frame.xml;
      delete frame.sourcePart;
      frame.edited = true;
      return true;
    case "image":
      return ctx.sourceMedia.has(frame.mediaId);
    case "media":
      if (frame.oversize || (frame.mediaId !== undefined && !ctx.sourceMedia.has(frame.mediaId))) {
        return false;
      }
      delete frame.sourceXml;
      if (frame.posterMediaId !== undefined && !ctx.sourceMedia.has(frame.posterMediaId)) {
        delete frame.posterMediaId;
      }
      return true;
    default:
      return true;
  }
}

/** Remaps every media reference in `frames` through `remap`; a picture
 * bullet whose image the source does not carry falls back to a plain bullet
 * rather than dangling. */
function remapMedia(frames: SlideFrame[], remap: (mediaId: string) => string, known: (mediaId: string) => boolean): void {
  for (const frame of frames) {
    for (const paragraph of paragraphsOf(frame)) {
      if (paragraph.bulletImageId !== undefined) {
        if (known(paragraph.bulletImageId)) {
          paragraph.bulletImageId = remap(paragraph.bulletImageId);
        } else {
          delete paragraph.bulletImageId;
        }
      }
    }
    if (frame.kind === "image") {
      frame.mediaId = remap(frame.mediaId);
    } else if (frame.kind === "media") {
      if (frame.mediaId !== undefined) {
        frame.mediaId = remap(frame.mediaId);
      }
      if (frame.posterMediaId !== undefined) {
        frame.posterMediaId = remap(frame.posterMediaId);
      }
    } else if (frame.kind === "group") {
      remapMedia(frame.children, remap, known);
    }
  }
}

/** A SPEC-094 scheme reference: a colour `{ theme: slot, …transforms }` or a
 * font `{ theme: "major" | "minor" }` (the only model objects keyed `theme`). */
function isThemeReference(value: unknown): value is { theme: string } {
  return (
    typeof value === "object" && value !== null && !Array.isArray(value) && typeof (value as { theme?: unknown }).theme === "string"
  );
}

/** Replaces every theme reference under `value` with a literal resolved
 * through `theme` ("keep source formatting"): a `{ theme: slot, … }` colour
 * through `resolveColor`, a `{ theme: "major" | "minor" }` font through
 * `resolveFont`. */
function resolveThemeRefs(value: unknown, theme: DocumentTheme): void {
  if (Array.isArray(value)) {
    for (const item of value) {
      resolveThemeRefs(item, theme);
    }
    return;
  }
  if (typeof value !== "object" || value === null) {
    return;
  }
  const record = value as Record<string, unknown>;
  for (const [key, child] of Object.entries(record)) {
    if (isThemeReference(child)) {
      record[key] =
        child.theme === "major" || child.theme === "minor"
          ? resolveFont(child as ModelFont, theme)
          : resolveColor(child as ThemeColorRef, theme);
    } else {
      resolveThemeRefs(child, theme);
    }
  }
}

/** The target layout a source `layoutKey` maps to under "use destination":
 * the target's first layout of the same kind; without source layout facts, the
 * key itself when the target knows it. Undefined = drop it. */
function mapLayoutKey(layoutKey: string, source: SlideImportSource, target: PresentationModel): string | undefined {
  const kind = source.layouts?.find((layout) => layout.key === layoutKey)?.kind;
  if (kind !== undefined) {
    return target.layouts?.find((layout) => layout.kind === kind)?.key;
  }
  return target.layouts?.some((layout) => layout.key === layoutKey) ? layoutKey : undefined;
}

function sameSize(a: number, b: number): boolean {
  return Math.abs(a - b) < 0.01;
}

/** Imports `source.slides[indexes]` into `target` per the import-repair rules
 * (see the module header), mutating `target` and returning the report. */
export function importSlides(
  target: PresentationModel,
  source: SlideImportSource,
  indexes: number[],
  options: SlideImportOptions = {},
): SlideImportReport {
  const report: SlideImportReport = {
    slides: 0,
    droppedFrames: 0,
    droppedFrameNames: [],
    lostTransitions: 0,
    lostAnimations: 0,
    mediaReused: 0,
    rescaled: false,
  };
  const formatting = options.formatting ?? "destination";
  // De-duplication compares base64 bytes, so a deferred encode must land first.
  flushPendingMedia(source.media);
  flushPendingMedia(target.media);
  const sourceMedia = new Map((source.media ?? []).map((part) => [part.mediaId, part]));
  const existingMedia = target.media ?? [];
  const media: MediaPart[] = [...existingMedia];
  const rehomed = new Map<string, string>();
  let mediaChanged = false;

  const remap = (mediaId: string): string => {
    const known = rehomed.get(mediaId);
    if (known !== undefined) {
      return known;
    }
    const part = sourceMedia.get(mediaId);
    if (!part) {
      return mediaId;
    }
    const existing = existingMedia.find((other) => other.contentType === part.contentType && other.data === part.data);
    let next: string;
    if (existing) {
      next = existing.mediaId;
    } else {
      const fresh = media.find((other) => other.contentType === part.contentType && other.data === part.data);
      next = fresh?.mediaId ?? mintMediaId({ ...target, media }, part.contentType);
      if (!fresh) {
        media.push({ mediaId: next, contentType: part.contentType, data: part.data });
        mediaChanged = true;
      }
    }
    rehomed.set(mediaId, next);
    return next;
  };
  const reusedIds = new Set(existingMedia.map((part) => part.mediaId));

  const scale =
    sameSize(source.slideWidthCm, target.slideWidthCm) && sameSize(source.slideHeightCm, target.slideHeightCm)
      ? undefined
      : Math.min(target.slideWidthCm / source.slideWidthCm, target.slideHeightCm / source.slideHeightCm);

  let at = Math.min(Math.max(options.at ?? target.slides.length, 0), target.slides.length);
  for (const index of indexes) {
    const original = source.slides[index];
    if (!original) {
      continue;
    }
    // 1. Clone + fresh ids.
    let slide: Slide = structuredClone(original);
    slide.id = crypto.randomUUID();
    const idMap = new Map<string, string>();
    remintFrameIds(slide.frames, idMap);

    // 2. Drop what cannot be re-homed.
    const ctx: DropContext = {
      crossFormat: source.format !== target.format,
      sourceMedia,
      name: options.frameName ?? defaultFrameName,
      dropped: [],
      droppedIds: new Set(),
    };
    slide.frames = keepRehomable(slide.frames, ctx);
    report.droppedFrames += ctx.dropped.length;
    report.droppedFrameNames.push(...ctx.dropped);

    // 3. Animations follow the id map; a dropped frame takes its own along.
    for (const [oldId, newId] of idMap) {
      if (ctx.droppedIds.has(newId)) {
        idMap.delete(oldId);
      }
    }
    if (slide.animations) {
      const kept = retargetAnimations(slide.animations, idMap) ?? [];
      report.lostAnimations += slide.animations.length - kept.length;
      slide.animations = kept;
    }
    const preserved = slide.preserved ?? {};
    if (slide.transition === undefined && (preserved["p:transition"] !== undefined || preserved["draw:transition"] !== undefined)) {
      report.lostTransitions += 1;
    }
    if (slide.animations === undefined && (preserved["p:timing"] !== undefined || preserved["anim"] !== undefined)) {
      report.lostAnimations += 1;
    }

    // 4. Rebuild from the model in either format.
    stripSlideParts(slide);
    slide.dirty = true;
    slide.notesDirty = true;

    // 5. Media, size, theme/layout.
    remapMedia(
      slide.frames,
      (mediaId) => {
        const next = remap(mediaId);
        if (reusedIds.has(next)) {
          report.mediaReused += 1;
        }
        return next;
      },
      (mediaId) => sourceMedia.has(mediaId),
    );
    if (scale !== undefined) {
      const offsetX = (target.slideWidthCm - source.slideWidthCm * scale) / 2;
      const offsetY = (target.slideHeightCm - source.slideHeightCm * scale) / 2;
      slide = scaleSlideGeometry(slide, scale, offsetX, offsetY);
      report.rescaled = true;
    }
    if (formatting === "source") {
      resolveThemeRefs(slide, source.theme ?? DEFAULT_THEME);
      delete slide.layoutKey;
    } else if (slide.layoutKey !== undefined) {
      const mapped = mapLayoutKey(slide.layoutKey, source, target);
      if (mapped === undefined) {
        delete slide.layoutKey;
      } else {
        slide.layoutKey = mapped;
      }
    }

    // 6. Splice in, joining the insertion point's section.
    const sectionKey = sectionKeyForInsertion(target, at);
    if (sectionKey === undefined) {
      delete slide.sectionKey;
    } else {
      slide.sectionKey = sectionKey;
    }
    target.slides.splice(at, 0, slide);
    at += 1;
    report.slides += 1;
  }
  if (mediaChanged) {
    // A new array, never an in-place push: the store's persistMedia latch
    // fires only when the reference changes.
    target.media = media;
  }
  normaliseSections(target);
  if (report.slides > 0) {
    target.structureDirty = true;
  }
  return report;
}
