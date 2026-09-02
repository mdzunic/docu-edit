// The direct-manipulation slide-object commands (SPEC-082), framework-free
// like every other presentation/ module: insert text boxes / images /
// catalogue shapes, move / resize / nudge / duplicate / delete them, reorder
// z-order, ungroup imported groups, and edit fills / strokes. Each command
// mutates one slide's `frames` in place, marks the slide edited (dirty; an odp
// page drops its verbatim `sourceXml` so the page serialiser regenerates it),
// and returns the frame id the view should select (or, for ungroup, the ids of
// the promoted children).
//
// Geometry stays in centimetres, colours lowercase `#rrggbb`. The commands
// never touch `opaque`/`odpContentEnvelope`; the writers treat `model.media`
// as the source of truth for inserted image parts (see the pptx/odp writers).
// This module imports only presentation siblings and src/lib/core — never the
// text product or a Vue module (boundary.spec.ts).
import { pxToCm } from "../core/units";
import { extensionForContentType, type MediaPart } from "../core/media";
import { LINE_GEOMETRIES, normalizeRotationDeg, type ShapeGeometry } from "../core/shapes";
import { rotatedBounds } from "../core/shapeGeometry";
import { resolvedShapeAdjustments } from "../core/shapeCatalogue";
import { MIN_COLUMN_WIDTH_CM } from "../core/tables";
import { markSlideEdited } from "./text";
import { createTableFrame, scaleTableGrid, setColumnWidths, setRowHeight } from "./tables";
import { removeAnimationsForFrame } from "./animations";
import type {
  FrameStroke,
  GroupFrame,
  ImageFrame,
  MediaFrame,
  PresentationFormat,
  PresentationModel,
  ShapeFrame,
  Slide,
  SlideFill,
  SlideFrame,
  SlideMediaKind,
  TextFrame,
} from "./model";

/** A frame's box in slide space (centimetres). */
export interface FrameRect {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** The deck-neutral default style a freshly drawn shape lands with — the
 * SPEC-051 defaults, restated here so this product owns them without importing
 * a text-product module. */
export const DECK_DEFAULT_SHAPE_FILL = "#4472c4";
export const DECK_DEFAULT_SHAPE_STROKE: { color: string; widthPt: number } = { color: "#2f528f", widthPt: 1 };

/** A click (no drag) with "Text box" armed drops a default 10×2.5cm box. */
export const DEFAULT_TEXT_BOX_SIZE_CM: { widthCm: number; heightCm: number } = { widthCm: 10, heightCm: 2.5 };

/** The smallest a resize may leave a frame (AC-24). */
export const MIN_FRAME_SIZE_CM = 0.5;

/** At least this much of a frame must stay on the slide after a move/nudge
 * (AC-18) — measured on the frame's axis-aligned box. */
export const FRAME_MIN_ON_SLIDE_CM = 1;

/** A duplicate lands offset by this much on each axis (AC-28). */
export const DUPLICATE_OFFSET_CM = 0.5;

/** Arrow-key nudge distances (AC-17). */
export const NUDGE_STEP_CM = 0.1;
export const NUDGE_LARGE_STEP_CM = 1;

/** Guide activation / snap distance (SPEC-072's 0.15cm convention, AC-19/20). */
export const GUIDE_SNAP_CM = 0.15;

/** Stroke width limits the style popover clamps to (AC-39). */
export const MIN_STROKE_WIDTH_PT = 0.25;
export const MAX_STROKE_WIDTH_PT = 8;

/** Where an ungroup command re-labels its foreign children so they remain
 * self-describing (AC-52's "foreign label"). */
const FOREIGN_LABEL = "Object";

function clamp(value: number, min: number, max: number): number {
  if (min > max) {
    return max;
  }
  return Math.min(Math.max(value, min), max);
}

/** cm rounded to 3 decimals — finer than the eye, coarser than the EMU/ODF
 * precision the writers round to, so geometry survives a build → re-parse
 * round trip within the ±0.01cm the ACs allow. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function slideAt(model: PresentationModel, slideIndex: number): Slide | null {
  return model.slides[slideIndex] ?? null;
}

function frameOf(slide: Slide, frameId: string): SlideFrame | null {
  return slide.frames.find((frame) => frame.id === frameId) ?? null;
}

/** Marks the slide edited after a command (dirty + odp sourceXml drop). */
function touch(model: PresentationModel, slide: Slide): void {
  markSlideEdited(slide, model.format);
}

/** A geometry change (move/resize/reorder/duplicate) invalidates a frame's
 * verbatim source container — the serialiser must regenerate it at the new
 * geometry. For a chart this drops `xml` but must NEVER set `edited` (the chart
 * PART is untouched, so it still round-trips byte-identical — SPEC-092 AC-16a,
 * D11); for a media frame it drops `sourceXml`, the frame-level analogue of
 * `markSlideEdited`'s odp rule (SPEC-095). A no-op for every other frame kind. */
function dropRebuiltContainerXml(frame: SlideFrame): void {
  if (frame.kind === "chart") {
    delete frame.xml;
  } else if (frame.kind === "media") {
    delete frame.sourceXml;
  }
}

// --- insertion --------------------------------------------------------------

function normaliseRect(rect: FrameRect): FrameRect {
  return {
    xCm: r3(rect.xCm),
    yCm: r3(rect.yCm),
    widthCm: r3(Math.max(rect.widthCm, 0)),
    heightCm: r3(Math.max(rect.heightCm, 0)),
  };
}

/** Inserts an empty free text box (no placeholder role) at `rect`, appended on
 * top in z-order. Returns its id (the view opens it into a text session). */
export function insertTextBoxFrame(model: PresentationModel, slideIndex: number, rect: FrameRect): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    ...normaliseRect(rect),
    kind: "text",
    paragraphs: [{ runs: [] }],
  };
  slide.frames.push(frame);
  touch(model, slide);
  return frame.id;
}

/** Inserts a `rows`×`cols` table centred at 80% of the slide width (SPEC-091
 * AC-5c), appended on top in z-order. Returns its id (the view selects it with
 * the caret in cell (0,0)). */
export function insertTableFrame(
  model: PresentationModel,
  slideIndex: number,
  rows: number,
  cols: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const frame = createTableFrame(rows, cols, model.slideWidthCm, model.slideHeightCm);
  slide.frames.push(frame);
  touch(model, slide);
  return frame.id;
}

/** Inserts a catalogue shape at `rect` with the deck-neutral default style
 * (AC-14); line geometries take a stroke only. Returns its id. */
export function insertShapeFrame(
  model: PresentationModel,
  slideIndex: number,
  geometry: ShapeGeometry,
  rect: FrameRect,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const frame: ShapeFrame = {
    id: crypto.randomUUID(),
    ...normaliseRect(rect),
    kind: "shape",
    geometry,
    stroke: { ...DECK_DEFAULT_SHAPE_STROKE },
  };
  if (!LINE_GEOMETRIES.has(geometry)) {
    frame.fill = { type: "solid", color: DECK_DEFAULT_SHAPE_FILL };
  }
  slide.frames.push(frame);
  touch(model, slide);
  return frame.id;
}

/** Inserts an image: adds `media` to `model.media` under a freshly minted,
 * collision-free `mediaId` (per the naming rule), lands an `ImageFrame` at
 * `rect`, and returns its id. The media part's own `mediaId` is overwritten
 * with the minted package path. */
export function insertImageFrame(
  model: PresentationModel,
  slideIndex: number,
  media: MediaPart,
  rect: FrameRect,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const mediaId = mintMediaId(model, media.contentType);
  const part: MediaPart = { mediaId, contentType: media.contentType, data: media.data };
  if (!model.media) {
    model.media = [];
  }
  model.media.push(part);
  const frame: ImageFrame = {
    id: crypto.randomUUID(),
    ...normaliseRect(rect),
    kind: "image",
    mediaId,
  };
  slide.frames.push(frame);
  touch(model, slide);
  return frame.id;
}

/** The package directory inserted media rides in, per format. */
export function mediaDirectory(model: PresentationModel): string {
  return model.format === "odp" ? "Pictures/" : "ppt/media/";
}

/** The box a freshly inserted video lands in: a 16 cm × 9 cm (16:9) frame
 * (SPEC-095). */
export const MEDIA_VIDEO_SIZE_CM: { widthCm: number; heightCm: number } = { widthCm: 16, heightCm: 9 };

/** The strip a freshly inserted audio clip lands in: 8 cm × 1.6 cm (SPEC-095). */
export const MEDIA_AUDIO_SIZE_CM: { widthCm: number; heightCm: number } = { widthCm: 8, heightCm: 1.6 };

/** The centred rect an inserted clip lands in (SPEC-095): the per-kind box
 * (video 16×9, audio 8×1.6), shrunk to the slide when it would overflow and
 * centred on the surface. */
function mediaInsertRect(model: PresentationModel, mediaKind: SlideMediaKind): FrameRect {
  const box = mediaKind === "video" ? MEDIA_VIDEO_SIZE_CM : MEDIA_AUDIO_SIZE_CM;
  const widthCm = Math.min(box.widthCm, model.slideWidthCm);
  const heightCm = Math.min(box.heightCm, model.slideHeightCm);
  return {
    xCm: r3((model.slideWidthCm - widthCm) / 2),
    yCm: r3((model.slideHeightCm - heightCm) / 2),
    widthCm: r3(widthCm),
    heightCm: r3(heightCm),
  };
}

/** Inserts an audio/video clip: adds `media` to `model.media` under a freshly
 * minted, collision-free `mediaId`, lands a `MediaFrame` at the per-kind box
 * (video 16×9 cm, audio 8×1.6 cm) centred and clamped into the slide, and
 * returns its id (SPEC-095). The media part's own `mediaId` is overwritten with
 * the minted package path. */
export function insertMediaFrame(
  model: PresentationModel,
  slideIndex: number,
  media: MediaPart,
  mediaKind: SlideMediaKind,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const mediaId = mintMediaId(model, media.contentType);
  const part: MediaPart = { mediaId, contentType: media.contentType, data: media.data };
  if (!model.media) {
    model.media = [];
  }
  model.media.push(part);
  const frame: MediaFrame = {
    id: crypto.randomUUID(),
    ...normaliseRect(mediaInsertRect(model, mediaKind)),
    kind: "media",
    media: mediaKind,
    mediaId,
    contentType: media.contentType,
  };
  slide.frames.push(frame);
  clampToSlide(frame, model.slideWidthCm, model.slideHeightCm);
  touch(model, slide);
  return frame.id;
}

/** The frame-level analogue of `markSlideEdited`'s odp `sourceXml` rule
 * (SPEC-095): a command that changes a modelled `MediaFrame` field drops the
 * verbatim source so the serialiser rebuilds it from the model, then marks the
 * slide edited (dirty; an odp page drops its own `sourceXml`). Called by the
 * Media panel's flag/poster edits — geometry commands drop `sourceXml` through
 * `dropRebuiltContainerXml` already. */
export function markFrameEdited(slide: Slide, frame: MediaFrame, format: PresentationFormat): void {
  delete frame.sourceXml;
  markSlideEdited(slide, format);
}

/** Mints a package part path for an inserted image that collides with no part
 * already in the package: `{dir}image{n}.{ext}`, `n` = 1 + the highest
 * `image{n}` numeric suffix across `opaque` part names and existing `media`
 * ids (arbitrary/non-numeric image names are ignored for the counter, then
 * guarded against by a final uniqueness pass). */
function mintMediaId(model: PresentationModel, contentType: string): string {
  const dir = mediaDirectory(model);
  const ext = extensionForContentType(contentType);
  const taken = new Set<string>();
  let max = 0;
  const scan = (name: string): void => {
    taken.add(name);
    const match = /image(\d+)\.[^./]+$/i.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  };
  for (const part of model.opaque ?? []) {
    scan(part.name);
  }
  for (const part of model.media ?? []) {
    scan(part.mediaId);
  }
  let n = max + 1;
  let candidate = `${dir}image${n}.${ext}`;
  while (taken.has(candidate)) {
    n += 1;
    candidate = `${dir}image${n}.${ext}`;
  }
  return candidate;
}

/** The display size an inserted image gets: its natural size (px → cm) scaled
 * down to fit inside a box of `maxWidthCm`×`maxHeightCm`, aspect preserved
 * (AC-5). A null natural size (an undecodable format, e.g. webp under jsdom)
 * falls back to the full fit box. */
export function fitImageWithinCm(
  natural: { widthPx: number; heightPx: number } | null,
  maxWidthCm: number,
  maxHeightCm: number,
): { widthCm: number; heightCm: number } {
  if (!natural || natural.widthPx <= 0 || natural.heightPx <= 0) {
    return { widthCm: r3(maxWidthCm), heightCm: r3(maxHeightCm) };
  }
  const naturalWidthCm = pxToCm(natural.widthPx);
  const naturalHeightCm = pxToCm(natural.heightPx);
  const scale = Math.min(1, maxWidthCm / naturalWidthCm, maxHeightCm / naturalHeightCm);
  return { widthCm: r3(naturalWidthCm * scale), heightCm: r3(naturalHeightCm * scale) };
}

/** The centred rect an inserted image lands in (AC-5): scaled to 50% of the
 * slide's width/height, centred on the surface. */
export function insertImageRect(
  model: PresentationModel,
  natural: { widthPx: number; heightPx: number } | null,
): FrameRect {
  const size = fitImageWithinCm(natural, model.slideWidthCm / 2, model.slideHeightCm / 2);
  return {
    xCm: r3((model.slideWidthCm - size.widthCm) / 2),
    yCm: r3((model.slideHeightCm - size.heightCm) / 2),
    widthCm: size.widthCm,
    heightCm: size.heightCm,
  };
}

// --- move / nudge -----------------------------------------------------------

/** A frame's axis-aligned box in slide space, folding in its rotation
 * (SPEC-096 AC — clamps, marquee intersection and guide activation all use
 * `rotatedBounds`). */
export function frameRotatedBox(frame: SlideFrame): FrameRect {
  const box = rotatedBounds(frame.widthCm, frame.heightCm, frame.rotationDeg ?? 0);
  return { xCm: frame.xCm - box.padX, yCm: frame.yCm - box.padY, widthCm: box.width, heightCm: box.height };
}

/** Clamps a frame's position so at least `FRAME_MIN_ON_SLIDE_CM` of it stays
 * on the slide (AC-18), on the frame's rotated axis-aligned box. */
function clampToSlide(frame: SlideFrame, slideWidthCm: number, slideHeightCm: number): void {
  const box = rotatedBounds(frame.widthCm, frame.heightCm, frame.rotationDeg ?? 0);
  const x = clamp(
    frame.xCm - box.padX,
    FRAME_MIN_ON_SLIDE_CM - box.width,
    slideWidthCm - FRAME_MIN_ON_SLIDE_CM,
  );
  const y = clamp(
    frame.yCm - box.padY,
    FRAME_MIN_ON_SLIDE_CM - box.height,
    slideHeightCm - FRAME_MIN_ON_SLIDE_CM,
  );
  translateFrame(frame, r3(x + box.padX) - frame.xCm, r3(y + box.padY) - frame.yCm);
}

/** Shifts a frame — and, for a group, its slide-space children with it — by a
 * raw delta. A group's children moving is a child-geometry change, so its
 * verbatim `xml` is deleted (SPEC-096: the only way that field ever changes);
 * a chart's verbatim container likewise. */
function translateFrame(frame: SlideFrame, dxCm: number, dyCm: number): void {
  if (dxCm === 0 && dyCm === 0) {
    return;
  }
  frame.xCm = r3(frame.xCm + dxCm);
  frame.yCm = r3(frame.yCm + dyCm);
  dropRebuiltContainerXml(frame);
  if (frame.kind === "group") {
    delete frame.xml;
    frame.children.forEach((child) => translateFrame(child, dxCm, dyCm));
  }
}

/** Moves a frame by `(dxCm, dyCm)`, clamped to the slide (AC-16/17/18).
 * Returns the moved frame's id. */
export function moveFrame(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  dxCm: number,
  dyCm: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame) {
    return null;
  }
  translateFrame(frame, dxCm, dyCm);
  clampToSlide(frame, model.slideWidthCm, model.slideHeightCm);
  dropRebuiltContainerXml(frame);
  touch(model, slide);
  return frame.id;
}

/** Moves a frame to an absolute position (a committed drag with snapping),
 * clamped to the slide. Returns the moved frame's id. */
export function setFramePosition(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  xCm: number,
  yCm: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame) {
    return null;
  }
  translateFrame(frame, xCm - frame.xCm, yCm - frame.yCm);
  clampToSlide(frame, model.slideWidthCm, model.slideHeightCm);
  dropRebuiltContainerXml(frame);
  touch(model, slide);
  return frame.id;
}

/** Moves several frames by one shared delta, clamped on the UNION of their
 * rotated boxes against `FRAME_MIN_ON_SLIDE_CM` — never per frame, so
 * relative positions cannot shift (SPEC-096 AC-13). Returns the moved ids. */
export function moveFrames(
  model: PresentationModel,
  slideIndex: number,
  frameIds: readonly string[],
  dxCm: number,
  dyCm: number,
): string[] | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const frames = slide.frames.filter((frame) => frameIds.includes(frame.id));
  if (frames.length === 0) {
    return null;
  }
  const boxes = frames.map(frameRotatedBox);
  const minX = Math.min(...boxes.map((b) => b.xCm));
  const minY = Math.min(...boxes.map((b) => b.yCm));
  const maxX = Math.max(...boxes.map((b) => b.xCm + b.widthCm));
  const maxY = Math.max(...boxes.map((b) => b.yCm + b.heightCm));
  const dx = clamp(minX + dxCm, FRAME_MIN_ON_SLIDE_CM - (maxX - minX), model.slideWidthCm - FRAME_MIN_ON_SLIDE_CM) - minX;
  const dy = clamp(minY + dyCm, FRAME_MIN_ON_SLIDE_CM - (maxY - minY), model.slideHeightCm - FRAME_MIN_ON_SLIDE_CM) - minY;
  for (const frame of frames) {
    translateFrame(frame, r3(dx), r3(dy));
  }
  touch(model, slide);
  return frames.map((frame) => frame.id);
}

// --- resize -----------------------------------------------------------------

/** Resizes a frame to `rect`, enforcing the 0.5×0.5cm minimum (AC-24). The
 * view computes the rect (aspect rules, handle anchoring); this clamps and
 * commits. Returns the resized frame's id. */
export function resizeFrame(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  rect: FrameRect,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame) {
    return null;
  }
  const nextWidth = r3(Math.max(rect.widthCm, MIN_FRAME_SIZE_CM));
  const nextHeight = r3(Math.max(rect.heightCm, MIN_FRAME_SIZE_CM));
  if (frame.kind === "table") {
    // Scale the grid proportionally so the columns/rows keep their relative
    // proportions (SPEC-091 AC-8c) — before overwriting widthCm/heightCm.
    scaleTableGrid(frame, nextWidth / frame.widthCm, nextHeight / frame.heightCm);
  }
  if (frame.kind === "group") {
    // Children live in slide space, so a group resize scales them with the
    // box — a child-geometry change that drops the verbatim xml (SPEC-096).
    scaleGroupChildren(frame, { xCm: r3(rect.xCm), yCm: r3(rect.yCm), widthCm: nextWidth, heightCm: nextHeight });
  }
  frame.xCm = r3(rect.xCm);
  frame.yCm = r3(rect.yCm);
  frame.widthCm = nextWidth;
  frame.heightCm = nextHeight;
  dropRebuiltContainerXml(frame);
  touch(model, slide);
  return frame.id;
}

/** Scales a group's slide-space children from the group's current box to
 * `next`, recursively, deleting every affected group's verbatim `xml`. */
function scaleGroupChildren(group: GroupFrame, next: FrameRect): void {
  const sx = group.widthCm > 0 ? next.widthCm / group.widthCm : 1;
  const sy = group.heightCm > 0 ? next.heightCm / group.heightCm : 1;
  delete group.xml;
  const scale = (frame: SlideFrame): void => {
    frame.xCm = r3(next.xCm + (frame.xCm - group.xCm) * sx);
    frame.yCm = r3(next.yCm + (frame.yCm - group.yCm) * sy);
    frame.widthCm = r3(frame.widthCm * sx);
    frame.heightCm = r3(frame.heightCm * sy);
    dropRebuiltContainerXml(frame);
    if (frame.kind === "group") {
      delete frame.xml;
      frame.children.forEach(scale);
    }
  };
  group.children.forEach(scale);
}

/** Live column-boundary drag (SPEC-091 AC-8a). `boundary` is the index of the
 * column left of the dragged internal boundary; the column right of it absorbs
 * exactly what this one gains, so the table's overall width is unchanged. Both
 * are floored at `MIN_COLUMN_WIDTH_CM`. `originWidthsCm` is the snapshot taken
 * at pointer-down so each move recomputes from a stable base. */
export function resizeTableColumnBoundary(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  boundary: number,
  originWidthsCm: number[],
  deltaCm: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || frame.kind !== "table") {
    return null;
  }
  const left = originWidthsCm[boundary];
  const right = originWidthsCm[boundary + 1];
  if (left === undefined || right === undefined) {
    return null;
  }
  const total = left + right;
  const nextLeft = Math.min(Math.max(left + deltaCm, MIN_COLUMN_WIDTH_CM), total - MIN_COLUMN_WIDTH_CM);
  const widths = [...originWidthsCm];
  widths[boundary] = nextLeft;
  widths[boundary + 1] = total - nextLeft;
  setColumnWidths(frame, widths);
  touch(model, slide);
  return frame.id;
}

/** Live row-boundary drag (SPEC-091 AC-8b): sets the height of the row above
 * the dragged boundary (a floor, not a fixed height — `setRowHeight` clamps to
 * `MIN_ROW_HEIGHT_CM`). `originHeightCm` is the row's height at pointer-down. */
export function resizeTableRowBoundary(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  boundary: number,
  originHeightCm: number,
  deltaCm: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || frame.kind !== "table") {
    return null;
  }
  setRowHeight(frame, boundary, originHeightCm + deltaCm);
  touch(model, slide);
  return frame.id;
}

// --- rotation / adjustments (SPEC-096) --------------------------------------

/** Sets a frame's rotation in whole degrees 0–359 (`normalizeRotationDeg`).
 * Zero deletes the field. Rotating a GROUP is rigid and composes entirely
 * onto its children ("Rotation maths"): each child's centre rotates about the
 * group centre and each child's own rotation gains the delta —
 * `GroupFrame.rotationDeg` itself is deleted rather than written, and the
 * group box is recomputed as the children's rotated union. Returns the id. */
export function setFrameRotation(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  rotationDeg: number,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame) {
    return null;
  }
  const deg = normalizeRotationDeg(rotationDeg);
  if (frame.kind === "group") {
    rotateGroupChildren(frame, deg - normalizeRotationDeg(frame.rotationDeg ?? 0));
  } else if (deg === 0) {
    delete frame.rotationDeg;
  } else {
    frame.rotationDeg = deg;
  }
  dropRebuiltContainerXml(frame);
  touch(model, slide);
  return frame.id;
}

/** Composes a rotation delta onto a group's children (SPEC-096 "Group
 * rotation"): rigid about the group centre, then the group's own rotation is
 * deleted and its box recomputed from the children. */
function rotateGroupChildren(group: GroupFrame, deltaDeg: number): void {
  delete group.xml;
  delete group.rotationDeg;
  if (deltaDeg === 0) {
    return;
  }
  const cx = group.xCm + group.widthCm / 2;
  const cy = group.yCm + group.heightCm / 2;
  const rad = (deltaDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const spin = (frame: SlideFrame): void => {
    const fcx = frame.xCm + frame.widthCm / 2 - cx;
    const fcy = frame.yCm + frame.heightCm / 2 - cy;
    const nx = cx + fcx * cos - fcy * sin;
    const ny = cy + fcx * sin + fcy * cos;
    frame.xCm = r3(nx - frame.widthCm / 2);
    frame.yCm = r3(ny - frame.heightCm / 2);
    dropRebuiltContainerXml(frame);
    if (frame.kind === "group") {
      delete frame.xml;
      frame.children.forEach(spin);
    } else {
      const deg = normalizeRotationDeg((frame.rotationDeg ?? 0) + deltaDeg);
      if (deg === 0) {
        delete frame.rotationDeg;
      } else {
        frame.rotationDeg = deg;
      }
    }
  };
  group.children.forEach(spin);
  const union = unionRotatedBox(group.children);
  if (union) {
    group.xCm = union.xCm;
    group.yCm = union.yCm;
    group.widthCm = union.widthCm;
    group.heightCm = union.heightCm;
  }
}

/** The union of the frames' rotated axis-aligned boxes, or null when empty. */
function unionRotatedBox(frames: readonly SlideFrame[]): FrameRect | null {
  if (frames.length === 0) {
    return null;
  }
  const boxes = frames.map(frameRotatedBox);
  const minX = Math.min(...boxes.map((b) => b.xCm));
  const minY = Math.min(...boxes.map((b) => b.yCm));
  const maxX = Math.max(...boxes.map((b) => b.xCm + b.widthCm));
  const maxY = Math.max(...boxes.map((b) => b.yCm + b.heightCm));
  return { xCm: r3(minX), yCm: r3(minY), widthCm: r3(maxX - minX), heightCm: r3(maxY - minY) };
}

/** Sets a shape's adjustment values, stored as the 0..1-style fractions
 * `resolvedShapeAdjustments` resolves (each clamped to its declared range).
 * A no-op returning null for a non-shape. Returns the frame id. */
export function setFrameAdjustments(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  adjustments: readonly number[],
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || frame.kind !== "shape") {
    return null;
  }
  frame.adjustments = resolvedShapeAdjustments(frame.geometry, adjustments);
  touch(model, slide);
  return frame.id;
}

// --- delete / duplicate -----------------------------------------------------

/** Removes a frame. Returns the id the view should select next (the previous
 * frame in z-order, or null when the slide is now empty). */
export function deleteFrame(model: PresentationModel, slideIndex: number, frameId: string): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const index = slide.frames.findIndex((frame) => frame.id === frameId);
  if (index < 0) {
    return null;
  }
  slide.frames.splice(index, 1);
  // Drop any animations that targeted the removed frame (SPEC-086 AC-18); undo
  // restores both because the undo stack snapshots the whole slide.
  slide.animations = removeAnimationsForFrame(slide.animations, frameId);
  touch(model, slide);
  const neighbour = slide.frames[index - 1] ?? slide.frames[index] ?? null;
  return neighbour ? neighbour.id : null;
}

/** Re-mints a frame's id (and, for a group, its children's) so a duplicate is
 * an independent selection/undo target. */
function remintIds(frame: SlideFrame): void {
  frame.id = crypto.randomUUID();
  if (frame.kind === "group") {
    frame.children.forEach(remintIds);
  }
}

/** Duplicates a frame offset by +0.5/+0.5cm, appended on top (AC-28). Returns
 * the copy's id. */
export function duplicateFrame(model: PresentationModel, slideIndex: number, frameId: string): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame) {
    return null;
  }
  const copy = structuredClone(frame);
  remintIds(copy);
  // A group's children shift with the copy (slide-space coordinates), which
  // deletes the copy's verbatim xml — the original keeps its own.
  translateFrame(copy, DUPLICATE_OFFSET_CM, DUPLICATE_OFFSET_CM);
  clampToSlide(copy, model.slideWidthCm, model.slideHeightCm);
  // A duplicated chart shares the source's `sourcePart` (legal — the writer
  // mints a fresh part for the later claimant, D8) but its verbatim container
  // now has the wrong geometry, so drop it (never `edited`; the part is intact).
  dropRebuiltContainerXml(copy);
  slide.frames.push(copy);
  touch(model, slide);
  return copy.id;
}

// --- z-order ----------------------------------------------------------------

export type ReorderTarget = "front" | "forward" | "backward" | "back";

/** Reorders a frame within its slide's z-order (AC-30). Returns the frame id
 * (kept selected). */
export function reorderFrame(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  target: ReorderTarget,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const from = slide.frames.findIndex((frame) => frame.id === frameId);
  if (from < 0) {
    return null;
  }
  const last = slide.frames.length - 1;
  let to = from;
  if (target === "front") {
    to = last;
  } else if (target === "back") {
    to = 0;
  } else if (target === "forward") {
    to = Math.min(from + 1, last);
  } else if (target === "backward") {
    to = Math.max(from - 1, 0);
  }
  if (to === from) {
    return frameId;
  }
  const [moved] = slide.frames.splice(from, 1);
  slide.frames.splice(to, 0, moved);
  dropRebuiltContainerXml(moved);
  touch(model, slide);
  return frameId;
}

/** Reorders several frames as ONE contiguous block (SPEC-096 AC-16/17/18):
 * lifts the selected frames out of z-order and re-inserts them together —
 * relative order preserved — at the target end/step. A no-op leaving the
 * order (and the slide) untouched when the block already sits at that end.
 * Returns the frame ids (kept selected). */
export function reorderFrames(
  model: PresentationModel,
  slideIndex: number,
  frameIds: readonly string[],
  target: ReorderTarget,
): string[] | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const wanted = new Set(frameIds);
  const block = slide.frames.filter((frame) => wanted.has(frame.id));
  if (block.length === 0) {
    return null;
  }
  const rest = slide.frames.filter((frame) => !wanted.has(frame.id));
  // How many non-selected frames sit below the block's bottom/top member.
  const indexOf = new Map(slide.frames.map((frame, i) => [frame.id, i]));
  const bottom = indexOf.get(block[0].id)!;
  const top = indexOf.get(block[block.length - 1].id)!;
  const restBelow = (position: number): number => rest.filter((frame) => indexOf.get(frame.id)! < position).length;
  let insertAt: number;
  if (target === "front") {
    insertAt = rest.length;
  } else if (target === "back") {
    insertAt = 0;
  } else if (target === "forward") {
    insertAt = Math.min(restBelow(top) + 1, rest.length);
  } else {
    insertAt = Math.max(restBelow(bottom) - 1, 0);
  }
  const next = [...rest.slice(0, insertAt), ...block, ...rest.slice(insertAt)];
  if (next.every((frame, i) => frame === slide.frames[i])) {
    return block.map((frame) => frame.id); // already at that end — untouched
  }
  slide.frames = next;
  block.forEach(dropRebuiltContainerXml);
  touch(model, slide);
  return block.map((frame) => frame.id);
}

// --- group / ungroup --------------------------------------------------------

/** Replaces two or more selected top-level frames with ONE authored
 * `GroupFrame` (SPEC-096 AC-40): the group sits at the children's rotated
 * union box, carries them as `children` in slide space (z-order preserved)
 * and NO `xml`; it lands at the topmost member's z-position. Any animation
 * targeting a grouped frame is dropped (grouped frames are unanimatable).
 * Returns the new group's id, or null for fewer than two matches. */
export function groupFrames(
  model: PresentationModel,
  slideIndex: number,
  frameIds: readonly string[],
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const wanted = new Set(frameIds);
  const children = slide.frames.filter((frame) => wanted.has(frame.id));
  if (children.length < 2) {
    return null;
  }
  const box = unionRotatedBox(children)!;
  const group: GroupFrame = {
    id: crypto.randomUUID(),
    ...box,
    kind: "group",
    children,
  };
  const topIndex = slide.frames.findIndex((frame) => frame === children[children.length - 1]);
  const removedBelow = children.length - 1; // every other member sits below the top one
  slide.frames = slide.frames.filter((frame) => !wanted.has(frame.id));
  slide.frames.splice(topIndex - removedBelow, 0, group);
  for (const child of children) {
    slide.animations = removeAnimationsForFrame(slide.animations, child.id);
  }
  touch(model, slide);
  return group.id;
}

// --- ungroup ----------------------------------------------------------------

/** Ungroups a `GroupFrame`: its children (already in slide-space geometry per
 * SPEC-079) replace it in place, the group's verbatim `xml` is dropped, and a
 * foreign child survives as a foreign frame with a stable label. Returns the
 * promoted children's ids (the view selects the first). A non-group frame is a
 * no-op returning null. */
export function ungroupFrame(model: PresentationModel, slideIndex: number, frameId: string): string[] | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const index = slide.frames.findIndex((frame) => frame.id === frameId);
  const group = index >= 0 ? slide.frames[index] : null;
  if (!group || group.kind !== "group") {
    return null;
  }
  const children = group.children.map((child) => {
    if (child.kind === "foreign" && !child.label) {
      return { ...child, label: FOREIGN_LABEL };
    }
    return child;
  });
  slide.frames.splice(index, 1, ...children);
  touch(model, slide);
  return children.map((child) => child.id);
}

// --- fill / stroke ----------------------------------------------------------

/** Sets a shape or text box's fill (AC-37), including "No fill". A no-op for a
 * frame kind without a fill. Returns the frame id. */
export function setFrameFill(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  fill: SlideFill,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || (frame.kind !== "shape" && frame.kind !== "text")) {
    return null;
  }
  frame.fill = fill;
  touch(model, slide);
  return frame.id;
}

/** Sets a shape, text box or image's stroke (AC-38/39/40); `null` clears it.
 * The width is clamped to 0.25–8pt. Returns the frame id. */
export function setFrameStroke(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  stroke: FrameStroke | null,
): string | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || (frame?.kind !== "shape" && frame?.kind !== "text" && frame?.kind !== "image")) {
    return null;
  }
  if (stroke) {
    frame.stroke = clampStroke(stroke);
  } else {
    delete frame.stroke;
  }
  touch(model, slide);
  return frame.id;
}

function clampStroke(stroke: FrameStroke): FrameStroke {
  return { color: stroke.color, widthPt: clamp(stroke.widthPt, MIN_STROKE_WIDTH_PT, MAX_STROKE_WIDTH_PT) };
}

// --- media panel edits (SPEC-095) -------------------------------------------

/** Applies `mutate` to the selected media frame's modelled flags and marks the
 * frame edited (drops its verbatim `sourceXml`, dirties the slide) — the model
 * half of the Media panel's flag edits. Returns true when a media frame was
 * found. The caller (`useSlideMedia`) records one undo entry and commits. */
export function editMediaFrame(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  mutate: (frame: MediaFrame) => void,
): boolean {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || frame.kind !== "media") {
    return false;
  }
  mutate(frame);
  markFrameEdited(slide, frame, model.format);
  return true;
}

/** Replaces a media frame's poster still (AC-28): mints a collision-free
 * `mediaId` for the image, adds it to `model.media`, points `posterMediaId` at
 * it, and marks the frame edited. Returns true on success. The caller bumps
 * media and commits. */
export function setMediaPoster(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  poster: MediaPart,
): boolean {
  const slide = slideAt(model, slideIndex);
  const frame = slide ? frameOf(slide, frameId) : null;
  if (!slide || !frame || frame.kind !== "media") {
    return false;
  }
  const mediaId = mintMediaId(model, poster.contentType);
  if (!model.media) {
    model.media = [];
  }
  model.media.push({ mediaId, contentType: poster.contentType, data: poster.data });
  frame.posterMediaId = mediaId;
  markFrameEdited(slide, frame, model.format);
  return true;
}
