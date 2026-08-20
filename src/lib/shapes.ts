// Shape/text-box structure commands (SPEC-051), framework-free like every
// lib module. A Shape is a block-level DrawingPlacement, addressed the same
// way images.ts addresses an Image — see textEditing.ts's blockLength,
// which treats it (like Table and Image) as a length-1 atom. This module
// owns inserting, drawing, resizing, rotating, moving/nudging, duplicating,
// z-ordering, grouping and re-anchoring shapes; images.ts keeps its own
// (near-identical) command set because SPEC-024 shipped first and this spec
// promises "no shipped call site changes" to it.
import {
  cloneContainers,
  isChart,
  isImage,
  isShape,
  splitParagraphAt,
  type Block,
  type Chart,
  type DocumentModel,
  type DrawingPlacement,
  type Image,
  type ImageWrapMode,
  type Paragraph,
  type Shape,
} from "./documentModel";
import { type ShapeGeometry, type ShapeOutline } from "./core/shapes";
import {
  anchoredParagraph,
  ordinalForParagraph,
  releaseAnchorOrdinal,
  resolveAnchorBlockIndex,
} from "./drawingAnchors";
import { blockLength, clampCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { LINE_GEOMETRIES } from "./core/shapes";
import {
  MIN_SHAPE_SIZE_CM,
  normalizeRotationDeg,
  NUDGE_FINE_PT_CM,
  NUDGE_PT_CM,
  type NudgeDirection,
} from "./core/shapes";

/** Where a shape command applies — identical shape to images.ts's
 * ImageLocation, restated here because the two modules are independent
 * (mirrors the images.ts/tables.ts precedent of not sharing a location
 * type across block kinds). */
export interface ShapeLocation {
  section: number;
  block: number;
}

/** Typed `null` seed for a shape-location ref in `.vue` scripts (espree bars
 * the annotation — see images.ts's noImageLocation). */
export function noShapeLocation(): ShapeLocation | null {
  return null;
}


function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

/** The shape a location addresses, or null. */
export function shapeAt(model: DocumentModel, location: ShapeLocation): Shape | null {
  const target = model.sections[location.section]?.blocks[location.block];
  return isShape(target) ? target : null;
}

/** Whether a shape's text frame holds any actual content — an absent
 * `text`, or one whose only paragraph is empty, counts as no text
 * ("Selection and hit testing"'s `hasText` only applies to a textFrame
 * *holding text*). Any non-paragraph block (a table, say) counts as
 * content without inspecting it further. */
export function shapeHasVisibleText(shape: Shape): boolean {
  if (!shape.text) {
    return false;
  }
  return shape.text.blocks.some((block) =>
    block.type === "paragraph" ? block.runs.some((run) => run.text.length > 0) : true,
  );
}

/** Every Shape block (not Image — hit testing this fixes is shape-specific,
 * see documentModel.ts's Shape doc comment), ordered topmost-first for
 * "Selection and hit testing": descending `zOrder`, ties broken by later
 * document position (later blocks paint on top of earlier ones at the same
 * z-index bucket — see renderFormatting.ts's shapeDisplayStyle). */
export function shapesInPaintOrder(model: DocumentModel): { location: ShapeLocation; shape: Shape }[] {
  const drawings: { location: ShapeLocation; shape: Shape }[] = [];
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isShape(block)) {
        drawings.push({ location: { section: sectionIndex, block: blockIndex }, shape: block });
      }
    });
  });
  return drawings.sort((a, b) => {
    const za = zOrderOf(a.shape);
    const zb = zOrderOf(b.shape);
    if (za !== zb) {
      return zb - za;
    }
    if (a.location.section !== b.location.section) {
      return b.location.section - a.location.section;
    }
    return b.location.block - a.location.block;
  });
}

function replaceShape(model: DocumentModel, location: ShapeLocation, next: Shape): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** The default anchored position for a non-inline drawing when no explicit
 * offset is given ("Default position" — identical for square/topAndBottom/
 * behind/infront, so a no-edit round trip of a position-less object stays
 * structurally equal). */
export const DEFAULT_DRAWING_POSITION = {
  relativeToH: "column" as const,
  relativeToV: "paragraph" as const,
  offsetXCm: 0,
  offsetYCm: 0,
};

/** New-shape geometry/placement, everything insertShape needs beyond where
 * to put it. `widthCm`/`heightCm` are the dragged (or default 3x2cm/3cm-line)
 * size (AC-4..AC-9); `offsetXCm`/`offsetYCm` place it precisely where drawn,
 * relative to the anchor paragraph the caller has already chosen (nearest
 * to the drop point) — absent falls back to DEFAULT_DRAWING_POSITION. */
export interface NewShape {
  geometry: ShapeGeometry;
  widthCm: number;
  heightCm: number;
  offsetXCm?: number;
  offsetYCm?: number;
  fill?: string;
  outline?: ShapeOutline;
  text?: Shape["text"];
}

/** Inserts a shape at the caret (AC-3..AC-9), splitting the caret's
 * paragraph around the insertion point exactly like images.ts's
 * insertImage. A freshly inserted shape floats (`wrap: "square"`) at the
 * given offset (or the shared anchored default) rather than flowing inline,
 * since draw mode always places a shape at an explicit page position. */
export function insertShape(
  model: DocumentModel,
  selection: ModelSelection,
  spec: NewShape,
): CommandResult {
  const next = cloneContainers(model);
  const at = clampCaret(next, selection.focus);
  if (at.cell) {
    return { model: next, selection: caretSelection(at) };
  }
  const blocks = next.sections[at.section].blocks;
  const target = blocks[at.block];
  const shape: Shape = {
    type: "shape",
    geometry: spec.geometry,
    widthCm: spec.widthCm,
    heightCm: spec.heightCm,
    wrap: "square",
    relativeToH: DEFAULT_DRAWING_POSITION.relativeToH,
    relativeToV: DEFAULT_DRAWING_POSITION.relativeToV,
    offsetXCm: spec.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm,
    offsetYCm: spec.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm,
  };
  if (spec.fill !== undefined) {
    shape.fill = spec.fill;
  }
  if (spec.outline !== undefined) {
    shape.outline = spec.outline;
  }
  if (spec.text !== undefined) {
    shape.text = spec.text;
  }

  const replacement: Block[] = [];
  let shapeBlockOffset: number;
  // The anchor pairing (SPEC-078): the shape pairs with the paragraph whose
  // top the caller measured its offsets from — the caret paragraph itself,
  // or the FIRST half when the insert splits it. A non-paragraph target
  // leaves the shape unpaired (nothing paragraph-shaped to measure from);
  // resolveAnchorBlockIndex's fallback then applies, as it does for every
  // pre-SPEC-078 document.
  if (target.type !== "paragraph") {
    if (at.offset === 0) {
      replacement.push(shape, target);
      shapeBlockOffset = 0;
    } else {
      replacement.push(target, shape, emptyParagraph());
      shapeBlockOffset = 1;
    }
  } else {
    const ordinal = ordinalForParagraph(next, target);
    shape.anchorOrdinal = ordinal;
    if (at.offset === 0) {
      replacement.push(shape, anchoredParagraph(target, ordinal));
      shapeBlockOffset = 0;
    } else if (at.offset >= blockLength(target)) {
      replacement.push(anchoredParagraph(target, ordinal), shape, emptyParagraph());
      shapeBlockOffset = 1;
    } else {
      const [before, after] = splitParagraphAt(target, at.offset);
      replacement.push(anchoredParagraph(before, ordinal), shape, after);
      shapeBlockOffset = 1;
    }
  }
  blocks.splice(at.block, 1, ...replacement);
  const shapeIndex = at.block + shapeBlockOffset;
  const landing: CaretPosition = { section: at.section, block: shapeIndex + 1, offset: 0 };
  return { model: next, selection: caretSelection(landing) };
}

/** "Insert -> Text box" (AC-9): a `rect` with no fill/outline and an empty
 * text frame, already positioned to enter text-editing mode. Not a
 * fourteenth geometry — see documentModel.ts's Shape doc comment. */
export function insertTextBox(model: DocumentModel, selection: ModelSelection): CommandResult {
  return insertShape(model, selection, {
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    text: { blocks: [emptyParagraph()], vAlign: "top" },
  });
}

/** Double-clicking a shape with no text frame yet gives it one, empty, so
 * the caret has somewhere to land (AC-17's "populating Shape.text if
 * absent") — a no-op once the shape already has one. */
export function ensureShapeTextFrame(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape || shape.text) {
    return { model: next, selection };
  }
  replaceShape(next, location, { ...shape, text: { blocks: [emptyParagraph()], vAlign: "top" } });
  return { model: next, selection };
}

/** Sets a shape's display size in cm (resize-handle drag on release, and the
 * shape toolbar's size fields). Clamped like images.ts's MIN_IMAGE_SIZE_CM
 * so a drag can never collapse a shape to nothing (a zero-height Line stays
 * allowed — that is a real horizontal line, not a collapsed shape; the
 * minimum itself is core/shapes.ts's MIN_SHAPE_SIZE_CM). */
export function setShapeSize(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  widthCm: number,
  heightCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const minH = LINE_GEOMETRIES.has(shape.geometry) ? 0 : MIN_SHAPE_SIZE_CM;
  replaceShape(next, location, {
    ...shape,
    widthCm: Math.max(MIN_SHAPE_SIZE_CM, widthCm),
    heightCm: Math.max(minH, heightCm),
  });
  return { model: next, selection };
}

/** Moves a shape by a page-space delta (a drag, or a keyboard nudge —
 * AC-40, AC-41), adjusting its own offset. Used directly for a shape whose
 * anchor did not change; a drag that crosses a paragraph boundary instead
 * goes through reanchorShape, which folds the same delta into the new
 * anchor's offset. */
export function moveShape(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  dxCm: number,
  dyCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  replaceShape(next, location, {
    ...shape,
    offsetXCm: (shape.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm) + dxCm,
    offsetYCm: (shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) + dyCm,
  });
  return { model: next, selection };
}

/** Arrow-key nudge (AC-40) / Ctrl+arrow fine nudge (AC-41): 1pt or 0.1pt in
 * the given direction. */
export function nudgeShape(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  direction: NudgeDirection,
  fine: boolean,
): CommandResult {
  const step = fine ? NUDGE_FINE_PT_CM : NUDGE_PT_CM;
  const dx = direction === "left" ? -step : direction === "right" ? step : 0;
  const dy = direction === "up" ? -step : direction === "down" ? step : 0;
  return moveShape(model, selection, location, dx, dy);
}

export function setShapeFill(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  fill: string | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const updated: Shape = { ...shape };
  if (fill) {
    updated.fill = fill;
  } else {
    delete updated.fill;
  }
  replaceShape(next, location, updated);
  return { model: next, selection };
}

export function setShapeOutline(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  outline: ShapeOutline | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const updated: Shape = { ...shape };
  if (outline) {
    updated.outline = outline;
  } else {
    delete updated.outline;
  }
  replaceShape(next, location, updated);
  return { model: next, selection };
}

export function setShapeRotation(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  rotationDeg: number,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const normalized = normalizeRotationDeg(rotationDeg);
  const updated: Shape = { ...shape };
  if (normalized === 0) {
    delete updated.rotationDeg;
  } else {
    updated.rotationDeg = normalized;
  }
  replaceShape(next, location, updated);
  return { model: next, selection };
}

export function setShapeWrap(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  wrap: ImageWrapMode,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const updated: Shape = { ...shape, wrap };
  if (wrap === "inline") {
    delete updated.offsetXCm;
    delete updated.offsetYCm;
    delete updated.relativeToH;
    delete updated.relativeToV;
    // An inline shape has no anchor (SPEC-078): drop the pairing and let
    // the paragraph's marker go with it when this was the last referrer.
    delete updated.anchorOrdinal;
    replaceShape(next, location, updated);
    releaseAnchorOrdinal(next, shape.anchorOrdinal);
    return { model: next, selection };
  }
  if (updated.offsetXCm === undefined) {
    updated.offsetXCm = DEFAULT_DRAWING_POSITION.offsetXCm;
    updated.offsetYCm = DEFAULT_DRAWING_POSITION.offsetYCm;
    updated.relativeToH = DEFAULT_DRAWING_POSITION.relativeToH;
    updated.relativeToV = DEFAULT_DRAWING_POSITION.relativeToV;
    // Newly floating (SPEC-078): pair with the paragraph the offsets will
    // be measured from — the adjacency-resolved one, since an inline shape
    // carries no ordinal to honour.
    const blocks = next.sections[location.section].blocks;
    const anchorIndex = resolveAnchorBlockIndex(blocks, location.block);
    const anchorBlock = anchorIndex === null ? undefined : blocks[anchorIndex];
    if (anchorIndex !== null && anchorBlock?.type === "paragraph") {
      const ordinal = ordinalForParagraph(next, anchorBlock);
      updated.anchorOrdinal = ordinal;
      blocks.splice(anchorIndex, 1, anchoredParagraph(anchorBlock, ordinal));
    }
  }
  replaceShape(next, location, updated);
  return { model: next, selection };
}

export function setShapeAltText(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  altText: string,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const updated: Shape = { ...shape };
  if (altText.length > 0) {
    updated.altText = altText;
  } else {
    delete updated.altText;
  }
  replaceShape(next, location, updated);
  return { model: next, selection };
}

/** Deletes a selected shape (AC-42): the caret lands where the shape was —
 * the same rule images.ts's deleteImage uses. */
export function deleteShape(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block, 1);
  // A paragraph carries anchorOrdinal only while it anchors something
  // (SPEC-078's release rule); the deleted shape may have been the last
  // referrer.
  releaseAnchorOrdinal(next, shape.anchorOrdinal);
  if (blocks.length === 0) {
    blocks.push(emptyParagraph());
  }
  const landing = Math.min(location.block, blocks.length - 1);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: landing, offset: 0 }),
  };
}

/** Every drawing (Shape, Image or Chart) block's placement + location, used
 * by z-order and grouping commands that must see every drawing to compute a
 * fresh ordinal/z-order value or a group's members — Chart participates here
 * "for free" as a third DrawingPlacement (SPEC-052 AC-24), the same way
 * Image joined Shape's own z-order/grouping machinery rather than growing a
 * duplicate of it. */
interface DrawingRef {
  location: ShapeLocation;
  block: Shape | Image | Chart;
}

function allDrawings(model: DocumentModel): DrawingRef[] {
  const drawings: DrawingRef[] = [];
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isShape(block) || isImage(block) || isChart(block)) {
        drawings.push({ location: { section: sectionIndex, block: blockIndex }, block });
      }
    });
  });
  return drawings;
}

function zOrderOf(block: DrawingPlacement): number {
  return block.zOrder ?? 0;
}

/** Duplicates a shape (AC-43): a copy offset 0.5cm down and right, one
 * z-order step nearer the reader than its source, inserted immediately
 * after the source and left selected. The copy keeps its source's
 * `anchorOrdinal` (the spread below carries it): two drawings, one anchor
 * paragraph, by construction — the arrangement SPEC-078's explicit pairing
 * exists to keep unambiguous. */
export const DUPLICATE_OFFSET_CM = 0.5;

export function duplicateShape(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const maxZ = Math.max(0, ...allDrawings(next).map((d) => zOrderOf(d.block)));
  const copy: Shape = {
    ...shape,
    offsetXCm: (shape.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm) + DUPLICATE_OFFSET_CM,
    offsetYCm: (shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) + DUPLICATE_OFFSET_CM,
    zOrder: Math.max(zOrderOf(shape) + 1, maxZ + 1),
  };
  delete (copy as { groupOrdinal?: number }).groupOrdinal;
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block + 1, 0, copy);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: location.block + 1, offset: 0 }),
  };
}

export type ZOrderCommand = "front" | "forward" | "backward" | "back";

/** Bring to front / bring forward / send backward / send to back (AC-29..
 * AC-33): reorders `location`'s drawing against every other drawing in the
 * document by `zOrder`. Forward/backward swap with the immediate neighbour
 * in the current stacking order; front/back jump to one past the current
 * extreme. Every drawing gets an explicit `zOrder` after this (no more
 * "absent means 0" ties), so the painted order this produces is stable. */
export function reorderShapeZ(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  command: ZOrderCommand,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const drawings = allDrawings(next).sort((a, b) => zOrderOf(a.block) - zOrderOf(b.block));
  const index = drawings.findIndex((d) => d.location.section === location.section && d.location.block === location.block);
  if (index < 0) {
    return { model: next, selection };
  }
  let targetIndex = index;
  if (command === "front") {
    targetIndex = drawings.length - 1;
  } else if (command === "back") {
    targetIndex = 0;
  } else if (command === "forward") {
    targetIndex = Math.min(drawings.length - 1, index + 1);
  } else {
    targetIndex = Math.max(0, index - 1);
  }
  const reordered = [...drawings];
  const [moved] = reordered.splice(index, 1);
  reordered.splice(targetIndex, 0, moved);
  reordered.forEach((d, i) => {
    const updated: Shape | Image | Chart = { ...d.block };
    updated.zOrder = i;
    next.sections[d.location.section].blocks.splice(d.location.block, 1, updated);
  });
  return { model: next, selection };
}

/** A fresh groupOrdinal not already used by any drawing in the document
 * (AC-35) — deliberately not named `id`, see documentModel.ts's Shape doc
 * comment. */
export function freshGroupOrdinal(model: DocumentModel): number {
  let max = 0;
  for (const { block } of allDrawings(model)) {
    if (block.groupOrdinal !== undefined) {
      max = Math.max(max, block.groupOrdinal);
    }
  }
  return max + 1;
}

/** Groups two or more drawings (AC-34..AC-36): assigns every member the
 * same fresh groupOrdinal. Members are always floating — an `inline` member
 * is first converted to `square` at the shared default position so it has
 * offsets to express a relative position ("Grouping — move, resize,
 * rotate"). */
export function groupShapes(
  model: DocumentModel,
  selection: ModelSelection,
  locations: ShapeLocation[],
): CommandResult {
  const next = cloneContainers(model);
  const ordinal = freshGroupOrdinal(next);
  for (const location of locations) {
    const block = next.sections[location.section]?.blocks[location.block];
    if (!isShape(block) && !isImage(block) && !isChart(block)) {
      continue;
    }
    const updated: Shape | Image | Chart = { ...block };
    updated.groupOrdinal = ordinal;
    if (updated.wrap === "inline") {
      updated.wrap = "square";
      updated.relativeToH = DEFAULT_DRAWING_POSITION.relativeToH;
      updated.relativeToV = DEFAULT_DRAWING_POSITION.relativeToV;
      updated.offsetXCm = DEFAULT_DRAWING_POSITION.offsetXCm;
      updated.offsetYCm = DEFAULT_DRAWING_POSITION.offsetYCm;
    }
    next.sections[location.section].blocks.splice(location.block, 1, updated);
  }
  return { model: next, selection };
}

/** Ungroups every drawing sharing `groupOrdinal` (AC-39): clears the field,
 * keeping each member's current position/size/rotation so nothing jumps. */
export function ungroupShapes(
  model: DocumentModel,
  selection: ModelSelection,
  groupOrdinal: number,
): CommandResult {
  const next = cloneContainers(model);
  for (const { location, block } of allDrawings(next)) {
    if (block.groupOrdinal !== groupOrdinal) {
      continue;
    }
    const updated: Shape | Image | Chart = { ...block };
    delete updated.groupOrdinal;
    next.sections[location.section].blocks.splice(location.block, 1, updated);
  }
  return { model: next, selection };
}

/** Every drawing sharing a groupOrdinal, in document order. */
export function groupMembers(model: DocumentModel, groupOrdinal: number): DrawingRef[] {
  return allDrawings(model).filter((d) => d.block.groupOrdinal === groupOrdinal);
}

/** Moves every member of a group by the same page-space delta ("Grouping —
 * move" — AC-36). Sizes are untouched. */
export function moveGroup(
  model: DocumentModel,
  selection: ModelSelection,
  groupOrdinal: number,
  dxCm: number,
  dyCm: number,
): CommandResult {
  let next = model;
  for (const { location } of groupMembers(model, groupOrdinal)) {
    next = moveShape(next, selection, location, dxCm, dyCm).model;
  }
  return { model: next, selection };
}

/** The page-space box spanning every member of a group — the pivot the
 * resize-handle drag scales about (AC-37). `null` for an ordinal with no
 * members (nothing to bound). */
export interface GroupBoundingBox {
  leftCm: number;
  topCm: number;
  widthCm: number;
  heightCm: number;
}

export function groupBoundingBoxCm(model: DocumentModel, groupOrdinal: number): GroupBoundingBox | null {
  const members = groupMembers(model, groupOrdinal);
  if (members.length === 0) {
    return null;
  }
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const { block } of members) {
    const x = block.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm;
    const y = block.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm;
    left = Math.min(left, x);
    top = Math.min(top, y);
    right = Math.max(right, x + block.widthCm);
    bottom = Math.max(bottom, y + block.heightCm);
  }
  return { leftCm: left, topCm: top, widthCm: right - left, heightCm: bottom - top };
}

/** Scales every group member's dimensions and offsets by the given per-axis
 * factor ("Grouping — resize", AC-37) — the caller (the group resize-handle
 * drag) computes `sx`/`sy` from the dragged box against the group's
 * measured bounding box, `sx === sy` when Shift is held. Member
 * `rotationDeg` is unchanged. */
export function resizeGroup(
  model: DocumentModel,
  selection: ModelSelection,
  groupOrdinal: number,
  sx: number,
  sy: number,
): CommandResult {
  const next = cloneContainers(model);
  for (const { location, block } of groupMembers(next, groupOrdinal)) {
    const updated: Shape | Image | Chart = { ...block };
    updated.widthCm = Math.max(MIN_SHAPE_SIZE_CM, updated.widthCm * sx);
    updated.heightCm = Math.max(0, updated.heightCm * sy);
    updated.offsetXCm = (updated.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm) * sx;
    updated.offsetYCm = (updated.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) * sy;
    next.sections[location.section].blocks.splice(location.block, 1, updated);
  }
  return { model: next, selection };
}

/** Adds `deltaDeg` to every group member's own rotationDeg ("Grouping —
 * rotate", AC-38) — each member spins about its own centre, which is what
 * "keeping each member's centre fixed relative to the group centre" means
 * (a member's own centre never moves when only its rotationDeg changes). */
export function rotateGroup(
  model: DocumentModel,
  selection: ModelSelection,
  groupOrdinal: number,
  deltaDeg: number,
): CommandResult {
  const next = cloneContainers(model);
  for (const { location, block } of groupMembers(next, groupOrdinal)) {
    const updated: Shape | Image | Chart = { ...block };
    const normalized = normalizeRotationDeg((block.rotationDeg ?? 0) + deltaDeg);
    if (normalized === 0) {
      delete updated.rotationDeg;
    } else {
      updated.rotationDeg = normalized;
    }
    next.sections[location.section].blocks.splice(location.block, 1, updated);
  }
  return { model: next, selection };
}

// --- re-anchoring ("the nearest paragraph") ---------------------------------

/** A candidate anchor paragraph's vertical span in page coordinates, plus
 * its address — supplied by the caller (the renderer knows layout, this
 * module stays framework-free/layout-free). */
export interface ParagraphSpan {
  section: number;
  block: number;
  topCm: number;
  bottomCm: number;
}

/** "Re-anchoring — the nearest paragraph": scores each candidate by the
 * distance from the dragged object's vertical centre `cy` to the nearer
 * edge of the paragraph's span (0 when `cy` falls inside it); the
 * lowest-scoring candidate wins, ties going to the earlier paragraph in
 * document order (`candidates` is assumed to already be in document order,
 * i.e. already filtered to the same container level as the current
 * anchor). Returns null when there are no candidates — the caller then
 * keeps the object's current anchor. */
export function nearestParagraph(cy: number, candidates: ParagraphSpan[]): ParagraphSpan | null {
  let best: ParagraphSpan | null = null;
  let bestScore = Infinity;
  for (const candidate of candidates) {
    const score = Math.max(candidate.topCm - cy, 0, cy - candidate.bottomCm);
    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return best;
}

/** Re-anchors a shape to the nearest candidate paragraph (AC-28): moves the
 * shape block to sit at the new anchor's position in the block list,
 * adjusts `offsetYCm` by the delta between the old and new anchor's tops so
 * the shape stays exactly where it was dropped (`relativeToH`/`relativeToV`
 * are unchanged), and re-points the anchor pairing (SPEC-078) — the new
 * anchor paragraph gets an `anchorOrdinal` if it lacked one, the shape
 * references it, and the old pairing is released when this shape was its
 * last referrer. Only the dragged shape's pairing changes: another drawing
 * anchored to either paragraph keeps its own ordinal, offsets and resolved
 * anchor untouched. No-ops (keeps the current anchor) when there is no
 * candidate, or the nearest candidate IS the current anchor. */
export function reanchorShape(
  model: DocumentModel,
  selection: ModelSelection,
  location: ShapeLocation,
  candidates: ParagraphSpan[],
  cy: number,
  oldAnchorTopCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const shape = shapeAt(next, location);
  if (!shape) {
    return { model: next, selection };
  }
  const nearest = nearestParagraph(cy, candidates);
  if (!nearest || (nearest.section === location.section && nearest.block === location.block)) {
    return { model: next, selection };
  }
  const anchorPara = next.sections[nearest.section]?.blocks[nearest.block];
  if (anchorPara?.type !== "paragraph") {
    // Candidates are paragraph spans measured against the same model; a
    // non-paragraph here means the caller's spans are stale — keep the
    // current anchor rather than pairing with the wrong block.
    return { model: next, selection };
  }
  const ordinal = ordinalForParagraph(next, anchorPara);
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block, 1);
  // The nearest candidate's index shifts down by one if it followed the
  // removed shape in the same section's block list.
  let targetBlock = nearest.block;
  if (nearest.section === location.section && nearest.block > location.block) {
    targetBlock -= 1;
  }
  const deltaCm = nearest.topCm - oldAnchorTopCm;
  const updated: Shape = {
    ...shape,
    offsetYCm: (shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) - deltaCm,
    anchorOrdinal: ordinal,
  };
  const targetBlocks = next.sections[nearest.section].blocks;
  if (anchorPara.anchorOrdinal !== ordinal) {
    targetBlocks.splice(targetBlock, 1, anchoredParagraph(anchorPara, ordinal));
  }
  // Lands immediately after its new anchor paragraph — the shape is a
  // top-level block "at the position it is anchored to", the same
  // convention insertShape/insertImage use for a floating object.
  const insertAt = targetBlock + 1;
  targetBlocks.splice(insertAt, 0, updated);
  // The updated shape already carries the new ordinal, so releasing the old
  // one is a no-op unless this shape really was its last referrer (and when
  // old === new, the shape itself keeps the pairing alive).
  releaseAnchorOrdinal(next, shape.anchorOrdinal);
  const newLocation: ShapeLocation = { section: nearest.section, block: insertAt };
  return {
    model: next,
    selection: caretSelection({ section: newLocation.section, block: newLocation.block + 1, offset: 0 }),
  };
}
