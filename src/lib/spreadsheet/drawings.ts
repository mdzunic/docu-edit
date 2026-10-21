// The sheet drawing layer's model logic (SPEC-131 "Anchors follow cells",
// "The drawing layer"): THIS FILE is the normative source for how an anchor
// becomes a rectangle and back, how anchors follow row and column edits, and
// every drawing command. Framework-free: the view (useSheetDrawings,
// useDrawingGestures) wraps each command in one SPEC-121 undo step.
//
// Geometry. A drawing's rectangle is in GridMetrics px at 100% zoom
// (EMU_PER_PX = 9525 at 96 dpi). A point sits at colLeft(col) +
// min(colOffEmu / 9525, colWidth(col)) (rows likewise): offsets are kept as
// read, even when larger than their cell, and clamped only here.
// | Behaviour | Anchors | Rectangle |
// | moveAndSize | twoCell, editAs absent or "twoCell" | from → to |
// | move | oneCell; twoCell editAs "oneCell" | from + extEmu |
// | fixed | absolute; twoCell editAs "absolute" | posEmu + extEmu |
// anchorFromRectPx is the inverse: the cell under each corner and the
// offset into it, always smaller than the cell (colAt never answers a
// zero-sized column, and the offset is clamped one EMU inside it).
//
// Structural edits (drawingsAfterEdit, the `drawings` reference adjuster,
// rows shown — columns are symmetric):
// | Edit | moveAndSize | move | fixed |
// | insert `count` at `at ≤ from.row` | from.row, to.row += count | from.row += count | — |
// | insert with from.row < at ≤ to.row | to.row += count (grows) | — | — |
// | delete band containing neither point | points below shift up | from below shifts up | — |
// | delete band containing to only | to = (at, 0) | — | — |
// | delete band containing from only | from = (at, 0), to shifts up | from = (at, 0) | — |
// | delete band containing both | drawing deleted | from = (at, 0) | — |
// A move or fixed drawing is never deleted by an edit, and only `from`
// moves: its size is `extEmu`. Hiding and resizing are not edits — the layer
// re-lays from the model, so moveAndSize drawings stretch and collapse with
// their cells while the others keep extEmu. A move of cells (SPEC-125) moves
// no drawing (Excel's "move objects with cells" is Backlog).
//
// Recomputed points. The writers derive a move anchor's `to` (and a fixed
// twoCell anchor's `from`/`to`) from extEmu/posEmu under the write-time
// geometry, and keep the source's point unless the derived one differs from
// it by more than one pixel (9,525 EMU) on either axis (recomputedPoint):
// Excel's own rounding never forces a rewrite.
//
// Commands. Each takes the model and the sheet's id and returns the new
// Sheet — a new `drawings` array, untouched drawings by identity — plus, for
// an inserted image, a new `media` array; applyDrawingResult installs it,
// setting `drawingsDirty` (xlsx) or `dirty` (ods: drawings live in the table
// XML). Drawing ids are minted "dr1", "dr2"… unique per sheet; a new
// drawing's name is "Picture {k}" or "{definition label} {k}", k = the
// cNvPr id the xlsx writer mints for it − 1 (nextDrawingNumber).
import { bytesToBase64, type OpaquePart } from "../core/packageParts";
import { extensionForContentType, type MediaPart } from "../core/media";
import { shapeDefinition } from "../core/shapeCatalogue";
import { DEFAULT_NEW_SHAPE_FILL, DEFAULT_NEW_SHAPE_OUTLINE, LINE_GEOMETRIES, type ShapeGeometry, type ShapeVisual } from "../core/shapes";
import type { CellRef } from "./address";
import { buildGridMetrics, type GridMetrics } from "./geometry";
import { MAX_COLS, MAX_ROWS, type AnchorPoint, type DrawingAnchor, type Sheet, type SheetDrawing, type SpreadsheetModel } from "./model";
import type { StructuralEdit } from "./structure";

/** EMU per CSS px at 96 dpi (914,400 / 96). */
export const EMU_PER_PX = 9525;
/** The smallest box a resize leaves, 100%-zoom px. */
export const MIN_DRAWING_SIZE_PX = 8;
/** Arrow-key nudges, 100%-zoom px (Shift: the large one). */
export const NUDGE_PX = 1;
export const NUDGE_LARGE_PX = 10;
/** An inserted image is scaled down to fit this fraction of the visible grid. */
export const IMAGE_FIT_FRACTION = 0.8;

export type AnchorBehaviour = "moveAndSize" | "move" | "fixed";

/** A rectangle in GridMetrics px at 100% zoom. */
export interface RectPx {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** How an anchor follows its cells (see the module header). */
export function anchorBehaviour(anchor: DrawingAnchor): AnchorBehaviour {
  if (anchor.kind === "oneCell") {
    return "move";
  }
  if (anchor.kind === "absolute") {
    return "fixed";
  }
  return anchor.editAs === "oneCell" ? "move" : anchor.editAs === "absolute" ? "fixed" : "moveAndSize";
}

// --- geometry ------------------------------------------------------------------------------

/** A point's position in px: its cell's corner plus the offset, clamped to the cell. */
export function pointPx(point: AnchorPoint, metrics: GridMetrics): { x: number; y: number } {
  return {
    x: metrics.colLeft(point.col) + Math.min(Math.max(0, point.colOffEmu) / EMU_PER_PX, metrics.colWidth(point.col)),
    y: metrics.rowTop(point.row) + Math.min(Math.max(0, point.rowOffEmu) / EMU_PER_PX, metrics.rowHeight(point.row)),
  };
}

/** The point under a px position: the cell holding it and the offset into
 * that cell, in whole EMU, always smaller than the cell. */
export function pointFromPx(x: number, y: number, metrics: GridMetrics): AnchorPoint {
  const px = Math.max(0, x);
  const py = Math.max(0, y);
  const col = metrics.colAt(px);
  const row = metrics.rowAt(py);
  return {
    col,
    colOffEmu: offsetInside(px - metrics.colLeft(col), metrics.colWidth(col)),
    row,
    rowOffEmu: offsetInside(py - metrics.rowTop(row), metrics.rowHeight(row)),
  };
}

/** An offset of `px` into a cell `size` px long, as whole EMU strictly
 * inside it (0 for a zero-sized cell). */
function offsetInside(px: number, size: number): number {
  const limit = Math.round(size * EMU_PER_PX) - 1;
  return Math.max(0, Math.min(Math.round(px * EMU_PER_PX), limit));
}

/** The anchor's rectangle (see the module header). */
export function drawingRectPx(anchor: DrawingAnchor, metrics: GridMetrics): RectPx {
  const behaviour = anchorBehaviour(anchor);
  if (behaviour === "fixed") {
    const pos = anchor.posEmu ?? (anchor.from ? emuOf(anchor.from, metrics) : { x: 0, y: 0 });
    const ext = anchor.extEmu ?? extentBetween(anchor, metrics);
    return { left: pos.x / EMU_PER_PX, top: pos.y / EMU_PER_PX, width: ext.cx / EMU_PER_PX, height: ext.cy / EMU_PER_PX };
  }
  const from = anchor.from ? pointPx(anchor.from, metrics) : { x: 0, y: 0 };
  if (behaviour === "move") {
    const ext = anchor.extEmu ?? extentBetween(anchor, metrics);
    return { left: from.x, top: from.y, width: ext.cx / EMU_PER_PX, height: ext.cy / EMU_PER_PX };
  }
  const to = anchor.to ? pointPx(anchor.to, metrics) : from;
  return { left: from.x, top: from.y, width: Math.max(0, to.x - from.x), height: Math.max(0, to.y - from.y) };
}

/** A point's absolute position in EMU. */
function emuOf(point: AnchorPoint, metrics: GridMetrics): { x: number; y: number } {
  const at = pointPx(point, metrics);
  return { x: Math.round(at.x * EMU_PER_PX), y: Math.round(at.y * EMU_PER_PX) };
}

/** The from → to distance of a twoCell anchor under `metrics`, EMU. */
function extentBetween(anchor: DrawingAnchor, metrics: GridMetrics): { cx: number; cy: number } {
  if (!anchor.from || !anchor.to) {
    return { cx: 0, cy: 0 };
  }
  const from = emuOf(anchor.from, metrics);
  const to = emuOf(anchor.to, metrics);
  return { cx: Math.max(0, to.x - from.x), cy: Math.max(0, to.y - from.y) };
}

/** A twoCell anchor as read, given its size truth: a move anchor gains
 * `extEmu` (from → to under the read-time geometry), a fixed one `posEmu`
 * and `extEmu`. Other anchors come back as they are. */
export function withDerivedSize(anchor: DrawingAnchor, metrics: GridMetrics): DrawingAnchor {
  if (anchor.kind !== "twoCell" || !anchor.from || !anchor.to) {
    return anchor;
  }
  const behaviour = anchorBehaviour(anchor);
  if (behaviour === "move") {
    return { ...anchor, extEmu: extentBetween(anchor, metrics) };
  }
  if (behaviour === "fixed") {
    return { ...anchor, posEmu: emuOf(anchor.from, metrics), extEmu: extentBetween(anchor, metrics) };
  }
  return anchor;
}

/** The anchor that puts a drawing of `behaviour` at `rect`. `like` (the
 * drawing's current anchor) keeps its kind and editAs; without it a
 * moveAndSize anchor is twoCell "twoCell", a move one twoCell "oneCell"
 * (Excel's own default for pictures) and a fixed one absolute. Every point
 * is pointFromPx's, so every offset is smaller than its cell. */
export function anchorFromRectPx(rect: RectPx, behaviour: AnchorBehaviour, metrics: GridMetrics, like?: DrawingAnchor): DrawingAnchor {
  const left = Math.max(0, rect.left);
  const top = Math.max(0, rect.top);
  const width = Math.max(0, rect.width);
  const height = Math.max(0, rect.height);
  const extEmu = { cx: Math.round(width * EMU_PER_PX), cy: Math.round(height * EMU_PER_PX) };
  const from = pointFromPx(left, top, metrics);
  const to = pointFromPx(left + width, top + height, metrics);
  const kind = like?.kind ?? (behaviour === "fixed" ? "absolute" : "twoCell");
  if (kind === "absolute") {
    return { kind: "absolute", posEmu: { x: Math.round(left * EMU_PER_PX), y: Math.round(top * EMU_PER_PX) }, extEmu };
  }
  if (kind === "oneCell") {
    return { kind: "oneCell", from, extEmu };
  }
  const editAs = like?.kind === "twoCell" ? (like.editAs ?? "twoCell") : behaviour === "move" ? "oneCell" : behaviour === "fixed" ? "absolute" : "twoCell";
  if (editAs === "oneCell") {
    return { kind: "twoCell", editAs, from, to, extEmu };
  }
  if (editAs === "absolute") {
    return { kind: "twoCell", editAs, from, to, posEmu: { x: Math.round(left * EMU_PER_PX), y: Math.round(top * EMU_PER_PX) }, extEmu };
  }
  return { kind: "twoCell", editAs, from, to };
}

/** The `from` and `to` a writer emits for a twoCell anchor under `metrics`
 * ("Recomputed points"): a moveAndSize anchor's own; a move anchor's `from`
 * and the `to` its extEmu reaches; a fixed one's both from posEmu + extEmu.
 * Each derived point is the source's (`source`, the points as they stand
 * in the drawing's source XML) while within one pixel of it. */
export function writtenPoints(
  anchor: DrawingAnchor,
  metrics: GridMetrics,
  source?: { from?: AnchorPoint; to?: AnchorPoint },
): { from: AnchorPoint; to: AnchorPoint } {
  const zero: AnchorPoint = { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 };
  const behaviour = anchorBehaviour(anchor);
  if (behaviour === "moveAndSize" || (!anchor.extEmu && behaviour === "move")) {
    const from = anchor.from ?? zero;
    return { from, to: anchor.to ?? from };
  }
  const rect = drawingRectPx(anchor, metrics);
  const from = behaviour === "move" ? (anchor.from ?? zero) : recomputedPoint(pointFromPx(rect.left, rect.top, metrics), source?.from, metrics);
  const to = recomputedPoint(pointFromPx(rect.left + rect.width, rect.top + rect.height, metrics), source?.to, metrics);
  return { from, to };
}

/** `derived`, unless `source` is within one pixel (9,525 EMU) of it on
 * both axes — then `source`, so rounding never rewrites a point. */
export function recomputedPoint(derived: AnchorPoint, source: AnchorPoint | undefined, metrics: GridMetrics): AnchorPoint {
  if (!source) {
    return derived;
  }
  const a = emuOf(derived, metrics);
  const b = emuOf(source, metrics);
  return Math.abs(a.x - b.x) > EMU_PER_PX || Math.abs(a.y - b.y) > EMU_PER_PX ? derived : source;
}

/** Whether a drawing paints: not `hidden`, and — a moveAndSize drawing —
 * not with every row or every column it spans hidden. */
export function isDrawingShown(drawing: SheetDrawing, metrics: GridMetrics): boolean {
  if (drawing.hidden) {
    return false;
  }
  const { anchor } = drawing;
  if (anchorBehaviour(anchor) !== "moveAndSize" || !anchor.from || !anchor.to) {
    return true;
  }
  const rows = spanSize(anchor.from.row, anchor.to.row, anchor.to.rowOffEmu, metrics.rowTop);
  const cols = spanSize(anchor.from.col, anchor.to.col, anchor.to.colOffEmu, metrics.colLeft);
  return rows > 0 && cols > 0;
}

/** The px an anchor's cells span along one axis: from its first index to
 * the last one it reaches into (`to`'s own cell only when its offset is past
 * the cell's start). */
function spanSize(first: number, lastPoint: number, lastOff: number, start: (i: number) => number): number {
  const last = Math.max(first, lastOff > 0 ? lastPoint : lastPoint - 1);
  return start(last + 1) - start(first);
}

/** The drawings that paint, back to front. */
export function visibleDrawings(sheet: Sheet, metrics: GridMetrics): SheetDrawing[] {
  return (sheet.drawings ?? []).filter((drawing) => isDrawingShown(drawing, metrics));
}

// --- structural edits --------------------------------------------------------------------

type Axis = "row" | "col";

const axisOf = (edit: StructuralEdit): Axis => (edit.kind === "insertRows" || edit.kind === "deleteRows" ? "row" : "col");
const isInsertEdit = (edit: StructuralEdit): boolean => edit.kind === "insertRows" || edit.kind === "insertCols";

/** A point with its index along `axis` set (offset `off` when given). */
function withIndex(point: AnchorPoint, axis: Axis, index: number, off?: number): AnchorPoint {
  if (axis === "row") {
    return { ...point, row: index, ...(off === undefined ? {} : { rowOffEmu: off }) };
  }
  return { ...point, col: index, ...(off === undefined ? {} : { colOffEmu: off }) };
}

const indexOf = (point: AnchorPoint, axis: Axis): number => (axis === "row" ? point.row : point.col);

/** A drawing after `edit` on its sheet (see the module header): the same
 * object when it does not move, null when the edit deletes it. */
export function drawingAfterEdit(drawing: SheetDrawing, edit: StructuralEdit): SheetDrawing | null {
  const { anchor } = drawing;
  const behaviour = anchorBehaviour(anchor);
  if (behaviour === "fixed" || !anchor.from) {
    return drawing;
  }
  const axis = axisOf(edit);
  const last = axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
  const { at, count } = edit;
  const from = indexOf(anchor.from, axis);
  const to = anchor.to && behaviour === "moveAndSize" ? indexOf(anchor.to, axis) : undefined;
  let next: DrawingAnchor = anchor;
  if (isInsertEdit(edit)) {
    if (at <= from) {
      next = { ...anchor, from: withIndex(anchor.from, axis, Math.min(from + count, last)) };
      if (to !== undefined) {
        next.to = withIndex(anchor.to!, axis, Math.min(to + count, last));
      }
    } else if (to !== undefined && at <= to) {
      next = { ...anchor, to: withIndex(anchor.to!, axis, Math.min(to + count, last)) };
    }
  } else {
    const end = at + count; // the first index after the band
    const inBand = (i: number): boolean => i >= at && i < end;
    const shifted = (point: AnchorPoint, i: number): AnchorPoint => (i >= end ? withIndex(point, axis, i - count) : point);
    if (to === undefined) {
      if (inBand(from)) {
        next = { ...anchor, from: withIndex(anchor.from, axis, at, 0) };
      } else if (from >= end) {
        next = { ...anchor, from: shifted(anchor.from, from) };
      }
    } else if (inBand(from) && inBand(to)) {
      return null;
    } else if (inBand(from)) {
      next = { ...anchor, from: withIndex(anchor.from, axis, at, 0), to: shifted(anchor.to!, to) };
    } else if (inBand(to)) {
      next = { ...anchor, to: withIndex(anchor.to!, axis, at, 0) };
    } else if (from >= end || to >= end) {
      next = { ...anchor, from: shifted(anchor.from, from), to: shifted(anchor.to!, to) };
    }
  }
  return next === anchor ? drawing : { ...drawing, anchor: next };
}

/** A sheet's drawings after `edit` on it: undefined when none moves or
 * goes (the array keeps its identity). */
export function drawingsAfterEdit(drawings: readonly SheetDrawing[], edit: StructuralEdit): SheetDrawing[] | undefined {
  let changed = false;
  const next: SheetDrawing[] = [];
  for (const drawing of drawings) {
    const moved = drawingAfterEdit(drawing, edit);
    changed ||= moved !== drawing;
    if (moved) {
      next.push(moved);
    }
  }
  return changed ? next : undefined;
}

/** The `drawings` reference adjuster's work on one model (structure.ts
 * registers it): the edited sheet's drawings follow per the table; an xlsx
 * sheet holding drawings is marked `drawingsDirty` — its written points are
 * re-derived under the new geometry even where no anchor moved — and the
 * SPEC-124 retained rows for drawing anchors are retired, so nothing is
 * adjusted twice. */
export function adjustDrawingsForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    if (sheet.id !== edit.sheetId || !sheet.drawings || sheet.drawings.length === 0) {
      return sheet;
    }
    const drawings = drawingsAfterEdit(sheet.drawings, edit);
    if (drawings === undefined && (model.format !== "xlsx" || sheet.drawingsDirty)) {
      return sheet;
    }
    changed = true;
    const next: Sheet = { ...sheet, drawings: drawings ?? sheet.drawings };
    markDrawingsChanged(model, next);
    return next;
  });
  if (changed) {
    model.sheets = sheets;
  }
}

// --- naming and minting -----------------------------------------------------------------

const CNVPR_ID = /<(?:[\w.-]+:)?cNvPr\b[^>]*?\sid\s*=\s*["'](\d+)["']/g;

/** The cNvPr ids a drawing's source XML declares. */
function sourceIds(drawing: SheetDrawing): number[] {
  return drawing.sourceXml === undefined ? [] : [...drawing.sourceXml.matchAll(CNVPR_ID)].map((match) => Number(match[1]));
}

/** The number a created drawing's name ends with (its k), or undefined. */
export function createdNumber(drawing: SheetDrawing): number | undefined {
  if (drawing.sourceXml !== undefined) {
    return undefined;
  }
  const match = /\s(\d+)$/.exec(drawing.name);
  return match ? Number(match[1]) : undefined;
}

/** The k of the next drawing created on the sheet: 1 + the highest of
 * every source cNvPr id − 1 and every created drawing's k. The xlsx writer
 * mints that drawing cNvPr id k + 1 (Excel's own numbering). */
export function nextDrawingNumber(sheet: Sheet): number {
  let max = 0;
  for (const drawing of sheet.drawings ?? []) {
    for (const id of sourceIds(drawing)) {
      max = Math.max(max, id - 1);
    }
    const k = createdNumber(drawing);
    if (k !== undefined) {
      max = Math.max(max, k);
    }
  }
  return max + 1;
}

/** The next drawing id of a sheet ("dr1", "dr2"…). */
export function nextDrawingId(sheet: Sheet): string {
  let max = 0;
  for (const drawing of sheet.drawings ?? []) {
    const match = /^dr(\d+)$/.exec(drawing.id);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `dr${max + 1}`;
}

/** The noun undo labels use: "Move image", "Delete shape"… */
export function drawingNoun(drawing: SheetDrawing): string {
  switch (drawing.content.kind) {
    case "image":
      return "image";
    case "shape":
      return "shape";
    case "group":
      return "group";
    case "chart":
      return "chart";
    case "foreign":
      return "object";
  }
}

/** The package path a new image part takes ("Media parts"): `xl/media/
 * image{n}.{ext}` (xlsx) or `Pictures/image{n}.{ext}` (ods), n = 1 + the
 * highest numeric suffix of the folder's images across `opaque` and `media`. */
export function nextMediaId(model: Pick<SpreadsheetModel, "format" | "opaque" | "media">, contentType: string): string {
  const folder = model.format === "ods" ? "Pictures/" : "xl/media/";
  const pattern = new RegExp(`^${folder.replace("/", "\\/")}image(\\d+)\\.[^/]*$`);
  let max = 0;
  const names = [...model.opaque.map((part: OpaquePart) => part.name), ...(model.media ?? []).map((part) => part.mediaId)];
  for (const name of names) {
    const match = pattern.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `${folder}image${max + 1}.${extensionForContentType(contentType)}`;
}

// --- commands --------------------------------------------------------------------------------

/** A drawing command's outcome: the new sheet and, for an inserted image,
 * the new media array; `drawingId` names the drawing to select. */
export interface DrawingCommandResult {
  sheet: Sheet;
  media?: MediaPart[];
  drawingId?: string;
}

/** Marks a sheet copy as changed in its drawings (see the module header). */
function markDrawingsChanged(model: Pick<SpreadsheetModel, "format">, sheet: Sheet): void {
  if (model.format === "xlsx") {
    sheet.drawingsDirty = true;
  } else {
    sheet.dirty = true;
  }
}

function sheetById(model: SpreadsheetModel, sheetId: string): Sheet | undefined {
  return model.sheets.find((sheet) => sheet.id === sheetId);
}

/** The sheet with `drawings` installed and marked changed. */
function withDrawings(model: SpreadsheetModel, sheet: Sheet, drawings: SheetDrawing[]): Sheet {
  const next: Sheet = { ...sheet, drawings };
  markDrawingsChanged(model, next);
  return next;
}

/** Installs a command's result into the model (the sheet by id; `media`
 * when it has one). */
export function applyDrawingResult(model: SpreadsheetModel, result: DrawingCommandResult): void {
  model.sheets = model.sheets.map((sheet) => (sheet.id === result.sheet.id ? result.sheet : sheet));
  if (result.media) {
    model.media = result.media;
  }
}

/** The rectangle clamped to the grid's top-left (a box never starts above
 * row 0 or left of column A). */
export function clampToGrid(rect: RectPx): RectPx {
  return { ...rect, left: Math.max(0, rect.left), top: Math.max(0, rect.top) };
}

/** A drawing moved or resized to `rect`, re-anchored with its own
 * behaviour (anchorFromRectPx); null when the sheet has no such drawing. */
export function setDrawingRect(model: SpreadsheetModel, sheetId: string, drawingId: string, rect: RectPx, metrics?: GridMetrics): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  const index = sheet?.drawings?.findIndex((drawing) => drawing.id === drawingId) ?? -1;
  if (!sheet || index < 0) {
    return null;
  }
  const drawing = sheet.drawings![index]!;
  const m = metrics ?? buildGridMetrics(sheet);
  const anchor = anchorFromRectPx(clampToGrid(rect), anchorBehaviour(drawing.anchor), m, drawing.anchor);
  const drawings = sheet.drawings!.slice();
  drawings[index] = { ...drawing, anchor };
  return { sheet: withDrawings(model, sheet, drawings), drawingId };
}

/** A drawing removed. */
export function deleteDrawing(model: SpreadsheetModel, sheetId: string, drawingId: string): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  if (!sheet?.drawings?.some((drawing) => drawing.id === drawingId)) {
    return null;
  }
  return { sheet: withDrawings(model, sheet, sheet.drawings.filter((drawing) => drawing.id !== drawingId)) };
}

export type ZOrderMove = "front" | "forward" | "backward" | "back";

/** The undo labels of the four z-order moves. */
export const Z_ORDER_LABELS: Record<ZOrderMove, string> = {
  front: "Bring to front",
  forward: "Bring forward",
  backward: "Send backward",
  back: "Send to back",
};

/** A drawing moved in the z-order (index 0 is the back); null when it is
 * already where the move would put it, or absent. */
export function reorderDrawing(model: SpreadsheetModel, sheetId: string, drawingId: string, move: ZOrderMove): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  const drawings = sheet?.drawings ?? [];
  const index = drawings.findIndex((drawing) => drawing.id === drawingId);
  if (!sheet || index < 0) {
    return null;
  }
  const last = drawings.length - 1;
  const target = move === "front" ? last : move === "back" ? 0 : move === "forward" ? Math.min(last, index + 1) : Math.max(0, index - 1);
  if (target === index) {
    return null;
  }
  const next = drawings.slice();
  const [drawing] = next.splice(index, 1);
  next.splice(target, 0, drawing!);
  return { sheet: withDrawings(model, sheet, next), drawingId };
}

/** A drawing's alt text set — an empty (or blank) text removes it. */
export function setDrawingAltText(model: SpreadsheetModel, sheetId: string, drawingId: string, text: string): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  const index = sheet?.drawings?.findIndex((drawing) => drawing.id === drawingId) ?? -1;
  if (!sheet || index < 0) {
    return null;
  }
  const drawing = sheet.drawings![index]!;
  const want = text.trim() === "" ? undefined : text;
  if (want === drawing.altText) {
    return null;
  }
  const next: SheetDrawing = { ...drawing };
  if (want === undefined) {
    delete next.altText;
  } else {
    next.altText = want;
  }
  const drawings = sheet.drawings!.slice();
  drawings[index] = next;
  return { sheet: withDrawings(model, sheet, drawings), drawingId };
}

/** What insertImage needs: the image's bytes and type, its natural size in
 * px (96 dpi), the cell its top-left lands on and the visible grid area it
 * is scaled down to fit IMAGE_FIT_FRACTION of. */
export interface ImageInsert {
  bytes: Uint8Array;
  contentType: string;
  naturalPx: { width: number; height: number };
  at: CellRef;
  visiblePx: { width: number; height: number };
}

/** The size an image is inserted at: natural, scaled down (aspect kept) to
 * fit IMAGE_FIT_FRACTION of the visible area; at least 1 px each way. */
export function insertedImageSizePx(natural: { width: number; height: number }, visible: { width: number; height: number }): { width: number; height: number } {
  const w = Math.max(1, natural.width);
  const h = Math.max(1, natural.height);
  const scale = Math.min(1, (visible.width * IMAGE_FIT_FRACTION) / w, (visible.height * IMAGE_FIT_FRACTION) / h);
  return { width: Math.max(1, w * scale), height: Math.max(1, h * scale) };
}

/** An image inserted on top (twoCell, editAs "oneCell" — Excel's picture
 * default), its bytes a new media part. */
export function insertImage(model: SpreadsheetModel, sheetId: string, image: ImageInsert, metrics?: GridMetrics): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  if (!sheet || sheet.kind !== "grid") {
    return null;
  }
  const m = metrics ?? buildGridMetrics(sheet);
  const size = insertedImageSizePx(image.naturalPx, image.visiblePx);
  const rect = { left: m.colLeft(image.at.col), top: m.rowTop(image.at.row), ...size };
  const mediaId = nextMediaId(model, image.contentType);
  const part: MediaPart = { mediaId, contentType: image.contentType, data: bytesToBase64(image.bytes) };
  const drawing: SheetDrawing = {
    id: nextDrawingId(sheet),
    anchor: anchorFromRectPx(rect, "move", m),
    name: `Picture ${nextDrawingNumber(sheet)}`,
    content: { kind: "image", mediaId },
  };
  return { sheet: withDrawings(model, sheet, [...(sheet.drawings ?? []), drawing]), media: [...(model.media ?? []), part], drawingId: drawing.id };
}

/** The visual a newly drawn shape takes: DEFAULT_NEW_SHAPE_FILL (none for
 * a line) and DEFAULT_NEW_SHAPE_OUTLINE. */
export function newShapeVisual(geometry: ShapeGeometry): ShapeVisual {
  return LINE_GEOMETRIES.has(geometry)
    ? { geometry, outline: { ...DEFAULT_NEW_SHAPE_OUTLINE } }
    : { geometry, fill: { ...DEFAULT_NEW_SHAPE_FILL }, outline: { ...DEFAULT_NEW_SHAPE_OUTLINE } };
}

/** A shape drawn on top at `rect` (twoCell, editAs "twoCell"). */
export function insertShape(model: SpreadsheetModel, sheetId: string, geometry: ShapeGeometry, rect: RectPx, metrics?: GridMetrics): DrawingCommandResult | null {
  const sheet = sheetById(model, sheetId);
  if (!sheet || sheet.kind !== "grid") {
    return null;
  }
  const m = metrics ?? buildGridMetrics(sheet);
  const label = shapeDefinition(geometry)?.label ?? "Shape";
  const drawing: SheetDrawing = {
    id: nextDrawingId(sheet),
    anchor: anchorFromRectPx(clampToGrid(rect), "moveAndSize", m),
    name: `${label} ${nextDrawingNumber(sheet)}`,
    content: { kind: "shape", visual: newShapeVisual(geometry) },
  };
  return { sheet: withDrawings(model, sheet, [...(sheet.drawings ?? []), drawing]), drawingId: drawing.id };
}
