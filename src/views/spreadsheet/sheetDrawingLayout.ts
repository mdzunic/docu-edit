// The sheet drawing layer's view geometry and paint decisions (SPEC-131 "The
// drawing layer"), framework-free so the pane split, the resize arithmetic
// and what each drawing paints are unit-testable without mounting a grid.
// drawings.ts owns the model — anchors, rectangles at 100% zoom, every
// command; this file owns where a rectangle shows and what fills it.
//
// Coordinates. Every rect is GridMetrics px at 100% zoom ("sheet px"). The
// grid scales its whole content box by the zoom (SheetGrid), so the layer
// places drawings in these units and the transform does the rest. The cell
// area (the layer's box, at the header offsets) shows the frozen bands in
// place and the rest offset by the scroll (useSheetViewport's sx/sy, content
// px): a sheet point (x, y) shows at x when its column is frozen, else x − sx,
// and at y when its row is frozen, else y − sy. Each pane is one clipping box
// of the cell area — corner [0, frozenW] × [0, frozenH] (origin 0, 0), top
// [frozenW, viewCols] × [0, frozenH] (origin frozenW + sx, 0), left
// [0, frozenW] × [frozenH, viewRows] (origin 0, frozenH + sy) and main (origin
// frozenW + sx, frozenH + sy) — and a drawing sits in a pane at
// (x − originX, y − originY). A drawing is mounted in every pane whose window
// (what the pane shows, plus one screen of overscan each way) its rectangle
// intersects, and the pane's overflow clips each copy, so a drawing straddling
// a freeze line shows split exactly at the line, the scrolled part moving with
// the scroll. One copy is `primary` — the pane holding the rectangle's
// top-left, else the first that mounts it: it alone is focusable, named and
// handled; the others are paint-only (aria-hidden).
//
// Resizing (resizedRect): a handle moves its edges by the pointer's delta,
// never past MIN_DRAWING_SIZE_PX (an axis that started at 0 — a straight
// line — may stay 0) nor past the grid's top-left; on a corner with
// `keepAspect` the larger relative change wins and the opposite corner stays.
import { DEFAULT_THEME, type DocumentTheme } from "../../lib/core/theme";
import type { ShapeVisual } from "../../lib/core/shapes";
import { MIN_DRAWING_SIZE_PX, type RectPx } from "../../lib/spreadsheet/drawings";
import type { DrawingText, SheetDrawing, SheetDrawingContent } from "../../lib/spreadsheet/model";
import { cellColorHex } from "../../lib/spreadsheet/styles";

export type DrawingPaneId = "corner" | "top" | "left" | "main";

/** What the pane split reads off the viewport (content px, unscaled). */
export interface PaneLayout {
  frozenW: number;
  frozenH: number;
  viewCols: number;
  viewRows: number;
  sx: number;
  sy: number;
}

/** One grid pane: its clipping box in the cell area and the sheet point at
 * its top-left (see the module header). */
export interface DrawingPane {
  id: DrawingPaneId;
  left: number;
  top: number;
  width: number;
  height: number;
  originX: number;
  originY: number;
}

/** A painted drawing and its rectangle (sheet px). */
export interface ShownDrawing {
  drawing: SheetDrawing;
  rect: RectPx;
}

/** One drawing's copy in a pane, placed in the pane's own px. */
export interface PaneItem {
  drawing: SheetDrawing;
  x: number;
  y: number;
  width: number;
  height: number;
  primary: boolean;
  selected: boolean;
}

export interface PaneView extends DrawingPane {
  items: PaneItem[];
}

/** The panes the viewport shows: up to four, an empty one omitted. */
export function drawingPanes(layout: PaneLayout): DrawingPane[] {
  const fw = Math.max(0, Math.min(layout.frozenW, layout.viewCols));
  const fh = Math.max(0, Math.min(layout.frozenH, layout.viewRows));
  const mainW = Math.max(0, layout.viewCols - fw);
  const mainH = Math.max(0, layout.viewRows - fh);
  const panes: DrawingPane[] = [
    { id: "corner", left: 0, top: 0, width: fw, height: fh, originX: 0, originY: 0 },
    { id: "top", left: fw, top: 0, width: mainW, height: fh, originX: fw + layout.sx, originY: 0 },
    { id: "left", left: 0, top: fh, width: fw, height: mainH, originX: 0, originY: fh + layout.sy },
    { id: "main", left: fw, top: fh, width: mainW, height: mainH, originX: fw + layout.sx, originY: fh + layout.sy },
  ];
  return panes.filter((pane) => pane.width > 0 && pane.height > 0);
}

/** The pane holding a rectangle's top-left point. */
export function primaryPaneOf(rect: RectPx, layout: PaneLayout): DrawingPaneId {
  const frozenRow = layout.frozenH > 0 && rect.top < layout.frozenH;
  const frozenCol = layout.frozenW > 0 && rect.left < layout.frozenW;
  if (frozenRow) {
    return frozenCol ? "corner" : "top";
  }
  return frozenCol ? "left" : "main";
}

/** Whether a rectangle (a zero-sized line included) meets the pane's window
 * — what it shows, grown by `overscan` px on each side. */
export function inPaneWindow(rect: RectPx, pane: DrawingPane, overscan: { x: number; y: number }): boolean {
  const left = pane.originX - overscan.x;
  const top = pane.originY - overscan.y;
  const right = pane.originX + pane.width + overscan.x;
  const bottom = pane.originY + pane.height + overscan.y;
  return rect.left <= right && rect.left + rect.width >= left && rect.top <= bottom && rect.top + rect.height >= top;
}

/** Every pane with the drawings it mounts (back to front), one screen of
 * overscan each way (see the module header). */
export function layoutDrawingPanes(shown: readonly ShownDrawing[], layout: PaneLayout, selectedId: string | null): PaneView[] {
  const panes = drawingPanes(layout);
  const overscan = { x: layout.viewCols, y: layout.viewRows };
  const mounted = panes.map((pane) => shown.filter((entry) => inPaneWindow(entry.rect, pane, overscan)));
  const primary = new Map<string, DrawingPaneId>();
  for (const entry of shown) {
    const want = primaryPaneOf(entry.rect, layout);
    const index = panes.findIndex((pane, i) => pane.id === want && mounted[i]!.includes(entry));
    const first = index >= 0 ? index : mounted.findIndex((entries) => entries.includes(entry));
    if (first >= 0) {
      primary.set(entry.drawing.id, panes[first]!.id);
    }
  }
  return panes.map((pane, i) => ({
    ...pane,
    items: mounted[i]!.map(({ drawing, rect }) => ({
      drawing,
      x: rect.left - pane.originX,
      y: rect.top - pane.originY,
      width: rect.width,
      height: rect.height,
      primary: primary.get(drawing.id) === pane.id,
      selected: drawing.id === selectedId,
    })),
  }));
}

/** A rectangle placed in a pane's own px (the draw gesture's preview). */
export function inPane(rect: RectPx, pane: DrawingPane): RectPx {
  return { left: rect.left - pane.originX, top: rect.top - pane.originY, width: rect.width, height: rect.height };
}

// --- handles and resizing ---------------------------------------------------------------

export type DrawingHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

/** The eight handles, clockwise from the top-left corner. */
export const DRAWING_HANDLES: readonly DrawingHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

export function isCornerHandle(handle: DrawingHandle): boolean {
  return handle.length === 2;
}

/** Whether a resize from `handle` keeps the aspect: images and groups do on
 * a corner unless Shift is held, shapes (and the rest) only with Shift. */
export function resizeKeepsAspect(content: SheetDrawingContent, handle: DrawingHandle, shift: boolean): boolean {
  if (!isCornerHandle(handle)) {
    return false;
  }
  const byDefault = content.kind === "image" || content.kind === "group";
  return byDefault ? !shift : shift;
}

/** `start` resized from `handle` by (dx, dy) sheet px (see the module header). */
export function resizedRect(start: RectPx, handle: DrawingHandle, dx: number, dy: number, keepAspect: boolean): RectPx {
  const west = handle.includes("w");
  const east = handle.includes("e");
  const north = handle.startsWith("n");
  const south = handle.startsWith("s");
  const minW = start.width > 0 ? MIN_DRAWING_SIZE_PX : 0;
  const minH = start.height > 0 ? MIN_DRAWING_SIZE_PX : 0;
  const right = start.left + start.width;
  const bottom = start.top + start.height;
  // West and north edges stop at the grid's top-left.
  let width = west ? Math.min(right, start.width - dx) : east ? start.width + dx : start.width;
  let height = north ? Math.min(bottom, start.height - dy) : south ? start.height + dy : start.height;
  width = Math.max(width, minW);
  height = Math.max(height, minH);
  if (keepAspect && start.width > 0 && start.height > 0) {
    const sw = width / start.width;
    const sh = height / start.height;
    let scale = Math.abs(sw - 1) >= Math.abs(sh - 1) ? sw : sh;
    scale = Math.max(scale, minW / start.width, minH / start.height);
    if (west) {
      scale = Math.min(scale, right / start.width);
    }
    if (north) {
      scale = Math.min(scale, bottom / start.height);
    }
    width = start.width * scale;
    height = start.height * scale;
  }
  return { left: west ? right - width : start.left, top: north ? bottom - height : start.top, width, height };
}

// --- what a drawing paints ---------------------------------------------------------------

/** A drawing's paint: an image, a shape (with its text), a group of member
 * paints, or the labelled dashed placeholder. */
export type DrawingPaint =
  | { kind: "image"; src: string; crop?: { l: number; t: number; r: number; b: number } }
  | { kind: "shape"; visual: ShapeVisual; text?: DrawingText }
  | { kind: "group"; members: { box: { x: number; y: number; w: number; h: number }; content: SheetDrawingContent }[] }
  | { kind: "placeholder"; label: string };

/** What `content` paints. `imageSrc` answers a media part's data URI, or ""
 * when the part is missing or of a type the browser cannot show (EMF, WMF,
 * TIFF) — that image paints the "Picture" placeholder. A chart (SPEC-132
 * widens it) and a face-less foreign drawing paint their placeholder; a
 * foreign drawing with a face paints the face. */
export function drawingPaint(content: SheetDrawingContent, imageSrc: (mediaId: string) => string): DrawingPaint {
  switch (content.kind) {
    case "image": {
      const src = imageSrc(content.mediaId);
      return src ? { kind: "image", src, ...(content.crop ? { crop: content.crop } : {}) } : { kind: "placeholder", label: "Picture" };
    }
    case "shape":
      return { kind: "shape", visual: content.visual, ...(content.text ? { text: content.text } : {}) };
    case "group":
      return { kind: "group", members: content.members.map(({ content: member, ...box }) => ({ box, content: member })) };
    case "chart":
      return { kind: "placeholder", label: "Chart" };
    case "foreign":
      return content.face ? drawingPaint(content.face, imageSrc) : { kind: "placeholder", label: content.label };
  }
}

/** The `<img>` box that shows a crop (fractions cut off each side) stretched
 * over a `width` × `height` drawing, inside a box that clips it. */
export function croppedImageBox(crop: { l: number; t: number; r: number; b: number } | undefined, width: number, height: number): RectPx {
  const keepX = 1 - (crop?.l ?? 0) - (crop?.r ?? 0);
  const keepY = 1 - (crop?.t ?? 0) - (crop?.b ?? 0);
  const fullW = keepX > 0 ? width / keepX : width;
  const fullH = keepY > 0 ? height / keepY : height;
  return { left: keepX > 0 ? -(crop?.l ?? 0) * fullW : 0, top: keepY > 0 ? -(crop?.t ?? 0) * fullH : 0, width: fullW, height: fullH };
}

/** A shape's text as the layer lays it over the canvas: the vertical
 * placement, the first run's look (px at 100% zoom; 11 pt in the theme's
 * dark text colour by default) and the paragraphs. */
export function drawingTextView(text: DrawingText, theme: DocumentTheme = DEFAULT_THEME): {
  justify: "flex-start" | "center" | "flex-end";
  style: Record<string, string>;
  paragraphs: { text: string; align: "left" | "center" | "right" }[];
} {
  const style: Record<string, string> = {
    "font-size": `${((text.sizePt ?? 11) * 96) / 72}px`,
    color: cellColorHex(text.color ?? { theme: "dk1" }, theme),
  };
  if (text.bold) {
    style["font-weight"] = "bold";
  }
  return {
    justify: text.anchor === "top" ? "flex-start" : text.anchor === "bottom" ? "flex-end" : "center",
    style,
    paragraphs: text.paragraphs.map((paragraph) => ({ text: paragraph.text, align: paragraph.align ?? "left" })),
  };
}

/** A painted drawing's role: an image is an img, everything else a group. */
export function drawingRole(drawing: SheetDrawing): "img" | "group" {
  return drawing.content.kind === "image" ? "img" : "group";
}

/** A painted drawing's accessible name: its alt text, else its name. */
export function drawingLabel(drawing: SheetDrawing): string {
  return drawing.altText?.trim() ? drawing.altText : drawing.name;
}

/** The CSS rotation a drawing renders with (whole degrees clockwise about
 * its centre): the drawing's own, else its shape's. */
export function drawingRotation(drawing: SheetDrawing): number {
  const own = drawing.rotationDeg ?? (drawing.content.kind === "shape" ? drawing.content.visual.rotationDeg : undefined);
  return own ?? 0;
}

// --- prop narrowing (the .vue parser bars casts; see AppShell.vue) ----------------------

export function asPaneView(value: unknown): PaneView {
  return value as PaneView;
}

export function asDrawingContent(value: unknown): SheetDrawingContent {
  return value as SheetDrawingContent;
}

export function asDocumentTheme(value: unknown): DocumentTheme {
  return (value as DocumentTheme | null) ?? DEFAULT_THEME;
}

export function asImageSrc(value: unknown): (mediaId: string) => string {
  return typeof value === "function" ? (value as (mediaId: string) => string) : () => "";
}

export function asRectPx(value: unknown): RectPx | null {
  return (value as RectPx | null) ?? null;
}

export function asDrawingHandle(value: unknown): DrawingHandle {
  return DRAWING_HANDLES.includes(value as DrawingHandle) ? (value as DrawingHandle) : "se";
}
