// Align and distribute for a drawing selection (SPEC-072). The pure geometry
// lives in src/lib/drawingAlign.ts (alignRects/distributeRects); this composable
// is the gesture/command layer that seam expects — it reads each selected
// drawing's laid-out rectangle (its offset + size, the unrotated frame the
// formats store), computes the reference rectangle, calls the pure operation,
// and writes the result back through one arrangeDrawings command so the
// snapshot history records exactly one labelled undo entry (AC-10). Undo
// restores model and selection for free: these commands only change offsets, so
// every drawing keeps its block position and the shape selection stays valid
// (AC-11).
//
// A multi-selection (two or more) aligns to its own bounding box and
// distributes its gaps (AC-1, AC-5, AC-6); a single drawing aligns to the page
// content area, switchable to the whole page or a column, remembered for the
// life of the view (AC-2, AC-3, AC-4).
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { CM_PER_PX } from "../../lib/core/units";
import {
  documentStart,
  caretSelection,
  type ModelSelection,
} from "../../lib/commands";
import {
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  type DocumentModel,
} from "../../lib/documentModel";
import {
  arrangeDrawings,
  shapeAt,
  type DrawingPositionUpdate,
  type ShapeLocation,
} from "../../lib/shapes";
import {
  alignRects,
  distributeRects,
  type AlignEdge,
  type AlignRect,
} from "../../lib/drawingAlign";
import type { EditorContext } from "./context";

/** Which reference a single-drawing align uses (AC-3). Session-scoped state,
 * never persisted (AC-4). */
export type AlignToReference = "margin" | "page" | "column";

export function asAlignToReference(value: string): AlignToReference {
  return value as AlignToReference;
}

export interface DrawingArrangeDeps extends EditorContext {
  /** The primary selected drawing, single-selection align + the position
   * dialog's target. */
  shapeSelection: Ref<ShapeLocation | null>;
  /** Shift-click multi-selection: 2+ entries drive bounding-box align and
   * distribute. */
  shapeMultiSelection: Ref<ShapeLocation[]>;
}

function rectOf(shape: { offsetXCm?: number; offsetYCm?: number; widthCm: number; heightCm: number }): AlignRect {
  return {
    xCm: shape.offsetXCm ?? 0,
    yCm: shape.offsetYCm ?? 0,
    wCm: shape.widthCm,
    hCm: shape.heightCm,
  };
}

/** The union of laid-out rectangles — the multi-selection's bounding box. */
function unionRect(rects: AlignRect[]): AlignRect {
  let minX = Infinity;
  let minY = Infinity;
  let maxR = -Infinity;
  let maxB = -Infinity;
  for (const r of rects) {
    minX = Math.min(minX, r.xCm);
    minY = Math.min(minY, r.yCm);
    maxR = Math.max(maxR, r.xCm + r.wCm);
    maxB = Math.max(maxB, r.yCm + r.hCm);
  }
  return { xCm: minX, yCm: minY, wCm: maxR - minX, hCm: maxB - minY };
}

/** The width of the drawing's column (SPEC-030). Uses the first column's width,
 * which is the whole story for the equal-width default; unequal-width columns
 * fall back to the first, since which column an anchor sits in is a layout fact
 * src/lib does not compute. */
function columnWidthCm(model: DocumentModel, section: number): number {
  const columns = sectionColumns(model, section);
  return resolveColumnWidthsCm(columns, sectionContentWidthCm(model, section))[0];
}

export interface DrawingArrangeApi {
  commands: Array<[string, () => void]>;
  alignToReference: Ref<AlignToReference>;
  setAlignToReference: (reference: AlignToReference) => void;
  /** True when the selected drawing's section has more than one column, so the
   * Column reference is offered (AC-3). */
  columnReferenceAvailable: ComputedRef<boolean>;
  /** "" when align is available, else a short reason for the disabled state
   * (AC-17). */
  alignDisabledReason: ComputedRef<string>;
  /** "" when distribute is available, else the disabled reason (AC-17). */
  distributeDisabledReason: ComputedRef<string>;
  /** Every Format → Drawing prop the menu bar reads, bundled so DocumentView
   * binds them with one v-bind (line-budget discipline). */
  menuBindings: ComputedRef<Record<string, unknown>>;
  /** Flip the persisted Snap to guides / Snap to grid preferences (AC-30/31). */
  toggleSnapToGuides: () => void;
  toggleSnapToGrid: () => void;
}

function locatedShapes(model: DocumentModel, locations: ShapeLocation[]) {
  return locations
    .map((location) => ({ location, shape: shapeAt(model, location) }))
    .filter((entry): entry is { location: ShapeLocation; shape: NonNullable<typeof entry.shape> } => entry.shape !== null);
}

export function useDrawingArrange(deps: DrawingArrangeDeps): DrawingArrangeApi {
  const { doc, shapeSelection, shapeMultiSelection, editable } = deps;

  const alignToReference = ref<AlignToReference>("margin");
  function setAlignToReference(reference: AlignToReference): void {
    alignToReference.value = reference;
  }

  const columnReferenceAvailable = computed<boolean>(() => {
    const model = doc.value;
    const location = shapeSelection.value;
    if (!model || !location) {
      return false;
    }
    return sectionColumns(model, location.section).count > 1;
  });

  const alignDisabledReason = computed<string>(() => {
    if (!editable.value) {
      return "Document is read-only";
    }
    if (shapeMultiSelection.value.length >= 2 || shapeSelection.value) {
      return "";
    }
    return "Select a drawing to align";
  });

  const distributeDisabledReason = computed<string>(() => {
    if (!editable.value) {
      return "Document is read-only";
    }
    if (shapeMultiSelection.value.length >= 3) {
      return "";
    }
    return "Select three or more drawings to distribute";
  });

  function beforeSelection(): ModelSelection {
    const root = deps.contentRoot.value as HTMLElement | null;
    return deps.currentSelection(root) ?? caretSelection(deps.getCaret() ?? documentStart());
  }

  function alignMulti(model: DocumentModel, locations: ShapeLocation[], edge: AlignEdge): DrawingPositionUpdate[] {
    const shapes = locatedShapes(model, locations);
    const rects = shapes.map((entry) => rectOf(entry.shape));
    const positions = alignRects(rects, edge, unionRect(rects));
    return shapes.map((entry, index) => ({
      location: entry.location,
      offsetXCm: positions[index].xCm,
      offsetYCm: positions[index].yCm,
    }));
  }

  /** An element's top-left in `.pages`-relative centimetres, through the live
   * zoom (SPEC-063's clientToContentPx). */
  function pageTopLeftCm(el: Element): { xCm: number; yCm: number } {
    const box = el.getBoundingClientRect();
    const point = deps.clientToContentPx(box.left, box.top);
    return { xCm: point.x * CM_PER_PX, yCm: point.y * CM_PER_PX };
  }

  /** An element's rectangle in `.pages`-relative centimetres. */
  function pageRectCm(el: Element): AlignRect {
    const box = el.getBoundingClientRect();
    const topLeft = deps.clientToContentPx(box.left, box.top);
    const bottomRight = deps.clientToContentPx(box.right, box.bottom);
    return {
      xCm: topLeft.x * CM_PER_PX,
      yCm: topLeft.y * CM_PER_PX,
      wCm: (bottomRight.x - topLeft.x) * CM_PER_PX,
      hCm: (bottomRight.y - topLeft.y) * CM_PER_PX,
    };
  }

  /** The single-drawing align reference rectangle, in `.pages`-relative cm,
   * measured from the page holding the drawing: the content area inside the
   * margins by default (AC-2), the whole page, or the first column (AC-3). It
   * must be measured, not computed from the model, because the offset written
   * back is anchor-relative (the renderer positions a drawing at its offset from
   * its anchor paragraph, not from the page origin) — so the target has to be
   * expressed in the same laid-out page space the drawing's own rectangle is. */
  function singleReferenceRect(model: DocumentModel, anchorEl: Element, section: number): AlignRect | null {
    const page =
      anchorEl.closest(".page-content") ??
      (deps.pageRef.value as Element | null)?.querySelector(".page-content") ??
      null;
    if (!page) {
      return null;
    }
    const reference = alignToReference.value;
    if (reference === "page") {
      return pageRectCm(page);
    }
    const marginEl = page.querySelector(".page-margin");
    if (!marginEl) {
      return null;
    }
    const content = pageRectCm(marginEl);
    // A column reference narrows the content area to the first column's width
    // (its left edge and height stay the content area's) — which column an
    // anchor sits in is a layout fact src/lib does not compute (AC-3).
    return reference === "column" ? { ...content, wCm: columnWidthCm(model, section) } : content;
  }

  function alignSingle(model: DocumentModel, location: ShapeLocation, edge: AlignEdge): DrawingPositionUpdate[] {
    const shape = shapeAt(model, location);
    const root = deps.contentRoot.value as Element | null;
    if (!shape || !root) {
      return [];
    }
    // The anchor span's top-left is the origin the drawing's offset measures
    // from (DocShapeBlock.vue) — reading it from the DOM lets a single-drawing
    // align land the drawing on the real page content-area / page / column edge,
    // independent of where its anchor paragraph sits down the page.
    const anchorEl = root.querySelector(
      `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"]`,
    );
    if (!anchorEl) {
      return [];
    }
    const reference = singleReferenceRect(model, anchorEl, location.section);
    if (!reference) {
      return [];
    }
    const origin = pageTopLeftCm(anchorEl);
    const offsetXCm = shape.offsetXCm ?? 0;
    const offsetYCm = shape.offsetYCm ?? 0;
    const drawingRect: AlignRect = {
      xCm: origin.xCm + offsetXCm,
      yCm: origin.yCm + offsetYCm,
      wCm: shape.widthCm,
      hCm: shape.heightCm,
    };
    const [target] = alignRects([drawingRect], edge, reference);
    // Only the aligned axis changes; the offset is the target page position back
    // in anchor space (target − anchor origin). No relativeToH/relativeToV is
    // stamped: the offset is anchor-relative, so re-labelling its reference edge
    // would misdescribe it (the renderer and the docx writer both read the
    // offset as measured from the anchor).
    const horizontal = edge === "left" || edge === "centerH" || edge === "right";
    return horizontal
      ? [{ location, offsetXCm: offsetXCm + (target.xCm - drawingRect.xCm) }]
      : [{ location, offsetYCm: offsetYCm + (target.yCm - drawingRect.yCm) }];
  }

  function align(edge: AlignEdge, label: string): void {
    const model = doc.value;
    if (!model || alignDisabledReason.value) {
      return;
    }
    const multi = shapeMultiSelection.value;
    const before = beforeSelection();
    const updates =
      multi.length >= 2
        ? alignMulti(model, multi, edge)
        : shapeSelection.value
          ? alignSingle(model, shapeSelection.value, edge)
          : [];
    if (updates.length === 0) {
      return;
    }
    deps.applyEdit(arrangeDrawings(model, before, updates), before, label);
  }

  function distribute(axis: "h" | "v", label: string): void {
    const model = doc.value;
    if (!model || distributeDisabledReason.value) {
      return;
    }
    const shapes = locatedShapes(model, shapeMultiSelection.value);
    if (shapes.length < 3) {
      return;
    }
    const before = beforeSelection();
    const positions = distributeRects(shapes.map((entry) => rectOf(entry.shape)), axis);
    const updates: DrawingPositionUpdate[] = shapes.map((entry, index) => ({
      location: entry.location,
      offsetXCm: positions[index].xCm,
      offsetYCm: positions[index].yCm,
    }));
    deps.applyEdit(arrangeDrawings(model, before, updates), before, label);
  }

  // Explicit literal tuples (not generated) so the commandCatalogue
  // completeness spec can read every id from source, as it does the other
  // spread-in command lists.
  const commands: Array<[string, () => void]> = [
    ["drawing.align-left", () => align("left", "align left")],
    ["drawing.align-center", () => align("centerH", "align centre")],
    ["drawing.align-right", () => align("right", "align right")],
    ["drawing.align-top", () => align("top", "align top")],
    ["drawing.align-middle", () => align("middleV", "align middle")],
    ["drawing.align-bottom", () => align("bottom", "align bottom")],
    ["drawing.distribute-horizontal", () => distribute("h", "distribute horizontally")],
    ["drawing.distribute-vertical", () => distribute("v", "distribute vertically")],
  ];

  // Snap-to-guides / snap-to-grid live in preferences (SPEC-066); the menu
  // toggles read and flip them here (AC-30/31/32).
  const { preferences, setPreference } = usePreferences();
  function toggleSnapToGuides(): void {
    setPreference("snapToGuides", !preferences.snapToGuides);
  }
  function toggleSnapToGrid(): void {
    setPreference("snapToGrid", !preferences.snapToGrid);
  }

  const menuBindings = computed<Record<string, unknown>>(() => ({
    drawingAlignReason: alignDisabledReason.value,
    drawingDistributeReason: distributeDisabledReason.value,
    alignToReference: alignToReference.value,
    columnReferenceAvailable: columnReferenceAvailable.value,
    snapToGuides: preferences.snapToGuides,
    snapToGrid: preferences.snapToGrid,
  }));

  return {
    commands,
    alignToReference,
    setAlignToReference,
    columnReferenceAvailable,
    alignDisabledReason,
    distributeDisabledReason,
    menuBindings,
    toggleSnapToGuides,
    toggleSnapToGrid,
  };
}
