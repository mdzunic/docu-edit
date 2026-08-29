// Alignment guides and snapping while a drawing is dragged, resized or created
// (SPEC-072). This is the gesture layer the pure src/lib/drawingSnap seam feeds:
// it measures the snap targets' laid-out rectangles from the DOM (other drawings
// on the page, the content-area margins and centre lines, and the page edges and
// centre), converts the 8-screen-pixel threshold to centimetres through the live
// zoom — the conversion that belongs here, not in src/lib (AC-54) — calls the
// pure matcher, and holds the transient guides for the overlay to render. Guides
// never touch the model; they clear on release (AC-55/56).
import { ref, type Ref } from "vue";
import { CM_PER_PX, PX_PER_CM } from "../../lib/core/units";
import { usePreferences } from "../../composables/usePreferences";
import {
  computeDrawingSnap,
  computeResizeSnap,
  type SnapCandidate,
  type SnapGuide,
} from "../../lib/drawingSnap";
import type { AlignRect } from "../../lib/drawingAlign";

export interface DrawingGuidesDeps {
  /** The `.pages` box — the frame every guide coordinate is measured in. */
  pageRef: Ref<Element | null>;
  /** The contenteditable canvas root, for querying drawing elements. */
  contentRoot: Ref<Element | null>;
  /** Maps a client point to unscaled content px relative to `.pages`. */
  clientToContentPx: (clientX: number, clientY: number) => { x: number; y: number };
  /** The live zoom fraction (SPEC-063), for the screen-pixel threshold. */
  zoom: Readonly<Ref<number>>;
}

const SNAP_THRESHOLD_PX = 8;

/** The vertical/horizontal candidate lines a laid-out rectangle contributes. */
function rectCandidates(x0: number, y0: number, x1: number, y1: number): SnapCandidate[] {
  return [
    { axis: "x", coordCm: x0, startCm: y0, endCm: y1 },
    { axis: "x", coordCm: (x0 + x1) / 2, startCm: y0, endCm: y1 },
    { axis: "x", coordCm: x1, startCm: y0, endCm: y1 },
    { axis: "y", coordCm: y0, startCm: x0, endCm: x1 },
    { axis: "y", coordCm: (y0 + y1) / 2, startCm: x0, endCm: x1 },
    { axis: "y", coordCm: y1, startCm: x0, endCm: x1 },
  ];
}

export function useDrawingGuides(deps: DrawingGuidesDeps) {
  const { pageRef, contentRoot, clientToContentPx, zoom } = deps;
  const { preferences } = usePreferences();

  /** The guides to render, in unscaled `.pages`-relative centimetres. Empty
   * except during an active snap. */
  const snapGuides = ref<SnapGuide[]>([]);

  /** An element's box in `.pages`-relative centimetres. */
  function boxCm(el: Element): { x0: number; y0: number; x1: number; y1: number } {
    const box = el.getBoundingClientRect();
    const topLeft = clientToContentPx(box.left, box.top);
    const bottomRight = clientToContentPx(box.right, box.bottom);
    return {
      x0: topLeft.x * CM_PER_PX,
      y0: topLeft.y * CM_PER_PX,
      x1: bottomRight.x * CM_PER_PX,
      y1: bottomRight.y * CM_PER_PX,
    };
  }

  /** Every other drawing's six alignment lines (AC-18). */
  function drawingCandidates(excludeEl: Element | null): SnapCandidate[] {
    const root = contentRoot.value;
    if (!root) return [];
    const candidates: SnapCandidate[] = [];
    for (const el of root.querySelectorAll(".doc-shape")) {
      if (el === excludeEl) continue;
      const { x0, y0, x1, y1 } = boxCm(el);
      candidates.push(...rectCandidates(x0, y0, x1, y1));
    }
    return candidates;
  }

  /** The `.page-content` holding the dragged drawing (or the first page). */
  function pageOf(excludeEl: Element | null): Element | null {
    return excludeEl?.closest(".page-content") ?? pageRef.value?.querySelector(".page-content") ?? null;
  }

  /** The content area's margins and centre lines, plus the page edges and centre
   * (AC-19/20), for the page holding the dragged drawing. */
  function frameCandidates(excludeEl: Element | null): SnapCandidate[] {
    const page = pageOf(excludeEl);
    if (!page) return [];
    const candidates: SnapCandidate[] = [];
    const pageBox = boxCm(page);
    candidates.push(...rectCandidates(pageBox.x0, pageBox.y0, pageBox.x1, pageBox.y1));
    const margin = page.querySelector(".page-margin");
    if (margin) {
      const m = boxCm(margin);
      candidates.push(...rectCandidates(m.x0, m.y0, m.x1, m.y1));
    }
    return candidates;
  }

  /** Grid lines within the content area (AC-33), origin at its top-left. */
  function gridSnapCandidates(excludeEl: Element | null): SnapCandidate[] {
    const step = preferences.gridSpacingCm;
    const page = pageOf(excludeEl);
    const margin = page?.querySelector(".page-margin");
    if (!margin || step <= 0) return [];
    const m = boxCm(margin);
    const candidates: SnapCandidate[] = [];
    for (let x = m.x0; x <= m.x1 + 1e-9; x += step) {
      candidates.push({ axis: "x", coordCm: x, startCm: m.y0, endCm: m.y1 });
    }
    for (let y = m.y0; y <= m.y1 + 1e-9; y += step) {
      candidates.push({ axis: "y", coordCm: y, startCm: m.x0, endCm: m.x1 });
    }
    return candidates;
  }

  /** Snaps a laid-out drawing rectangle (in `.pages` cm) and records the guides
   * for the overlay. Returns the snapped top-left. Alt suppresses snapping
   * entirely (AC-25); so does having both snap preferences off. */
  function applySnap(rect: AlignRect, excludeEl: Element | null, altKey: boolean): { xCm: number; yCm: number } {
    if (altKey || (!preferences.snapToGuides && !preferences.snapToGrid)) {
      clearGuides();
      return { xCm: rect.xCm, yCm: rect.yCm };
    }
    const candidates: SnapCandidate[] = [];
    if (preferences.snapToGuides) {
      candidates.push(...drawingCandidates(excludeEl), ...frameCandidates(excludeEl));
    }
    if (preferences.snapToGrid) {
      candidates.push(...gridSnapCandidates(excludeEl));
    }
    const thresholdCm = (SNAP_THRESHOLD_PX / Math.max(zoom.value, 0.01)) * CM_PER_PX;
    const result = computeDrawingSnap(rect, candidates, thresholdCm);
    snapGuides.value = result.guides;
    return { xCm: result.xCm, yCm: result.yCm };
  }

  /** Snaps a drawing being resized (AC-27): the top-left of `rect` is the fixed
   * offset corner and `wCm`/`hCm` the size the pointer has produced so far. Only
   * the moving right/bottom edges snap; guides are recorded for the overlay just
   * as a move does. Alt or both snap preferences off suppress it (AC-25). */
  function applyResizeSnap(
    rect: AlignRect,
    excludeEl: Element | null,
    altKey: boolean,
  ): { widthCm: number; heightCm: number } {
    if (altKey || (!preferences.snapToGuides && !preferences.snapToGrid)) {
      clearGuides();
      return { widthCm: rect.wCm, heightCm: rect.hCm };
    }
    const candidates: SnapCandidate[] = [];
    if (preferences.snapToGuides) {
      candidates.push(...drawingCandidates(excludeEl), ...frameCandidates(excludeEl));
    }
    if (preferences.snapToGrid) {
      candidates.push(...gridSnapCandidates(excludeEl));
    }
    const thresholdCm = (SNAP_THRESHOLD_PX / Math.max(zoom.value, 0.01)) * CM_PER_PX;
    const result = computeResizeSnap(rect, candidates, thresholdCm);
    snapGuides.value = result.guides;
    return { widthCm: result.widthCm, heightCm: result.heightCm };
  }

  function clearGuides(): void {
    if (snapGuides.value.length > 0) {
      snapGuides.value = [];
    }
  }

  /** A guide in on-screen (unscaled `.pages`) pixels, for the overlay. */
  function guidePx(guide: SnapGuide): { left: number; top: number; length: number; vertical: boolean } {
    const vertical = guide.axis === "x";
    return {
      left: guide.coordCm * PX_PER_CM,
      top: guide.startCm * PX_PER_CM,
      length: (guide.endCm - guide.startCm) * PX_PER_CM,
      vertical,
    };
  }

  return { snapGuides, applySnap, applyResizeSnap, clearGuides, guidePx };
}
