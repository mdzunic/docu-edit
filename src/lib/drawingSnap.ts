// Pure snapping candidate-matching for drawings (SPEC-072), framework- and
// layout-free like drawingAlign.ts. The gesture layer lays the dragged drawing
// and the snap targets out into page rectangles, converts the screen-pixel
// threshold to centimetres through the live zoom (that conversion belongs in
// the gesture layer, not here — AC-54), and calls this. It never reads the
// model, never touches the DOM and never affects the model — guides are a
// transient render artefact the caller draws and discards on release.
import type { AlignRect } from "./drawingAlign";

export type SnapAxis = "x" | "y";

/** One alignment line to test the dragged drawing against: a vertical (axis
 * "x") or horizontal (axis "y") line at `coordCm`. `startCm`/`endCm` are the
 * target's extent along the perpendicular axis, so a matched guide can be drawn
 * spanning from the dragged drawing to the target (AC-21). */
export interface SnapCandidate {
  axis: SnapAxis;
  coordCm: number;
  startCm: number;
  endCm: number;
}

/** A guide line to render through a matched feature (AC-21, AC-22). */
export interface SnapGuide {
  axis: SnapAxis;
  coordCm: number;
  startCm: number;
  endCm: number;
}

/** Typed pass-through for the guides array a `.vue` template holds untyped
 * (same espree idiom as connectors' asConnectorRenders). */
export function asSnapGuides(value: unknown[]): SnapGuide[] {
  return value as SnapGuide[];
}

export interface SnapResult {
  /** The snapped top-left, equal to the input when nothing matched. */
  xCm: number;
  yCm: number;
  /** The guides for every candidate the snapped drawing lines up with. */
  guides: SnapGuide[];
}

const EPS = 1e-9;

/** Snaps a laid-out drawing rectangle onto the nearest candidate alignment
 * within `thresholdCm` and returns the snapped top-left plus the guides to
 * draw. The drawing's own six features are tested against each candidate: its
 * left / horizontal-centre / right against vertical candidates, its top /
 * vertical-centre / bottom against horizontal candidates. The closest match on
 * each axis sets that axis's snap; both axes can snap at once, and every
 * candidate sharing a matched coordinate contributes a guide (multiple
 * simultaneous matches => multiple guides, AC-22). */
export function computeDrawingSnap(
  rect: AlignRect,
  candidates: SnapCandidate[],
  thresholdCm: number,
): SnapResult {
  const featuresX = [rect.xCm, rect.xCm + rect.wCm / 2, rect.xCm + rect.wCm];
  const featuresY = [rect.yCm, rect.yCm + rect.hCm / 2, rect.yCm + rect.hCm];

  function bestMatch(axis: SnapAxis, features: number[]): { delta: number; coordCm: number } | null {
    let best: { delta: number; coordCm: number } | null = null;
    for (const candidate of candidates) {
      if (candidate.axis !== axis) continue;
      for (const feature of features) {
        const delta = candidate.coordCm - feature;
        if (Math.abs(delta) <= thresholdCm && (best === null || Math.abs(delta) < Math.abs(best.delta))) {
          best = { delta, coordCm: candidate.coordCm };
        }
      }
    }
    return best;
  }

  const matchX = bestMatch("x", featuresX);
  const matchY = bestMatch("y", featuresY);
  const xCm = rect.xCm + (matchX ? matchX.delta : 0);
  const yCm = rect.yCm + (matchY ? matchY.delta : 0);

  const guides: SnapGuide[] = [];
  function collect(axis: SnapAxis, matchedCoord: number | null, perpStart: number, perpEnd: number): void {
    if (matchedCoord === null) return;
    for (const candidate of candidates) {
      if (candidate.axis === axis && Math.abs(candidate.coordCm - matchedCoord) <= EPS) {
        guides.push({
          axis,
          coordCm: candidate.coordCm,
          startCm: Math.min(perpStart, candidate.startCm),
          endCm: Math.max(perpEnd, candidate.endCm),
        });
      }
    }
  }
  // A vertical (x) guide spans the dragged drawing's vertical extent; a
  // horizontal (y) guide spans its horizontal extent — both taken at the
  // snapped position so the line passes through the aligned edge.
  collect("x", matchX ? matchX.coordCm : null, yCm, yCm + rect.hCm);
  collect("y", matchY ? matchY.coordCm : null, xCm, xCm + rect.wCm);

  return { xCm, yCm, guides };
}

/** The vertical grid candidate coordinates covering [0, extentCm] at `spacing`,
 * as snap candidates spanning the given perpendicular range. Snap-to-grid uses
 * the nearest of these (SPEC-072); the gesture layer decides whether grid
 * snapping is on. */
export function gridCandidates(
  axis: SnapAxis,
  spacingCm: number,
  extentCm: number,
  perpStartCm: number,
  perpEndCm: number,
): SnapCandidate[] {
  const result: SnapCandidate[] = [];
  if (spacingCm <= 0) return result;
  for (let coord = 0; coord <= extentCm + EPS; coord += spacingCm) {
    result.push({ axis, coordCm: coord, startCm: perpStartCm, endCm: perpEndCm });
  }
  return result;
}
