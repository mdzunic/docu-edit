// Pure cm-rect algebra for align / distribute / snap (SPEC-096) — the one
// product-neutral home SPEC-072 consumes for the text side too. Cm rectangles
// in, cm rectangles out: no Vue, no product types, core siblings only
// (boundary.spec.ts). The presentation stage drives it through
// src/views/presentation/stageMath.ts; nothing here touches a model.

/** An axis-aligned rectangle in centimetres. */
export interface Rect {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

export type AlignEdge = "left" | "centerX" | "right" | "top" | "middleY" | "bottom";

/** What the aligned edge lines up against: the selection's own bounding box,
 * the first rect in the list, or the container (the slide). */
export type AlignReference = "selection" | "first" | "container";

/** The union bounding box of `rects`, or null for an empty list. */
export function boundingRect(rects: readonly Rect[]): Rect | null {
  if (rects.length === 0) {
    return null;
  }
  const minX = Math.min(...rects.map((r) => r.xCm));
  const minY = Math.min(...rects.map((r) => r.yCm));
  const maxX = Math.max(...rects.map((r) => r.xCm + r.widthCm));
  const maxY = Math.max(...rects.map((r) => r.yCm + r.heightCm));
  return { xCm: minX, yCm: minY, widthCm: maxX - minX, heightCm: maxY - minY };
}

/** The reference box an edge aligns against. */
function referenceBox(rects: readonly Rect[], reference: AlignReference, container: Rect): Rect {
  if (reference === "container") {
    return container;
  }
  if (reference === "first") {
    return rects[0] ?? container;
  }
  return boundingRect(rects) ?? container;
}

/** Aligns every rect's `edge` to the reference box, returning new rects in
 * the input order with only the aligned axis moved. */
export function alignRects(
  rects: readonly Rect[],
  edge: AlignEdge,
  reference: AlignReference,
  container: Rect,
): Rect[] {
  const box = referenceBox(rects, reference, container);
  return rects.map((r) => {
    const out = { ...r };
    switch (edge) {
      case "left":
        out.xCm = box.xCm;
        break;
      case "centerX":
        out.xCm = box.xCm + box.widthCm / 2 - r.widthCm / 2;
        break;
      case "right":
        out.xCm = box.xCm + box.widthCm - r.widthCm;
        break;
      case "top":
        out.yCm = box.yCm;
        break;
      case "middleY":
        out.yCm = box.yCm + box.heightCm / 2 - r.heightCm / 2;
        break;
      case "bottom":
        out.yCm = box.yCm + box.heightCm - r.heightCm;
        break;
    }
    return out;
  });
}

/** Distributes ≥3 rects along `axis`: sorts by the axis' centre, holds the
 * outermost two in place, and equalises the *gaps* between neighbours (not
 * the centres — ragged sizes still end up evenly spaced). Fewer than three
 * rects come back unchanged. Returned in the input order. */
export function distributeRects(rects: readonly Rect[], axis: "x" | "y"): Rect[] {
  if (rects.length < 3) {
    return rects.map((r) => ({ ...r }));
  }
  const pos = axis === "x" ? (r: Rect) => r.xCm : (r: Rect) => r.yCm;
  const size = axis === "x" ? (r: Rect) => r.widthCm : (r: Rect) => r.heightCm;
  const order = rects
    .map((r, index) => ({ r, index }))
    .toSorted((a, b) => pos(a.r) + size(a.r) / 2 - (pos(b.r) + size(b.r) / 2));
  const first = order[0].r;
  const last = order[order.length - 1].r;
  const span = pos(last) + size(last) - pos(first);
  const total = order.reduce((sum, o) => sum + size(o.r), 0);
  const gap = (span - total) / (order.length - 1);
  const out = rects.map((r) => ({ ...r }));
  let cursor = pos(first) + size(first) + gap;
  for (const o of order.slice(1, -1)) {
    if (axis === "x") {
      out[o.index].xCm = cursor;
    } else {
      out[o.index].yCm = cursor;
    }
    cursor += size(o.r) + gap;
  }
  return out;
}

/** One line a dragged rect may snap to. `kind` is what the line is (an edge
 * or a centreline), used for matching: a centre candidate attracts only the
 * rect's own centre, an edge candidate either of its edges. */
export interface GuideCandidate {
  axis: "x" | "y";
  line: number;
  kind: "edge" | "centre";
}

/** The snap candidates for a drag: the container's centrelines and edges
 * FIRST (the priority tier — snapRect fires the earliest container hit
 * before any frame line), then every other rect's edges and centres. */
export function guideCandidates(others: readonly Rect[], container: Rect): GuideCandidate[] {
  const out: GuideCandidate[] = [
    { axis: "x", line: container.xCm + container.widthCm / 2, kind: "centre" },
    { axis: "y", line: container.yCm + container.heightCm / 2, kind: "centre" },
    { axis: "x", line: container.xCm, kind: "edge" },
    { axis: "x", line: container.xCm + container.widthCm, kind: "edge" },
    { axis: "y", line: container.yCm, kind: "edge" },
    { axis: "y", line: container.yCm + container.heightCm, kind: "edge" },
  ];
  for (const o of others) {
    out.push(
      { axis: "x", line: o.xCm, kind: "edge" },
      { axis: "x", line: o.xCm + o.widthCm, kind: "edge" },
      { axis: "x", line: o.xCm + o.widthCm / 2, kind: "centre" },
      { axis: "y", line: o.yCm, kind: "edge" },
      { axis: "y", line: o.yCm + o.heightCm, kind: "edge" },
      { axis: "y", line: o.yCm + o.heightCm / 2, kind: "centre" },
    );
  }
  return out;
}

/** How many leading candidates per axis form the container tier that
 * `guideCandidates` emits (centre, near edge, far edge). */
const CONTAINER_TIER = 3;

interface AxisHit {
  candidate: GuideCandidate;
  shift: number;
  dist: number;
  tierIndex: number;
}

function axisHits(
  rect: Rect,
  candidates: readonly GuideCandidate[],
  axis: "x" | "y",
  thresholdCm: number,
): AxisHit[] {
  const start = axis === "x" ? rect.xCm : rect.yCm;
  const extent = axis === "x" ? rect.widthCm : rect.heightCm;
  const hits: AxisHit[] = [];
  let tierIndex = 0;
  for (const candidate of candidates) {
    if (candidate.axis !== axis) {
      continue;
    }
    const lines = candidate.kind === "centre" ? [start + extent / 2] : [start, start + extent];
    let best: { shift: number; dist: number } | null = null;
    for (const line of lines) {
      const shift = candidate.line - line;
      const dist = Math.abs(shift);
      if (dist <= thresholdCm && (!best || dist < best.dist)) {
        best = { shift, dist };
      }
    }
    if (best) {
      hits.push({ candidate, shift: best.shift, dist: best.dist, tierIndex });
    }
    tierIndex += 1;
  }
  return hits;
}

/** Snaps `rect` against the candidates, firing at most one per axis: the
 * first container-tier hit (centre before edges — `guideCandidates`' order),
 * else the nearest frame line. Returns the snapped rect and the fired
 * candidates. */
export function snapRect(
  rect: Rect,
  candidates: readonly GuideCandidate[],
  thresholdCm: number,
): { rect: Rect; fired: GuideCandidate[] } {
  const out = { ...rect };
  const fired: GuideCandidate[] = [];
  for (const axis of ["x", "y"] as const) {
    const hits = axisHits(rect, candidates, axis, thresholdCm);
    const container = hits.find((h) => h.tierIndex < CONTAINER_TIER);
    const frame = hits
      .filter((h) => h.tierIndex >= CONTAINER_TIER)
      .toSorted((a, b) => a.dist - b.dist)[0];
    const chosen = container ?? frame;
    if (!chosen) {
      continue;
    }
    if (axis === "x") {
      out.xCm += chosen.shift;
    } else {
      out.yCm += chosen.shift;
    }
    fired.push(chosen.candidate);
  }
  return { rect: out, fired };
}
