// Pure alignment and distribution of laid-out drawing rectangles (SPEC-072),
// framework- and layout-free like the rest of src/lib. The gesture/command
// layer lays drawings out into page rectangles (SPEC-057's established seam)
// and passes them here; these functions never compute layout, never touch the
// model and never read screen pixels — they take rectangles in centimetres and
// return new top-left positions in centimetres, in input order.

/** The six edges/axes a group of drawings can be aligned on. */
export type AlignEdge =
  | "left"
  | "centerH"
  | "right"
  | "top"
  | "middleV"
  | "bottom";

/** A laid-out drawing rectangle in centimetres: top-left (x, y) and size.
 *  A rotated drawing contributes its unrotated frame rectangle (SPEC-057). */
export interface AlignRect {
  xCm: number;
  yCm: number;
  wCm: number;
  hCm: number;
}

/** Aligns every rectangle to one edge/centre of the `reference` rectangle,
 * returning each rectangle's new top-left in input order. Only the coordinate
 * on the aligned axis changes; the other is passed through unchanged. Each new
 * coordinate is a single assigned reference value (or reference ± a size), so
 * results are exact — no accumulation. The reference is chosen by the caller:
 * the selection's bounding-box union (multi-selection), the page content area,
 * page, column, or the container box (interior selection). Because the caller
 * passes the union, the outermost drawing on the aligned edge is already at the
 * reference and does not move. */
export function alignRects(
  rects: AlignRect[],
  edge: AlignEdge,
  reference: AlignRect,
): Array<{ xCm: number; yCm: number }> {
  return rects.map((rect) => {
    switch (edge) {
      case "left":
        return { xCm: reference.xCm, yCm: rect.yCm };
      case "centerH":
        return {
          xCm: reference.xCm + (reference.wCm - rect.wCm) / 2,
          yCm: rect.yCm,
        };
      case "right":
        return { xCm: reference.xCm + reference.wCm - rect.wCm, yCm: rect.yCm };
      case "top":
        return { xCm: rect.xCm, yCm: reference.yCm };
      case "middleV":
        return {
          xCm: rect.xCm,
          yCm: reference.yCm + (reference.hCm - rect.hCm) / 2,
        };
      case "bottom":
        return {
          xCm: rect.xCm,
          yCm: reference.yCm + reference.hCm - rect.hCm,
        };
    }
  });
}

/** Equalises the gaps between rectangles along one axis, leaving the outermost
 * two in place, and returns each rectangle's new top-left in input order. The
 * gap is the space between one rectangle's trailing edge and the next's leading
 * edge — not the distance between centres; the two differ whenever the
 * rectangles are not the same size, and gap-equalisation is what looks right.
 * Fewer than three rectangles are returned unchanged. */
export function distributeRects(
  rects: AlignRect[],
  axis: "h" | "v",
): Array<{ xCm: number; yCm: number }> {
  const result = rects.map((r) => ({ xCm: r.xCm, yCm: r.yCm }));
  if (rects.length < 3) return result;

  const pos = (r: AlignRect) => (axis === "h" ? r.xCm : r.yCm);
  const size = (r: AlignRect) => (axis === "h" ? r.wCm : r.hCm);

  // Order by leading edge; the first and last in that order are the outermost
  // two and stay put. We reposition only the interior rectangles.
  const order = rects.map((_, i) => i).toSorted((a, b) => pos(rects[a]) - pos(rects[b]));
  const first = rects[order[0]];
  const last = rects[order[order.length - 1]];

  const start = pos(first);
  const span = pos(last) + size(last) - start;
  let totalSize = 0;
  for (const r of rects) totalSize += size(r);
  const gap = (span - totalSize) / (rects.length - 1);

  let cursor = start;
  for (const i of order) {
    if (axis === "h") result[i].xCm = cursor;
    else result[i].yCm = cursor;
    cursor += size(rects[i]) + gap;
  }
  return result;
}
