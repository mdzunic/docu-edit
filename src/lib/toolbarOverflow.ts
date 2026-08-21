// SPEC-061 overflow planning — the pure geometry behind the toolbar's
// collapse-by-band behaviour, kept framework-free so it is unit-testable
// without a DOM. The strip groups its controls into eight fixed bands
// (History, Style, Type, Character, Colour, Paragraph, Lists, Paragraph
// extras). When the measured band widths do not fit the available strip
// width, whole bands collapse into a trailing "More formatting" menu, and
// the strip never scrolls (AC-12, AC-13).

/** The order bands collapse in: highest band number first. Bands 1–4 never
 * collapse — history, style, type and character formatting justify permanent
 * space (SPEC-049's own test for what stays in the strip). */
export const COLLAPSE_ORDER = [8, 7, 6, 5];

export interface PlanCollapseParams {
  /** Measured natural width of each band, in px, keyed by band number. Taken
   * once while every band is expanded and cached; a missing band counts as 0. */
  bandWidths: Record<number, number>;
  /** The band numbers present in the strip, in left-to-right order. */
  bandOrder: number[];
  /** Flex gap between adjacent bands, in px. */
  gap: number;
  /** Content width available to the strip, in px (clientWidth minus any
   * reserved space for open panes). */
  available: number;
  /** Width the trailing overflow button occupies once any band collapses,
   * including the gap before it, in px. */
  overflowWidth: number;
  /** How many bands are collapsed right now — used for hysteresis so the strip
   * does not flap one band in and out at the exact boundary. */
  previousCollapsed: number;
  /** Extra slack (px) required before a collapsed band is re-expanded. */
  hysteresis: number;
}

/** Width the strip needs to show all but `collapsedCount` of the collapsible
 * bands, plus the overflow button when anything is collapsed. */
function requiredWidth(
  collapsedCount: number,
  { bandWidths, bandOrder, gap, overflowWidth }: PlanCollapseParams,
): number {
  const collapsed = new Set(COLLAPSE_ORDER.slice(0, collapsedCount));
  const visible = bandOrder.filter((band) => !collapsed.has(band));
  let width = visible.reduce((sum, band) => sum + (bandWidths[band] ?? 0), 0);
  if (visible.length > 0) {
    width += gap * (visible.length - 1);
  }
  if (collapsedCount > 0) {
    width += overflowWidth;
  }
  return width;
}

/** How many bands (from the tail of COLLAPSE_ORDER) should collapse so the
 * strip fits `available`. Returns 0..COLLAPSE_ORDER.length. Chooses the
 * fewest collapses that fit; when that means re-expanding a band relative to
 * the current state, it holds unless there is `hysteresis` px of extra slack. */
export function planCollapse(params: PlanCollapseParams): number {
  const max = COLLAPSE_ORDER.length;
  let fit = max;
  for (let count = 0; count <= max; count++) {
    if (requiredWidth(count, params) <= params.available) {
      fit = count;
      break;
    }
  }
  if (fit < params.previousCollapsed) {
    // Expanding a band: only do it with slack to spare, else hold the current
    // count, stepping down only as far as a count that fits with margin.
    let hold = params.previousCollapsed;
    for (let count = fit; count < params.previousCollapsed; count++) {
      if (requiredWidth(count, params) + params.hysteresis <= params.available) {
        hold = count;
        break;
      }
    }
    fit = hold;
  }
  return fit;
}

/** The set of band numbers collapsed for a given collapse count. */
export function collapsedBands(collapsedCount: number): Set<number> {
  return new Set(COLLAPSE_ORDER.slice(0, collapsedCount));
}
