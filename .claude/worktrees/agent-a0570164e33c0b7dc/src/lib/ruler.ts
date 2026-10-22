// Geometry for the horizontal ruler above the page (SPEC-014 AC-22, AC-23,
// AC-25): maps the current paragraph's logical indents (plus the page
// margins) onto physical marker positions, and a dragged marker's physical
// position back onto the logical value it sets. Framework-free so the
// mapping — especially the RTL mirroring — is testable without mounting the
// component.
//
// All positions are centimetres measured from the page's PHYSICAL LEFT
// edge, matching CSS `left` on the ruler's markers. The paragraph's
// direction decides which physical edge the logical start/end indents hang
// off: for RTL the origin is the right margin and start grows leftward.
import type { ParagraphFormatting } from "./documentModel";

export interface RulerContext {
  pageWidthCm: number;
  marginLeftCm: number;
  marginRightCm: number;
}

/** Physical marker positions, in cm from the page's left edge. */
export interface RulerMarkerPositions {
  start: number;
  firstLine: number;
  end: number;
  marginLeft: number;
  marginRight: number;
}

export type RulerMarkerKind = "start" | "first-line" | "end" | "margin-left" | "margin-right";

/** The model field a marker drag updates, with its new (logical) value. */
export interface RulerUpdate {
  kind: "indentStartCm" | "indentEndCm" | "indentFirstLineCm" | "marginLeftCm" | "marginRightCm";
  valueCm: number;
}

/** Drag positions snap to 0.05 cm — fine enough to feel continuous, coarse
 * enough to hit round values; the command layer re-snaps onto the exact
 * cross-format grid before the model stores anything. */
export const RULER_DRAG_STEP_CM = 0.05;

function snapStep(cm: number): number {
  return Math.round(cm / RULER_DRAG_STEP_CM) * RULER_DRAG_STEP_CM;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

export function markerPositions(
  context: RulerContext,
  formatting: ParagraphFormatting,
): RulerMarkerPositions {
  const rtl = formatting.direction === "rtl";
  const start = formatting.indentStartCm ?? 0;
  const end = formatting.indentEndCm ?? 0;
  const firstLine = formatting.indentFirstLineCm ?? 0;
  const startPos = rtl
    ? context.pageWidthCm - context.marginRightCm - start
    : context.marginLeftCm + start;
  const endPos = rtl
    ? context.marginLeftCm + end
    : context.pageWidthCm - context.marginRightCm - end;
  return {
    start: startPos,
    // First-line is start-relative; it grows in the text direction.
    firstLine: rtl ? startPos - firstLine : startPos + firstLine,
    end: endPos,
    marginLeft: context.marginLeftCm,
    marginRight: context.pageWidthCm - context.marginRightCm,
  };
}

/** The logical value a marker dragged to `positionCm` (cm from the page's
 * left edge) implies. Indents stay non-negative and inside the printable
 * width; the first-line indent may go negative (a hanging indent) as far as
 * the near page edge. Margins stay within the page, leaving room for the
 * opposite margin. */
export function valueFromMarkerPosition(
  kind: RulerMarkerKind,
  positionCm: number,
  context: RulerContext,
  formatting: ParagraphFormatting,
): RulerUpdate {
  const rtl = formatting.direction === "rtl";
  const printableCm = Math.max(0, context.pageWidthCm - context.marginLeftCm - context.marginRightCm);
  const position = clamp(positionCm, 0, context.pageWidthCm);
  switch (kind) {
    case "start": {
      const raw = rtl
        ? context.pageWidthCm - context.marginRightCm - position
        : position - context.marginLeftCm;
      return { kind: "indentStartCm", valueCm: clamp(snapStep(raw), 0, printableCm) };
    }
    case "end": {
      const raw = rtl
        ? position - context.marginLeftCm
        : context.pageWidthCm - context.marginRightCm - position;
      return { kind: "indentEndCm", valueCm: clamp(snapStep(raw), 0, printableCm) };
    }
    case "first-line": {
      const start = formatting.indentStartCm ?? 0;
      const startPos = rtl
        ? context.pageWidthCm - context.marginRightCm - start
        : context.marginLeftCm + start;
      const raw = rtl ? startPos - position : position - startPos;
      const nearMargin = rtl ? context.marginRightCm : context.marginLeftCm;
      return {
        kind: "indentFirstLineCm",
        valueCm: clamp(snapStep(raw), -(start + nearMargin), printableCm - start),
      };
    }
    case "margin-left":
      return {
        kind: "marginLeftCm",
        valueCm: clamp(snapStep(position), 0, Math.max(0, context.pageWidthCm - context.marginRightCm)),
      };
    case "margin-right":
      return {
        kind: "marginRightCm",
        valueCm: clamp(
          snapStep(context.pageWidthCm - position),
          0,
          Math.max(0, context.pageWidthCm - context.marginLeftCm),
        ),
      };
  }
}
