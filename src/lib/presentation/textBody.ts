// Text-body derivations for slide frames (SPEC-097): the effective insets and
// anchor a frame renders with, the deterministic autofit ladder, and the
// commands that edit `TextFrame.body`/`ShapeFrame.body`. Framework-free and
// DOM-free — the MEASURING half of autofit lives in
// src/components/presentation/slideTextFit.ts (a rendering concern), which
// hands `fitScale` its `measureHeightCm` callback.
//
// `TextBodyProps` is declared in ./model (beside SlideTransition/
// SlideAnimation) and only re-exported here: this module's derivations take
// `TextFrame`/`ShapeFrame`, so declaring the type downstream would invert the
// layering into a cycle (animations.ts:13 is the precedent).
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { ShapeFrame, SlideFrame, TextAnchor, TextAutofit, TextBodyProps, TextFrame } from "./model";
import { MIN_FRAME_SIZE_CM } from "./frames";

export type { TextAnchor, TextAutofit, TextBodyProps } from "./model";

/** A frame that carries a text body. */
export type TextBodyFrame = TextFrame | ShapeFrame;

/** The pptx `a:bodyPr` default insets (91440/45720/91440/45720 EMU), in cm. */
export const DEFAULT_INSETS_CM: Readonly<{ left: number; top: number; right: number; bottom: number }> = {
  left: 0.254,
  top: 0.127,
  right: 0.254,
  bottom: 0.127,
};

/** The column-count range (pptx `numCol`). */
export const MAX_TEXT_COLUMNS = 16;

/** Whether a frame carries a text body (text frames and shapes). */
export function isTextBodyFrame(frame: SlideFrame | null | undefined): frame is TextBodyFrame {
  return !!frame && (frame.kind === "text" || frame.kind === "shape");
}

/** The insets a frame renders with: its own, else the pptx defaults. */
export function effectiveInsetsCm(frame: TextBodyFrame): { left: number; top: number; right: number; bottom: number } {
  return { ...DEFAULT_INSETS_CM, ...(frame.body?.insetsCm ?? {}) };
}

/** The vertical anchor: stored, else "top" for a text frame and "middle" for
 * a shape (SPEC-079's renderer assumption, now a stored default). */
export function effectiveAnchor(frame: TextBodyFrame): TextAnchor {
  return frame.body?.anchor ?? (frame.kind === "shape" ? "middle" : "top");
}

/** The autofit mode: stored, else "none". */
export function effectiveAutofit(frame: TextBodyFrame): TextAutofit {
  return frame.body?.autofit ?? "none";
}

/** The column layout: stored, else one column. */
export function effectiveColumns(frame: TextBodyFrame): { count: number; gapCm: number } {
  return frame.body?.columns ?? { count: 1, gapCm: 0 };
}

/** The deterministic shrink ladder (font scale %), largest first. */
export const AUTOFIT_STEPS = [100, 92.5, 85, 77.5, 70, 65, 60, 55, 50, 45, 40, 35, 30, 25] as const;

/** The line-spacing scale paired with a font scale on the ladder: 80 at
 * ≤ 50, 90 at ≤ 70, else 100. */
export function lineScaleFor(fontScalePct: number): number {
  if (fontScalePct <= 50) return 80;
  if (fontScalePct <= 70) return 90;
  return 100;
}

/** Walks the ladder against a measurer: the first step whose measured content
 * height fits `availableCm` (the frame's inset-reduced height), else the last
 * step. Pure — the DOM measurer is injected (slideTextFit.ts). */
export function fitScale(
  measureHeightCm: (fontScalePct: number, lineScalePct: number) => number,
  availableCm: number,
): { fontScalePct: number; lineScalePct: number } {
  for (const step of AUTOFIT_STEPS) {
    const line = lineScaleFor(step);
    if (measureHeightCm(step, line) <= availableCm + 1e-6) {
      return { fontScalePct: step, lineScalePct: line };
    }
  }
  const last = AUTOFIT_STEPS[AUTOFIT_STEPS.length - 1];
  return { fontScalePct: last, lineScalePct: lineScaleFor(last) };
}

/** The height available to text inside a frame (its height less the top and
 * bottom insets), never negative. */
export function availableTextHeightCm(frame: TextBodyFrame): number {
  const insets = effectiveInsetsCm(frame);
  return Math.max(0, frame.heightCm - insets.top - insets.bottom);
}

const r3 = (value: number): number => Math.round(value * 1000) / 1000;

/** "resizeShape": the frame height that holds `contentHeightCm` plus the
 * insets, floored at MIN_FRAME_SIZE_CM and rounded r3. */
export function resizedHeightCm(frame: TextBodyFrame, contentHeightCm: number): number {
  const insets = effectiveInsetsCm(frame);
  return r3(Math.max(MIN_FRAME_SIZE_CM, contentHeightCm + insets.top + insets.bottom));
}

/** The stored scale a renderer starts from: the frame's own (shrink only),
 * else 100/100. */
export function storedScale(frame: TextBodyFrame): { fontScalePct: number; lineScalePct: number } {
  if (effectiveAutofit(frame) !== "shrink") return { fontScalePct: 100, lineScalePct: 100 };
  return { fontScalePct: frame.body?.fontScalePct ?? 100, lineScalePct: frame.body?.lineScalePct ?? 100 };
}

/** pptx stores `fontScale`/`lnSpcReduction` in 1000ths of a percent — the one
 * sanctioned round-trip delta (SPEC-097): a scale quantised to that grid. */
export function quantizeScalePct(pct: number): number {
  return Math.round(pct * 1000) / 1000;
}

// --- commands ----------------------------------------------------------------

/** A text-body edit; `null` clears a field back to its default. */
export interface TextBodyPatch {
  autofit?: TextAutofit | null;
  anchor?: TextAnchor | null;
  insetsCm?: { left: number; top: number; right: number; bottom: number } | null;
  columns?: { count: number; gapCm: number } | null;
  wrap?: false | null;
  /** Only meaningful with autofit "shrink" (dropped otherwise). */
  fontScalePct?: number | null;
  lineScalePct?: number | null;
}

const clamp = (v: number, lo: number, hi: number): number => Math.min(Math.max(v, lo), hi);

/** Applies a body patch to a frame, normalising to the model's minimal form
 * (defaults dropped, an empty `body` removed). Editing autofit or anchor
 * clears the frame's preserved `pptxBodyPr`/`odpBodyProps` — the source's
 * `vert`/`upright`/`anchorCtr`/unrecognised autofit shadow exactly those
 * (Table.docxTblPr's rule); insets/columns/wrap shadow nothing preserved. */
export function applyTextBody(frame: TextBodyFrame, patch: TextBodyPatch): void {
  const body: TextBodyProps = { ...(frame.body ?? {}) };
  if (patch.autofit !== undefined || patch.anchor !== undefined) {
    delete frame.pptxBodyPr;
    delete frame.odpBodyProps;
  }
  if (patch.autofit !== undefined) {
    if (patch.autofit === null || patch.autofit === "none") delete body.autofit;
    else body.autofit = patch.autofit;
  }
  if (patch.anchor !== undefined) {
    const fallback = frame.kind === "shape" ? "middle" : "top";
    if (patch.anchor === null || patch.anchor === fallback) delete body.anchor;
    else body.anchor = patch.anchor;
  }
  if (patch.insetsCm !== undefined) {
    if (patch.insetsCm === null) delete body.insetsCm;
    else {
      const next = {
        left: r3(clamp(patch.insetsCm.left, 0, 50)),
        top: r3(clamp(patch.insetsCm.top, 0, 50)),
        right: r3(clamp(patch.insetsCm.right, 0, 50)),
        bottom: r3(clamp(patch.insetsCm.bottom, 0, 50)),
      };
      const isDefault =
        next.left === DEFAULT_INSETS_CM.left &&
        next.top === DEFAULT_INSETS_CM.top &&
        next.right === DEFAULT_INSETS_CM.right &&
        next.bottom === DEFAULT_INSETS_CM.bottom;
      if (isDefault) delete body.insetsCm;
      else body.insetsCm = next;
    }
  }
  if (patch.columns !== undefined) {
    if (patch.columns === null) delete body.columns;
    else {
      const count = Math.round(clamp(patch.columns.count, 1, MAX_TEXT_COLUMNS));
      const gapCm = r3(clamp(patch.columns.gapCm, 0, 50));
      if (count === 1 && gapCm === 0) delete body.columns;
      else body.columns = { count, gapCm };
    }
  }
  if (patch.wrap !== undefined) {
    if (patch.wrap === false) body.wrap = false;
    else delete body.wrap;
  }
  if (patch.fontScalePct !== undefined) {
    if (patch.fontScalePct === null) delete body.fontScalePct;
    else body.fontScalePct = quantizeScalePct(clamp(patch.fontScalePct, 10, 100));
  }
  if (patch.lineScalePct !== undefined) {
    if (patch.lineScalePct === null) delete body.lineScalePct;
    else body.lineScalePct = quantizeScalePct(clamp(patch.lineScalePct, 80, 100));
  }
  if (body.autofit !== "shrink") {
    delete body.fontScalePct;
    delete body.lineScalePct;
  }
  if (Object.keys(body).length === 0) delete frame.body;
  else frame.body = body;
}
