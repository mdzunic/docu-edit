// The DOM half of slide text autofit (SPEC-097): measuring a text body's
// content height at a given font/line scale, and the inline style a text
// body renders with (insets, anchor, columns, the fit scale). Kept beside
// slideTextDom.ts under the same header rule — a rendering concern, NOT under
// src/lib/presentation: src/lib/presentation/textBody.ts#fitScale walks the
// deterministic ladder against the `measureBodyHeightCm` callback this module
// supplies, and stays DOM-free and unit-testable with a fake measurer.
//
// The scale travels as two CSS custom properties on the text-body element —
// `--fit-scale` (font) and `--line-scale` (line spacing) — which the run and
// paragraph styles in slideRender.ts multiply through (`calc(Npx *
// var(--fit-scale, 1))`). So one rendered body re-scales without
// re-rendering its runs, and an off-flow clone measures any step of the
// ladder by changing two properties.
//
// Measuring NEVER writes the model (SPEC-097 "measuring never dirties a
// slide"): the stage re-renders at a measured scale locally; only a command
// that has already called markSlideEdited (useSlideTextDepth.ts) stores one.
import { cmPx } from "./slideRender";
import { pxToCm } from "../../lib/core/units";
import type { ShapeFrame, TextFrame } from "../../lib/presentation/model";
import {
  effectiveAnchor,
  effectiveColumns,
  effectiveInsetsCm,
  fitScale,
  availableTextHeightCm,
  effectiveAutofit,
} from "../../lib/presentation/textBody";

/** A text-carrying frame. */
type BodyFrame = TextFrame | ShapeFrame;

/** A font/line scale pair, in percent. */
export interface FitScale {
  fontScalePct: number;
  lineScalePct: number;
}

/** The unscaled pair. */
export const FULL_SCALE: FitScale = { fontScalePct: 100, lineScalePct: 100 };

/** The class the renderer gives a frame's text-body element, and the data
 * attribute carrying the frame id on the frame root — the hooks
 * `frameTextHost` finds a rendered body by. A frame in a text session renders
 * its editing body (TEXT_EDIT_BODY_CLASS) instead. */
export const TEXT_BODY_CLASS = "text-body";
export const TEXT_EDIT_BODY_CLASS = "text-edit-body";
export const FRAME_ID_ATTRIBUTE = "data-frame-id";

const JUSTIFY_FOR_ANCHOR = { top: "flex-start", middle: "center", bottom: "flex-end" } as const;

/** The inline style of a frame's text-body element: insets as padding, the
 * anchor as flex `justify-content`, columns, wrap, and the fit scale as the
 * two custom properties. Pure (no DOM) — SlideFrameView binds it, and
 * `applyFitStyle` assigns it to a clone. */
export function bodyStyle(frame: BodyFrame, scale: FitScale = FULL_SCALE): Record<string, string> {
  const insets = effectiveInsetsCm(frame);
  const columns = effectiveColumns(frame);
  const style: Record<string, string> = {
    "--fit-scale": String(scale.fontScalePct / 100),
    "--line-scale": String(scale.lineScalePct / 100),
    padding: `${cmPx(insets.top)} ${cmPx(insets.right)} ${cmPx(insets.bottom)} ${cmPx(insets.left)}`,
    "justify-content": JUSTIFY_FOR_ANCHOR[effectiveAnchor(frame)],
  };
  if (columns.count > 1) {
    style["column-count"] = String(columns.count);
    style["column-gap"] = cmPx(columns.gapCm);
    // A multi-column body flows as one block, so the flex column that anchors
    // a single column becomes block flow (anchor still applies via padding).
    style.display = "block";
  }
  if (frame.body?.wrap === false) {
    style["white-space"] = "pre";
  }
  return style;
}

/** Assigns `bodyStyle` to an element (the measuring clone, or a body the
 * stage re-scales without a Vue re-render). */
export function applyFitStyle(el: HTMLElement, frame: BodyFrame, scale: FitScale = FULL_SCALE): void {
  for (const [key, value] of Object.entries(bodyStyle(frame, scale))) {
    el.style.setProperty(key, value);
  }
}

/** The rendered text-body element of a frame on the stage (`root` defaults
 * to the document; `bodyClass` picks the editing body), or null when it is
 * not rendered. */
export function frameTextHost(frameId: string, root: ParentNode = document, bodyClass = TEXT_BODY_CLASS): HTMLElement | null {
  // Matched by attribute value rather than an escaped selector: ids are
  // arbitrary strings, and `CSS.escape` is absent in jsdom.
  const frameEl = Array.from(root.querySelectorAll(`[${FRAME_ID_ATTRIBUTE}]`)).find(
    (el) => el.getAttribute(FRAME_ID_ATTRIBUTE) === frameId,
  );
  const host = frameEl?.querySelector(`.${bodyClass}`);
  return host instanceof HTMLElement ? host : null;
}

/** Measures the content height (cm) of a text body at a scale: an off-flow
 * clone of `host` — inserted beside it so it inherits the stage's own font
 * stack and scoped styles — laid out at `widthCm` with the scale applied and
 * no insets, height auto. Layout px (offsetHeight), so a CSS-transformed
 * stage measures the same as an unscaled one. */
export function measureBodyHeightCm(host: HTMLElement, widthCm: number, fontScalePct: number, lineScalePct: number): number {
  const clone = host.cloneNode(true) as HTMLElement;
  clone.style.setProperty("--fit-scale", String(fontScalePct / 100));
  clone.style.setProperty("--line-scale", String(lineScalePct / 100));
  clone.style.position = "absolute";
  clone.style.visibility = "hidden";
  clone.style.pointerEvents = "none";
  clone.style.left = "0";
  clone.style.top = "0";
  // A body that is itself absolutely inset (shape text: `inset: 0`) would
  // otherwise stretch to its frame — release the far edges so height is auto.
  clone.style.right = "auto";
  clone.style.bottom = "auto";
  clone.style.width = cmPx(widthCm);
  clone.style.height = "auto";
  clone.style.minHeight = "0";
  clone.style.maxHeight = "none";
  clone.style.padding = "0";
  clone.style.overflow = "visible";
  clone.setAttribute("aria-hidden", "true");
  const parent = host.parentElement ?? host.ownerDocument.body;
  parent.appendChild(clone);
  try {
    return pxToCm(clone.offsetHeight);
  } finally {
    clone.remove();
  }
}

/** The text width (cm) inside a frame: its width less the side insets. */
export function textWidthCm(frame: BodyFrame): number {
  const insets = effectiveInsetsCm(frame);
  return Math.max(0, frame.widthCm - insets.left - insets.right);
}

/** What a rendered body measures to: the shrink scale the ladder settles on
 * (100/100 unless autofit is "shrink") and the natural content height at
 * full scale (for "resizeShape"). Null when the body is not rendered or the
 * host cannot lay out (jsdom: every height measures 0, so the ladder would
 * trivially "fit" — callers treat null as "keep what is stored"). */
export function measureFrameFit(frame: BodyFrame, host: HTMLElement | null): { scale: FitScale; contentHeightCm: number } | null {
  if (!host) return null;
  const width = textWidthCm(frame);
  const contentHeightCm = measureBodyHeightCm(host, width, 100, 100);
  if (contentHeightCm <= 0) return null;
  if (effectiveAutofit(frame) !== "shrink") return { scale: FULL_SCALE, contentHeightCm };
  const scale = fitScale((font, line) => measureBodyHeightCm(host, width, font, line), availableTextHeightCm(frame));
  return { scale, contentHeightCm };
}

/** What a frame's EDITING body measures to, during a text session. Its runs
 * are plain px at full scale with the stage's own paragraph styles, so it
 * lays out "resizeShape"'s input (the full-scale content height) but no step
 * of the shrink ladder: null for any other mode, or where nothing lays out —
 * the session's end re-measures the rendered body. */
export function measureEditingFit(frame: BodyFrame, editBody: HTMLElement | null): { scale: FitScale; contentHeightCm: number } | null {
  if (!editBody || effectiveAutofit(frame) !== "resizeShape") return null;
  const contentHeightCm = measureBodyHeightCm(editBody, textWidthCm(frame), 100, 100);
  return contentHeightCm > 0 ? { scale: FULL_SCALE, contentHeightCm } : null;
}
