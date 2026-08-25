// Pure render helpers for the slide components (SPEC-079). The `.vue` scripts
// here follow this repo's ESLint-espree convention — runtime `defineProps`,
// no TS type syntax (generics, annotations, casts, type imports) — so every
// typed transformation (prop coercion, per-frame style computation, frame
// narrowing) lives in this `.ts` module and the templates call into it. Fully
// framework-agnostic: plain data in, inline-style records / booleans out.
import { cmToPx } from "../../lib/core/units";
import { mediaDataUri, type MediaPart } from "../../lib/core/media";
import type { Slide, SlideFrame, SlideParagraph, SlideRun } from "../../lib/presentation/model";
import { slideOutline } from "../../lib/presentation/slideModel";
import type { ShapeVisual } from "../../lib/core/shapes";

/** pt → px at 96/72, the ratio the text editor uses. */
const PT_TO_PX = 96 / 72;
const LEVEL_INDENT_CM = 0.95;
const DEFAULT_FONT_SIZE_PT = 18;

/** A centimetre length as a px string (the model's own cm → CSS px). */
export function cmPx(cm: number): string {
  return `${cmToPx(cm)}px`;
}

// --- prop coercion (the `as*` pattern, like documentModel's asShape) --------
export function asSlide(value: unknown): Slide {
  return value as Slide;
}
export function asSlideFrame(value: unknown): SlideFrame {
  return value as SlideFrame;
}
export function asMediaParts(value: unknown): MediaPart[] {
  return (value ?? []) as MediaPart[];
}
export function asSlides(value: unknown): Slide[] {
  return (value ?? []) as Slide[];
}
export function asResolver(value: unknown): (mediaId: string) => string {
  return value as (mediaId: string) => string;
}
export function asNumber(value: unknown): number {
  return Number(value) || 0;
}
export function asString(value: unknown): string {
  return value === undefined || value === null ? "" : String(value);
}

// --- media resolution -------------------------------------------------------
/** A resolver from `mediaId` to a `data:` URI over a media-part list. */
export function mediaResolver(media: MediaPart[]): (mediaId: string) => string {
  const map = new Map(media.map((part) => [part.mediaId, part]));
  return (mediaId) => mediaDataUri(map.get(mediaId));
}

// --- frame narrowing (templates cannot narrow a discriminated union) --------
export function frameKind(frame: SlideFrame): SlideFrame["kind"] {
  return frame.kind;
}
export function isForeignVisible(frame: SlideFrame): boolean {
  return frame.kind === "foreign" && frame.noBounds !== true;
}
export function imageSrc(frame: SlideFrame, resolveMedia: (mediaId: string) => string): string {
  return frame.kind === "image" ? resolveMedia(frame.mediaId) : "";
}
export function foreignLabel(frame: SlideFrame): string {
  return frame.kind === "foreign" ? frame.label : "";
}
export function groupChildren(frame: SlideFrame): SlideFrame[] {
  return frame.kind === "group" ? frame.children : [];
}

// --- positioning & fills ----------------------------------------------------
export function framePositionStyle(frame: SlideFrame): Record<string, string> {
  const style: Record<string, string> = {
    position: "absolute",
    left: cmPx(frame.xCm),
    top: cmPx(frame.yCm),
    width: cmPx(frame.widthCm),
    height: cmPx(frame.heightCm),
  };
  if (frame.rotationDeg) {
    style.transform = `rotate(${frame.rotationDeg}deg)`;
    style.transformOrigin = "center";
  }
  return style;
}

/** A text frame's own background colour, or "transparent". */
export function textFrameFill(frame: SlideFrame): string {
  if (frame.kind === "text" && frame.fill && frame.fill.type === "solid") {
    return frame.fill.color;
  }
  return "transparent";
}

/** The slide surface background — always light (white unless the deck says
 * otherwise). */
export function surfaceBackground(slide: Slide): string {
  const bg = slide.background;
  return bg && bg.type === "solid" ? bg.color : "#ffffff";
}

// --- text ------------------------------------------------------------------
/** A text/shape frame's paragraphs annotated with their bullet marker
 * (numbered bullets restart per contiguous same-level run). */
export function markedParagraphs(frame: SlideFrame): { para: SlideParagraph; marker: string }[] {
  const paragraphs = frame.kind === "text" ? frame.paragraphs : frame.kind === "shape" ? frame.paragraphs ?? [] : [];
  const counts: Record<number, number> = {};
  let prev: SlideParagraph | undefined;
  return paragraphs.map((para) => {
    let marker = "";
    if (para.bullet === "number") {
      const level = para.level ?? 0;
      if (prev && prev.bullet === "number" && (prev.level ?? 0) === level) {
        counts[level] = (counts[level] ?? 0) + 1;
      } else {
        counts[level] = 1;
      }
      marker = `${counts[level]}.`;
    } else if (para.bullet === "disc") {
      marker = "•";
    }
    prev = para;
    return { para, marker };
  });
}

export function paragraphStyle(para: SlideParagraph): Record<string, string> {
  return {
    textAlign: para.align ?? "left",
    marginLeft: para.level ? cmPx(para.level * LEVEL_INDENT_CM) : "0",
  };
}

export function runStyle(run: SlideRun): Record<string, string> {
  const style: Record<string, string> = {};
  if (run.bold) style.fontWeight = "700";
  if (run.italic) style.fontStyle = "italic";
  if (run.underline) style.textDecoration = "underline";
  if (run.fontFamily) style.fontFamily = `'${run.fontFamily}', sans-serif`;
  style.fontSize = `${(run.sizePt ?? DEFAULT_FONT_SIZE_PT) * PT_TO_PX}px`;
  style.color = run.color ?? "#000000";
  return style;
}

// --- shapes -----------------------------------------------------------------
/** A shape frame as a ShapeVisual for `paintShapeIntoCanvas`, or null for a
 * non-shape. Rotation is left to the frame's CSS transform (AC-36), so the
 * canvas paints unrotated. */
export function shapeVisualOf(frame: SlideFrame): ShapeVisual | null {
  if (frame.kind !== "shape") {
    return null;
  }
  const visual: ShapeVisual = { geometry: frame.geometry };
  if (frame.fill && frame.fill.type === "solid") {
    visual.fill = frame.fill.color;
  }
  if (frame.stroke) {
    visual.outline = { color: frame.stroke.color, widthPt: frame.stroke.widthPt, dash: "solid" };
  }
  if (frame.adjustments && frame.adjustments.length > 0) {
    visual.adjustments = frame.adjustments;
  }
  return visual;
}

// --- filmstrip --------------------------------------------------------------
export function filmstripLabel(slide: Slide, index: number, total: number): string {
  const outline = slideOutline(slide);
  const name = outline.title && outline.title.trim().length > 0 ? outline.title : "untitled";
  return `Slide ${index + 1} of ${total} — ${name}`;
}

export function filmstripThumbScale(widthCm: number, thumbWidthPx: number): number {
  return thumbWidthPx / Math.max(1, cmToPx(widthCm));
}

// --- filmstrip menu / drag DOM helpers (SPEC-080) --------------------------
// The .vue script is espree-linted (no TS casts/generics), so the typed DOM
// queries the context menu needs live here.

/** Focuses the first element matching `selector` inside `container`. */
export function focusSelector(container: Element | null | undefined, selector: string): void {
  const el = container?.querySelector(selector);
  if (el instanceof HTMLElement) {
    el.focus();
  }
}

/** The focusable menu-item buttons inside a rendered menu, in DOM order. */
export function menuButtons(container: Element | null | undefined): HTMLElement[] {
  const nodes = container ? Array.from(container.querySelectorAll('[role="menuitem"]')) : [];
  return nodes.filter((node): node is HTMLElement => node instanceof HTMLElement);
}

/** A viewport anchor point just inside an element (for a keyboard-opened
 * menu), or the origin when it can't be found. */
export function anchorPoint(container: Element | null | undefined, selector: string): { x: number; y: number } {
  const el = container?.querySelector(selector);
  if (el instanceof HTMLElement) {
    const rect = el.getBoundingClientRect();
    return { x: rect.left + 24, y: rect.bottom };
  }
  return { x: 0, y: 0 };
}
