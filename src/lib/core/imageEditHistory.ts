// The image editor's history list (SPEC-112): the edits of a session in the
// order the user made them — a record of intent — of which the stack
// (normaliseStack in imageEdits.ts) is the canonical program. Pure helpers
// over immutable arrays: each returns a NEW history and never mutates its
// input, so an undo snapshot is just the array it held.
//
// A history holds at most one entry per single-valued kind (straighten,
// crop, rotate, resize, adjust; flip once per axis) — refining a crop or
// turning again updates that entry in place, so the list reads as one row
// per thing the user changed — while redactions and annotations accumulate.
import { isNeutralAdjust, type AdjustChannel } from "./imageAdjust";
import {
  MAX_EDIT_STACK,
  normaliseStack,
  resolveEditGeometry,
  type ImageEdit,
  type PixelSize,
} from "./imageEdits";

type EditOf<K extends ImageEdit["kind"]> = Extract<ImageEdit, { kind: K }>;

/** The history's entry of a single-valued kind, if any. */
export function findEntry<K extends ImageEdit["kind"]>(history: readonly ImageEdit[], kind: K): EditOf<K> | undefined {
  return history.find((edit): edit is EditOf<K> => edit.kind === kind);
}

/** `edit` replacing the history's entry of its kind (straighten, crop or
 * resize), or appended when there is none. */
export function upsertEntry(history: readonly ImageEdit[], edit: ImageEdit): ImageEdit[] {
  const index = history.findIndex((entry) => entry.kind === edit.kind);
  if (index < 0) {
    return [...history, edit];
  }
  return history.flatMap((entry, i) => (i === index ? [edit] : entry.kind === edit.kind ? [] : [entry]));
}

/** The history without its entry of `kind`. */
export function removeKind(history: readonly ImageEdit[], kind: ImageEdit["kind"]): ImageEdit[] {
  return history.filter((entry) => entry.kind !== kind);
}

/** One adjust channel set, merged into the adjust entry (dropped when every
 * channel is back at neutral). */
export function setAdjustChannel(
  history: readonly ImageEdit[],
  channel: AdjustChannel,
  value: number | boolean,
): ImageEdit[] {
  const current = findEntry(history, "adjust") ?? { kind: "adjust" as const };
  const next: EditOf<"adjust"> = { ...current };
  if (value === 0 || value === false) {
    delete next[channel];
  } else if (channel === "grayscale" || channel === "invert") {
    next[channel] = value === true;
  } else {
    next[channel] = Number(value);
  }
  return isNeutralAdjust(next) ? removeKind(history, "adjust") : upsertEntry(history, next);
}

/** Turns the rotate entry by `delta` quarter turns clockwise (−1 is left);
 * a full turn removes it. */
export function rotateBy(history: readonly ImageEdit[], delta: number): ImageEdit[] {
  const current = findEntry(history, "rotate")?.quarterTurns ?? 0;
  const turns = ((((current + delta) % 4) + 4) % 4) as 0 | 1 | 2 | 3;
  return turns === 0 ? removeKind(history, "rotate") : upsertEntry(history, { kind: "rotate", quarterTurns: turns });
}

/** Flips one axis: flipping an axis already flipped removes that flip. */
export function toggleFlip(history: readonly ImageEdit[], axis: "horizontal" | "vertical"): ImageEdit[] {
  const index = history.findIndex((entry) => entry.kind === "flip" && entry.axis === axis);
  return index >= 0 ? history.filter((_, i) => i !== index) : [...history, { kind: "flip", axis }];
}

/** How many redactions and annotations the history holds. */
export function objectCount(history: readonly ImageEdit[]): number {
  return history.filter((entry) => entry.kind === "redact" || entry.kind === "annotate").length;
}

/** Whether another redaction or annotation still fits the stack bound. */
export function canAddObject(history: readonly ImageEdit[]): boolean {
  return normaliseStack(history).length < MAX_EDIT_STACK;
}

/** The history with entry `index` replaced. */
export function replaceAt(history: readonly ImageEdit[], index: number, edit: ImageEdit): ImageEdit[] {
  return history.map((entry, i) => (i === index ? edit : entry));
}

/** The history without entry `index`. */
export function removeAt(history: readonly ImageEdit[], index: number): ImageEdit[] {
  return history.filter((_, i) => i !== index);
}

/** Keeps a resize's scale when the size it scales changes under it: when
 * `next` changed the crop, straighten or rotation (so the pre-resize size
 * moved) but left the resize entry as it was, that entry is rebased to the
 * same per-axis factors over the new size — a rotation swaps it, a tighter
 * crop shrinks it — rather than stretching the new picture to the old
 * pixel count. */
export function keepResizeScale(
  previous: readonly ImageEdit[],
  next: readonly ImageEdit[],
  naturalPx: PixelSize,
): ImageEdit[] {
  const before = findEntry(previous, "resize");
  const after = findEntry(next, "resize");
  if (!before || !after || before.widthPx !== after.widthPx || before.heightPx !== after.heightPx) {
    return [...next];
  }
  const was = resolveEditGeometry(previous, naturalPx).preResizePx;
  const now = resolveEditGeometry(next, naturalPx).preResizePx;
  if (was.widthPx === now.widthPx && was.heightPx === now.heightPx) {
    return [...next];
  }
  const swapped = was.widthPx === now.heightPx && was.heightPx === now.widthPx && was.widthPx !== was.heightPx;
  const widthPx = swapped ? after.heightPx : Math.max(1, Math.round((now.widthPx * after.widthPx) / was.widthPx));
  const heightPx = swapped ? after.widthPx : Math.max(1, Math.round((now.heightPx * after.heightPx) / was.heightPx));
  return upsertEntry(next, { ...after, widthPx, heightPx });
}
