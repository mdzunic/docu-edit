// The image editor's rotate, adjust and straighten commands (SPEC-112): thin
// wrappers turning the options bars' controls into history mutations
// through the session's gesture API. A slider is a gesture — its first
// `input` records ("Adjust brightness"), every later one updates the
// preview, and its `change` (the release, or one arrow-key step) commits —
// so a drag across a slider is one undo entry; a typed number or a button
// is one `apply`. Every label reads as the history list does.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { AdjustChannel } from "../../lib/core/imageAdjust";
import { findEntry, removeKind, rotateBy, setAdjustChannel, toggleFlip, upsertEntry } from "../../lib/core/imageEditHistory";
import { MAX_EDIT_EDGE_PX, MAX_STRAIGHTEN_DEGREES, type ImageEdit } from "../../lib/core/imageEdits";
import type { EditSession } from "./editSession";

export interface EditCommands {
  rotate: (delta: 1 | -1) => void;
  flip: (axis: "horizontal" | "vertical") => void;
  /** The adjust entry's value for a channel (0 / false when neutral). */
  adjustValue: (channel: AdjustChannel) => number | boolean;
  /** A slider's `input`: begins the gesture on the first one. */
  adjustInput: (channel: AdjustChannel, value: number) => void;
  /** A one-shot change: a typed number or a toggle. */
  setAdjust: (channel: AdjustChannel, value: number | boolean) => void;
  straighten: ComputedRef<number>;
  straightenInput: (degrees: number) => void;
  setStraighten: (degrees: number) => void;
  /** Ends a slider gesture (its `change`). */
  release: () => void;
  resizeLocked: Ref<boolean>;
  resample: Ref<"nearest" | "bilinear">;
  /** The size a resize would start from (the output before any resize). */
  resizeBase: ComputedRef<{ widthPx: number; heightPx: number } | null>;
  /** The output size now. */
  outputSize: ComputedRef<{ widthPx: number; heightPx: number } | null>;
  setResizeWidth: (widthPx: number) => void;
  setResizeHeight: (heightPx: number) => void;
  setResizePercent: (percent: number) => void;
  setResample: (resample: "nearest" | "bilinear") => void;
  clearResize: () => void;
}

function clampEdge(value: number): number {
  return Math.min(MAX_EDIT_EDGE_PX, Math.max(1, Math.round(Number.isFinite(value) ? value : 1)));
}

export function useEditCommands(session: EditSession): EditCommands {
  const resizeLocked = ref(true);
  const resample = ref<"nearest" | "bilinear">("bilinear");
  let sliding: string | null = null;

  function release(): void {
    if (sliding !== null) {
      sliding = null;
      session.end();
    }
  }

  function slide(label: string, next: () => ImageEdit[]): void {
    if (sliding !== label || !session.active()) {
      session.begin(label);
      sliding = label;
    }
    session.update(next());
  }

  function once(label: string, next: () => ImageEdit[]): void {
    release();
    session.apply(label, next());
  }

  const adjustEntry = computed(() => findEntry(session.stack.value, "adjust"));
  function adjustValue(channel: AdjustChannel): number | boolean {
    const value = adjustEntry.value?.[channel];
    return channel === "grayscale" || channel === "invert" ? value === true : typeof value === "number" ? value : 0;
  }

  const straighten = computed(() => findEntry(session.stack.value, "straighten")?.degrees ?? 0);
  function straightened(degrees: number): ImageEdit[] {
    const value = Math.round(Math.min(MAX_STRAIGHTEN_DEGREES, Math.max(-MAX_STRAIGHTEN_DEGREES, Number(degrees) || 0)) * 100) / 100;
    return value === 0
      ? removeKind(session.history.value, "straighten")
      : upsertEntry(session.history.value, { kind: "straighten", degrees: value });
  }

  const outputSize = computed(() => {
    const geometry = session.geometry.value;
    return geometry ? { widthPx: geometry.widthPx, heightPx: geometry.heightPx } : null;
  });
  const resizeBase = computed(() => session.geometry.value?.preResizePx ?? null);

  function resized(widthPx: number, heightPx: number): ImageEdit[] {
    const base = resizeBase.value;
    if (base && widthPx === base.widthPx && heightPx === base.heightPx) {
      return removeKind(session.history.value, "resize");
    }
    return upsertEntry(session.history.value, { kind: "resize", widthPx, heightPx, resample: resample.value });
  }

  function setResizeWidth(value: number): void {
    const base = resizeBase.value;
    const current = outputSize.value;
    if (!base || !current) return;
    const widthPx = clampEdge(value);
    const heightPx = resizeLocked.value ? clampEdge((widthPx * base.heightPx) / base.widthPx) : current.heightPx;
    once("Resize", () => resized(widthPx, heightPx));
  }

  function setResizeHeight(value: number): void {
    const base = resizeBase.value;
    const current = outputSize.value;
    if (!base || !current) return;
    const heightPx = clampEdge(value);
    const widthPx = resizeLocked.value ? clampEdge((heightPx * base.widthPx) / base.heightPx) : current.widthPx;
    once("Resize", () => resized(widthPx, heightPx));
  }

  function setResizePercent(percent: number): void {
    const base = resizeBase.value;
    if (!base || !(percent > 0)) return;
    once("Resize", () => resized(clampEdge((base.widthPx * percent) / 100), clampEdge((base.heightPx * percent) / 100)));
  }

  function setResample(value: "nearest" | "bilinear"): void {
    resample.value = value;
    const entry = findEntry(session.history.value, "resize");
    if (entry && entry.resample !== value) {
      once("Resize", () => upsertEntry(session.history.value, { ...entry, resample: value }));
    }
  }

  return {
    rotate: (delta) => once(delta > 0 ? "Rotate right" : "Rotate left", () => rotateBy(session.history.value, delta)),
    flip: (axis) => once(axis === "horizontal" ? "Flip horizontal" : "Flip vertical", () => toggleFlip(session.history.value, axis)),
    adjustValue,
    adjustInput: (channel, value) =>
      slide(`Adjust ${channel}`, () => setAdjustChannel(session.history.value, channel, Number(value) || 0)),
    setAdjust: (channel, value) => once(`Adjust ${channel}`, () => setAdjustChannel(session.history.value, channel, value)),
    straighten,
    straightenInput: (degrees) => slide("Straighten", () => straightened(degrees)),
    setStraighten: (degrees) => once("Straighten", () => straightened(degrees)),
    release,
    resizeLocked,
    resample,
    resizeBase,
    outputSize,
    setResizeWidth,
    setResizeHeight,
    setResizePercent,
    setResample,
    clearResize: () => once("Reset size", () => removeKind(session.history.value, "resize")),
  };
}
