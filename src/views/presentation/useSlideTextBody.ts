// The text-body half of SPEC-097's slide text depth (composed by
// useSlideTextDepth): the Text panel's autofit / anchor / insets / columns
// values and commands over the edited-or-selected text or shape frame, the
// frame-level command runner the link commands share, and the AUTOFIT
// WRITE-BACK RULE. A scale/height reaches the model ONLY inside a mutation
// that has already called markSlideEdited: picking a mode in text-autofit
// ("shrink" stores the measured fontScalePct/lineScalePct, "resizeShape" sets
// heightCm = resizedHeightCm(frame, contentHeightCm)), or a later command of
// the text-depth composables on that frame, re-measured once the stage has
// re-rendered (nextTick) and folded into the same edit. Rendering and reading
// never write — slideTextFit.ts measures, only a command stores.
import { computed, nextTick, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import { markSlideEdited } from "../../lib/presentation/text";
import {
  applyTextBody,
  effectiveAnchor,
  effectiveAutofit,
  effectiveColumns,
  effectiveInsetsCm,
  isTextBodyFrame,
  resizedHeightCm,
  type TextAnchor,
  type TextAutofit,
  type TextBodyFrame,
  type TextBodyPatch,
} from "../../lib/presentation/textBody";
import { stageFit, type FrameMeasurer } from "./slideTextDepthFields";

export interface SlideTextBodyDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  record: (label?: string) => void;
  commit: () => void;
  /** The frame the body controls edit (the edited frame, else the one
   * selected text/shape frame). */
  bodyFrame: ComputedRef<TextBodyFrame | null>;
  /** Measures a rendered body; defaults to the stage DOM (slideTextFit.ts). */
  measureFit?: FrameMeasurer;
}

export function useSlideTextBody(deps: SlideTextBodyDeps) {
  const measure = deps.measureFit ?? stageFit;

  function frameById(id: string): SlideFrame | null {
    return deps.currentSlide.value?.frames.find((frame) => frame.id === id) ?? null;
  }

  const bodyEnabled = computed(() => deps.bodyFrame.value !== null);
  const autofit = computed<TextAutofit | "">(() => (deps.bodyFrame.value ? effectiveAutofit(deps.bodyFrame.value) : ""));
  const anchor = computed<TextAnchor | "">(() => (deps.bodyFrame.value ? effectiveAnchor(deps.bodyFrame.value) : ""));
  const insetsCm = computed(() => (deps.bodyFrame.value ? effectiveInsetsCm(deps.bodyFrame.value) : null));
  const columns = computed(() => (deps.bodyFrame.value ? effectiveColumns(deps.bodyFrame.value) : null));

  /** Stores what the measurer reports for a frame's autofit mode — the shrink
   * scale (100 stored as absent), or the resized height — and returns whether
   * anything changed. Called ONLY inside a mutation that already called
   * markSlideEdited. A body that is not rendered keeps what is stored. */
  function writeFit(frame: TextBodyFrame): boolean {
    const mode = effectiveAutofit(frame);
    if (mode === "none") return false;
    const fit = measure(frame);
    if (!fit) return false;
    if (mode === "shrink") {
      const before = `${frame.body?.fontScalePct}/${frame.body?.lineScalePct}`;
      applyTextBody(frame, {
        fontScalePct: fit.scale.fontScalePct < 100 ? fit.scale.fontScalePct : null,
        lineScalePct: fit.scale.lineScalePct < 100 ? fit.scale.lineScalePct : null,
      });
      return `${frame.body?.fontScalePct}/${frame.body?.lineScalePct}` !== before;
    }
    const height = resizedHeightCm(frame, fit.contentHeightCm);
    if (height === frame.heightCm) return false;
    frame.heightCm = height;
    return true;
  }

  /** After a command on an autofitting frame, re-measure once the stage has
   * re-rendered and fold the result into the SAME edit — no new undo entry
   * (the command's record() already snapshotted the pre-edit state). */
  async function refit(frameId: string | null): Promise<void> {
    if (!frameId) return;
    await nextTick();
    const frame = frameById(frameId);
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!isTextBodyFrame(frame) || !slide || !model || effectiveAutofit(frame) === "none") return;
    if (writeFit(frame)) {
      markSlideEdited(slide, model.format);
      deps.commit();
    }
  }

  /** A frame-level command as ONE undo entry: record → mutate →
   * markSlideEdited → (measure an autofitting body) → commit → refit. */
  function frameCommand(label: string, frame: SlideFrame | null, mutate: (frame: SlideFrame) => void): boolean {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!frame || !slide || !model) return false;
    deps.record(label);
    mutate(frame);
    markSlideEdited(slide, model.format);
    if (isTextBodyFrame(frame)) writeFit(frame);
    deps.commit();
    void refit(frame.id);
    return true;
  }

  function bodyCommand(label: string, patch: TextBodyPatch): void {
    frameCommand(label, deps.bodyFrame.value, (frame) => {
      if (isTextBodyFrame(frame)) applyTextBody(frame, patch);
    });
  }

  /** Picking a mode in text-autofit — the one place a scale is first stored. */
  function setAutofit(mode: TextAutofit): void {
    bodyCommand("Text autofit", { autofit: mode });
  }
  function setAnchor(value: TextAnchor): void {
    bodyCommand("Text anchor", { anchor: value });
  }
  function setInset(side: "left" | "top" | "right" | "bottom", cm: number): void {
    const current = insetsCm.value;
    if (!current || !Number.isFinite(cm)) return;
    bodyCommand("Text insets", { insetsCm: { ...current, [side]: cm } });
  }
  function setColumnCount(count: number): void {
    const current = columns.value;
    if (!current || !Number.isFinite(count)) return;
    bodyCommand("Text columns", { columns: { count, gapCm: current.gapCm } });
  }
  function setColumnGap(cm: number): void {
    const current = columns.value;
    if (!current || !Number.isFinite(cm)) return;
    bodyCommand("Text columns", { columns: { count: current.count, gapCm: cm } });
  }

  return {
    bodyEnabled,
    autofit,
    anchor,
    insetsCm,
    columns,
    refit,
    frameCommand,
    setAutofit,
    setAnchor,
    setInset,
    setColumnCount,
    setColumnGap,
  };
}
