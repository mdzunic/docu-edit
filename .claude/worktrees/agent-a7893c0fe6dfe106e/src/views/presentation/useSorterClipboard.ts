// The slide sorter's clipboard (SPEC-099), lifted out of useSlideSorter so
// the controller stays within its composable budget: Cmd/Ctrl+C/X/V and the
// sorter's Copy/Cut/Paste buttons write and read slides through
// ClipboardEvent.clipboardData exactly as SPEC-016's DOC_MODEL_MIME does —
// `text/plain` (the outline; all another application ever gets) plus the
// private SLIDES_MIME flavour, with the in-tab store standing in above the
// size cap and for the async API's text-only read. A paste carrying no slides
// is a no-op that says so; a paste into a text input or an open dialog is left
// to its own handler (SPEC-081's text session owns its contenteditable).
import { computed, onBeforeUnmount, onMounted, type ComputedRef, type ShallowRef } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { importSlides, type SlideImportReport } from "../../lib/presentation/slideImport";
import {
  latestSlideCopy,
  resolveSlidePaste,
  SLIDES_MIME,
  slideClipboardFlavours,
  type SlideClipboardPayload,
} from "../../lib/presentation/slideClipboard";
import { frameAccessibleName } from "../../components/presentation/slideRender";
import { isInsideOverlay, isTypingTarget, NO_SLIDES_MESSAGE, plural } from "./sorterFields";
import type { SlideSelection } from "./useSlideSelection";

export interface SorterClipboardDeps {
  model: ShallowRef<PresentationModel | null>;
  /** Whether the sorter is showing — the clipboard events are its only then. */
  open: ComputedRef<boolean>;
  selection: SlideSelection;
  /** The sorter's record(label) → mutate → commit() runner. */
  run: <T>(label: string, apply: (model: PresentationModel) => T) => T | undefined;
  /** Deletes the selection ("Delete slides") — a cut's second half. */
  remove: () => void;
  announce: (message: string) => void;
  showReport: (result: SlideImportReport, at: number, verb: string) => void;
}

export function useSorterClipboard(deps: SorterClipboardDeps) {
  const { selection, run, remove, announce, showReport, open } = deps;

  /** Writes the selection's two flavours to a clipboard event (SPEC-016's
   * discipline: only ClipboardEvent.clipboardData carries the private type). */
  function writeSelection(data: DataTransfer): boolean {
    const model = deps.model.value;
    if (!model || selection.indexes.value.length === 0) {
      return false;
    }
    const flavours = slideClipboardFlavours(model, selection.indexes.value);
    data.setData("text/plain", flavours.text);
    if (flavours.payload !== null) {
      data.setData(SLIDES_MIME, flavours.payload);
    }
    return true;
  }

  /** Inserts a payload's slides after the selection (one undo step). */
  function pasteFrom(payload: SlideClipboardPayload): void {
    const at = Math.max(...selection.indexes.value, -1) + 1;
    const result = run("Paste slides", (model) =>
      importSlides(model, payload, payload.slides.map((_, index) => index), { at, frameName: frameAccessibleName }),
    );
    if (!result || result.slides === 0) {
      announce(NO_SLIDES_MESSAGE);
      return;
    }
    const pasted = Array.from({ length: result.slides }, (_, i) => at + i);
    selection.selectRange(pasted);
    announce(`${plural(result.slides, "slide")} pasted`);
    if (result.droppedFrames + result.lostTransitions + result.lostAnimations > 0 || result.rescaled) {
      showReport(result, at, "Pasted");
    }
  }

  /** Cmd/Ctrl+C / X / V and the sorter's buttons land here while the sorter
   * is open: copy and cut write both flavours; paste takes the slides payload
   * (the in-tab copy winning by nonce) or announces there are none. */
  function onClipboardEvent(event: ClipboardEvent): void {
    if (!open.value || isTypingTarget(event.target) || isInsideOverlay(event.target) || !event.clipboardData) {
      return;
    }
    if (event.type === "copy" || event.type === "cut") {
      if (writeSelection(event.clipboardData)) {
        event.preventDefault();
        if (event.type === "cut") {
          removeAsCut();
        } else {
          announce(`${plural(selection.indexes.value.length, "slide")} copied`);
        }
      }
      return;
    }
    if (event.type === "paste") {
      event.preventDefault();
      const payload = resolveSlidePaste({
        slides: event.clipboardData.getData(SLIDES_MIME),
        text: event.clipboardData.getData("text/plain"),
      });
      if (payload) {
        pasteFrom(payload);
      } else {
        announce(NO_SLIDES_MESSAGE);
      }
    }
  }

  function removeAsCut(): void {
    const cutCount = selection.indexes.value.length;
    remove();
    announce(`${plural(cutCount, "slide")} cut`);
  }

  /** The Copy/Cut buttons: a real copy/cut via execCommand — the only
   * synchronous route from a click to clipboardData (and the private type);
   * the event lands in onClipboardEvent. Without it, the outline text goes
   * through the async API and the in-tab store still pastes the slides. */
  function copy(): void {
    if (!document.execCommand?.("copy")) {
      const model = deps.model.value;
      if (model) {
        const flavours = slideClipboardFlavours(model, selection.indexes.value);
        void navigator.clipboard?.writeText(flavours.text).catch(() => {});
        announce(`${plural(selection.indexes.value.length, "slide")} copied`);
      }
    }
  }

  function cut(): void {
    if (!document.execCommand?.("cut")) {
      const model = deps.model.value;
      if (model) {
        const flavours = slideClipboardFlavours(model, selection.indexes.value);
        void navigator.clipboard?.writeText(flavours.text).catch(() => {});
        removeAsCut();
      }
    }
  }

  /** The Paste button: the async API cannot see the private type, so the
   * clipboard's text is matched against this tab's own copies; when the
   * clipboard cannot be read at all, the latest in-tab copy is used. */
  async function paste(): Promise<void> {
    let payload: SlideClipboardPayload | null;
    const api = navigator.clipboard;
    if (!api?.readText) {
      payload = latestSlideCopy() ?? null;
    } else {
      try {
        payload = resolveSlidePaste({ text: await api.readText() });
      } catch {
        payload = latestSlideCopy() ?? null;
      }
    }
    if (payload) {
      pasteFrom(payload);
    } else {
      announce(NO_SLIDES_MESSAGE);
    }
  }

  const canPaste = computed(() => deps.model.value !== null);

  onMounted(() => {
    for (const type of ["copy", "cut", "paste"]) {
      window.addEventListener(type, onClipboardEvent as EventListener);
    }
  });
  onBeforeUnmount(() => {
    for (const type of ["copy", "cut", "paste"]) {
      window.removeEventListener(type, onClipboardEvent as EventListener);
    }
  });

  return { copy, cut, paste, canPaste, onClipboardEvent };
}
