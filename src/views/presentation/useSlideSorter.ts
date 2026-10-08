// The slide sorter's controller (SPEC-099): a second view over the same deck,
// `currentIndex` and slide commands the filmstrip drives — a grid of every
// slide under its section headers, the shared multi-selection, block
// drag-reorder, duplicate/delete of the selection, section commands, the
// slide clipboard and "Reuse slides". Every mutation is record(label) →
// mutate → commit() with the labels SPEC-099 names, so each is one undo step.
// Built once by PresentationView (one `useSlideSorter({...})` call); it builds
// useSlideSelection (shared with the filmstrip) and useReuseSlides inside, as
// usePresentationEditing builds useFrameSelection.
import { computed, onBeforeUnmount, onMounted, ref, type Ref, type ShallowRef } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import {
  addSection as addSectionCommand,
  cleanSectionName,
  moveSection as moveSectionCommand,
  removeSection as removeSectionCommand,
  sectionsOf,
} from "../../lib/presentation/sections";
import { deleteSlides, duplicateSlides, isNoopMove, moveSlides } from "../../lib/presentation/slides";
import { importSlides, type SlideImportReport } from "../../lib/presentation/slideImport";
import {
  latestSlideCopy,
  resolveSlidePaste,
  SLIDES_MIME,
  slideClipboardFlavours,
  type SlideClipboardPayload,
} from "../../lib/presentation/slideClipboard";
import { slideOutline } from "../../lib/presentation/slideModel";
import { frameAccessibleName } from "../../components/presentation/slideRender";
import { useSlideSelection } from "./useSlideSelection";
import { useReuseSlides } from "./useReuseSlides";

export interface SlideSorterDeps {
  model: ShallowRef<PresentationModel | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  reportError: (message: string) => void;
}

/** One card in the grid. */
export interface SorterCard {
  index: number;
  slide: Slide;
}

/** One run of cards under its header; `key` null = a deck with no sections
 * (one headerless group). */
export interface SorterGroup {
  key: string | null;
  name: string;
  collapsed: boolean;
  startIndex: number;
  endIndex: number;
  cards: SorterCard[];
}

/** The import report as the sorter shows it. */
export interface SorterReport {
  title: string;
  lines: string[];
}

/** What the sorter says when a paste carries no slides (AC-9). */
export const NO_SLIDES_MESSAGE = "The clipboard doesn't hold any slides";

function plural(n: number, noun: string): string {
  return `${n} ${noun}${n === 1 ? "" : "s"}`;
}

/** The honest summary of an import: what arrived (by title), what was reused,
 * and what could not come along. */
export function describeImport(report: SlideImportReport, titles: string[], verb: string): SorterReport {
  const named = titles.map((title) => `“${title}”`).join(", ");
  const lines = [`Kept ${plural(report.slides, "slide")}: ${named}.`];
  if (report.mediaReused > 0) {
    lines.push(`Reused ${plural(report.mediaReused, "image")} already in this deck.`);
  }
  if (report.rescaled) {
    lines.push("Resized to fit this deck's slide size.");
  }
  if (report.droppedFrames > 0) {
    lines.push(`Left out ${plural(report.droppedFrames, "object")} this deck can't hold: ${report.droppedFrameNames.join(", ")}.`);
  }
  if (report.lostTransitions > 0) {
    lines.push(`${plural(report.lostTransitions, "transition")} couldn't come along.`);
  }
  if (report.lostAnimations > 0) {
    lines.push(`${plural(report.lostAnimations, "animation")} couldn't come along.`);
  }
  if (report.droppedFrames + report.lostTransitions + report.lostAnimations === 0) {
    lines.push("Nothing was left out.");
  }
  return { title: `${verb} ${plural(report.slides, "slide")}`, lines };
}

function isTypingTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.tagName === "SELECT" || target.isContentEditable)
  );
}

function isInsideOverlay(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[role="dialog"], [role="menu"]') !== null;
}

export function useSlideSorter(deps: SlideSorterDeps) {
  const { preferences, setPreference } = usePreferences();
  const open = computed(() => preferences.presentationSorterOpen);
  const announcement = ref("");
  const report = ref<SorterReport | null>(null);
  /** The section whose name is being edited (the new one, after "New section"). */
  const renaming = ref<string | null>(null);

  function toggle(): void {
    setPreference("presentationSorterOpen", !open.value);
  }

  const count = (): number => deps.model.value?.slides.length ?? 0;
  const selection = useSlideSelection({
    count,
    currentIndex: deps.currentIndex,
    select: (index) => {
      if (index >= 0 && index < count()) {
        deps.currentIndex.value = index;
      }
    },
  });

  /** The grid's groups, a fresh array per commit (the useAnimations rule). */
  const rows = computed<SorterGroup[]>(() => {
    const model = deps.model.value;
    if (!model) {
      return [];
    }
    const cards = (start: number, end: number): SorterCard[] =>
      model.slides.slice(start, end).map((slide, offset) => ({ index: start + offset, slide }));
    const runs = sectionsOf(model);
    if (runs.length === 0) {
      return [{ key: null, name: "", collapsed: false, startIndex: 0, endIndex: model.slides.length, cards: cards(0, model.slides.length) }];
    }
    return runs.map((run) => ({
      key: run.section.key,
      name: run.section.name,
      collapsed: run.section.collapsed === true,
      startIndex: run.startIndex,
      endIndex: run.endIndex,
      cards: cards(run.startIndex, run.endIndex),
    }));
  });

  /** An odp deck carrying sections: Impress keeps them but shows them nowhere,
   * which the sorter says once (AC-8). */
  const odpNote = computed(() => deps.model.value?.format === "odp" && (deps.model.value.sections?.length ?? 0) > 0);

  function announce(message: string): void {
    // Re-set even when unchanged so a repeated result is announced again.
    announcement.value = "";
    announcement.value = message;
  }

  function run<T>(label: string, apply: (model: PresentationModel) => T): T | undefined {
    const model = deps.model.value;
    if (!model) {
      return undefined;
    }
    deps.record(label);
    const result = apply(model);
    deps.commit();
    return result;
  }

  function selectAt(index: number, event?: MouseEvent | KeyboardEvent): void {
    selection.click(index, event ? { shiftKey: event.shiftKey, ctrlKey: event.ctrlKey, metaKey: event.metaKey } : {});
    announce(`${plural(selection.indexes.value.length, "slide")} selected`);
  }

  function selectAll(): void {
    selection.selectAll();
    announce(`${plural(selection.indexes.value.length, "slide")} selected`);
  }

  function clear(): void {
    selection.collapse();
    announce("1 slide selected");
  }

  // --- sections ------------------------------------------------------------

  function addSection(atIndex = selection.indexes.value[0] ?? 0, name = "New section"): void {
    const key = run("New section", (model) => addSectionCommand(model, atIndex, name));
    if (key) {
      renaming.value = key;
      announce(`Section added`);
    }
  }

  /** Commits a rename; false (nothing recorded) for a name empty after trim. */
  function renameSection(key: string, name: string): boolean {
    const model = deps.model.value;
    const clean = cleanSectionName(name);
    const section = model?.sections?.find((entry) => entry.key === key);
    renaming.value = null;
    if (!model || !section || clean === null) {
      return false;
    }
    if (clean !== section.name) {
      run("Rename section", () => {
        section.name = clean;
        model.structureDirty = true;
      });
    }
    return true;
  }

  function moveSection(key: string, delta: number): void {
    const start = run("Move section", (model) => moveSectionCommand(model, key, delta));
    if (start !== undefined && start >= 0) {
      selectRun(key);
      announce(`Section moved ${delta < 0 ? "up" : "down"}`);
    }
  }

  function removeSection(key: string, mode: "sectionOnly" | "withSlides"): void {
    const next = run("Remove section", (model) => removeSectionCommand(model, key, mode));
    if (next !== undefined && next >= 0) {
      deps.currentIndex.value = next;
      selection.collapse();
      announce(mode === "withSlides" ? "Section and its slides removed" : "Section removed");
    }
  }

  /** Collapses or expands a run — editor view state only: no undo step and
   * nothing serialised (`DeckSection.collapsed`), just persisted with the deck. */
  function toggleCollapsed(key: string): void {
    const section = deps.model.value?.sections?.find((entry) => entry.key === key);
    if (!section) {
      return;
    }
    if (section.collapsed) {
      delete section.collapsed;
    } else {
      section.collapsed = true;
    }
    deps.commit();
  }

  function selectRun(key: string): void {
    const group = rows.value.find((row) => row.key === key);
    if (group && group.endIndex > group.startIndex) {
      selection.selectRange(
        group.cards.map((card) => card.index),
        group.startIndex,
      );
    }
  }

  // --- slide commands ----------------------------------------------------------

  function duplicate(): void {
    const copies = run("Duplicate slides", (model) => duplicateSlides(model, selection.indexes.value));
    if (copies && copies.length > 0) {
      selection.selectRange(copies);
      announce(`${plural(copies.length, "slide")} duplicated`);
    }
  }

  function remove(): void {
    const picked = selection.indexes.value;
    if (picked.length === 0) {
      return;
    }
    const next = run("Delete slides", (model) => deleteSlides(model, picked));
    if (next !== undefined) {
      deps.currentIndex.value = next;
      selection.collapse();
      announce(`${plural(picked.length, "slide")} deleted`);
    }
  }

  /** Drops the selection as a block at insertion point `to`, joining
   * `sectionKey` when the drop names one (a slide dragged past a header adopts
   * the section it lands in). A drop that changes nothing records nothing. */
  function reorder(to: number, sectionKey?: string): void {
    const model = deps.model.value;
    const picked = selection.indexes.value;
    if (!model || isNoopMove(model, picked, to, sectionKey)) {
      return;
    }
    const moved = run("Move slides", (deck) => moveSlides(deck, picked, to, sectionKey));
    if (moved && moved.length > 0) {
      selection.selectRange(moved);
      announce(
        moved.length === 1
          ? `Slide moved to position ${moved[0] + 1}`
          : `${plural(moved.length, "slide")} moved to positions ${moved[0] + 1}–${moved[moved.length - 1] + 1}`,
      );
    }
  }

  // --- clipboard ---------------------------------------------------------------

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

  function showReport(result: SlideImportReport, at: number, verb: string): void {
    const model = deps.model.value;
    const titles = (model?.slides.slice(at, at + result.slides) ?? []).map(
      (slide, i) => slideOutline(slide).title ?? `Slide ${at + i + 1}`,
    );
    report.value = describeImport(result, titles, verb);
  }

  function dismissReport(): void {
    report.value = null;
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
    const count = selection.indexes.value.length;
    remove();
    announce(`${plural(count, "slide")} cut`);
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
    let payload: SlideClipboardPayload | null = null;
    try {
      const text = (await navigator.clipboard?.readText()) ?? "";
      payload = resolveSlidePaste({ text });
    } catch {
      payload = latestSlideCopy() ?? null;
    }
    if (payload) {
      pasteFrom(payload);
    } else {
      announce(NO_SLIDES_MESSAGE);
    }
  }

  const canPaste = computed(() => deps.model.value !== null);

  // --- Reuse slides ------------------------------------------------------------

  const reuse = useReuseSlides({
    model: deps.model,
    record: deps.record,
    commit: deps.commit,
    reportError: deps.reportError,
    insertAt: () => Math.max(...selection.indexes.value, -1) + 1,
    onImported: (result, at) => {
      selection.selectRange(Array.from({ length: result.slides }, (_, i) => at + i));
      showReport(result, at, "Inserted");
      announce(`${plural(result.slides, "slide")} inserted`);
    },
  });

  // --- keyboard ----------------------------------------------------------------

  /** The sorter's keys while it is open (registered ahead of the stage's
   * window shortcuts): Cmd/Ctrl+A/D, Delete, and the selection keys. Other
   * keys are kept from the hidden stage's handlers, except undo/redo and the
   * Cmd/Ctrl+Shift shortcuts (notes, present), which pass through; Cmd/Ctrl+
   * C/X/V are left alone so the browser fires the clipboard events. */
  function onGlobalKeydown(event: KeyboardEvent): void {
    if (!open.value || isTypingTarget(event.target) || isInsideOverlay(event.target)) {
      return;
    }
    const mod = event.metaKey || event.ctrlKey;
    const key = event.key.toLowerCase();
    if (mod && !event.shiftKey && key === "d") {
      event.preventDefault();
      event.stopImmediatePropagation();
      duplicate();
      return;
    }
    if (!mod && (event.key === "Delete" || event.key === "Backspace")) {
      event.preventDefault();
      event.stopImmediatePropagation();
      remove();
      return;
    }
    if (selection.onKeydown(event)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      announce(`${plural(selection.indexes.value.length, "slide")} selected`);
      return;
    }
    const passThrough =
      (mod && (event.shiftKey || ["z", "y", "c", "x", "v"].includes(key))) || /^F\d+$/.test(event.key) || event.key === "Escape";
    if (!passThrough) {
      event.stopImmediatePropagation();
    }
  }

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

  return {
    open,
    toggle,
    rows,
    selection,
    selectAt,
    selectAll,
    clear,
    addSection,
    renameSection,
    moveSection,
    removeSection,
    toggleCollapsed,
    renaming,
    copy,
    cut,
    paste,
    canPaste,
    duplicate,
    remove,
    reorder,
    reuse,
    report,
    dismissReport,
    announcement,
    odpNote,
    onClipboardEvent,
    onGlobalKeydown,
  };
}

export type SlideSorter = ReturnType<typeof useSlideSorter>;
