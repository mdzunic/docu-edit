// The slide sorter's controller (SPEC-099): a second view over the same deck,
// `currentIndex` and slide commands the filmstrip drives — a grid of every
// slide under its section headers, the shared multi-selection, block
// drag-reorder, duplicate/delete of the selection, section commands, the
// slide clipboard and "Reuse slides". Every mutation is record(label) →
// mutate → commit() with the labels SPEC-099 names, so each is one undo step.
// Built once by PresentationView (one `useSlideSorter({...})` call); it builds
// useSlideSelection (shared with the filmstrip) and useReuseSlides inside, as
// usePresentationEditing builds useFrameSelection.
import { computed, ref, type Ref, type ShallowRef } from "vue";
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
import type { SlideImportReport } from "../../lib/presentation/slideImport";
import { slideOutline } from "../../lib/presentation/slideModel";
import { describeImport, isInsideOverlay, isTypingTarget, plural, type SorterReport } from "./sorterFields";
import { useSlideSelection } from "./useSlideSelection";
import { useReuseSlides } from "./useReuseSlides";
import { useSorterClipboard } from "./useSorterClipboard";

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
    const sections = sectionsOf(model);
    if (sections.length === 0) {
      return [{ key: null, name: "", collapsed: false, startIndex: 0, endIndex: model.slides.length, cards: cards(0, model.slides.length) }];
    }
    return sections.map((section) => ({
      key: section.section.key,
      name: section.section.name,
      collapsed: section.section.collapsed === true,
      startIndex: section.startIndex,
      endIndex: section.endIndex,
      cards: cards(section.startIndex, section.endIndex),
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

  // --- clipboard: copy/cut/paste of slides (useSorterClipboard) -----------------

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

  const clipboard = useSorterClipboard({ model: deps.model, open, selection, run, remove, announce, showReport });
  const { copy, cut, paste, canPaste, onClipboardEvent } = clipboard;

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
