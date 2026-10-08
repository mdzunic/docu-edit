// The slide multi-selection (SPEC-099), shared verbatim by the slide sorter
// and the filmstrip: one instance, built by useSlideSorter, drives both views'
// `aria-selected` and their gestures, so a selection made in one is the
// selection in the other. The selection always contains the current slide —
// a plain click selects one slide and makes it current; Cmd/Ctrl+click
// toggles; Shift+click extends from the anchor; Cmd/Ctrl+A selects all;
// Escape collapses back to the current slide; arrows/Home/End move the
// roving selection and Shift+arrow extends it. Moving the current slide from
// anywhere else (PageDown on the stage, a command's returned index) collapses
// the selection onto it.
import { computed, ref, watch, type Ref } from "vue";

export interface SlideSelectionDeps {
  /** How many slides the deck has. */
  count: () => number;
  currentIndex: Ref<number>;
  /** Makes `index` the current slide (the deck navigation's selectSlide). */
  select: (index: number) => void;
}

/** The modifier state of a pointer gesture: Shift extends, Cmd/Ctrl toggles. */
export interface SelectionModifiers {
  shiftKey?: boolean;
  ctrlKey?: boolean;
  metaKey?: boolean;
}

function range(from: number, to: number): number[] {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  return Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
}

export function useSlideSelection(deps: SlideSelectionDeps) {
  // The explicit multi-selection; [] means "just the current slide".
  const picked = ref<number[]>([]);
  const anchor = ref(-1);

  /** The selected slide indexes, ascending; never empty while the deck has
   * slides. */
  const indexes = computed<number[]>(() => {
    const total = deps.count();
    const current = deps.currentIndex.value;
    const list = picked.value.filter((index) => index >= 0 && index < total);
    if (list.length > 0) {
      return list;
    }
    return total > 0 ? [Math.min(Math.max(current, 0), total - 1)] : [];
  });

  /** The current slide moved somewhere the selection does not cover → the
   * selection collapses onto it. */
  watch(deps.currentIndex, (current) => {
    if (!picked.value.includes(current)) {
      picked.value = [];
      anchor.value = current;
    }
  });

  function set(list: number[], current: number): void {
    const total = deps.count();
    picked.value = [...new Set(list)].filter((index) => index >= 0 && index < total).toSorted((a, b) => a - b);
    deps.select(current);
  }

  function anchorIndex(): number {
    return anchor.value >= 0 && anchor.value < deps.count() ? anchor.value : deps.currentIndex.value;
  }

  /** A click on slide `index` with the gesture's modifiers. */
  function click(index: number, modifiers: SelectionModifiers = {}): void {
    if (index < 0 || index >= deps.count()) {
      return;
    }
    if (modifiers.shiftKey) {
      set(range(anchorIndex(), index), index);
      return;
    }
    if (modifiers.metaKey || modifiers.ctrlKey) {
      const current = indexes.value;
      if (current.includes(index)) {
        const rest = current.filter((other) => other !== index);
        if (rest.length > 0) {
          anchor.value = index;
          set(rest, rest.includes(deps.currentIndex.value) ? deps.currentIndex.value : rest[rest.length - 1]);
        }
        return;
      }
      anchor.value = index;
      set([...current, index], index);
      return;
    }
    anchor.value = index;
    set([index], index);
  }

  function selectAll(): void {
    const total = deps.count();
    if (total > 0) {
      set(range(0, total - 1), deps.currentIndex.value);
    }
  }

  /** Back to just the current slide. */
  function collapse(): void {
    picked.value = [];
    anchor.value = deps.currentIndex.value;
  }

  /** Selects exactly `list` (a command's result — moved, duplicated or pasted
   * slides), making `current` (default: the first) the current slide. */
  function selectRange(list: number[], current = list[0]): void {
    if (list.length === 0 || current === undefined) {
      return;
    }
    anchor.value = list[0];
    set(list, current);
  }

  /** Moves the roving selection by `delta` (or to an end); Shift extends from
   * the anchor. */
  function move(target: number, extend: boolean): void {
    const total = deps.count();
    const next = Math.min(Math.max(target, 0), total - 1);
    if (total === 0) {
      return;
    }
    if (extend) {
      set(range(anchorIndex(), next), next);
    } else {
      anchor.value = next;
      set([next], next);
    }
  }

  /** The selection keys; true when the event was one of them (the caller then
   * prevents the default and keeps it from other handlers). */
  function onKeydown(event: KeyboardEvent): boolean {
    const mod = event.metaKey || event.ctrlKey;
    const current = deps.currentIndex.value;
    if (mod && !event.shiftKey && !event.altKey && (event.key === "a" || event.key === "A")) {
      selectAll();
      return true;
    }
    if (mod || event.altKey) {
      return false;
    }
    switch (event.key) {
      case "Escape":
        if (picked.value.length > 1) {
          collapse();
          return true;
        }
        return false;
      case "ArrowDown":
      case "ArrowRight":
        move(current + 1, event.shiftKey);
        return true;
      case "ArrowUp":
      case "ArrowLeft":
        move(current - 1, event.shiftKey);
        return true;
      case "Home":
        move(0, event.shiftKey);
        return true;
      case "End":
        move(deps.count() - 1, event.shiftKey);
        return true;
      default:
        return false;
    }
  }

  return { indexes, anchor, click, selectAll, collapse, selectRange, onKeydown };
}

export type SlideSelection = ReturnType<typeof useSlideSelection>;
