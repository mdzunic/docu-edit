// The tray as a page organiser (SPEC-111): entries selected as a set —
// Shift+click for a range, Cmd/Ctrl+click to add or remove one, Cmd/Ctrl+A
// for all, or every page of one source — and the page actions on them:
// quarter turns left or right (onto each page's own /Rotate, mod 360),
// delete (from the PLAN only, never the stored source), duplicate, and a
// blank page inserted after the selection. An action on an entry applies to
// the whole selection when that entry is part of it. Each change is said
// through the tray's `announcement`. useAssemblyTray composes it over the
// same entries.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { pageCountLabel } from "../../lib/assembly";
import { MAX_MERGE_PAGES } from "../../lib/core/pdfDocument";
import { rotationLabel, turned } from "../../lib/pdfMerge";
import { blankEntry, entryLabel, freshKey, isPageEntry, PAGES_FULL_MESSAGE, type TrayEntry } from "./trayEntries";

/** What a selection click carries: Shift for a range, Cmd/Ctrl to toggle. */
export interface PickModifiers {
  range?: boolean;
  toggle?: boolean;
}

export interface TrayPagesDeps {
  entries: Ref<TrayEntry[]>;
  /** The tray's own selected entry, which a click also sets. */
  selectedKey: Ref<string>;
  announcement: Ref<string>;
  refusal: Ref<string>;
}

export interface TrayPagesController {
  /** The selected entries' keys, in tray order. */
  selectedKeys: ComputedRef<string[]>;
  isSelected: (key: string) => boolean;
  /** A click on an entry: alone it selects just that entry; with `toggle`
   * (Cmd/Ctrl) it adds or removes it; with `range` (Shift) it selects from
   * the last plain or toggled click to it. */
  pick: (key: string, modifiers?: PickModifiers) => void;
  /** Cmd/Ctrl+A. */
  selectAll: () => void;
  /** "Select all pages from this source": every page of `key`'s PDF. */
  selectSource: (key: string) => void;
  /** No more pages fit (MAX_MERGE_PAGES). */
  isPagesFull: ComputedRef<boolean>;
  /** page-rotate-left/right: `delta` (±90) turns onto `key`'s page — or
   * every selected page, when `key` is one of them. */
  rotate: (key: string, delta: number) => void;
  /** page-delete: from the plan only, never from the stored source. */
  removePages: (key: string) => void;
  /** page-duplicate: a copy of each page right after it. */
  duplicatePages: (key: string) => boolean;
  /** assembly-blank-page: after the selection, else at the end. */
  addBlank: () => boolean;
}

/** The tray's own hold on the selection, beside the controller it shares. */
export interface TraySelection {
  /** Selects `key` alone ("" selects nothing). */
  only: (key: string) => void;
  /** Forgets keys that left the tray. */
  forget: (keys: Iterable<string>) => void;
}

export function useTrayPages(deps: TrayPagesDeps): { pages: TrayPagesController; selection: TraySelection } {
  const { entries, selectedKey, announcement, refusal } = deps;
  // The selection, and the entry a Shift+click ranges from.
  const selection = ref<ReadonlySet<string>>(new Set());
  let anchor = "";

  const selectedKeys = computed(() => entries.value.filter((entry) => selection.value.has(entry.key)).map((entry) => entry.key));
  const isPagesFull = computed(() => entries.value.length >= MAX_MERGE_PAGES);

  function indexOf(key: string): number {
    return entries.value.findIndex((entry) => entry.key === key);
  }

  function setSelection(keys: Iterable<string>): void {
    selection.value = new Set(keys);
  }

  /** The keys an action on `key` applies to: the selection when `key` is in
   * it, else `key` alone — in tray order. */
  function targets(key: string): string[] {
    return selection.value.has(key) ? selectedKeys.value : indexOf(key) >= 0 ? [key] : [];
  }

  function pick(key: string, modifiers: PickModifiers = {}): void {
    const index = indexOf(key);
    if (index < 0) {
      return;
    }
    if (modifiers.range) {
      const from = Math.max(0, indexOf(anchor));
      const [lo, hi] = from <= index ? [from, index] : [index, from];
      const range = entries.value.slice(lo, hi + 1).map((entry) => entry.key);
      setSelection(modifiers.toggle ? [...selection.value, ...range] : range);
      selectedKey.value = key;
      return;
    }
    anchor = key;
    if (modifiers.toggle) {
      const next = new Set(selection.value);
      if (next.delete(key)) {
        selectedKey.value = selectedKey.value === key ? "" : selectedKey.value;
      } else {
        next.add(key);
        selectedKey.value = key;
      }
      setSelection(next);
      return;
    }
    setSelection([key]);
    selectedKey.value = key;
  }

  function selectAll(): void {
    setSelection(entries.value.map((entry) => entry.key));
    announcement.value = `Selected all ${entries.value.length}.`;
  }

  function selectSource(key: string): void {
    const entry = entries.value[indexOf(key)];
    if (!entry?.page) {
      return;
    }
    const keys = entries.value.filter((candidate) => candidate.page && candidate.id === entry.id).map((candidate) => candidate.key);
    setSelection(keys);
    anchor = key;
    selectedKey.value = key;
    announcement.value = `Selected ${pageCountLabel(keys.length)} of “${entry.title}”.`;
  }

  function rotate(key: string, delta: number): void {
    const keys = new Set(targets(key));
    const next = entries.value.map((entry) =>
      keys.has(entry.key) && isPageEntry(entry) ? { ...entry, rotateBy: turned(entry.rotateBy ?? 0, delta) } : entry,
    );
    const changed = next.filter((entry, index) => entry !== entries.value[index]);
    if (changed.length === 0) {
      return;
    }
    entries.value = next;
    const way = delta < 0 ? "left" : "right";
    announcement.value =
      changed.length === 1
        ? `Turned ${entryLabel(changed[0])} ${way}: ${rotationLabel({ rotate: changed[0].page?.rotate ?? 0, rotateBy: changed[0].rotateBy ?? 0 })}.`
        : `Turned ${changed.length} pages ${way}.`;
  }

  function removePages(key: string): void {
    const keys = new Set(targets(key));
    const gone = entries.value.filter((entry) => keys.has(entry.key));
    if (gone.length === 0) {
      return;
    }
    entries.value = entries.value.filter((entry) => !keys.has(entry.key));
    setSelection([...selection.value].filter((candidate) => !keys.has(candidate)));
    if (keys.has(selectedKey.value)) {
      selectedKey.value = "";
    }
    refusal.value = "";
    announcement.value = gone.length === 1 ? `Removed ${entryLabel(gone[0])} from the plan.` : `Removed ${gone.length} pages from the plan.`;
  }

  function duplicatePages(key: string): boolean {
    const keys = new Set(targets(key));
    if (keys.size === 0) {
      return false;
    }
    if (entries.value.length + keys.size > MAX_MERGE_PAGES) {
      refusal.value = PAGES_FULL_MESSAGE;
      return false;
    }
    const next: TrayEntry[] = [];
    for (const entry of entries.value) {
      next.push(entry);
      if (keys.has(entry.key)) {
        next.push({ ...entry, key: freshKey() });
      }
    }
    entries.value = next;
    refusal.value = "";
    const [first] = entries.value.filter((entry) => keys.has(entry.key));
    announcement.value = keys.size === 1 ? `Duplicated ${entryLabel(first)}.` : `Duplicated ${keys.size} pages.`;
    return true;
  }

  function addBlank(): boolean {
    if (isPagesFull.value) {
      refusal.value = PAGES_FULL_MESSAGE;
      return false;
    }
    const last = selectedKeys.value.length > 0 ? indexOf(selectedKeys.value[selectedKeys.value.length - 1]) : -1;
    const at = last >= 0 ? last + 1 : entries.value.length;
    entries.value = [...entries.value.slice(0, at), blankEntry(), ...entries.value.slice(at)];
    refusal.value = "";
    announcement.value = `Added a blank page at position ${at + 1} of ${entries.value.length}.`;
    return true;
  }

  return {
    pages: {
      selectedKeys,
      isSelected: (key: string) => selection.value.has(key),
      pick,
      selectAll,
      selectSource,
      isPagesFull,
      rotate,
      removePages,
      duplicatePages,
      addBlank,
    },
    selection: {
      only: (key: string) => {
        setSelection(key ? [key] : []);
        anchor = key;
      },
      forget: (keys: Iterable<string>) => {
        const gone = new Set(keys);
        if ([...selection.value].some((key) => gone.has(key))) {
          setSelection([...selection.value].filter((key) => !gone.has(key)));
        }
        if (gone.has(anchor)) {
          anchor = "";
        }
      },
    },
  };
}
