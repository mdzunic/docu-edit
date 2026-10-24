// The assembly tray (SPEC-109): the chosen library images in the order they
// will be laid out — and nothing but this order decides the layout order.
// Holds up to MAX_ASSEMBLY_ITEMS entries; adds (from the library picker or
// "Combine…"), removes, duplicates and reorders them (drag, or Alt+↑/Alt+↓
// on a focused row), saying each change through `announcement`, which the
// tray renders in its role="status" live region. A duplicate is its own
// entry with its own key (and, for overlay, its own transform) over the same
// library image.
//
// SPEC-111 makes it a page organiser too: a PDF added to the tray expands
// into one entry per PAGE (with the quarter turns the user adds to it), a
// blank page can be inserted, and entries can be selected as a set —
// Shift+click for a range, Cmd/Ctrl+click to add or remove one, Cmd/Ctrl+A
// for all, or every page of one source — so a rotation, a duplicate or a
// delete applies to the whole selection. Deleting a page removes it from
// the PLAN only. A tray holds up to MAX_ASSEMBLY_ITEMS images and
// MAX_MERGE_PAGES entries in all.
//
// Only ever READS the library: an entry carries the image's (or PDF's) id,
// title and size; the source document is never touched.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { moveEntry, pageCountLabel, trayMoveAnnouncement } from "../../lib/assembly";
import { MAX_ASSEMBLY_ITEMS, normaliseTransform, type OverlayTransform } from "../../lib/core/assemblyLayout";
import { MAX_MERGE_PAGES } from "../../lib/core/pdfDocument";
import { capTrayItems } from "../../lib/assemblyPersist";
import {
  entrySentenceLabel,
  isPageEntry,
  libraryImage,
  libraryPdf,
  PAGES_FULL_MESSAGE,
  pageEntries,
  TRAY_FULL_MESSAGE,
  trayEntry,
  type TrayEntry,
} from "./trayEntries";
import { useTrayPages, type TrayPagesController } from "./useTrayPages";

// The entries themselves (and the page organiser's controller) live beside
// this file; the tray's importers keep reaching them through here.
export {
  asTrayEntries,
  blankEntry,
  entryLabel,
  isCombinableFormat,
  isPageEntry,
  libraryImage,
  libraryPdf,
  pageCardLabel,
  pageEntries,
  PAGES_FULL_MESSAGE,
  TRAY_FULL_MESSAGE,
  trayEntry,
  type TrayEntry,
  type TrayImage,
  type TrayPage,
  type TrayPdf,
} from "./trayEntries";
export type { PickModifiers } from "./useTrayPages";

export interface AssemblyTrayController extends TrayPagesController {
  entries: Ref<TrayEntry[]>;
  /** The selected entry's key ("" when none) — the inspector edits it. */
  selectedKey: Ref<string>;
  selected: ComputedRef<TrayEntry | null>;
  /** The live region's words for the last change. */
  announcement: Ref<string>;
  /** Why the last add or duplicate was refused or cut short ("" when not). */
  refusal: Ref<string>;
  isFull: ComputedRef<boolean>;
  select: (key: string) => void;
  /** Adds the library images (and, SPEC-111, PDFs as their pages) `ids`, in
   * order, up to the limits. Returns how many entries were added. */
  add: (ids: readonly string[]) => number;
  remove: (key: string) => void;
  duplicate: (key: string) => boolean;
  /** Moves an entry `delta` places (the keyboard's Alt+↑ / Alt+↓). */
  move: (key: string, delta: number) => boolean;
  /** Moves an entry to `index` (a drop). */
  moveTo: (key: string, index: number) => boolean;
  setTransform: (key: string, transform: OverlayTransform) => void;
  /** Replaces the whole tray (a restored plan, Start over). */
  replace: (entries: TrayEntry[]) => void;
}

export function useAssemblyTray(): AssemblyTrayController {
  const entries = ref<TrayEntry[]>([]);
  const selectedKey = ref("");
  const announcement = ref("");
  const refusal = ref("");
  const { pages, selection } = useTrayPages({ entries, selectedKey, announcement, refusal });

  const selected = computed(() => entries.value.find((entry) => entry.key === selectedKey.value) ?? null);
  const imageCount = computed(() => entries.value.filter((entry) => !isPageEntry(entry)).length);
  const isFull = computed(() => imageCount.value >= MAX_ASSEMBLY_ITEMS || entries.value.length >= MAX_MERGE_PAGES);

  function indexOf(key: string): number {
    return entries.value.findIndex((entry) => entry.key === key);
  }

  function add(ids: readonly string[]): number {
    const added: TrayEntry[] = [];
    let images = imageCount.value;
    let total = entries.value.length;
    let imagesLeft = 0;
    let pagesLeft = 0;
    let anyPdf = false;
    for (const id of ids) {
      const image = libraryImage(id);
      if (image) {
        if (images < MAX_ASSEMBLY_ITEMS && total < MAX_MERGE_PAGES) {
          added.push(trayEntry(image));
          images += 1;
          total += 1;
        } else {
          imagesLeft += 1;
        }
        continue;
      }
      const pdf = libraryPdf(id);
      if (pdf) {
        anyPdf = true;
        const pdfPages = pageEntries(pdf);
        const fit = pdfPages.slice(0, Math.max(0, MAX_MERGE_PAGES - total));
        added.push(...fit);
        total += fit.length;
        pagesLeft += pdfPages.length - fit.length;
      }
    }
    refusal.value = [
      imagesLeft > 0 ? `${TRAY_FULL_MESSAGE} ${imagesLeft} ${imagesLeft === 1 ? "image was" : "images were"} not added.` : "",
      pagesLeft > 0 ? `${PAGES_FULL_MESSAGE} ${pagesLeft} ${pagesLeft === 1 ? "page was" : "pages were"} not added.` : "",
    ]
      .filter(Boolean)
      .join(" ");
    if (added.length > 0) {
      entries.value = [...entries.value, ...added];
      announcement.value = anyPdf
        ? `Added ${pageCountLabel(added.length)}.`
        : added.length === 1
          ? `Added “${added[0].title}”.`
          : `Added ${added.length} images.`;
    }
    return added.length;
  }

  function remove(key: string): void {
    const index = indexOf(key);
    if (index < 0) {
      return;
    }
    const [gone] = entries.value.splice(index, 1);
    entries.value = [...entries.value];
    if (selectedKey.value === key) {
      selectedKey.value = "";
    }
    selection.forget([key]);
    refusal.value = "";
    announcement.value = `Removed “${gone.title}”.`;
  }

  function duplicate(key: string): boolean {
    const index = indexOf(key);
    if (index < 0) {
      return false;
    }
    if (isFull.value) {
      refusal.value = TRAY_FULL_MESSAGE;
      return false;
    }
    const source = entries.value[index];
    const copy = trayEntry(source, source.transform);
    entries.value = [...entries.value.slice(0, index + 1), copy, ...entries.value.slice(index + 1)];
    refusal.value = "";
    announcement.value = `Duplicated “${source.title}” at position ${index + 2} of ${entries.value.length}.`;
    return true;
  }

  function moveTo(key: string, target: number): boolean {
    const from = indexOf(key);
    const to = Math.min(Math.max(0, target), entries.value.length - 1);
    if (from < 0 || to === from) {
      return false;
    }
    entries.value = moveEntry(entries.value, from, to);
    const moved = entries.value[to];
    announcement.value = isPageEntry(moved)
      ? `${entrySentenceLabel(moved)} moved to position ${to + 1} of ${entries.value.length}.`
      : trayMoveAnnouncement(moved.title, to + 1, entries.value.length);
    return true;
  }

  function move(key: string, delta: number): boolean {
    const from = indexOf(key);
    if (from < 0) {
      return false;
    }
    const to = from + delta;
    if (to < 0 || to >= entries.value.length) {
      const entry = entries.value[from];
      announcement.value = `${entrySentenceLabel(entry)} is already ${to < 0 ? "first" : "last"}.`;
      return false;
    }
    return moveTo(key, to);
  }

  function setTransform(key: string, transform: OverlayTransform): void {
    const index = indexOf(key);
    if (index >= 0) {
      const next = [...entries.value];
      next[index] = { ...next[index], transform: normaliseTransform(transform) };
      entries.value = next;
    }
  }

  function replace(next: TrayEntry[]): void {
    entries.value = capTrayItems(next, isPageEntry);
    selectedKey.value = "";
    selection.only("");
    refusal.value = "";
  }

  return {
    ...pages,
    entries,
    selectedKey,
    selected,
    announcement,
    refusal,
    isFull,
    select: (key: string) => {
      selectedKey.value = key;
      selection.only(key);
    },
    add,
    remove,
    duplicate,
    move,
    moveTo,
    setTransform,
    replace,
  };
}
