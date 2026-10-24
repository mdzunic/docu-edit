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
import { getMedia, getStoredRecord } from "../../lib/documentStore";
import { isImageFormat } from "../../lib/libraryMenu";
import { moveEntry, pageCountLabel, trayMoveAnnouncement } from "../../lib/assembly";
import { DEFAULT_OVERLAY_TRANSFORM, MAX_ASSEMBLY_ITEMS, normaliseTransform, type OverlayTransform } from "../../lib/core/assemblyLayout";
import { MAX_MERGE_PAGES } from "../../lib/core/pdfDocument";
import { capTrayItems } from "../../lib/assemblyPersist";
import { BLANK_PAGE_TITLE, rotationLabel, turned } from "../../lib/pdfMerge";
import type { PdfSummary } from "../../lib/media/model";
import { sourceItemPx } from "../library/useConvertToPdf";

/** Whether a library entry of `format` can go in the tray: an image, or
 * (SPEC-111) a PDF, as its pages. */
export function isCombinableFormat(format: string | undefined): boolean {
  return isImageFormat(format) || format === "pdf";
}

/** A library image as the tray needs it. */
export interface TrayImage {
  id: string;
  title: string;
  format: string;
  lastModified: number;
  /** 0 when the image's size could not be read. */
  widthPx: number;
  heightPx: number;
}

/** SPEC-111: what a page entry knows of its page, from the PDF's summary. */
export interface TrayPage {
  /** 0-based index into the PDF's pages. */
  index: number;
  /** How many pages the PDF has. */
  count: number;
  widthPt: number;
  heightPt: number;
  /** The page's own /Rotate. */
  rotate: 0 | 90 | 180 | 270;
}

export interface TrayEntry extends TrayImage {
  /** Unique within the tray: a duplicate shares its source's `id`. */
  key: string;
  /** overlay only. */
  transform: OverlayTransform;
  /** SPEC-111: a page of the library PDF `id`. */
  page?: TrayPage;
  /** SPEC-111: an empty page at the sheet's size (`id` is ""). */
  blank?: boolean;
  /** SPEC-111: quarter turns added to a page, mod 360 (0 when absent). */
  rotateBy?: 0 | 90 | 180 | 270;
}

/** Whether an entry is a page of a PDF plan (a PDF page or a blank page)
 * rather than an image to lay out. */
export function isPageEntry(entry: Pick<TrayEntry, "page" | "blank">): boolean {
  return entry.page !== undefined || entry.blank === true;
}

/** How an announcement names an entry: “shot”, page 2 of “contract”, the
 * blank page. */
export function entryLabel(entry: Pick<TrayEntry, "title" | "page" | "blank">): string {
  if (entry.blank) {
    return "the blank page";
  }
  return entry.page ? `page ${entry.page.index + 1} of “${entry.title}”` : `“${entry.title}”`;
}

/** A page card's accessible name: "Page 3: page 2 of “contract”, rotate 90,
 * selected". */
export function pageCardLabel(entry: TrayEntry, position: number, selected: boolean): string {
  const rotation = rotationLabel({ rotate: entry.page?.rotate ?? 0, rotateBy: entry.rotateBy ?? 0 });
  return `Page ${position}: ${entryLabel(entry)}, ${rotation}${selected ? ", selected" : ""}`;
}

/** A library PDF as the tray expands it: the source, and its pages from the
 * stored summary (no payload is read) — null when it is missing, trashed or
 * has no summary. */
export interface TrayPdf {
  source: TrayImage;
  pages: PdfSummary["pages"];
}

export function libraryPdf(id: string): TrayPdf | null {
  const record = getStoredRecord(id);
  const model = getMedia(id);
  if (!record || !model || record.trashedAt !== undefined || model.format !== "pdf" || !model.pdf) {
    return null;
  }
  return {
    source: { id, title: model.title, format: "pdf", lastModified: record.lastModified, widthPx: 0, heightPx: 0 },
    pages: model.pdf.pages,
  };
}

/** The library image stored under `id` — null when it is missing, trashed
 * or not an image. */
export function libraryImage(id: string): TrayImage | null {
  const record = getStoredRecord(id);
  const model = getMedia(id);
  if (!record || !model || record.trashedAt !== undefined || !isImageFormat(model.format)) {
    return null;
  }
  let size = null;
  try {
    size = sourceItemPx(model);
  } catch {
    size = null;
  }
  return {
    id,
    title: model.title,
    format: model.format,
    lastModified: record.lastModified,
    widthPx: size?.widthPx ?? 0,
    heightPx: size?.heightPx ?? 0,
  };
}

let keySeed = 0;

function freshKey(): string {
  keySeed += 1;
  return `tray-${keySeed}`;
}

/** A tray entry over `image`, with its own key. */
export function trayEntry(image: TrayImage, transform?: Partial<OverlayTransform>): TrayEntry {
  return { ...image, key: freshKey(), transform: normaliseTransform(transform ?? DEFAULT_OVERLAY_TRANSFORM) };
}

/** Page entries for `pdf`: every page in order by default, or the pages
 * `items` name (each with its turns); an index the PDF lacks is skipped. */
export function pageEntries(
  pdf: TrayPdf,
  items: ReadonlyArray<{ index: number; rotateBy?: 0 | 90 | 180 | 270 }> = pdf.pages.map((_, index) => ({ index })),
): TrayEntry[] {
  return items
    .filter((item) => Number.isInteger(item.index) && item.index >= 0 && item.index < pdf.pages.length)
    .map((item) => {
      const page = pdf.pages[item.index];
      return {
        ...trayEntry(pdf.source),
        page: { index: item.index, count: pdf.pages.length, widthPt: page.widthPt, heightPt: page.heightPt, rotate: page.rotate },
        ...(item.rotateBy ? { rotateBy: item.rotateBy } : {}),
      };
    });
}

/** A blank page entry. */
export function blankEntry(rotateBy: 0 | 90 | 180 | 270 = 0): TrayEntry {
  return {
    ...trayEntry({ id: "", title: BLANK_PAGE_TITLE, format: "pdf", lastModified: 0, widthPx: 0, heightPx: 0 }),
    blank: true,
    ...(rotateBy ? { rotateBy } : {}),
  };
}

/** Typed pass-through for entries read off a `.vue` prop (runtime-typed
 * `Object`; espree bars the annotation in `.vue` scripts). */
export function asTrayEntries(value: unknown): TrayEntry[] {
  return Array.isArray(value) ? (value as TrayEntry[]) : [];
}

export const TRAY_FULL_MESSAGE = `A tray holds at most ${MAX_ASSEMBLY_ITEMS} images.`;
export const PAGES_FULL_MESSAGE = `A PDF plan holds at most ${MAX_MERGE_PAGES} pages.`;

/** What a selection click carries: Shift for a range, Cmd/Ctrl to toggle. */
export interface PickModifiers {
  range?: boolean;
  toggle?: boolean;
}

export interface AssemblyTrayController {
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
  /** Adds the library images `ids`, in order, up to the limit. Returns how
   * many were added. */
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
  /** SPEC-111: the selected entries' keys, in tray order. */
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

export function useAssemblyTray(): AssemblyTrayController {
  const entries = ref<TrayEntry[]>([]);
  const selectedKey = ref("");
  const announcement = ref("");
  const refusal = ref("");
  // The page selection, and the entry a Shift+click ranges from.
  const selection = ref<ReadonlySet<string>>(new Set());
  let anchor = "";

  const selected = computed(() => entries.value.find((entry) => entry.key === selectedKey.value) ?? null);
  const imageCount = computed(() => entries.value.filter((entry) => !isPageEntry(entry)).length);
  const isFull = computed(() => imageCount.value >= MAX_ASSEMBLY_ITEMS || entries.value.length >= MAX_MERGE_PAGES);
  const isPagesFull = computed(() => entries.value.length >= MAX_MERGE_PAGES);
  const selectedKeys = computed(() => entries.value.filter((entry) => selection.value.has(entry.key)).map((entry) => entry.key));

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
        const pages = pageEntries(pdf);
        const fit = pages.slice(0, Math.max(0, MAX_MERGE_PAGES - total));
        added.push(...fit);
        total += fit.length;
        pagesLeft += pages.length - fit.length;
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
    if (selection.value.has(key)) {
      setSelection([...selection.value].filter((candidate) => candidate !== key));
    }
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
      ? `${capitalised(entryLabel(moved))} moved to position ${to + 1} of ${entries.value.length}.`
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
      announcement.value = `${capitalised(entryLabel(entry))} is already ${to < 0 ? "first" : "last"}.`;
      return false;
    }
    return moveTo(key, to);
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
    setSelection([]);
    anchor = "";
    refusal.value = "";
  }

  return {
    entries,
    selectedKey,
    selected,
    announcement,
    refusal,
    isFull,
    select: (key: string) => {
      selectedKey.value = key;
      setSelection(key ? [key] : []);
      anchor = key;
    },
    add,
    remove,
    duplicate,
    move,
    moveTo,
    setTransform,
    replace,
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
  };
}

function capitalised(text: string): string {
  return text.charAt(0).toUpperCase() + text.slice(1);
}
