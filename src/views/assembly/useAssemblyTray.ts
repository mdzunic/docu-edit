// The assembly tray (SPEC-109): the chosen library images in the order they
// will be laid out — and nothing but this order decides the layout order.
// Holds up to MAX_ASSEMBLY_ITEMS entries; adds (from the library picker or
// "Combine…"), removes, duplicates and reorders them (drag, or Alt+↑/Alt+↓
// on a focused row), saying each change through `announcement`, which the
// tray renders in its role="status" live region. A duplicate is its own
// entry with its own key (and, for overlay, its own transform) over the same
// library image.
//
// Only ever READS the library: an entry carries the image's id, title and
// pixel size; the source document is never touched.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { getMedia, getStoredRecord } from "../../lib/documentStore";
import { isImageFormat } from "../../lib/libraryMenu";
import { moveEntry, trayMoveAnnouncement } from "../../lib/assembly";
import { DEFAULT_OVERLAY_TRANSFORM, MAX_ASSEMBLY_ITEMS, normaliseTransform, type OverlayTransform } from "../../lib/core/assemblyLayout";
import { sourceItemPx } from "../library/useConvertToPdf";

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

export interface TrayEntry extends TrayImage {
  /** Unique within the tray: a duplicate shares its source's `id`. */
  key: string;
  /** overlay only. */
  transform: OverlayTransform;
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

/** A tray entry over `image`, with its own key. */
export function trayEntry(image: TrayImage, transform?: Partial<OverlayTransform>): TrayEntry {
  keySeed += 1;
  return { ...image, key: `tray-${keySeed}`, transform: normaliseTransform(transform ?? DEFAULT_OVERLAY_TRANSFORM) };
}

export const TRAY_FULL_MESSAGE = `A tray holds at most ${MAX_ASSEMBLY_ITEMS} images.`;

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
}

export function useAssemblyTray(): AssemblyTrayController {
  const entries = ref<TrayEntry[]>([]);
  const selectedKey = ref("");
  const announcement = ref("");
  const refusal = ref("");

  const selected = computed(() => entries.value.find((entry) => entry.key === selectedKey.value) ?? null);
  const isFull = computed(() => entries.value.length >= MAX_ASSEMBLY_ITEMS);

  function indexOf(key: string): number {
    return entries.value.findIndex((entry) => entry.key === key);
  }

  function add(ids: readonly string[]): number {
    const images = ids.map(libraryImage).filter((image) => image !== null);
    const room = Math.max(0, MAX_ASSEMBLY_ITEMS - entries.value.length);
    const added = images.slice(0, room).map((image) => trayEntry(image));
    const left = images.length - added.length;
    refusal.value = left > 0 ? `${TRAY_FULL_MESSAGE} ${left} ${left === 1 ? "image was" : "images were"} not added.` : "";
    if (added.length > 0) {
      entries.value = [...entries.value, ...added];
      announcement.value = added.length === 1 ? `Added “${added[0].title}”.` : `Added ${added.length} images.`;
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
    announcement.value = trayMoveAnnouncement(entries.value[to].title, to + 1, entries.value.length);
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
      announcement.value = `“${entry.title}” is already ${to < 0 ? "first" : "last"}.`;
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
    entries.value = next.slice(0, MAX_ASSEMBLY_ITEMS);
    selectedKey.value = "";
    refusal.value = "";
  }

  return {
    entries,
    selectedKey,
    selected,
    announcement,
    refusal,
    isFull,
    add,
    remove,
    duplicate,
    move,
    moveTo,
    setTransform,
    replace,
  };
}
