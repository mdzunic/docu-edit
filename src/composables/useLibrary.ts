// The document library's view state (SPEC-067): search, sort, filters,
// multi-select, the trash toggle and the storage-usage figures. It ties the
// reactive UI state to the pure derivations in lib/libraryView and the data
// layer in lib/documentStore, so HomeView.vue stays a thin wiring hub.
//
// View mode and sort are persisted (AC-2, AC-15); nothing about selection or
// the filters is (spec: "Selection state lives in the view").
import { computed, ref, shallowRef } from "vue";
import type { DocumentFormat, LibraryEntry } from "../lib/documentStore";
import {
  deleteDocument,
  duplicateDocument,
  duplicateStoredDocument,
  libraryStorageUsage,
  listDocuments,
  listTrashedDocuments,
  restoreDocument,
  setStarred,
  trashDocument,
} from "../lib/documentStore";
import type { DocumentKind } from "../lib/documentKinds";
import {
  filterCounts,
  filterLibrary,
  searchLibrary,
  sortLibrary,
  toggleChip,
  type FormatChip,
  type LibraryFilter,
} from "../lib/libraryView";
import {
  loadLibrarySort,
  loadLibraryView,
  saveLibrarySort,
  saveLibraryView,
  type LibrarySort,
  type LibrarySortKey,
  type LibraryViewMode,
} from "../lib/libraryPrefs";
import { downloadStoredDocument } from "../lib/libraryDownload";
import { isImageFormat } from "../lib/libraryMenu";

/** The default sort direction for each key: newest/largest first for time and
 * size, A–Z for title (AC-11, AC-12). */
const DEFAULT_DIRECTION: Record<LibrarySortKey, "asc" | "desc"> = {
  modified: "desc",
  title: "asc",
  size: "desc",
};

/** Bulk download (AC-27..29): each entry through its own product's writer,
 * dispatched by kind in lib/libraryDownload.ts (see documentKinds.ts). */
async function downloadDocuments(ids: string[]): Promise<void> {
  for (const id of ids) {
    await downloadStoredDocument(id);
  }
}

export function useLibrary() {
  const view = ref<LibraryViewMode>(loadLibraryView());
  const sort = ref<LibrarySort>(loadLibrarySort());
  const search = ref("");
  const activeFormats = shallowRef<Set<DocumentFormat>>(new Set());
  const starredOnly = ref(false);
  const showTrash = ref(false);

  const entries = shallowRef<LibraryEntry[]>([]);
  const trashed = shallowRef<LibraryEntry[]>([]);
  const usage = shallowRef(libraryStorageUsage());
  const now = ref(Date.now());

  // Selection (AC-23..26): ids, plus the anchor a Shift-click ranges from.
  const selected = shallowRef<Set<string>>(new Set());
  const selectionAnchor = ref("");
  // The row/card that has roving focus for keyboard navigation (AC-57).
  const focusedId = ref("");

  /** Re-reads everything derived from the store. Called after any mutation. */
  function refresh(): void {
    entries.value = listDocuments();
    trashed.value = listTrashedDocuments();
    usage.value = libraryStorageUsage();
    now.value = Date.now();
    // Drop selection entries that no longer exist in the active view.
    const live = new Set((showTrash.value ? trashed.value : entries.value).map((e) => e.id));
    const next = new Set([...selected.value].filter((id) => live.has(id)));
    if (next.size !== selected.value.size) {
      selected.value = next;
    }
  }

  const sizeById = computed(() => {
    const map = new Map<string, number>();
    for (const doc of usage.value.byDocument) {
      map.set(doc.id, doc.bytes);
    }
    return map;
  });
  const sizeOf = (id: string): number => sizeById.value.get(id) ?? 0;

  const filter = computed<LibraryFilter>(() => ({
    search: search.value,
    formats: activeFormats.value,
    starredOnly: starredOnly.value,
  }));

  /** The library documents after search + filters + sort (AC-8, AC-16, AC-21). */
  const visible = computed(() =>
    sortLibrary(filterLibrary(entries.value, filter.value), sort.value, sizeOf),
  );

  /** Trashed documents after the search filter, newest-trashed first. */
  const trashVisible = computed(() => searchLibrary(trashed.value, search.value));

  /** The list currently on screen — the one keyboard nav and select-all act on. */
  const displayed = computed(() => (showTrash.value ? trashVisible.value : visible.value));

  const counts = computed(() => filterCounts(entries.value, search.value));

  const filterActive = computed(
    () =>
      search.value.trim().length > 0 ||
      activeFormats.value.size > 0 ||
      starredOnly.value,
  );

  /** No documents match the current search/filter, though the library is not
   * empty (AC-54) — distinct from an empty library. */
  const noResults = computed(
    () => !showTrash.value && entries.value.length > 0 && visible.value.length === 0,
  );

  /** True when the library has no documents at all (first-run, AC-51). */
  const libraryEmpty = computed(() => entries.value.length === 0);
  // SPEC-109: "New › Combined image" is inert without an image to combine.
  const hasImages = computed(() => entries.value.some((entry) => isImageFormat(entry.format)));

  const trashCount = computed(() => trashed.value.length);

  const entryById = computed(() => {
    const map = new Map<string, LibraryEntry>();
    for (const entry of [...entries.value, ...trashed.value]) {
      map.set(entry.id, entry);
    }
    return map;
  });

  /** The largest documents across the whole library (AC-48), for the storage
   * panel — from libraryStorageUsage, which spans trashed and untrashed. */
  const largestDocuments = computed(() =>
    usage.value.byDocument
      .slice(0, 5)
      .map((doc) => ({
        id: doc.id,
        bytes: doc.bytes,
        title: entryById.value.get(doc.id)?.title ?? "Untitled document",
      })),
  );

  // --- preferences ---------------------------------------------------------

  function setView(next: LibraryViewMode): void {
    view.value = next;
    saveLibraryView(next);
  }

  function setSortKey(key: string): void {
    if (key !== "modified" && key !== "title" && key !== "size") {
      return;
    }
    sort.value = { key, direction: DEFAULT_DIRECTION[key] };
    saveLibrarySort(sort.value);
  }

  function toggleSortDirection(): void {
    sort.value = {
      key: sort.value.key,
      direction: sort.value.direction === "asc" ? "desc" : "asc",
    };
    saveLibrarySort(sort.value);
  }

  /** Setters the library sub-components call in place of assigning to the
   * controller's refs directly — those live behind the `lib` prop there, so a
   * bare `lib.x.value = …` trips vue/no-mutating-props. HomeView, which owns
   * the controller as a local, still assigns to the refs directly. */
  function setSearch(value: string): void {
    search.value = value;
  }

  function setShowTrash(value: boolean): void {
    showTrash.value = value;
  }

  function setFocusedId(id: string): void {
    focusedId.value = id;
  }

  // --- filters -------------------------------------------------------------

  function toggleFormat(format: DocumentFormat): void {
    const next = new Set(activeFormats.value);
    if (next.has(format)) {
      next.delete(format);
    } else {
      next.add(format);
    }
    activeFormats.value = next;
  }

  /** Toggles a whole filter chip (SPEC-107): one chip can stand for several
   * formats ("Images"), so toggling routes through the chip helpers. */
  function toggleFormatChip(chip: FormatChip): void {
    activeFormats.value = toggleChip(chip, activeFormats.value);
  }

  function toggleStarredFilter(): void {
    starredOnly.value = !starredOnly.value;
  }

  function clearFilters(): void {
    search.value = "";
    activeFormats.value = new Set();
    starredOnly.value = false;
  }

  // --- starring (AC-18..22) ------------------------------------------------

  function toggleStar(id: string): void {
    const entry = entryById.value.get(id);
    setStarred(id, !(entry?.starred ?? false));
    refresh();
  }

  // --- selection (AC-23..26) -----------------------------------------------

  function isSelected(id: string): boolean {
    return selected.value.has(id);
  }

  function toggleSelection(id: string): void {
    const next = new Set(selected.value);
    if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    selected.value = next;
    selectionAnchor.value = id;
  }

  /** Makes `id` the sole selection — the grid's plain-click behaviour: the
   * clicked card becomes the quick-look subject without toggling into (or
   * accumulating onto) the checkbox multi-selection. */
  function selectOnly(id: string): void {
    selected.value = new Set([id]);
    selectionAnchor.value = id;
  }

  function selectRange(id: string): void {
    const order = displayed.value.map((entry) => entry.id);
    const anchor = selectionAnchor.value || order[0] || id;
    const from = order.indexOf(anchor);
    const to = order.indexOf(id);
    if (from === -1 || to === -1) {
      toggleSelection(id);
      return;
    }
    const [lo, hi] = from <= to ? [from, to] : [to, from];
    const next = new Set(selected.value);
    for (let i = lo; i <= hi; i += 1) {
      next.add(order[i]!);
    }
    selected.value = next;
  }

  /** Selection click with modifiers (AC-24, AC-25): Shift ranges from the
   * anchor, Ctrl/Cmd toggles a single row, a plain checkbox toggle is the
   * default. */
  function selectWithModifiers(
    id: string,
    modifiers: { shift?: boolean; ctrlOrMeta?: boolean },
  ): void {
    if (modifiers.shift) {
      selectRange(id);
    } else {
      toggleSelection(id);
    }
  }

  function selectAll(): void {
    // Ctrl/Cmd+A selects everything matching the current filter (AC-26).
    selected.value = new Set(displayed.value.map((entry) => entry.id));
  }

  function clearSelection(): void {
    selected.value = new Set();
    selectionAnchor.value = "";
  }

  const selectedIds = computed(() => [...selected.value]);
  const selectionCount = computed(() => selected.value.size);
  const hasSelection = computed(() => selected.value.size > 0);

  // --- trash + undo (AC-30, AC-33..39, AC-44..46) --------------------------

  /** Moves documents to the trash and returns the batch of ids, so the caller
   * can raise the Undo toast (AC-33, AC-34). */
  function trashDocuments(ids: string[]): string[] {
    for (const id of ids) {
      trashDocument(id);
    }
    clearSelection();
    refresh();
    return ids;
  }

  /** Restores a batch of documents — the toast's Undo (AC-35) and the Trash
   * view's Restore (AC-38) both call this. */
  function restoreDocuments(ids: string[]): void {
    for (const id of ids) {
      restoreDocument(id);
    }
    refresh();
  }

  /** Permanently deletes documents — the ONLY path, with Empty trash, that
   * reaches the destructive store operation (AC-46). */
  function deletePermanently(ids: string[]): void {
    for (const id of ids) {
      deleteDocument(id);
    }
    refresh();
  }

  function emptyTrash(): void {
    deletePermanently(trashed.value.map((entry) => entry.id));
  }

  // --- bulk actions (AC-27..29) --------------------------------------------

  function duplicateDocuments(ids: string[]): void {
    // SPEC-137: workbooks duplicate whole (loaded, then released) one at a time.
    const workbooks = ids.filter((id) => kindOf(id) === "spreadsheet");
    for (const id of ids.filter((other) => !workbooks.includes(other))) {
      duplicateDocument(id);
    }
    clearSelection();
    refresh();
    if (workbooks.length > 0) {
      void (async () => {
        for (const id of workbooks) {
          await duplicateStoredDocument(id);
        }
        refresh();
      })();
    }
  }

  function starDocuments(ids: string[]): void {
    // Star all if any is unstarred, else unstar all — a single predictable toggle.
    const anyUnstarred = ids.some((id) => !(entryById.value.get(id)?.starred ?? false));
    for (const id of ids) {
      setStarred(id, anyUnstarred);
    }
    refresh();
  }

  // --- keyboard navigation (AC-57..61) -------------------------------------

  function focusNext(delta: number): void {
    const order = displayed.value.map((entry) => entry.id);
    if (order.length === 0) {
      return;
    }
    const current = order.indexOf(focusedId.value);
    const next = current === -1 ? 0 : Math.min(order.length - 1, Math.max(0, current + delta));
    focusedId.value = order[next]!;
  }

  function focusFirst(): void {
    focusedId.value = displayed.value[0]?.id ?? "";
  }

  /** The title of any document (library or trashed), for toasts and dialogs. */
  function titleOf(id: string): string {
    return entryById.value.get(id)?.title ?? "Untitled document";
  }

  /** Which product an entry belongs to — drives each row's route (see
   * documentKinds.editorPathFor). Unknown ids read as text, the safe
   * default (/doc/:id shows its own not-found state). */
  function kindOf(id: string): DocumentKind {
    return entryById.value.get(id)?.kind ?? "text";
  }

  /** An entry's format ("png", "pdf", "docx", …) — which per-format row
   * actions apply (SPEC-108's Convert to PDF). Undefined for unknown ids. */
  function formatOf(id: string): DocumentFormat | undefined {
    return entryById.value.get(id)?.format;
  }

  /** Roving tabindex for a row/card (AC-57): the focused row is tabbable;
   * with no focus yet, the first row is. Shared by the list, grid and trash
   * renderings so keyboard entry behaves identically in each. */
  function rowTabindex(id: string, index: number): number {
    const focused = focusedId.value;
    if (focused === id) {
      return 0;
    }
    return focused === "" && index === 0 ? 0 : -1;
  }

  return {
    // state
    view,
    sort,
    search,
    activeFormats,
    starredOnly,
    showTrash,
    now,
    focusedId,
    // setters (prop-safe for sub-components)
    setSearch,
    setShowTrash,
    setFocusedId,
    // derived
    visible,
    trashVisible,
    displayed,
    counts,
    filterActive,
    noResults,
    libraryEmpty,
    hasImages,
    trashCount,
    usage,
    largestDocuments,
    sizeOf,
    // selection
    selected,
    selectedIds,
    selectionCount,
    hasSelection,
    isSelected,
    toggleSelection,
    selectOnly,
    selectRange,
    selectWithModifiers,
    selectAll,
    clearSelection,
    // prefs
    setView,
    setSortKey,
    toggleSortDirection,
    // filters
    toggleFormat,
    toggleFormatChip,
    toggleStarredFilter,
    clearFilters,
    // starring
    toggleStar,
    // trash
    trashDocuments,
    restoreDocuments,
    deletePermanently,
    emptyTrash,
    // bulk
    duplicateDocuments,
    starDocuments,
    downloadDocuments,
    // keyboard
    focusNext,
    focusFirst,
    titleOf,
    kindOf,
    formatOf,
    rowTabindex,
    // lifecycle
    refresh,
  };
}

/** The reactive controller the library components receive as their `lib`
 * prop (runtime-typed `Object` in the .vue files — the espree constraint —
 * so this alias is where the real shape lives for .ts callers). */
export type LibraryController = ReturnType<typeof useLibrary>;
