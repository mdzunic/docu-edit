// The workbook view's load path (SPEC-116): resolves the routed id to its
// SpreadsheetModel, guards the kind (any other kind's id hands off to its own
// view through the registry, mirroring the router's guardKind), runs the
// reader-version upgrade for a record stored by an older reader, and owns
// which sheet is showing — the file's active sheet on open, then the tab the
// user picks. A missing id leaves `workbook` null (the not-found state).
// SPEC-137: a stored workbook is read back from storage on open
// (loadSpreadsheet, which also runs the reader-version upgrade) — `loading`
// is true meanwhile, so the view shows its loading state, not the not-found
// one — and released when the view goes (releaseSpreadsheet); a load that
// came back incomplete sets `storageIncomplete` (read-only).
import { computed, onScopeDispose, provide, shallowRef, ref, watch, type ComputedRef, type InjectionKey, type Ref } from "vue";
import type { Router } from "vue-router";
import { editorPathFor } from "../../lib/documentKinds";
import {
  documentKindOf,
  getSpreadsheet,
  getSpreadsheetShell,
  initLibrary,
  isSpreadsheetStorageIncomplete,
  loadSpreadsheet,
  releaseSpreadsheet,
} from "../../lib/documentStore";
import { formatBadgeLabel, spreadsheetFormatLabel } from "../../lib/format";
import { downloadSpreadsheet } from "../../lib/spreadsheet/download";
import { foreignSheetMessage } from "../../lib/spreadsheet/foreign";
import type { SheetPreview, SpreadsheetModel } from "../../lib/spreadsheet/model";

/** One sheet tab: the sheet's index in the model, its id and name. */
export interface SheetTab {
  index: number;
  id: string;
  name: string;
}

/** Narrows SheetTabs' runtime-typed `tabs` prop (the .vue parser bars casts). */
export function asSheetTabs(value: unknown): SheetTab[] {
  return Array.isArray(value) ? (value as SheetTab[]) : [];
}

/** Narrows SheetMiniature's runtime-typed `preview` prop; an absent preview
 * draws as an empty grid. */
export function asSheetPreview(value: unknown): SheetPreview {
  const preview = value as Partial<SheetPreview> | undefined;
  return { sheetName: preview?.sheetName ?? "", cells: Array.isArray(preview?.cells) ? preview.cells : [] };
}

/** Narrows WorkbookBody's runtime-typed `book` prop (useWorkbookLoad's
 * return; the .vue parser bars casts). */
export function asWorkbookBook(value: unknown): ReturnType<typeof useWorkbookLoad> {
  return value as ReturnType<typeof useWorkbookLoad>;
}

/** SPEC-137: the view's "reading the workbook back from storage" flag,
 * provided to WorkbookNotFound so it shows the loading state instead. */
export const WORKBOOK_LOADING_KEY: InjectionKey<Readonly<Ref<boolean>>> = Symbol("workbook-loading");

export interface WorkbookLoadDeps {
  workbookId: ComputedRef<string>;
  router: Router;
}

export function useWorkbookLoad(deps: WorkbookLoadDeps) {
  const workbook = shallowRef<SpreadsheetModel | null>(null);
  /** The model index of the sheet on show. */
  const activeIndex = ref(0);
  /** Bumped whenever a different model object is shown (another workbook,
   * or the reader-version upgrade's re-read), so the grid re-mounts and
   * re-applies that model's view settings (SPEC-117). */
  const generation = ref(0);
  /** Bumped when a recalculation (SPEC-120) rewrote cell values: it replaces
   * the touched sheets copy-on-write inside the same model, so the sheet on
   * show re-reads without a re-mount. */
  const revision = ref(0);

  function show(model: SpreadsheetModel | null): void {
    if (model === workbook.value) {
      return;
    }
    workbook.value = model;
    activeIndex.value = model?.activeSheetIndex ?? 0;
    generation.value += 1;
  }

  /** SPEC-137: the stored workbook is being read back (loadSpreadsheet). */
  const loading = ref(false);
  /** SPEC-137: its load came back incomplete — read-only, with
   * WORKBOOK_STORAGE_INCOMPLETE_MESSAGE. */
  const storageIncomplete = ref(false);
  provide(WORKBOOK_LOADING_KEY, loading);
  /** The workbook this view made resident, released when it goes. */
  let held: string | null = null;
  let disposed = false;

  function hold(id: string | null): void {
    if (held !== null && held !== id) {
      void releaseSpreadsheet(held);
    }
    held = id;
  }

  function load(): void {
    const id = deps.workbookId.value;
    const kind = documentKindOf(id);
    if (kind !== undefined && kind !== "spreadsheet") {
      hold(null);
      void deps.router.replace(editorPathFor(kind, id));
      return;
    }
    if (kind === undefined) {
      hold(null);
      loading.value = false;
      show(null);
      return;
    }
    hold(id);
    // A resident workbook (just opened) shows at once; one read back from
    // storage shows once loadSpreadsheet resolves — which also brings a
    // record stored by an older reader up to date.
    const resident = getSpreadsheet(id);
    if (resident) {
      show(resident);
    }
    loading.value = resident === undefined;
    void loadSpreadsheet(id).then((model) => {
      if (disposed || deps.workbookId.value !== id) {
        return;
      }
      loading.value = false;
      storageIncomplete.value = isSpreadsheetStorageIncomplete(id);
      show(model ?? null);
    });
  }

  // main.ts awaits initLibrary() before mounting, so the re-sync after it
  // resolves matters only when the view is mounted directly (tests).
  load();
  void initLibrary().then(() => {
    if (!disposed) {
      load();
    }
  });
  watch(deps.workbookId, load);
  onScopeDispose(() => {
    disposed = true;
    hold(null);
  });

  const activeSheet = computed(() => {
    void revision.value;
    return workbook.value?.sheets[activeIndex.value] ?? null;
  });

  /** A recalculation pass rewrote values in the model on show. */
  function sheetsChanged(): void {
    revision.value += 1;
  }
  /** One tab per visible sheet; hidden and very-hidden sheets have none. */
  const tabs = computed((): SheetTab[] =>
    (workbook.value?.sheets ?? []).flatMap((sheet, index) =>
      sheet.visibility === "visible" ? [{ index, id: sheet.id, name: sheet.name }] : [],
    ),
  );

  function selectSheet(index = 0): void {
    if (tabs.value.some((tab) => tab.index === index)) {
      activeIndex.value = index;
    }
  }

  /** Ctrl+PageUp/PageDown: the previous/next visible sheet, wrapping. */
  function cycleSheet(delta = 1): void {
    const order = tabs.value.map((tab) => tab.index);
    if (order.length === 0) {
      return;
    }
    const at = Math.max(0, order.indexOf(activeIndex.value));
    activeIndex.value = order[(at + delta + order.length) % order.length]!;
  }

  /** Shows a re-read of the workbook on show in its place (SPEC-126's
   * import dialog "Apply"): the grid re-mounts on the new model. */
  function replace(model: SpreadsheetModel): void {
    show(model);
  }

  function download(): void {
    if (workbook.value) {
      void downloadSpreadsheet(workbook.value);
    }
  }

  function close(): void {
    // SPEC-060: the library scrolls to and highlights the row just closed.
    void deps.router.push({ path: "/", state: { closedDocId: deps.workbookId.value } });
  }

  return {
    workbook,
    activeIndex,
    generation,
    /** Bumped by sheetsChanged — every change to the model on show: a
     * command, undo or redo, a recalculation pass, a chart refresh — which
     * mutates it in place (SPEC-136's print preview and page-break lines
     * follow it). */
    revision,
    activeSheet,
    tabs,
    loading,
    storageIncomplete,
    title: computed(() => workbook.value?.title ?? (loading.value ? (getSpreadsheetShell(deps.workbookId.value)?.title ?? "") : "No document")),
    badge: computed(() => (workbook.value ? formatBadgeLabel(workbook.value.format, workbook.value.packageVariant) : "")),
    formatLabel: computed(() =>
      workbook.value ? spreadsheetFormatLabel(workbook.value.format, workbook.value.packageVariant) : "",
    ),
    foreignMessage: computed(() =>
      workbook.value && activeSheet.value ? foreignSheetMessage(workbook.value, activeSheet.value) : "",
    ),
    selectSheet,
    cycleSheet,
    sheetsChanged,
    replace,
    download,
    close,
  };
}
