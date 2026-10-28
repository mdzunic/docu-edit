// A workbook's view state (SPEC-117 "View state"): which sheet is on show
// and, per sheet, its selection, zoom, gridlines toggle and top-left cell —
// and SPEC-119's View ▸ Show formulas toggle.
// It is remembered across reloads but is never a change to the file:
//
// - The fields live here, keyed by sheet id and seeded from the model the
//   view loaded (the file's own view settings, or what an earlier session
//   left). A change schedules a write; writes happen at most once per
//   VIEW_STATE_PERSIST_DELAY_MS, and are flushed on close (unmount) and when
//   the page goes away. That flush starts at `beforeunload`, as the
//   editor's autosave does, while the page still has the lifetime for the
//   store's yield-then-put to land (a write begun at `pagehide` is lost to
//   a reload); `pagehide` is the fallback where `beforeunload` never fires.
// - A write builds a new model whose changed sheets are NEW sheet objects
//   carrying the new view fields — sheet objects are never mutated in place,
//   and every sheet keeps its `rows` array, so the store's row split writes
//   no rows record — and hands it to storeSpreadsheetViewState, which keeps
//   the record's lastModified. Nothing here sets `dirty` (SPEC-121) or calls
//   touchDocument, and nothing touches `opaque`: an untouched workbook still
//   downloads byte-identical.
// - SPEC-138: no write while `writable` says no — another tab holds the
//   workbook, or an earlier version is being previewed (its model must never
//   reach the store); the change waits for the next flush. And the session's
//   `sheetView` — the tab, the active cell and the top-left cell a draft or a
//   commit records — is read from here (sheetView) and put back here
//   (applySheetView, through withSheetView: into the model as view state too,
//   so a re-seed from it agrees).
import { computed, onBeforeUnmount, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { storeSpreadsheetViewState } from "../../lib/documentStore";
import type { CellRef } from "../../lib/spreadsheet/address";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection, SheetViewState as RecordedSheetView } from "../../lib/spreadsheet/selection";

export const VIEW_STATE_PERSIST_DELAY_MS = 1000;

/** One sheet's view fields, in the model's own encoding: `zoomPercent`
 * absent = 100, `showGridLines` absent = shown, `showFormulas` absent =
 * values shown (SPEC-119). */
export interface SheetViewFields {
  selection?: SheetSelection;
  zoomPercent?: number;
  showGridLines?: false;
  topLeftCell?: CellRef;
  showFormulas?: true;
}

const VIEW_KEYS = ["selection", "zoomPercent", "showGridLines", "topLeftCell", "showFormulas"] as const;

/** A sheet's view fields as stored on the sheet. */
export function viewFieldsOf(sheet: Sheet): SheetViewFields {
  const fields: SheetViewFields = {};
  if (sheet.selection !== undefined) {
    fields.selection = sheet.selection;
  }
  if (sheet.zoomPercent !== undefined && sheet.zoomPercent !== 100) {
    fields.zoomPercent = sheet.zoomPercent;
  }
  if (sheet.showGridLines === false) {
    fields.showGridLines = false;
  }
  if (sheet.topLeftCell !== undefined) {
    fields.topLeftCell = sheet.topLeftCell;
  }
  if (sheet.showFormulas === true) {
    fields.showFormulas = true;
  }
  return fields;
}

/** `fields` normalised to the model's encoding (100%, shown gridlines and
 * shown values are absences). */
function normalised(fields: SheetViewFields): SheetViewFields {
  const out: SheetViewFields = { ...fields };
  if (out.zoomPercent === undefined || out.zoomPercent === 100) {
    delete out.zoomPercent;
  }
  if (out.showGridLines !== false) {
    delete out.showGridLines;
  }
  if (out.showFormulas !== true) {
    delete out.showFormulas;
  }
  for (const key of VIEW_KEYS) {
    if (out[key] === undefined) {
      delete out[key];
    }
  }
  return out;
}

function sameFields(a: SheetViewFields, b: SheetViewFields): boolean {
  return VIEW_KEYS.every((key) => JSON.stringify(a[key]) === JSON.stringify(b[key]));
}

/** A new sheet object with `fields` as its view fields — the same object
 * when nothing differs. `rows` (and every other field) keeps its identity. */
export function withViewFields(sheet: Sheet, fields: SheetViewFields): Sheet {
  if (sameFields(viewFieldsOf(sheet), fields)) {
    return sheet;
  }
  const next: Sheet = { ...sheet };
  for (const key of VIEW_KEYS) {
    delete next[key];
  }
  return Object.assign(next, fields);
}

/** The view fields a recorded view puts on its sheet: the selection
 * collapsed to its active cell, and its top-left cell. */
function recordedFields(base: SheetViewFields, view: RecordedSheetView): SheetViewFields {
  const { row, col } = view.active;
  return normalised({
    ...base,
    selection: { active: { row, col }, ranges: [{ top: row, left: col, bottom: row, right: col }], activeRangeIndex: 0 },
    topLeftCell: view.topRow > 0 || view.leftCol > 0 ? { row: view.topRow, col: view.leftCol } : undefined,
  });
}

/** Puts a recorded view (SPEC-138's sheetView) into `model` as view state —
 * its sheet becomes the active one, carrying the view's selection and
 * top-left cell (a new sheet object; `rows` keeps its identity). False, and
 * nothing changed, when the view names a sheet the model has no visible tab
 * for (a view naming a sheet that no longer exists is ignored). */
export function withSheetView(model: SpreadsheetModel, view: RecordedSheetView | undefined): boolean {
  const index = view ? model.sheets.findIndex((sheet) => sheet.id === view.sheetId && sheet.visibility === "visible") : -1;
  if (!view || index < 0) {
    return false;
  }
  const sheet = model.sheets[index]!;
  const next = withViewFields(sheet, recordedFields(viewFieldsOf(sheet), view));
  if (next !== sheet) {
    model.sheets = model.sheets.map((each, at) => (at === index ? next : each));
  }
  model.activeSheetIndex = index;
  return true;
}

export interface SheetViewStateDeps {
  workbook: ShallowRef<SpreadsheetModel | null>;
  /** The model index of the sheet on show (useWorkbookLoad's). */
  activeIndex: Ref<number>;
  /** SPEC-138: whether a write may reach the store now; absent = always. */
  writable?: () => boolean;
}

export function useSheetViewState(deps: SheetViewStateDeps) {
  /** sheet id → view fields. Replaced, never mutated. */
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  let timer: ReturnType<typeof setTimeout> | null = null;
  let pending = false;
  /** The active sheet index the store holds for this workbook. */
  let storedActiveIndex = 0;

  function cancel(): void {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  // A different model (another workbook, or the reader-version upgrade's
  // re-read) brings its own view settings: drop anything pending for the old
  // object — writing it would put the stale model back in the store.
  watch(
    deps.workbook,
    (model) => {
      cancel();
      pending = false;
      storedActiveIndex = model?.activeSheetIndex ?? 0;
      fields.value = Object.fromEntries((model?.sheets ?? []).map((sheet) => [sheet.id, viewFieldsOf(sheet)]));
    },
    { immediate: true },
  );

  const activeSheet = computed(() => deps.workbook.value?.sheets[deps.activeIndex.value] ?? null);
  /** The view fields of the sheet on show. */
  const active = computed((): SheetViewFields => {
    const sheet = activeSheet.value;
    return sheet ? (fields.value[sheet.id] ?? {}) : {};
  });

  /** The model a write stores: the loaded model with the current view
   * fields and active sheet. */
  function viewedModel(model: SpreadsheetModel): SpreadsheetModel {
    return {
      ...model,
      activeSheetIndex: deps.activeIndex.value,
      sheets: model.sheets.map((sheet) => withViewFields(sheet, fields.value[sheet.id] ?? {})),
    };
  }

  function flush(): void {
    cancel();
    const model = deps.workbook.value;
    if (!pending || !model || deps.writable?.() === false) {
      return;
    }
    pending = false;
    storedActiveIndex = deps.activeIndex.value;
    void storeSpreadsheetViewState(viewedModel(model));
  }

  /** Marks the view state changed: written at most once per delay window. */
  function schedule(): void {
    pending = true;
    if (timer === null) {
      timer = setTimeout(flush, VIEW_STATE_PERSIST_DELAY_MS);
    }
  }

  /** Merges `patch` into a sheet's view fields; a no-op change schedules
   * nothing. */
  function update(sheetId: string, patch: SheetViewFields): void {
    const current = fields.value[sheetId] ?? {};
    const next = normalised({ ...current, ...patch });
    if (sameFields(current, next)) {
      return;
    }
    fields.value = { ...fields.value, [sheetId]: next };
    schedule();
  }

  /** Merges `patch` into the sheet on show. */
  function updateActive(patch: SheetViewFields): void {
    const sheet = activeSheet.value;
    if (sheet) {
      update(sheet.id, patch);
    }
  }

  /** SPEC-121: writes the current view fields and active sheet into the
   * live model before a command records its undo step (so a snapshot holds
   * each sheet's selection) — new sheet objects, every `rows` kept. */
  function syncInto(model: SpreadsheetModel): void {
    model.activeSheetIndex = deps.activeIndex.value;
    model.sheets = model.sheets.map((sheet) => withViewFields(sheet, fields.value[sheet.id] ?? viewFieldsOf(sheet)));
  }

  /** SPEC-138: where the view stands, as a draft or a commit records it —
   * the sheet on show, its active cell and its top-left cell (indices). */
  function sheetView(): RecordedSheetView | undefined {
    const sheet = activeSheet.value;
    if (!sheet) {
      return undefined;
    }
    const view = fields.value[sheet.id] ?? {};
    return {
      sheetId: sheet.id,
      active: view.selection?.active ?? { row: 0, col: 0 },
      topRow: view.topLeftCell?.row ?? 0,
      leftCol: view.topLeftCell?.col ?? 0,
    };
  }

  /** SPEC-138: puts a recorded view back on the model on show (withSheetView)
   * and into the fields and the active sheet. Returns whether it applied. */
  function applySheetView(view: RecordedSheetView | undefined): boolean {
    const model = deps.workbook.value;
    if (!model || !withSheetView(model, view)) {
      return false;
    }
    const sheet = model.sheets[model.activeSheetIndex]!;
    fields.value = { ...fields.value, [sheet.id]: viewFieldsOf(sheet) };
    storedActiveIndex = model.activeSheetIndex;
    deps.activeIndex.value = model.activeSheetIndex;
    return true;
  }

  watch(deps.activeIndex, (next) => {
    if (deps.workbook.value && next !== storedActiveIndex) {
      schedule();
    }
  });

  const onLeave = (): void => flush();
  if (typeof window !== "undefined") {
    window.addEventListener("beforeunload", onLeave);
    window.addEventListener("pagehide", onLeave);
  }
  onBeforeUnmount(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("beforeunload", onLeave);
      window.removeEventListener("pagehide", onLeave);
    }
    flush();
  });

  return { fields, active, activeSheet, update, updateActive, flush, syncInto, sheetView, applySheetView };
}

export type SheetViewState = ReturnType<typeof useSheetViewState>;
