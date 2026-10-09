// A workbook's view state (SPEC-117 "View state"): which sheet is on show
// and, per sheet, its selection, zoom, gridlines toggle and top-left cell.
// It is remembered across reloads but is never a change to the file:
//
// - The fields live here, keyed by sheet id and seeded from the model the
//   view loaded (the file's own view settings, or what an earlier session
//   left). A change schedules a write; writes happen at most once per
//   VIEW_STATE_PERSIST_DELAY_MS, and are flushed on close (unmount) and on
//   `pagehide`.
// - A write builds a new model whose changed sheets are NEW sheet objects
//   carrying the new view fields — sheet objects are never mutated in place,
//   and every sheet keeps its `rows` array, so the store's row split writes
//   no rows record — and hands it to storeSpreadsheetViewState, which keeps
//   the record's lastModified. Nothing here sets `dirty` (SPEC-121) or calls
//   touchDocument, and nothing touches `opaque`: an untouched workbook still
//   downloads byte-identical.
import { computed, onBeforeUnmount, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { storeSpreadsheetViewState } from "../../lib/documentStore";
import type { CellRef } from "../../lib/spreadsheet/address";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";

export const VIEW_STATE_PERSIST_DELAY_MS = 1000;

/** One sheet's view fields, in the model's own encoding: `zoomPercent`
 * absent = 100, `showGridLines` absent = shown. */
export interface SheetViewFields {
  selection?: SheetSelection;
  zoomPercent?: number;
  showGridLines?: false;
  topLeftCell?: CellRef;
}

const VIEW_KEYS = ["selection", "zoomPercent", "showGridLines", "topLeftCell"] as const;

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
  return fields;
}

/** `fields` normalised to the model's encoding (100% and shown gridlines
 * are absences). */
function normalised(fields: SheetViewFields): SheetViewFields {
  const out: SheetViewFields = { ...fields };
  if (out.zoomPercent === undefined || out.zoomPercent === 100) {
    delete out.zoomPercent;
  }
  if (out.showGridLines !== false) {
    delete out.showGridLines;
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

export interface SheetViewStateDeps {
  workbook: ShallowRef<SpreadsheetModel | null>;
  /** The model index of the sheet on show (useWorkbookLoad's). */
  activeIndex: Ref<number>;
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
    if (!pending || !model) {
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

  watch(deps.activeIndex, (next) => {
    if (deps.workbook.value && next !== storedActiveIndex) {
      schedule();
    }
  });

  const onPageHide = (): void => flush();
  if (typeof window !== "undefined") {
    window.addEventListener("pagehide", onPageHide);
  }
  onBeforeUnmount(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("pagehide", onPageHide);
    }
    flush();
  });

  return { fields, active, activeSheet, update, updateActive, flush };
}

export type SheetViewState = ReturnType<typeof useSheetViewState>;
