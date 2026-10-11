// Recalculation in the workbook view (SPEC-120 "Opening a workbook", "Time
// slicing", "Stale cells", "Cycles") — lib/spreadsheet/formula/recalc.ts does
// the work; this drives it and owns what the view shows of it:
//
// - Opening: cached values paint first. A workbook whose file asks for it
//   (`calc.fullCalcOnLoad`) or that has a formula without a cached value gets
//   a full "open" pass; any other gets, once the browser is idle, one "open"
//   pass that recomputes the volatile cells and their dependents (TODAY is
//   today, as in Excel). A workbook saved in manual calculation mode is
//   recalculated automatically all the same and raises one toast; its file's
//   setting stays as it is. No open-time pass sets `dirty` (SPEC-121) or
//   touches `opaque`/`sourceXml`, so an untouched workbook still downloads
//   byte-identical.
// - Passes run in RECALC_SLICE_MS slices through scheduleEagerWork; the first
//   slice runs in the triggering task, so a pass that fits shows no flicker,
//   and "Calculating…" shows only once a pass has run longer than
//   CALCULATING_STATUS_DELAY_MS. A pass started while another runs cancels it
//   and inherits the cells it had not evaluated.
// - SPEC-121's edit pipeline drives `edit(changed, done)`: an "edit" pass
//   whose `done` receives the pass's changedSheetIds when it completes, and
//   `settle()`, which completes a running pass at once so the next command,
//   undo or redo never interleaves with it.
// - "Recalculate" (Data menu) and F9 run a full recalculation of every
//   formula cell, ignoring cached values. F9 from an input, textarea, select
//   or dialog is not a shortcut (isOutsideDocumentInput, SPEC-021), leaving
//   it to SPEC-122's cell editor.
// - The stale set is view state only — never persisted, never written to the
//   file: a stale cell's `sheet-cell` carries data-stale="true" and a title
//   naming why it couldn't be recalculated.
// - The circular status names the first cell of the first cycle ("Circular
//   reference: Loop!A1", the sheet name quoted per SPEC-119) while any cycle
//   exists — found from the graph at open too, where cycle cells keep their
//   cached values. An iterating workbook meant its cycles: no status.
import { computed, onScopeDispose, ref, shallowRef, watch, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { runWhenIdle, scheduleEagerWork, type IdleHandle } from "../../lib/core/idleScheduler";
import { isOutsideDocumentInput } from "../../lib/shortcuts";
import { formatA1 } from "../../lib/spreadsheet/address";
import { cellKey, findCycles, graphFor, parseCellKey, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import type { UncomputableReason } from "../../lib/spreadsheet/formula/evaluate";
import { quoteSheetName } from "../../lib/spreadsheet/formula/print";
import {
  CALCULATING_STATUS_DELAY_MS,
  needsFullCalcOnOpen,
  RECALC_SLICE_MS,
  startRecalc,
  uncomputableReasonOf,
  type RecalcOptions,
  type RecalcPass,
} from "../../lib/spreadsheet/formula/recalc";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";

export const MANUAL_MODE_MESSAGE = "This workbook was set to calculate manually. It recalculates automatically here.";
export const CALCULATING_TEXT = "Calculating…";
export const STALE_TITLE_PREFIX = "This value couldn't be recalculated here — ";
/** How long the idle-time open pass may wait under sustained load. */
export const OPEN_PASS_IDLE_TIMEOUT_MS = 1000;

/** A stale cell's title (SPEC-120 "Stale cells"). */
export function staleTitle(reason: UncomputableReason): string {
  const suffix = ((): string => {
    switch (reason.kind) {
      case "function":
        return `${reason.name} isn't supported yet`;
      case "name":
        return "named ranges aren't supported yet";
      case "structured":
        return "table references aren't supported yet";
      case "external":
        return "it links to another workbook";
      case "dynamicArray":
        return "dynamic array formulas aren't supported yet";
      case "unparsed":
        return "its formula couldn't be read";
      case "iterative":
        return "iterative calculation isn't supported";
      case "regex":
        return "regular-expression criteria aren't supported";
      case "unsupportedCase":
        return `this use of ${reason.name} isn't supported yet`;
      default:
        return "data tables aren't supported";
    }
  })();
  return STALE_TITLE_PREFIX + suffix;
}

/** "Circular reference: Loop!A1" for a cycle's first cell. */
export function circularText(model: SpreadsheetModel, key: CellKey): string {
  const { sheetId, row, col } = parseCellKey(key);
  const sheet = model.sheets.find((candidate) => candidate.id === sheetId);
  return `Circular reference: ${quoteSheetName(sheet?.name ?? "")}!${formatA1({ row, col })}`;
}

/** F9 alone: a full recalculation. */
export function isRecalculateShortcut(event: Pick<KeyboardEvent, "key" | "ctrlKey" | "shiftKey" | "altKey" | "metaKey">): boolean {
  return event.key === "F9" && !event.ctrlKey && !event.shiftKey && !event.altKey && !event.metaKey;
}

export interface RecalcDeps {
  /** useWorkbookLoad's: the model on show and the hook that has the view
   * re-read its sheets after a pass rewrote values in place. */
  book: { workbook: ShallowRef<SpreadsheetModel | null>; sheetsChanged(): void };
  /** The user's number locale (text coerces to numbers in it). */
  locale: () => SpreadsheetLocale;
}

export function useRecalc(deps: RecalcDeps) {
  const { addToast } = useToasts();
  /** The circular-reference status text ("" when there is no cycle). */
  const circular = ref("");
  const calculating = ref(false);
  /** Stale cells and why — view state only. */
  const stale = shallowRef<ReadonlyMap<CellKey, UncomputableReason>>(new Map());
  let pass: RecalcPass | null = null;
  let passModel: SpreadsheetModel | null = null;
  /** SPEC-121: the edit pipeline's callback for the running pass. */
  let passDone: ((changedSheetIds: string[]) => void) | null = null;
  /** SPEC-129: the running pass's own keys, and who hears each result. */
  let passKeys: CellKey[] = [];
  const resultListeners: ((changed: CellKey[]) => void)[] = [];
  let slot: IdleHandle | null = null;
  let statusTimer: ReturnType<typeof setTimeout> | null = null;

  function stopDriving(): void {
    slot?.cancel();
    slot = null;
    if (statusTimer !== null) {
      clearTimeout(statusTimer);
      statusTimer = null;
    }
    calculating.value = false;
  }

  /** The cycles the model's graph holds, as the status shows them. */
  function reportCycles(model: SpreadsheetModel): void {
    const cycles = model.calc?.iterate ? [] : findCycles(graphFor(model));
    circular.value = cycles.length > 0 ? circularText(model, cycles[0]![0]!) : "";
  }

  function finish(current: RecalcPass, model: SpreadsheetModel): void {
    pass = null;
    passModel = null;
    stopDriving();
    if (current.result.stale.length > 0) {
      const next = new Map(stale.value);
      for (const key of current.result.stale) {
        next.set(key, uncomputableReasonOf(model, key) ?? { kind: "unparsed" });
      }
      stale.value = next;
    }
    reportCycles(model);
    const changed = [...passKeys, ...current.result.valueChanged];
    passKeys = [];
    for (const listener of resultListeners) {
      listener(changed);
    }
    const done = passDone;
    passDone = null;
    done?.(current.result.changedSheetIds);
  }

  /** One slice; true when the pass is done. */
  function slice(current: RecalcPass, model: SpreadsheetModel): boolean {
    const sheets = model.sheets;
    let done: boolean;
    try {
      done = current.step(RECALC_SLICE_MS);
    } catch (error) {
      // A defect in the engine must not take the view down: stop the pass
      // and keep what is on screen.
      console.error("Recalculation failed", error);
      current.cancel();
      pass = null;
      passModel = null;
      stopDriving();
      const callback = passDone;
      passDone = null;
      callback?.([]);
      return true;
    }
    if (model.sheets !== sheets && deps.book.workbook.value === model) {
      deps.book.sheetsChanged();
    }
    if (done) {
      finish(current, model);
    }
    return done;
  }

  function drive(current: RecalcPass, model: SpreadsheetModel): void {
    slot = scheduleEagerWork(() => {
      slot = null;
      if (pass === current && !slice(current, model)) {
        drive(current, model);
      }
    }, RECALC_SLICE_MS);
  }

  /** Starts a pass — cancelling a running one, whose unevaluated cells it
   * inherits (and whose edit callback it carries) — and runs its first
   * slice now. */
  function run(changed: CellKey[], options: RecalcOptions, done?: (changedSheetIds: string[]) => void): void {
    const model = deps.book.workbook.value;
    if (!model) {
      done?.([]);
      return;
    }
    let inherited: CellKey[] = [];
    let carried: ((changedSheetIds: string[]) => void) | null = null;
    if (pass && passModel === model) {
      inherited = pass.cancel();
      carried = passDone;
    }
    stopDriving();
    passKeys = [...new Set([...inherited, ...changed])];
    const current = startRecalc(model, passKeys, { ...options, settings: { locale: deps.locale() } });
    pass = current;
    passModel = model;
    passDone =
      carried && done
        ? (ids) => {
            carried(ids);
            done(ids);
          }
        : (done ?? carried);
    if (slice(current, model)) {
      return;
    }
    statusTimer = setTimeout(() => {
      statusTimer = null;
      if (pass === current) {
        calculating.value = true;
      }
    }, CALCULATING_STATUS_DELAY_MS);
    drive(current, model);
  }

  /** Data ▸ Recalculate and F9: every formula cell, cached values ignored. */
  function recalculate(): void {
    run([], { reason: "manual", full: true });
  }

  /** SPEC-121's edit pipeline: a pass of reason "edit" over `changed`
   * (possibly sliced) that calls `done` with its changedSheetIds when it
   * completes. */
  function edit(changed: CellKey[], done: (changedSheetIds: string[]) => void): void {
    run(changed, { reason: "edit" }, done);
  }

  /** Completes a running pass now — the next command, undo or redo waits
   * for it, so steps never interleave. */
  function settle(): void {
    const current = pass;
    const model = passModel;
    if (!current || !model) {
      return;
    }
    stopDriving();
    while (pass === current && !slice(current, model)) {
      // each slice ends at a flush; keep going
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !isRecalculateShortcut(event) || isOutsideDocumentInput(event.target) || !deps.book.workbook.value) {
      return;
    }
    event.preventDefault();
    recalculate();
  }

  /** The open-time policy (see the module header). */
  function open(model: SpreadsheetModel): void {
    if (needsFullCalcOnOpen(model)) {
      slot = scheduleEagerWork(() => {
        slot = null;
        if (deps.book.workbook.value === model) {
          run([], { reason: "open", full: true });
        }
      }, RECALC_SLICE_MS);
    } else {
      runWhenIdle(() => {
        if (deps.book.workbook.value === model && pass === null) {
          run([], { reason: "open" });
        }
      }, OPEN_PASS_IDLE_TIMEOUT_MS);
    }
    if (model.calc?.manual) {
      addToast(MANUAL_MODE_MESSAGE);
    }
  }

  watch(
    () => deps.book.workbook.value,
    (model, previous) => {
      if (model === previous) {
        return;
      }
      pass?.cancel();
      pass = null;
      passModel = null;
      passDone = null;
      stopDriving();
      stale.value = new Map();
      circular.value = "";
      if (model) {
        open(model);
      }
    },
    { immediate: true },
  );

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown);
  }
  onScopeDispose(() => {
    pass?.cancel();
    pass = null;
    stopDriving();
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown);
    }
  });

  /** A stale cell's title on a sheet, if it is stale. */
  function staleTitleAt(sheetId: string, row: number, col: number): string | undefined {
    const reason = stale.value.get(cellKey(sheetId, row, col));
    return reason ? staleTitle(reason) : undefined;
  }

  /** SPEC-129: calls `listener` with every completed pass's changed keys —
   * the cells it was started on and those whose value it changed. */
  function onResult(listener: (changed: CellKey[]) => void): void {
    resultListeners.push(listener);
  }

  return {
    circular,
    calculating,
    stale,
    onResult,
    /** SpreadsheetStatusBar's recalculation props, bound as one object. */
    status: computed(() => ({ circular: circular.value, calculating: calculating.value })),
    staleTitleAt,
    recalculate,
    edit,
    settle,
    onKeydown,
  };
}

export type Recalc = ReturnType<typeof useRecalc>;
