// SPEC-120 — useRecalc, the view's side of recalculation: the stale titles
// and the circular-reference text; the open-time policy (an idle volatile
// pass, or a full pass when the file asks or a formula has no cached
// value); the manual-mode toast, once; F9 and Data ▸ Recalculate as a full
// recalculation, F9 ignored from an input (SPEC-021's focus guard);
// "Calculating…" only past CALCULATING_STATUS_DELAY_MS; a pass started while
// another runs cancels it and inherits its unevaluated cells. The idle
// scheduler is replaced so slices run when the test says; the clock is
// pinned with setClock — nothing reads the wall clock.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, shallowRef } from "vue";
import { setClock } from "../../lib/core/clock";
import { calcWorkbook } from "../../lib/spreadsheet/fixtures/calcWorkbook";
import { getCell } from "../../lib/spreadsheet/address";
import type { RecalcPass } from "../../lib/spreadsheet/formula/recalc";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { useToasts } from "../../composables/useToasts";
import { CALCULATING_TEXT, circularText, isRecalculateShortcut, MANUAL_MODE_MESSAGE, staleTitle, useRecalc } from "./useRecalc";

const scheduler = vi.hoisted(() => ({ eager: [] as Array<() => void>, idle: [] as Array<() => void> }));
const passes = vi.hoisted(() => ({ next: null as RecalcPass | null, calls: [] as string[][] }));

vi.mock("../../lib/core/idleScheduler", () => ({
  scheduleEagerWork: (callback: (budget: number) => void, budget: number) => {
    let cancelled = false;
    scheduler.eager.push(() => {
      if (!cancelled) {
        callback(budget);
      }
    });
    return {
      cancel() {
        cancelled = true;
      },
    };
  },
  runWhenIdle: (callback: () => void) => {
    scheduler.idle.push(callback);
  },
}));

vi.mock("../../lib/spreadsheet/formula/recalc", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../lib/spreadsheet/formula/recalc")>();
  return {
    ...actual,
    startRecalc: (...args: Parameters<typeof actual.startRecalc>) => {
      passes.calls.push(args[1]);
      const injected = passes.next;
      passes.next = null;
      return injected ?? actual.startRecalc(...args);
    },
  };
});

/** Runs queued eager slots (and those they queue), at most `limit`. */
function drainEager(limit = 1000): void {
  for (let i = 0; i < limit && scheduler.eager.length > 0; i += 1) {
    scheduler.eager.shift()!();
  }
}

function drainIdle(): void {
  while (scheduler.idle.length > 0) {
    scheduler.idle.shift()!();
  }
}

function valueOf(model: SpreadsheetModel, row: number, col: number, sheet = 0) {
  return getCell(model.sheets[sheet]!, row, col)?.value ?? null;
}

function mount(model: SpreadsheetModel) {
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const sheetsChanged = vi.fn();
  const scope = effectScope();
  const recalc = scope.run(() => useRecalc({ book: { workbook, sheetsChanged }, locale: () => SPREADSHEET_LOCALES["en-US"] }))!;
  return { recalc, sheetsChanged, workbook, scope };
}

/** A pass that runs until released. */
function longPass(left: string[] = ["sheet-Sheet1!0,0"]) {
  let release = false;
  const pass: RecalcPass & { cancelled: boolean } = {
    cancelled: false,
    result: { valueChanged: [], changedSheetIds: [], circular: [], stale: [], evaluated: 0 },
    step: () => release,
    cancel() {
      pass.cancelled = true;
      return left;
    },
  };
  return { pass, release: () => (release = true) };
}

beforeEach(() => {
  setClock(new Date(2026, 9, 8, 9, 30));
  scheduler.eager.length = 0;
  scheduler.idle.length = 0;
  passes.calls.length = 0;
  useToasts().toasts.value = [];
});

afterEach(() => {
  setClock(null);
  vi.useRealTimers();
});

describe("the texts", () => {
  it("titles a stale cell by its reason, per the \"Stale cells\" table", () => {
    const prefix = "This value couldn't be recalculated here — ";
    expect(staleTitle({ kind: "function", name: "_XLL.CUSTOM" })).toBe(`${prefix}_XLL.CUSTOM isn't supported yet`);
    expect(staleTitle({ kind: "name", name: "Rate" })).toBe(`${prefix}named ranges aren't supported yet`);
    expect(staleTitle({ kind: "structured" })).toBe(`${prefix}table references aren't supported yet`);
    expect(staleTitle({ kind: "external" })).toBe(`${prefix}it links to another workbook`);
    expect(staleTitle({ kind: "dynamicArray" })).toBe(`${prefix}dynamic array formulas aren't supported yet`);
    expect(staleTitle({ kind: "unparsed" })).toBe(`${prefix}its formula couldn't be read`);
    expect(staleTitle({ kind: "iterative" })).toBe(`${prefix}iterative calculation isn't supported`);
    expect(staleTitle({ kind: "regex" })).toBe(`${prefix}regular-expression criteria aren't supported`);
    expect(staleTitle({ kind: "dataTable" })).toBe(`${prefix}data tables aren't supported`);
  });

  it("names a cycle's first cell with the sheet name quoted per SPEC-119", () => {
    const model = calcWorkbook({ Loop: {}, "My Sheet": {} });
    expect(circularText(model, "sheet-Loop!0,0")).toBe("Circular reference: Loop!A1");
    expect(circularText(model, "sheet-My Sheet!1,1")).toBe("Circular reference: 'My Sheet'!B2");
    expect(CALCULATING_TEXT).toBe("Calculating…");
  });

  it("takes F9 alone as the shortcut", () => {
    expect(isRecalculateShortcut({ key: "F9", ctrlKey: false, shiftKey: false, altKey: false, metaKey: false })).toBe(true);
    expect(isRecalculateShortcut({ key: "F9", ctrlKey: false, shiftKey: true, altKey: false, metaKey: false })).toBe(false);
  });
});

describe("opening a workbook", () => {
  it("paints the cache, then — once idle — recomputes only the volatile cells, and reports a cycle", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "TODAY()", value: 45000 }, B1: { formula: "1+1", value: 999 }, C1: { formula: "D1", value: 0 }, D1: { formula: "C1", value: 0 } } });
    const { recalc, sheetsChanged } = mount(model);
    expect(valueOf(model, 0, 0)).toBe(45000);
    expect(scheduler.eager).toEqual([]);
    drainIdle();
    expect(valueOf(model, 0, 0)).toBe(46303);
    expect(valueOf(model, 0, 1)).toBe(999);
    expect(sheetsChanged).toHaveBeenCalled();
    expect(recalc.circular.value).toBe("Circular reference: Sheet1!C1");
    expect(useToasts().toasts.value).toEqual([]);
  });

  it("runs a full pass when a formula has no cached value (or the file asks), never the idle one", () => {
    const model = calcWorkbook({ Sheet1: { A1: 2, B1: "=A1*21", C1: { formula: "A1+1", value: 999 } } });
    mount(model);
    expect(scheduler.idle).toEqual([]);
    drainEager();
    expect(valueOf(model, 0, 1)).toBe(42);
    expect(valueOf(model, 0, 2)).toBe(3);
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });

  it("raises the manual-mode toast once for a workbook saved to calculate manually — and recalculates it all the same", () => {
    const model = calcWorkbook({ Sheet1: { A1: "=1+1" } }, { calc: { manual: true } });
    mount(model);
    drainEager();
    expect(useToasts().toasts.value.map((toast) => toast.message)).toEqual([MANUAL_MODE_MESSAGE]);
    expect(valueOf(model, 0, 0)).toBe(2);
  });
});

describe("F9 and Data ▸ Recalculate", () => {
  it("F9 recalculates every formula cell, cached values ignored, and marks the stale ones", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: { formula: "A1*2", value: 99 }, C1: { formula: "_xll.CUSTOM(B1)", value: 42 } } });
    const { recalc } = mount(model);
    drainIdle();
    expect(valueOf(model, 0, 1)).toBe(99);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "F9", bubbles: true }));
    expect(valueOf(model, 0, 1)).toBe(2);
    expect(valueOf(model, 0, 2)).toBe(42);
    expect(recalc.staleTitleAt("sheet-Sheet1", 0, 2)).toBe("This value couldn't be recalculated here — _XLL.CUSTOM isn't supported yet");
    expect(recalc.staleTitleAt("sheet-Sheet1", 0, 1)).toBeUndefined();
  });

  it("ignores F9 typed into an input, and runs the same full pass from the menu", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: { formula: "A1*2", value: 99 } } });
    const { recalc } = mount(model);
    drainIdle();
    const input = document.createElement("input");
    document.body.append(input);
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "F9", bubbles: true }));
    input.remove();
    expect(valueOf(model, 0, 1)).toBe(99);
    recalc.recalculate();
    expect(valueOf(model, 0, 1)).toBe(2);
  });
});

describe("long passes", () => {
  it("show Calculating… only once a pass has run longer than CALCULATING_STATUS_DELAY_MS", () => {
    vi.useFakeTimers();
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc } = mount(model);
    const { pass, release } = longPass();
    passes.next = pass;
    recalc.recalculate();
    vi.advanceTimersByTime(249);
    drainEager(3);
    expect(recalc.calculating.value).toBe(false);
    expect(recalc.status.value).toEqual({ circular: "", calculating: false });
    vi.advanceTimersByTime(1);
    expect(recalc.calculating.value).toBe(true);
    release();
    drainEager();
    expect(recalc.calculating.value).toBe(false);
  });

  it("a pass started while another runs cancels it and inherits its unevaluated cells", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc } = mount(model);
    const first = longPass(["sheet-Sheet1!0,0"]);
    passes.next = first.pass;
    recalc.recalculate();
    recalc.recalculate();
    expect(first.pass.cancelled).toBe(true);
    expect(passes.calls.at(-1)).toEqual(["sheet-Sheet1!0,0"]);
  });
});
