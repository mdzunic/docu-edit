// SPEC-120 — useRecalc, the view's side of recalculation: the stale titles
// and the circular-reference text; the open-time policy (an idle volatile
// pass, or a full pass when the file asks or a formula has no cached
// value); the manual-mode toast, once; F9 and Data ▸ Recalculate as a full
// recalculation, F9 ignored from an input (SPEC-021's focus guard);
// "Calculating…" only past CALCULATING_STATUS_DELAY_MS; a pass started while
// another runs cancels it and inherits its unevaluated cells; SPEC-137: a
// scheduled slice's view work (the re-read, the pass's completion) in a task
// after its step's, run at once by settle() or a newer pass. The idle
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

/** A pass whose every step rewrites the model's sheets (copy-on-write, as
 * the engine's flush does) and that is done at step `steps`. */
function rewritingPass(model: SpreadsheetModel, steps: number) {
  let calls = 0;
  let cancelled = false;
  const pass: RecalcPass = {
    result: { valueChanged: [], changedSheetIds: [], circular: [], stale: [], evaluated: 0 },
    step: () => {
      calls += 1;
      model.sheets = [...model.sheets];
      return calls >= steps;
    },
    cancel() {
      cancelled = true;
      return [];
    },
  };
  return { pass, calls: () => calls, cancelled: () => cancelled };
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

describe("the calc state (SPEC-137)", () => {
  it("is calculating from a pass's start — before the status delay — and idle once it completes", () => {
    vi.useFakeTimers();
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc } = mount(model);
    drainIdle();
    expect(recalc.calcState.value).toBe("idle");
    const { pass, release } = longPass();
    passes.next = pass;
    recalc.recalculate();
    expect(recalc.calcState.value).toBe("calculating");
    expect(recalc.calculating.value).toBe(false);
    drainEager(3);
    expect(recalc.calcState.value).toBe("calculating");
    release();
    drainEager();
    expect(recalc.calcState.value).toBe("idle");
  });

  it("reads calculating during the first slice of a pass that finishes in it, and idle after", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: { formula: "A1*2", value: 99 } } });
    const { recalc } = mount(model);
    drainIdle();
    const seen: string[] = [];
    const quick: RecalcPass = {
      result: { valueChanged: [], changedSheetIds: [], circular: [], stale: [], evaluated: 0 },
      step: () => {
        seen.push(recalc.calcState.value);
        return true;
      },
      cancel: () => [],
    };
    passes.next = quick;
    recalc.recalculate();
    expect(seen).toEqual(["calculating"]);
    expect(recalc.calcState.value).toBe("idle");
    expect(scheduler.eager).toEqual([]);
  });

  it("stays calculating while a cancelled pass's successor runs, until the successor completes", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc } = mount(model);
    drainIdle();
    const first = longPass();
    passes.next = first.pass;
    recalc.recalculate();
    const second = longPass();
    passes.next = second.pass;
    recalc.recalculate();
    expect(first.pass.cancelled).toBe(true);
    first.release();
    drainEager(5);
    expect(recalc.calcState.value).toBe("calculating");
    second.release();
    drainEager();
    expect(recalc.calcState.value).toBe("idle");
  });

  it("runs a scheduled slice's step alone in its task — the re-read of what it wrote, and the pass's completion, follow in a task of their own", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc, sheetsChanged } = mount(model);
    drainIdle();
    sheetsChanged.mockClear();
    const steps = rewritingPass(model, 3);
    passes.next = steps.pass;
    const done = vi.fn();
    recalc.edit(["sheet-Sheet1!0,0"], done);
    // The first slice keeps its view work in the triggering task (SPEC-120).
    expect(steps.calls()).toBe(1);
    expect(sheetsChanged).toHaveBeenCalledTimes(1);
    drainEager(1);
    expect(steps.calls()).toBe(2);
    expect(sheetsChanged).toHaveBeenCalledTimes(1);
    drainEager(1);
    expect(sheetsChanged).toHaveBeenCalledTimes(2);
    expect(steps.calls()).toBe(2);
    drainEager(1);
    expect(steps.calls()).toBe(3);
    expect(sheetsChanged).toHaveBeenCalledTimes(2);
    expect(recalc.calcState.value).toBe("calculating");
    expect(done).not.toHaveBeenCalled();
    drainEager(1);
    expect(sheetsChanged).toHaveBeenCalledTimes(3);
    expect(done).toHaveBeenCalledTimes(1);
    expect(recalc.calcState.value).toBe("idle");
    expect(scheduler.eager).toEqual([]);
  });

  it("settle() runs a scheduled slice's pending view work now — a pass it finished completes", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc, sheetsChanged } = mount(model);
    drainIdle();
    sheetsChanged.mockClear();
    const steps = rewritingPass(model, 2);
    passes.next = steps.pass;
    const done = vi.fn();
    recalc.edit(["sheet-Sheet1!0,0"], done);
    drainEager(1);
    expect(steps.calls()).toBe(2);
    expect(done).not.toHaveBeenCalled();
    recalc.settle();
    expect(sheetsChanged).toHaveBeenCalledTimes(2);
    expect(done).toHaveBeenCalledTimes(1);
    expect(recalc.calcState.value).toBe("idle");
    drainEager();
    expect(done).toHaveBeenCalledTimes(1);
    expect(steps.calls()).toBe(2);
  });

  it("a newer pass completes a pass whose completion is pending instead of cancelling it", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } });
    const { recalc } = mount(model);
    drainIdle();
    const first = rewritingPass(model, 2);
    passes.next = first.pass;
    const firstDone = vi.fn();
    recalc.edit(["sheet-Sheet1!0,0"], firstDone);
    drainEager(1);
    const second = longPass();
    passes.next = second.pass;
    recalc.edit(["sheet-Sheet1!1,1"], vi.fn());
    expect(firstDone).toHaveBeenCalledTimes(1);
    expect(first.cancelled()).toBe(false);
    expect(passes.calls.at(-1)).toEqual(["sheet-Sheet1!1,1"]);
    expect(recalc.calcState.value).toBe("calculating");
  });

  it("covers the deferred open pass: calculating from open until it has run", () => {
    const full = mount(calcWorkbook({ Sheet1: { A1: 2, B1: "=A1*21" } }));
    expect(full.recalc.calcState.value).toBe("calculating");
    drainEager();
    expect(full.recalc.calcState.value).toBe("idle");
    const idle = mount(calcWorkbook({ Sheet1: { A1: { formula: "TODAY()", value: 45000 } } }));
    expect(idle.recalc.calcState.value).toBe("calculating");
    drainIdle();
    drainEager();
    expect(idle.recalc.calcState.value).toBe("idle");
  });
});
