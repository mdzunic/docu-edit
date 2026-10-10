// SPEC-124 useHeaderResize: the boundary hit-test within RESIZE_HIT_PX (by
// zoom), the cursor, the drag's guide and its "Width: 8.43 (64 pixels)" /
// "Height: 15.00 (20 pixels)" tooltip, Escape cancelling, the release as one
// resize of the dragged column — or of every selected one when it belongs to
// a whole-column selection — and a double-click's autofit.
import { afterEach, describe, expect, it } from "vitest";
import { computed, effectScope, ref, shallowRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { Band, BandAxis } from "../../lib/spreadsheet/bands";
import { buildGridMetrics } from "../../lib/spreadsheet/geometry";
import { MAX_COLS, MAX_ROWS, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { RESIZE_HIT_PX, useHeaderResize } from "./useHeaderResize";

const scopes: EffectScope[] = [];
afterEach(() => scopes.splice(0).forEach((scope) => scope.stop()));

function harness(ranges: GridRange[] = [{ top: 0, left: 0, bottom: 0, right: 0 }], zoomPercent = 100) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, columns: [{ min: 1, max: 1, widthPx: 100 }] }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>({ active: { row: ranges[0]!.top, col: ranges[0]!.left }, ranges, activeRangeIndex: 0 });
  const resized: [BandAxis, Band[], number][] = [];
  const autofitted: [BandAxis, Band[]][] = [];
  const scope = effectScope();
  scopes.push(scope);
  const resize = scope.run(() =>
    useHeaderResize({
      book: { workbook, activeSheet },
      structure: {
        resizeBands: (axis, bands, px) => resized.push([axis, [...bands], px]),
        autofitBands: (axis, bands) => autofitted.push([axis, [...bands]]),
      },
      selection: { selection: computed(() => selection.value) },
      sheetMetrics: { metrics: computed(() => buildGridMetrics(workbook.value!.sheets[0]!)) },
      zoom: { percent: computed(() => zoomPercent) },
    }),
  )!;
  return { resize, resized, autofitted, selection };
}

describe("SPEC-124 useHeaderResize", () => {
  it("finds the boundary within 3 screen px of a column's or row's trailing edge", () => {
    const { resize } = harness();
    // A is 0–64, B (100 px) 64–164.
    expect(resize.boundaryAt("col", 64)).toBe(0);
    expect(resize.boundaryAt("col", 64 - RESIZE_HIT_PX)).toBe(0);
    expect(resize.boundaryAt("col", 64 + RESIZE_HIT_PX)).toBe(0);
    expect(resize.boundaryAt("col", 70)).toBeNull();
    expect(resize.boundaryAt("col", 163)).toBe(1);
    expect(resize.boundaryAt("row", 20)).toBe(0);
    expect(resize.boundaryAt("row", 30)).toBeNull();
    // At 200% zoom 3 screen px are 1.5 content px.
    expect(harness(undefined, 200).resize.boundaryAt("col", 66)).toBeNull();
  });

  it("shows the resize cursor while hovering a boundary and while dragging", () => {
    const { resize } = harness();
    expect(resize.cursor.value).toBeUndefined();
    resize.setHover("col");
    expect(resize.cursor.value).toBe("col-resize");
    resize.setHover(null);
    resize.begin("row", 0, 20, 20, 10);
    expect(resize.cursor.value).toBe("row-resize");
  });

  it("drags with a guide on the boundary and the size tooltip, and commits one resize on release", () => {
    const { resize, resized } = harness();
    resize.begin("col", 0, 64, 64, 10);
    expect(resize.tooltip.value?.text).toBe("Width: 8.43 (64 pixels)");
    resize.move(100, 100, 12);
    expect(resize.guide.value).toEqual({ axis: "col", at: 100 });
    expect(resize.tooltip.value?.text).toBe("Width: 13.57 (100 pixels)");
    resize.finish();
    expect(resized).toEqual([["col", [{ start: 0, end: 0 }], 100]]);
    expect(resize.guide.value).toBeNull();
    expect(resize.active.value).toBe(false);
  });

  it("names a row's height in points", () => {
    const { resize } = harness();
    resize.begin("row", 0, 20, 20, 5);
    expect(resize.tooltip.value?.text).toBe("Height: 15.00 (20 pixels)");
    resize.move(30, 30, 5);
    expect(resize.tooltip.value?.text).toBe("Height: 22.50 (30 pixels)");
  });

  it("keeps the guide at the leading edge once the size reaches 0 — a release under 4 px is a resize the structure turns into a hide", () => {
    const { resize, resized } = harness();
    resize.begin("col", 1, 164, 164, 10);
    resize.move(40, 40, 10);
    expect(resize.guide.value).toEqual({ axis: "col", at: 64 });
    resize.finish();
    expect(resized).toEqual([["col", [{ start: 1, end: 1 }], 0]]);
  });

  it("Escape cancels the drag; a release that changes nothing records nothing", () => {
    const { resize, resized } = harness();
    resize.begin("col", 0, 64, 64, 10);
    resize.move(90, 90, 10);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));
    expect(resize.guide.value).toBeNull();
    resize.finish();
    resize.begin("col", 0, 64, 64, 10);
    resize.finish();
    expect(resized).toEqual([]);
  });

  it("resizes and autofits every selected column when the dragged header belongs to a whole-column selection", () => {
    const columns = { top: 0, bottom: MAX_ROWS - 1 };
    const { resize, resized, autofitted } = harness([
      { ...columns, left: 0, right: 2 },
      { ...columns, left: 5, right: 5 },
    ]);
    resize.begin("col", 1, 164, 164, 10);
    resize.move(184, 184, 10);
    resize.finish();
    expect(resized).toEqual([["col", [{ start: 0, end: 2 }, { start: 5, end: 5 }], 120]]);
    // Outside the selection only the dragged column.
    resize.autofit("col", 7);
    resize.autofit("col", 2);
    expect(autofitted).toEqual([
      ["col", [{ start: 7, end: 7 }]],
      ["col", [{ start: 0, end: 2 }, { start: 5, end: 5 }]],
    ]);
    expect(MAX_COLS).toBeGreaterThan(5);
  });
});
