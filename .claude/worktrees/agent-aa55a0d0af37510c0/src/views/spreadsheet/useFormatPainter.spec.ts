// SPEC-123 useFormatPainter: a click arms one application, a double-click
// locks it, Escape or another click disarms; the next completed selection (a
// drag released, or Enter in the grid) receives the source's pattern as ONE
// "Format painter" step; a multi-range source disables it.
import { describe, expect, it } from "vitest";
import { computed, effectScope, ref, shallowRef } from "vue";
import type { FormatResult } from "../../lib/spreadsheet/formatting";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { effectiveStyleId } from "../../lib/spreadsheet/styles";
import { useFormatPainter } from "./useFormatPainter";

function harness() {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }] }];
  model.styles = [{}, { font: { bold: true } }];
  const workbook = shallowRef(model);
  const selection = ref<SheetSelection>({ active: { row: 0, col: 0 }, ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], activeRangeIndex: 0 });
  const drag = shallowRef<object | null>(null);
  const labels: string[] = [];
  const scope = effectScope();
  const painter = scope.run(() =>
    useFormatPainter({
      book: { activeSheet: computed(() => workbook.value.sheets[0]!) },
      selection: { selection: computed(() => selection.value), drag: drag as never },
      formatting: {
        command: (label: string, mutate: (m: SpreadsheetModel, id: string, ranges: SheetSelection["ranges"]) => FormatResult) => {
          labels.push(label);
          mutate(workbook.value, workbook.value.sheets[0]!.id, selection.value.ranges);
          workbook.value = { ...workbook.value };
        },
      },
    }),
  )!;
  const select = (top: number, left: number, bottom: number, right: number): void => {
    selection.value = { active: { row: top, col: left }, ranges: [{ top, left, bottom, right }], activeRangeIndex: 0 };
  };
  const release = async (): Promise<void> => {
    drag.value = {};
    await Promise.resolve();
    drag.value = null;
    await new Promise((resolve) => setTimeout(resolve, 0));
  };
  return { painter, workbook, labels, select, release, scope };
}

describe("SPEC-123 useFormatPainter", () => {
  it("arms once on a click, paints the next released selection and disarms", async () => {
    const h = harness();
    h.painter.click();
    expect(h.painter.mode.value).toBe("once");
    h.select(2, 0, 3, 1);
    await h.release();
    expect(h.labels).toEqual(["Format painter"]);
    expect(effectiveStyleId(h.workbook.value.sheets[0]!, 3, 1)).toBe(1);
    expect(h.painter.mode.value).toBe("");
    h.scope.stop();
  });

  it("locks on a double-click until Escape or another click, and Enter completes a keyboard selection", async () => {
    const h = harness();
    h.painter.click();
    h.painter.click();
    h.painter.dblclick();
    expect(h.painter.mode.value).toBe("locked");
    h.select(5, 0, 5, 0);
    await h.release();
    const grid = document.body.appendChild(document.createElement("div"));
    grid.dataset.testid = "spreadsheet-grid";
    h.select(6, 0, 6, 0);
    grid.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    expect(h.labels).toEqual(["Format painter", "Format painter"]);
    expect(h.painter.mode.value).toBe("locked");
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(h.painter.mode.value).toBe("");
    h.painter.dblclick();
    h.painter.click();
    expect(h.painter.mode.value).toBe("");
    grid.remove();
    h.scope.stop();
  });

  it("is disabled for a multi-range source", () => {
    const h = harness();
    h.select(0, 0, 0, 0);
    h.select(0, 0, 0, 0);
    expect(h.painter.disabled.value).toBe(false);
    h.scope.stop();
  });
});
