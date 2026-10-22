// SPEC-136's usePageSetup, useSheetPrint and useSheetPdfExport over a real
// undo stack: the dialog's OK is one "Page setup" step and Cancel nothing;
// each break and print-area command one labelled step; the preview, its
// scope and range and a refused export never write the model; Stop ends an
// export with no download.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { MAX_SHEET_PDF_PAGES, PDF_LIMIT_MESSAGE } from "../../lib/spreadsheet/print";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { usePageSetup } from "./usePageSetup";
import { useSheetPdfExport } from "./useSheetPdfExport";
import { useSheetPrint, type PlacedPage } from "./useSheetPrint";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const triggerBlobDownload = vi.hoisted(() => vi.fn());
vi.mock("../../lib/core/triggerDownload", () => ({ triggerBlobDownload }));

const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  triggerBlobDownload.mockClear();
});

function harness(at: GridRange = { top: 29, left: 0, bottom: 29, right: 0 }) {
  const model: SpreadsheetModel = createBlankWorkbook();
  const rows: Row[] = Array.from({ length: 70 }, (_, index) => ({ index, cells: [{ col: 0, value: index }] }));
  model.sheets = [{ ...model.sheets[0]!, rows }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>({ active: { row: at.top, col: at.left }, ranges: [at], activeRangeIndex: 0 });
  const scope = effectScope();
  scopes.push(scope);
  const parts = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const book = { workbook, activeSheet } as unknown as Parameters<typeof usePageSetup>[0]["book"];
    const pageSetup = usePageSetup({ book, viewState: { syncInto: vi.fn() }, selection: { selection: computed(() => selection.value) }, editing: { undo, cell: { editing: computed(() => false) } } as never });
    const print = useSheetPrint({ book, selection: { selection: computed(() => selection.value) }, pageSetup, locale: () => SPREADSHEET_LOCALES["en-US"] });
    const pdf = useSheetPdfExport({ book, print });
    return { undo, pageSetup, print, pdf };
  })!;
  const sheet = () => workbook.value!.sheets[0]!;
  return { ...parts, sheet, workbook };
}

describe("usePageSetup", () => {
  it("the dialog's OK is one 'Page setup' step; Cancel changes nothing", () => {
    const h = harness();
    h.pageSetup.dialog.show();
    h.pageSetup.dialog.form.value!.orientation = "landscape";
    h.pageSetup.dialog.cancel();
    expect(h.sheet().pageSetup).toBeUndefined();
    expect(h.undo.canUndo.value).toBe(false);
    h.pageSetup.dialog.show();
    const form = h.pageSetup.dialog.form.value!;
    form.orientation = "landscape";
    form.scale = "500";
    expect(h.pageSetup.dialog.problem.value?.message).toBe("Enter a scale between 10% and 400%");
    form.scale = "80";
    form.margins.top = "40";
    expect(h.pageSetup.dialog.problem.value?.message).toBe("Margins are too large for this page size");
    form.margins.top = "2";
    form.header.center = "Page &P";
    h.pageSetup.dialog.ok();
    expect(h.sheet().pageSetup).toMatchObject({ orientation: "landscape", scale: 80, headerFooter: { odd: { header: { center: "Page &P" } } } });
    expect(h.undo.undoLabel.value).toBe("Page setup");
    h.undo.undo();
    expect(h.sheet().pageSetup).toBeUndefined();
  });

  it("Insert / Remove / Reset page breaks and Set / Clear print area are one labelled step each", () => {
    const h = harness();
    h.pageSetup.insertBreak();
    expect([h.sheet().pageSetup?.rowBreaks, h.undo.undoLabel.value]).toEqual([[29], "Insert page break"]);
    h.pageSetup.removeBreak();
    expect([h.sheet().pageSetup, h.undo.undoLabel.value]).toEqual([undefined, "Remove page break"]);
    h.pageSetup.insertBreak();
    h.pageSetup.resetBreaks();
    expect(h.undo.undoLabel.value).toBe("Reset page breaks");
    h.pageSetup.setArea();
    expect([h.workbook.value!.names.at(-1)?.formula, h.undo.undoLabel.value]).toEqual(["Sheet1!$A$30", "Set print area"]);
    h.pageSetup.clearArea();
    expect([h.workbook.value!.names, h.undo.undoLabel.value]).toEqual([[], "Clear print area"]);
  });
});

describe("useSheetPrint and useSheetPdfExport", () => {
  it("the preview's job, scope and range never write the model; closing it turns page-break lines on", () => {
    const h = harness();
    const model = h.workbook.value!;
    const sheets = model.sheets;
    h.print.open();
    expect(h.print.pageCount.value).toBe(2);
    h.print.rangeMode.value = "custom";
    h.print.rangeText.value = "3";
    expect(h.print.rangeError.value).toBe("Pages go from 1 to 2");
    h.print.rangeText.value = "2";
    expect(h.print.chosen.value.map((each) => each.index)).toEqual([1]);
    h.print.setScope("workbook");
    expect(h.print.pageCount.value).toBe(2);
    h.print.close();
    expect(h.pageSetup.pageBreaks.shown.value).toBe(true);
    expect(model.sheets).toBe(sheets);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("an export over 2,000 pages is refused; Stop ends one at the next page with no download", async () => {
    const h = harness();
    const page = h.print.jobFor("sheet").pages[0]!;
    const many: PlacedPage[] = Array.from({ length: MAX_SHEET_PDF_PAGES + 1 }, (_, index) => ({ page, index }));
    await h.pdf.exportPdf(many);
    expect(h.pdf.notice.value).toBe(PDF_LIMIT_MESSAGE);
    expect(h.pdf.busy.value).toBe(false);
    const running = h.pdf.exportPdf([{ page, index: 0 }]);
    expect(h.pdf.busy.value).toBe(true);
    expect(h.pdf.progress.value).toEqual({ done: 0, total: 1 });
    h.pdf.abort();
    await running;
    expect(triggerBlobDownload).not.toHaveBeenCalled();
    expect(h.pdf.busy.value).toBe(false);
    expect(h.undo.canUndo.value).toBe(false);
  });
});
