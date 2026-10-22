// SPEC-136's usePageSetup, useSheetPrint and useSheetPdfExport over a real
// undo stack: the dialog's OK is one "Page setup" step, OK on an unchanged
// form and Cancel nothing; each break and print-area command one labelled
// step; the preview, its scope and range and a refused export never write
// the model; the preview prints the sheet on show and follows every change
// to the model, as the page-break lines do; Stop ends an export with no
// download, and exported pages print their job's total.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { Cell, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { MAX_SHEET_PDF_PAGES, PDF_LIMIT_MESSAGE } from "../../lib/spreadsheet/print";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook, emptySheet } from "../../lib/spreadsheet/sheets";
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

function harness(at: GridRange = { top: 29, left: 0, bottom: 29, right: 0 }, others: Sheet[] = []) {
  const model: SpreadsheetModel = createBlankWorkbook();
  const rows: Row[] = Array.from({ length: 70 }, (_, index) => ({ index, cells: [{ col: 0, value: index }] }));
  model.sheets = [{ ...model.sheets[0]!, rows }, ...others];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  // useWorkbookLoad's book: the sheet on show is the view's activeIndex (a
  // tab switch writes nothing to the model), re-read on each revision.
  const activeIndex = ref(0);
  const revision = ref(0);
  const activeSheet = computed(() => {
    void revision.value;
    return workbook.value?.sheets[activeIndex.value] ?? null;
  });
  const selection = ref<SheetSelection>({ active: { row: at.top, col: at.left }, ranges: [at], activeRangeIndex: 0 });
  const scope = effectScope();
  scopes.push(scope);
  const parts = scope.run(() => {
    // useWorkbookEditing's onApplied: triggerRef, then sheetsChanged.
    const onApplied = () => {
      triggerRef(workbook);
      revision.value += 1;
    };
    const undo = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied, activateSheet: () => {} });
    const book = { workbook, activeIndex, revision, activeSheet } as unknown as Parameters<typeof usePageSetup>[0]["book"];
    const pageSetup = usePageSetup({ book, viewState: { syncInto: vi.fn() }, selection: { selection: computed(() => selection.value) }, editing: { undo, cell: { editing: computed(() => false) } } as never });
    const print = useSheetPrint({ book, selection: { selection: computed(() => selection.value) }, pageSetup, locale: () => SPREADSHEET_LOCALES["en-US"] });
    const pdf = useSheetPdfExport({ book, print });
    return { undo, pageSetup, print, pdf };
  })!;
  const sheet = () => workbook.value!.sheets[0]!;
  /** One command writing `cells` into the first sheet's rows. */
  const write = (cells: Record<number, Cell[]>) => {
    parts.undo.record("Typing", { sheetId: sheet().id });
    const written = new Set(Object.keys(cells).map(Number));
    const kept = sheet().rows.filter((row) => !written.has(row.index));
    const added = [...written].map((index) => ({ index, cells: cells[index]! }));
    workbook.value!.sheets = [{ ...sheet(), rows: [...kept, ...added].toSorted((a, b) => a.index - b.index) }, ...workbook.value!.sheets.slice(1)];
    parts.undo.commit();
  };
  return { ...parts, sheet, workbook, activeIndex, revision, write };
}

/** The idle callback the page-break lines recompute on (setTimeout under jsdom). */
async function idle(): Promise<void> {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));
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

  it("OK on the form as shown writes nothing: no step, no paper pinned on a sheet that names none", () => {
    const h = harness();
    const sheet = h.sheet();
    h.pageSetup.dialog.show();
    h.pageSetup.dialog.ok();
    expect(h.pageSetup.dialog.open.value).toBe(false);
    expect(h.sheet()).toBe(sheet);
    expect(h.sheet().pageSetup).toBeUndefined();
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("page-break lines recompute after every model change: a print area (a name — the sheet object untouched), a cell, a recalculation", async () => {
    const h = harness();
    const rowBreaks = () => h.pageSetup.pageBreaks.lines.value.filter((line) => line.axis === "row").map((line) => line.index);
    h.pageSetup.pageBreaks.toggle();
    await idle();
    const breaks = rowBreaks();
    expect(breaks).toHaveLength(1);
    const sheet = h.sheet();
    h.pageSetup.setArea(); // A30 alone: one page
    expect(h.sheet()).toBe(sheet);
    await idle();
    expect(rowBreaks()).toEqual([]);
    h.pageSetup.clearArea();
    await idle();
    expect(rowBreaks()).toEqual(breaks);
    h.write({ 299: [{ col: 0, value: "far below" }] });
    await idle();
    expect(rowBreaks().length).toBeGreaterThan(2);
    // A recalculation pass rewrites the sheet in place and says so through sheetsChanged alone.
    h.workbook.value!.sheets = [{ ...h.sheet(), rows: h.sheet().rows.filter((row) => row.index !== 299) }];
    h.revision.value += 1;
    await idle();
    expect(rowBreaks()).toEqual(breaks);
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

  it("the preview prints the sheet on show — a tab switch, which writes nothing to the model — and an empty one has nothing to print", () => {
    const h = harness(undefined, [emptySheet("Notes")]);
    const first = h.sheet();
    h.print.open();
    expect(new Set(h.print.job.value.pages.map((page) => page.sheetId))).toEqual(new Set([first.id]));
    h.print.close();
    h.activeIndex.value = 1;
    expect(h.workbook.value!.activeSheetIndex).toBe(0);
    h.print.open();
    expect(h.print.pageCount.value).toBe(0);
    expect(h.print.canPrint.value).toBe(false);
    expect(h.print.emptyMessage).toBe("This sheet has nothing to print.");
    // The download menu's PDF prints the same sheet.
    expect(h.print.jobFor("sheet").pages).toEqual([]);
    // A switch with the preview open shows the other sheet's pages.
    h.activeIndex.value = 0;
    expect(h.print.pageCount.value).toBe(2);
  });

  it("pages laid out before a change show it: a command's cell and its link, a recalculation's value", () => {
    const h = harness();
    h.print.open();
    const a1 = () => h.print.layoutFor(h.print.chosen.value[0]!, h.print.pageCount.value).cells.find((cell) => cell.ref === "A1" && cell.heading === undefined);
    expect(a1()?.content).toBe("0");
    h.write({ 0: [{ col: 0, value: "Example", formula: 'HYPERLINK("https://example.com","Example")' }] });
    expect(a1()).toMatchObject({ content: "Example", href: "https://example.com" });
    // A recalculation pass rewrites values in place and says so through sheetsChanged alone.
    h.workbook.value!.sheets = [{ ...h.sheet(), rows: h.sheet().rows.map((row) => (row.index === 0 ? { ...row, cells: [{ col: 0, value: "Recalculated" }] } : row)) }];
    h.revision.value += 1;
    expect(a1()?.content).toBe("Recalculated");
  });

  it("an export of part of the job prints the job's total, as the preview does", async () => {
    const h = harness();
    h.print.open();
    h.print.rangeMode.value = "custom";
    h.print.rangeText.value = "2";
    const running = h.pdf.exportPdf();
    expect(h.pdf.progress.value).toEqual({ done: 0, total: 1 });
    expect(h.pdf.jobTotal.value).toBe(2);
    expect(h.pdf.sheetPages.value.map((each) => each.index)).toEqual([1]);
    h.pdf.abort();
    await running;
    expect(h.pdf.jobTotal.value).toBe(0);
    expect(triggerBlobDownload).not.toHaveBeenCalled();
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
