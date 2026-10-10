// SPEC-125 useSheetClipboard: Copy and Cut write both flavours (TSV and the
// HTML table carrying the payload) and show the marching ants without
// touching the model; refusals toast; Paste of the view's own copy is one
// "Paste" step that keeps the ants, of its cut one "Move cells" step that
// ends them (once); Escape and any edit end the ants and cancel a cut; the
// cell editor keeps its own paste while editing; the cell menu's Paste reads
// the async clipboard (SPEC-016's snapshot rule; an unavailable clipboard
// toasts); Paste special reads it too (a refused read toasts) and pastes as
// one "Paste special" step; external HTML and plain text paste; another
// workbook's formulas naming sheets this one lacks paste as values with a
// toast. jsdom has no ClipboardEvent: the handlers get stand-in events.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { getCell } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import {
  buildClipboardCopy,
  CLIPBOARD_BLOCKED_MESSAGE,
  CLIPBOARD_UNAVAILABLE_MESSAGE,
  copyFlavours,
  MULTIPLE_SELECTIONS_MESSAGE,
  PAYLOAD_ATTRIBUTE,
  TOO_MANY_TO_CHANGE_MESSAGE,
} from "../../lib/spreadsheet/clipboard";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useSheetClipboard } from "./useSheetClipboard";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  document.body.innerHTML = "";
});

const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

/** A stand-in for ClipboardEvent's DataTransfer. */
class FakeData {
  readonly data = new Map<string, string>();
  setData(type: string, value: string): void {
    this.data.set(type, value);
  }
  getData(type: string): string {
    return this.data.get(type) ?? "";
  }
}

/** A stand-in clipboard event whose target is inside the grid. */
function clipboardEvent(data: FakeData, target: Element) {
  return { clipboardData: data, target, preventDefault: vi.fn(), stopPropagation: vi.fn() } as unknown as ClipboardEvent & { preventDefault: ReturnType<typeof vi.fn> };
}

/** An async clipboard holding text/html and text/plain (or refusing). Its
 * items' blobs are stand-ins too: jsdom's Blob has no text(). */
function asyncClipboard(flavours: { html?: string; text?: string } | "refused"): Clipboard {
  return {
    read: async () => {
      if (flavours === "refused") {
        throw new DOMException("denied", "NotAllowedError");
      }
      const types = Object.entries({ "text/html": flavours.html, "text/plain": flavours.text }).filter(([, value]) => value !== undefined);
      return [{ types: types.map(([type]) => type), getType: async (type: string) => ({ text: async () => String(Object.fromEntries(types)[type]) }) }];
    },
  } as unknown as Clipboard;
}

/** Sheet1: B2 = 1 (bold), C2 = 2, B3 = =B2*10 (10); Sheet2 empty. */
function rowsOf(): Row[] {
  return [
    { index: 1, cells: [{ col: 1, value: 1, styleId: 1 }, { col: 2, value: 2 }] },
    { index: 2, cells: [{ col: 1, value: 10, formula: "B2*10" }] },
  ];
}

function harness(options: { clipboard?: Clipboard; rows?: Row[] } = {}) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.styles = [{}, { font: { bold: true } }];
  model.sheets = [{ ...model.sheets[0]!, name: "Sheet1", rows: options.rows ?? rowsOf() }, { ...model.sheets[0]!, id: "other", name: "Sheet2", rows: [] }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => workbook.value?.sheets[activeIndex.value] ?? null);
  const selection = ref<SheetSelection>({ active: { row: 1, col: 1 }, ranges: [range(1, 1, 2, 2)], activeRangeIndex: 0 });
  const editing = ref(false);
  const toasts: string[] = [];
  const grid = document.createElement("div");
  grid.dataset.testid = "spreadsheet-grid";
  const target = document.createElement("div");
  grid.append(target);
  document.body.append(grid);
  const scope = effectScope();
  scopes.push(scope);
  const { undo, clipboard } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const transfer = useSheetClipboard({
      book: { workbook, activeSheet },
      viewState: { syncInto: vi.fn() },
      selection: {
        selection: computed(() => selection.value),
        set: (next: SheetSelection) => {
          selection.value = next;
        },
        formatContext: computed(() => ({ locale: EN_US, dateSystem: "1900" as const })),
      },
      editing: { undo: stack, cell: { editing: computed(() => editing.value), focusCell: vi.fn() } },
      locale: () => EN_US,
      toast: (message) => toasts.push(message),
      measurer: fixedMeasurer(7, 20),
      clipboard: () => options.clipboard,
    });
    return { undo: stack, clipboard: transfer };
  })!;
  const select = (r: GridRange): void => {
    selection.value = { active: { row: r.top, col: r.left }, ranges: [r], activeRangeIndex: 0 };
  };
  const cellAt = (r: number, c: number, sheet = 0) => getCell(workbook.value!.sheets[sheet]!, r, c);
  return { clipboard, undo, workbook, selection, select, editing, toasts, target, cellAt, activeIndex };
}

describe("SPEC-125 useSheetClipboard — copy and cut", () => {
  it("Copy writes TSV and the HTML table carrying the payload, shows the ants, and leaves the model untouched", () => {
    const h = harness();
    const sheets = h.workbook.value!.sheets;
    const data = new FakeData();
    const event = clipboardEvent(data, h.target);
    h.clipboard.onCopy(event);
    expect(event.preventDefault).toHaveBeenCalled();
    expect(data.getData("text/plain")).toBe("1\t2\r\n10\t");
    expect(data.getData("text/html")).toContain(`<table ${PAYLOAD_ATTRIBUTE}="`);
    expect(h.clipboard.marquee.value).toEqual({ sheetId: sheets[0]!.id, range: range(1, 1, 2, 2), mode: "copy" });
    expect(h.workbook.value!.sheets).toBe(sheets);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("refuses a multi-range selection with a toast, writing nothing", () => {
    const h = harness();
    h.selection.value = { active: { row: 1, col: 1 }, ranges: [range(1, 1, 1, 1), range(4, 4, 4, 4)], activeRangeIndex: 0 };
    const data = new FakeData();
    const event = clipboardEvent(data, h.target);
    h.clipboard.onCopy(event);
    expect(h.toasts).toEqual([MULTIPLE_SELECTIONS_MESSAGE]);
    expect(event.preventDefault).not.toHaveBeenCalled();
    expect(data.data.size).toBe(0);
  });

  it("ignores events outside the grid and while a cell is edited (the editor's own copy and paste)", () => {
    const h = harness();
    const outside = document.createElement("input");
    document.body.append(outside);
    const data = new FakeData();
    h.clipboard.onCopy(clipboardEvent(data, outside));
    expect(data.data.size).toBe(0);
    h.editing.value = true;
    h.clipboard.onCopy(clipboardEvent(data, h.target));
    const paste = new FakeData();
    paste.setData("text/plain", "typed");
    const event = clipboardEvent(paste, h.target);
    h.clipboard.onPaste(event);
    expect(data.data.size).toBe(0);
    expect(event.preventDefault).not.toHaveBeenCalled();
  });
});

describe("SPEC-125 useSheetClipboard — paste", () => {
  function copied(h: ReturnType<typeof harness>, cut = false): FakeData {
    const data = new FakeData();
    (cut ? h.clipboard.onCut : h.clipboard.onCopy)(clipboardEvent(data, h.target));
    return data;
  }

  it("pastes the view's own copy at the selection as one 'Paste' step, selects it, and keeps the ants", async () => {
    const h = harness();
    const data = copied(h);
    h.select(range(5, 4, 5, 4));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.cellAt(5, 4)).toEqual({ col: 4, value: 1, styleId: 1 });
    expect(h.cellAt(6, 4)).toEqual({ col: 4, value: 10, formula: "E6*10" });
    expect(h.undo.undoLabel.value).toBe("Paste");
    expect(h.selection.value.ranges).toEqual([range(5, 4, 6, 5)]);
    await nextTick();
    expect(h.clipboard.marquee.value?.mode).toBe("copy");
  });

  it("moves the cells of its own cut — 'Move cells', the ants gone — once; then it pastes as a copy", async () => {
    const h = harness();
    const data = copied(h, true);
    expect(h.clipboard.marquee.value?.mode).toBe("cut");
    expect(h.cellAt(1, 1)?.value).toBe(1); // a cut leaves the model untouched
    h.select(range(8, 1, 8, 1));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.undo.undoLabel.value).toBe("Move cells");
    expect(h.cellAt(1, 1)).toBeUndefined();
    expect(h.cellAt(8, 1)).toEqual({ col: 1, value: 1, styleId: 1 });
    expect(h.cellAt(9, 1)?.formula).toBe("B9*10");
    expect(h.clipboard.marquee.value).toBeNull();
    h.select(range(20, 1, 20, 1));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.undo.undoLabel.value).toBe("Paste");
    expect(h.cellAt(8, 1)?.value).toBe(1);
    expect(h.cellAt(20, 1)?.value).toBe(1);
  });

  it("a cut pasted onto another sheet records the workbook and moves its cells there", () => {
    const h = harness();
    const data = copied(h, true);
    h.activeIndex.value = 1;
    h.select(range(0, 0, 0, 0));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.cellAt(0, 0, 1)).toEqual({ col: 0, value: 1, styleId: 1 });
    expect(h.cellAt(1, 1, 0)).toBeUndefined();
  });

  it("Escape ends the ants and cancels a pending cut; so does any edit", async () => {
    const h = harness();
    const data = copied(h, true);
    // The composable listens on the window: dispatch inside the grid.
    h.target.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(h.clipboard.marquee.value).toBeNull();
    h.select(range(8, 1, 8, 1));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.undo.undoLabel.value).toBe("Paste"); // cancelled: a copy
    expect(h.cellAt(1, 1)?.value).toBe(1);
    copied(h, true);
    expect(h.clipboard.marquee.value?.mode).toBe("cut");
    h.undo.undo();
    await nextTick();
    expect(h.clipboard.marquee.value).toBeNull();
  });

  it("an external paste is an edit like any other: it ends the ants and cancels a pending cut", async () => {
    const h = harness();
    const data = copied(h, true);
    const text = new FakeData();
    text.setData("text/plain", "external");
    h.select(range(12, 0, 12, 0));
    h.clipboard.onPaste(clipboardEvent(text, h.target));
    expect(h.cellAt(12, 0)?.value).toBe("external");
    await nextTick();
    expect(h.clipboard.marquee.value).toBeNull();
    h.select(range(8, 1, 8, 1));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.undo.undoLabel.value).toBe("Paste"); // cancelled: a copy
    expect(h.cellAt(1, 1)?.value).toBe(1);
    expect(h.cellAt(8, 1)?.value).toBe(1);
  });

  it("refuses an external table whose spans pass MAX_COPY_CELLS with a toast, changing nothing", () => {
    const h = harness();
    const sheets = h.workbook.value!.sheets;
    const data = new FakeData();
    data.setData("text/html", '<table><tr><td rowspan="20000" colspan="1000">x</td></tr></table>');
    data.setData("text/plain", "x");
    h.select(range(10, 0, 10, 0));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.toasts).toEqual([TOO_MANY_TO_CHANGE_MESSAGE]);
    expect(h.workbook.value!.sheets).toBe(sheets);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("imports an external HTML table, and plain text through parseInput", () => {
    const h = harness();
    const html = new FakeData();
    html.setData("text/html", "<table><tr><td><b>Bold</b></td><td>1,234.50</td></tr></table>");
    h.select(range(10, 0, 10, 0));
    h.clipboard.onPaste(clipboardEvent(html, h.target));
    expect(h.cellAt(10, 0)?.value).toBe("Bold");
    expect(h.cellAt(10, 1)?.value).toBe(1234.5);
    const text = new FakeData();
    text.setData("text/plain", "=a1+1\tplain\r\n");
    h.select(range(12, 0, 12, 0));
    h.clipboard.onPaste(clipboardEvent(text, h.target));
    expect(h.cellAt(12, 0)?.formula).toBe("A1+1");
    expect(h.cellAt(12, 1)?.value).toBe("plain");
  });

  it("toasts another workbook's formulas that name sheets this one lacks, pasted as values", () => {
    const h = harness();
    const other = createBlankWorkbook();
    other.sheets = [{ ...other.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: 7, formula: "Data!A1" }] }] }];
    const copy = buildClipboardCopy(other, other.sheets[0]!.id, range(0, 0, 0, 0));
    const data = new FakeData();
    data.setData("text/html", copyFlavours(other, other.sheets[0]!.id, copy).html);
    h.select(range(4, 4, 4, 4));
    h.clipboard.onPaste(clipboardEvent(data, h.target));
    expect(h.cellAt(4, 4)).toEqual({ col: 4, value: 7 });
    expect(h.toasts).toEqual(["1 formula referring to other sheets was pasted as a value"]);
  });
});

describe("SPEC-125 useSheetClipboard — the cell menu and Paste special", () => {
  it("the menu's Paste toasts an unavailable clipboard", async () => {
    const h = harness({ clipboard: asyncClipboard("refused") });
    h.clipboard.menu("paste");
    await vi.waitFor(() => expect(h.toasts).toEqual([CLIPBOARD_UNAVAILABLE_MESSAGE]));
  });

  it("the menu's Paste recognises the last copy by its TSV when the HTML lost the attribute (SPEC-016's snapshot rule)", async () => {
    let flavours = { html: "", text: "" };
    const h = harness({ clipboard: { read: async () => asyncClipboard(flavours).read() } as unknown as Clipboard });
    const data = new FakeData();
    h.clipboard.onCopy(clipboardEvent(data, h.target));
    flavours = { html: "<table><tr><td>1</td><td>2</td></tr><tr><td>10</td><td></td></tr></table>", text: data.getData("text/plain") };
    h.select(range(6, 6, 6, 6));
    h.clipboard.menu("paste");
    await vi.waitFor(() => expect(h.cellAt(7, 6)?.formula).toBe("G7*10"));
    expect(h.cellAt(6, 6)?.styleId).toBe(1);
  });

  it("the menu's Copy and Cut fire the native events through execCommand", () => {
    const h = harness();
    const execCommand = vi.fn();
    Object.defineProperty(document, "execCommand", { value: execCommand, configurable: true });
    h.clipboard.menu("copy");
    h.clipboard.menu("cut");
    expect(execCommand.mock.calls).toEqual([["copy"], ["cut"]]);
    Reflect.deleteProperty(document, "execCommand");
  });

  it("Paste special: a refused read toasts; a read opens the dialog with what the source supplies, and OK pastes one 'Paste special' step", async () => {
    const refused = harness({ clipboard: asyncClipboard("refused") });
    await refused.clipboard.openPasteSpecial();
    expect(refused.toasts).toEqual([CLIPBOARD_BLOCKED_MESSAGE]);
    const h = harness({ clipboard: asyncClipboard({ text: "1\t2\r\n3\t4" }) });
    await h.clipboard.openPasteSpecial();
    expect(h.clipboard.pasteSpecial.value?.capabilities).toEqual({ formulas: false, formats: false, numberFormats: true, transpose: true });
    h.select(range(10, 10, 10, 10));
    h.clipboard.confirmPasteSpecial({ what: "values", transpose: true });
    expect(h.clipboard.pasteSpecial.value).toBeNull();
    expect(h.undo.undoLabel.value).toBe("Paste special");
    expect([h.cellAt(10, 10)?.value, h.cellAt(10, 11)?.value, h.cellAt(11, 10)?.value]).toEqual([1, 3, 2]);
  });

  it("Paste special of a pending cut pastes a copy and is an edit: the ants end and the cut is cancelled; of the view's own copy the ants stay", async () => {
    let flavours: { html?: string; text?: string } = {};
    const h = harness({ clipboard: { read: async () => asyncClipboard(flavours).read() } as unknown as Clipboard });
    const cut = new FakeData();
    h.clipboard.onCut(clipboardEvent(cut, h.target));
    flavours = { html: cut.getData("text/html"), text: cut.getData("text/plain") };
    h.select(range(20, 0, 20, 0));
    await h.clipboard.openPasteSpecial();
    h.clipboard.confirmPasteSpecial({ what: "values" });
    expect(h.undo.undoLabel.value).toBe("Paste special");
    expect([h.cellAt(20, 0)?.value, h.cellAt(1, 1)?.value]).toEqual([1, 1]);
    await nextTick();
    expect(h.clipboard.marquee.value).toBeNull();
    h.select(range(8, 1, 8, 1));
    h.clipboard.onPaste(clipboardEvent(cut, h.target));
    expect(h.undo.undoLabel.value).toBe("Paste"); // cancelled: a copy
    expect(h.cellAt(1, 1)?.value).toBe(1);

    h.select(range(1, 1, 2, 2));
    const copy = new FakeData();
    h.clipboard.onCopy(clipboardEvent(copy, h.target));
    flavours = { html: copy.getData("text/html"), text: copy.getData("text/plain") };
    h.select(range(30, 0, 30, 0));
    await h.clipboard.openPasteSpecial();
    h.clipboard.confirmPasteSpecial({ what: "values" });
    await nextTick();
    expect(h.clipboard.marquee.value?.mode).toBe("copy");
  });

  it("Ctrl+Shift+V and Ctrl+Alt+V in the grid open Paste special (jsdom is not a Mac)", async () => {
    for (const keys of [{ shiftKey: true }, { altKey: true }]) {
      const h = harness({ clipboard: asyncClipboard({ text: "x" }) });
      const event = new KeyboardEvent("keydown", { key: "V", code: "KeyV", ctrlKey: true, ...keys, bubbles: true, cancelable: true });
      h.target.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      await vi.waitFor(() => expect(h.clipboard.pasteSpecial.value).not.toBeNull());
      scopes.splice(0).forEach((scope) => scope.stop());
    }
  });
});
