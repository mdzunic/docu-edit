// SPEC-122 "Editor DOM": the in-cell editor — always mounted (transparent
// and empty in Ready, named "B3, Revenue"), the first typed character
// landing once through the browser's native insertion, normalise-after-input
// re-rendering one coloured span per reference, nothing re-rendered while an
// IME composition runs, paste as plain text, the formula error under it, and
// a double-click's caret at the clicked character.
import { afterEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { computed, defineComponent, h, nextTick, ref, shallowRef } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { cellEditorView } from "../../views/spreadsheet/cellEditorView";
import { FORMULA_ERROR_MESSAGE } from "../../views/spreadsheet/cellEntry";
import { useCellEditing, type CellEditing } from "../../views/spreadsheet/useCellEditing";
import { useSheetMetrics } from "../../views/spreadsheet/useSheetMetrics";
import { useSheetSelection } from "../../views/spreadsheet/useSheetSelection";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import type { SheetViewFields } from "../../views/spreadsheet/useSheetViewState";
import CellEditor from "./CellEditor.vue";

const wrappers: VueWrapper[] = [];
afterEach(() => {
  for (const wrapper of wrappers.splice(0)) {
    wrapper.unmount();
  }
  document.body.innerHTML = "";
});

const enUS = () => SPREADSHEET_LOCALES["en-US"];

/** B3 "Revenue", B4 1200 (#,##0) on a blank workbook. */
function book(): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.styles = [...model.styles, { numberFormat: "#,##0" }];
  model.sheets = [{ ...model.sheets[0]!, rows: [{ index: 2, cells: [{ col: 1, value: "Revenue" }] }, { index: 3, cells: [{ col: 1, value: 1200, styleId: 1 }] }] }];
  return model;
}

/** CellEditor bound to a real useCellEditing, its label from cellEditorView
 * as the grid computes it; attached so focus and the selection work. */
function mountEditor(model = book()) {
  let editing!: CellEditing;
  let select!: (row: number, col: number) => void;
  const wrapper = mount(
    defineComponent({
      setup() {
        const workbook = shallowRef<SpreadsheetModel | null>(model);
        const activeIndex = ref(0);
        const activeSheet = computed(() => workbook.value?.sheets[activeIndex.value] ?? null);
        const fields = shallowRef<SheetViewFields>({});
        const viewState = {
          active: computed(() => fields.value),
          updateActive: (patch: SheetViewFields) => {
            fields.value = { ...fields.value, ...patch };
          },
        };
        const sheetMetrics = useSheetMetrics(() => activeSheet.value);
        const selection = useSheetSelection({ sheet: () => activeSheet.value, sheetMetrics, viewState, model: () => workbook.value, locale: enUS });
        const undo = useSpreadsheetUndo({ model: workbook, recalc: (_, done) => done([]), onApplied: () => {}, activateSheet: () => {} });
        editing = useCellEditing({
          book: { workbook, activeIndex, activeSheet, selectSheet: () => {} },
          selection,
          sheetMetrics,
          undo,
          locale: enUS,
          toast: () => {},
          measurer: fixedMeasurer(7, 20),
        });
        select = (row, col) => {
          selection.press({ row, col }, { shift: false, mod: false });
          selection.release();
        };
        const view = computed(() => {
          const sheet = activeSheet.value!;
          const active = selection.selection.value.active;
          return cellEditorView({
            sheet,
            active,
            style: selection.styleAt(sheet, active.row, active.col) ?? {},
            rect: { key: "a", left: 64, top: 40, width: 64, height: 20 },
            room: 600,
            theme: DEFAULT_THEME,
            format: selection.formatContext.value,
          });
        });
        return () => h(CellEditor, { editor: editing, boxStyle: view.value.boxStyle, textStyle: view.value.textStyle, label: view.value.label });
      },
    }),
    { attachTo: document.body },
  );
  wrappers.push(wrapper);
  const host = () => wrapper.get('[data-testid="cell-editor"]').element as HTMLElement;
  return { wrapper, editing: () => editing, select, host, model };
}

/** Places the caret at `offset` in `node` (the browser's own caret). */
function caret(node: Node, offset: number): void {
  const range = document.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  document.getSelection()!.removeAllRanges();
  document.getSelection()!.addRange(range);
}

/** The browser's native insertion of `text` at the end of `host`, then its input event. */
function typeNatively(host: HTMLElement, text: string, init: InputEventInit = {}): void {
  const node = host.lastChild?.nodeType === 3 ? (host.lastChild as Text) : host.appendChild(document.createTextNode(""));
  node.data += text;
  caret(node, node.data.length);
  host.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text, ...init }));
}

describe("CellEditor", () => {
  it("is mounted in Ready: transparent, empty, contenteditable, named by the cell and its displayed value", async () => {
    const { wrapper, select, host } = mountEditor();
    select(2, 1);
    await nextTick();
    const el = host();
    expect(el.getAttribute("contenteditable")).toBe("true");
    expect(el.getAttribute("role")).toBe("textbox");
    expect(el.getAttribute("aria-label")).toBe("B3, Revenue");
    expect(el.getAttribute("data-mode")).toBe("ready");
    expect(el.textContent).toBe("");
    // No paper style in Ready (the scoped CSS keeps it transparent).
    expect(el.getAttribute("style")).toBeNull();
    expect(wrapper.find(".cell-editor-box").classes()).not.toContain("is-editing");
    select(3, 1);
    await nextTick();
    expect(host().getAttribute("aria-label")).toBe("B4, 1,200");
    select(9, 9);
    await nextTick();
    expect(host().getAttribute("aria-label")).toBe("J10");
  });

  it("the first typed character lands once — the browser inserts it, the editor reads it back", async () => {
    const { wrapper, editing, select, host } = mountEditor();
    select(2, 1);
    host().focus();
    typeNatively(host(), "R");
    await nextTick();
    expect(editing().mode.value).toBe("enter");
    expect(editing().text.value).toBe("R");
    expect(host().textContent).toBe("R");
    expect(host().getAttribute("data-mode")).toBe("enter");
    expect(wrapper.find(".cell-editor-box").classes()).toContain("is-editing");
    typeNatively(host(), "e");
    await nextTick();
    expect([editing().text.value, host().textContent, editing().caret.value]).toEqual(["Re", "Re", 2]);
  });

  it("re-renders a formula with one coloured span per reference (--sheet-ref-N by first appearance)", async () => {
    const { editing, select, host } = mountEditor();
    select(0, 3);
    host().focus();
    typeNatively(host(), "=A1+B2*A1");
    await nextTick();
    const spans = [...host().querySelectorAll("span.sheet-ref")] as HTMLElement[];
    expect(spans.map((span) => [span.textContent, span.dataset.refIndex, span.style.color])).toEqual([
      ["A1", "0", "var(--sheet-ref-1)"],
      ["B2", "1", "var(--sheet-ref-2)"],
      ["A1", "0", "var(--sheet-ref-1)"],
    ]);
    expect(host().textContent).toBe("=A1+B2*A1");
    expect(editing().caret.value).toBe(9);
  });

  it("never re-renders mid-composition; the composed text is read back once it ends", async () => {
    const { editing, select, host } = mountEditor();
    select(0, 0);
    host().focus();
    host().dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
    await nextTick();
    expect(editing().mode.value).toBe("enter");
    const composing = host().appendChild(document.createTextNode("に"));
    caret(composing, 1);
    host().dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true, inputType: "insertCompositionText", data: "に" }));
    // A key during the composition is ignored by the editor's table.
    const enter = new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true, isComposing: true });
    host().dispatchEvent(enter);
    await nextTick();
    expect(host().firstChild).toBe(composing);
    expect([editing().mode.value, editing().text.value, enter.defaultPrevented]).toEqual(["enter", "", false]);
    composing.data = "日本";
    caret(composing, 2);
    host().dispatchEvent(new CompositionEvent("compositionend", { bubbles: true, data: "日本" }));
    await nextTick();
    expect([editing().text.value, host().textContent]).toEqual(["日本", "日本"]);
  });

  it("paste inserts the clipboard's plain text at the caret", async () => {
    const { editing, select, host } = mountEditor();
    select(2, 1);
    host().focus();
    host().dispatchEvent(new KeyboardEvent("keydown", { key: "F2", bubbles: true, cancelable: true }));
    await nextTick();
    const paste = new Event("paste", { bubbles: true, cancelable: true });
    const clipboard: Record<string, string> = { "text/plain": " up\r\nnext", "text/html": "<b>up</b>" };
    Object.defineProperty(paste, "clipboardData", { value: { getData: (type: string) => clipboard[type] ?? "" } });
    host().dispatchEvent(paste);
    await nextTick();
    expect(paste.defaultPrevented).toBe(true);
    expect(editing().text.value).toBe("Revenue up\nnext");
    expect(host().querySelector("b")).toBeNull();
    expect(host().textContent).toBe("Revenue up\nnext");
  });

  it("shows a formula error beneath it in cell-editor-error and stays open", async () => {
    const { wrapper, editing, select, host } = mountEditor();
    select(0, 4);
    host().focus();
    typeNatively(host(), "=1+");
    host().dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    await nextTick();
    const error = wrapper.get('[data-testid="cell-editor-error"]');
    expect(error.text()).toBe(FORMULA_ERROR_MESSAGE);
    expect(error.attributes("role")).toBe("alert");
    expect(host().getAttribute("aria-invalid")).toBe("true");
    expect(editing().mode.value).toBe("enter");
    host().dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    await nextTick();
    expect(wrapper.find('[data-testid="cell-editor-error"]').exists()).toBe(false);
    expect(host().textContent).toBe("");
  });

  it("a double-click opens Edit mode with the caret at the clicked character", async () => {
    const { editing, select, host } = mountEditor();
    select(2, 1);
    host().focus();
    // jsdom has no layout: the browser's hit test, answering offset 3 of the rendered text.
    Object.defineProperty(document, "caretRangeFromPoint", {
      configurable: true,
      value: () => {
        const range = document.createRange();
        range.setStart(host().firstChild ?? host(), host().firstChild ? 3 : 0);
        return range;
      },
    });
    try {
      expect(editing().openEdit({ point: { x: 10, y: 10 } })).toBe(true);
      await nextTick();
      await nextTick();
      expect([editing().mode.value, editing().text.value, editing().caret.value]).toEqual(["edit", "Revenue", 3]);
    } finally {
      Reflect.deleteProperty(document, "caretRangeFromPoint");
    }
  });
});
