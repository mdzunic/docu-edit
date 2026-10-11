// SPEC-117's read-only formula bar: the Name Box is a read-only input named
// "Name Box" showing the active cell's address (or the drag size), and the
// content is a read-only textbox showing the active cell's text with its line
// breaks preserved.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { computed, defineComponent, h, nextTick, ref, shallowRef } from "vue";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { Cell, CellStyle, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../lib/spreadsheet/numberFormat/locale";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useCellEditing, type CellEditing } from "../../views/spreadsheet/useCellEditing";
import { useFormulaDisplay } from "../../views/spreadsheet/useFormulaDisplay";
import { useSheetMetrics } from "../../views/spreadsheet/useSheetMetrics";
import { useSheetSelection } from "../../views/spreadsheet/useSheetSelection";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import type { SheetViewFields } from "../../views/spreadsheet/useSheetViewState";
import CellEditor from "./CellEditor.vue";
import FormulaBar from "./FormulaBar.vue";

describe("FormulaBar", () => {
  it("shows the active cell's address in a read-only Name Box", () => {
    const wrapper = mount(FormulaBar, { props: { address: "B3", content: "10" } });
    expect(wrapper.find('[data-testid="formula-bar"]').exists()).toBe(true);
    const nameBox = wrapper.get('[data-testid="name-box"]');
    expect(nameBox.element.tagName).toBe("INPUT");
    expect((nameBox.element as HTMLInputElement).readOnly).toBe(true);
    expect(nameBox.attributes("aria-label")).toBe("Name Box");
    expect((nameBox.element as HTMLInputElement).value).toBe("B3");
  });

  it("shows the drag size while a drag is in progress", async () => {
    const wrapper = mount(FormulaBar, { props: { address: "B3", content: "" } });
    await wrapper.setProps({ address: "3R × 2C" });
    expect((wrapper.get('[data-testid="name-box"]').element as HTMLInputElement).value).toBe("3R × 2C");
  });

  it("shows the content as a read-only textbox with line breaks preserved", () => {
    const wrapper = mount(FormulaBar, { props: { address: "A12", content: "First  line\ttabbed\nSecond\nthird" } });
    const content = wrapper.get('[data-testid="formula-bar-content"]');
    expect(content.attributes("role")).toBe("textbox");
    expect(content.attributes("aria-readonly")).toBe("true");
    expect(content.element.textContent).toBe("First  line\ttabbed\nSecond\nthird");
    // pre-wrap keeps the breaks on screen (scoped CSS is not applied in jsdom,
    // so the class carrying it is what the test can pin).
    expect(content.classes()).toContain("formula-bar-content");
  });

  it("mutes a dynamic array member's formula (SPEC-119): data-spill-member and the muted class, only then", async () => {
    const wrapper = mount(FormulaBar, { props: { address: "H3", content: "=SORT(A2:A6)", spillMember: true } });
    const content = wrapper.get('[data-testid="formula-bar-content"]');
    expect(content.text()).toBe("=SORT(A2:A6)");
    expect(content.attributes("data-spill-member")).toBe("true");
    expect(content.classes()).toContain("is-spill-member");
    await wrapper.setProps({ address: "H2", spillMember: false });
    expect(content.attributes("data-spill-member")).toBeUndefined();
    expect(content.classes()).not.toContain("is-spill-member");
  });

  it("is empty for an empty active cell", () => {
    const wrapper = mount(FormulaBar, { props: { address: "Z99", content: "" } });
    expect(wrapper.get('[data-testid="formula-bar-content"]').element.textContent).toBe("");
  });
});

// SPEC-118: what the bar shows is the active cell's edit text — editText
// through the cell's effective style, in the view's locale — wired the way
// the workbook view wires it (since SPEC-119, useFormulaDisplay's barText:
// formulaBarText falls back to editText for a cell with no formula).
describe("FormulaBar edit text (SPEC-118)", () => {
  const styles: CellStyle[] = [
    {},
    { builtinFormatId: 14, numberFormat: "m/d/yyyy" },
    { numberFormat: "0.00%" },
    { quotePrefix: true },
    { numberFormat: "[h]:mm:ss" },
    { numberFormat: "#,##0.00" },
  ];
  const cells: Array<[number, Cell]> = [
    [0, { col: 0, value: 46303.5729166667, styleId: 1 }], // A1: a date and a time under id 14
    [1, { col: 0, value: 0.125, styleId: 2 }], // A2
    [2, { col: 0, value: "007", styleId: 3 }], // A3: quote-prefixed text
    [3, { col: 0, value: 1234.5, styleId: 5 }], // A4: grouping is never shown
    [4, { col: 0, value: 1.5, styleId: 4 }], // A5: elapsed time over a day
    [5, { col: 0, value: true }], // A6
    [6, { col: 0, value: { error: "#DIV/0!" } }], // A7
    [7, { col: 0, value: 46303, styleId: 1 }], // A8: a date at midnight
  ];
  const sheet: Sheet = {
    id: "s1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: cells.map(([index, cell]) => ({ index, cells: [cell] })),
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
  };
  const model: SpreadsheetModel = {
    kind: "spreadsheet",
    id: "w",
    title: "W",
    format: "xlsx",
    readerVersion: 3,
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles,
    defaultStyleId: 0,
  };

  /** The bar over `model`, its content useFormulaDisplay's barText. */
  function mountBar() {
    const locale = shallowRef<SpreadsheetLocaleTag>("en-US");
    const fields = shallowRef<SheetViewFields>({});
    let select!: (row: number) => void;
    const wrapper = mount(
      defineComponent({
        setup() {
          const sheetMetrics = useSheetMetrics(() => sheet);
          const viewState = {
            active: computed(() => fields.value),
            updateActive: (patch: SheetViewFields) => {
              fields.value = { ...fields.value, ...patch };
            },
          };
          const selection = useSheetSelection({
            sheet: () => sheet,
            sheetMetrics,
            viewState,
            model: () => model,
            locale: () => SPREADSHEET_LOCALES[locale.value],
          });
          select = (row) => {
            selection.press({ row, col: 0 }, { shift: false, mod: false });
            selection.release();
          };
          const formulas = useFormulaDisplay({ sheet: () => sheet, model: () => model, viewState, selection });
          return () => h(FormulaBar, { address: selection.nameBoxText.value, content: formulas.barText.value });
        },
      }),
    );
    const content = async (row: number, tag: SpreadsheetLocaleTag = "en-US"): Promise<string | null> => {
      locale.value = tag;
      select(row);
      await nextTick();
      return wrapper.get('[data-testid="formula-bar-content"]').element.textContent;
    };
    return { wrapper, content };
  }

  it("shows a date as the locale's short date plus long time when it has a time part", async () => {
    const { wrapper, content } = mountBar();
    expect(await content(0)).toBe("10/8/2026 1:45:00 PM");
    expect(await content(0, "de-DE")).toBe("08.10.2026 13:45:00");
    expect(await content(7)).toBe("10/8/2026");
    expect(await content(7, "en-GB")).toBe("08/10/2026");
    wrapper.unmount();
  });

  it("shows percentages ×100 with %, numbers plain with the locale's decimal, elapsed times as a date and time", async () => {
    const { wrapper, content } = mountBar();
    expect(await content(1)).toBe("12.5%");
    expect(await content(3)).toBe("1234.5");
    expect(await content(3, "fr-FR")).toBe("1234,5");
    expect(await content(4)).toBe("1/1/1900 12:00:00 PM");
    wrapper.unmount();
  });

  it("shows text with its quote prefix, and booleans and errors in English in every locale", async () => {
    const { wrapper, content } = mountBar();
    expect(await content(2)).toBe("'007");
    expect(await content(5, "de-DE")).toBe("TRUE");
    expect(await content(6, "fr-FR")).toBe("#DIV/0!");
    wrapper.unmount();
  });
});

const enUS = () => SPREADSHEET_LOCALES["en-US"];

/** The bar and the in-cell editor over one useCellEditing: A1 "Revenue", A2 =A1. */
function mountEditing() {
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: "Revenue" }] }, { index: 1, cells: [{ col: 0, value: "Revenue", formula: "A1" }] }] }];
  let editing!: CellEditing;
  let selectionText!: () => string;
  const editable = ref(true);
  const canInsertFunction = ref(true);
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
        const formulas = useFormulaDisplay({ sheet: () => activeSheet.value, model: () => workbook.value, viewState, selection });
        selectionText = () => selection.nameBoxText.value;
        return () =>
          h("div", [
            h(FormulaBar, {
              address: selection.nameBoxText.value,
              content: formulas.barText.value,
              editor: editing,
              editable: editable.value,
              canInsertFunction: canInsertFunction.value,
            }),
            h(CellEditor, { editor: editing, label: selection.nameBoxText.value }),
          ]);
      },
    }),
    { attachTo: document.body },
  );
  const bar = () => wrapper.get('[data-testid="formula-bar-content"]').element as HTMLElement;
  const cell = () => wrapper.get('[data-testid="cell-editor"]').element as HTMLElement;
  return { wrapper, model, editable, canInsertFunction, editing: () => editing, bar, cell, active: () => selectionText() };
}

function caret(node: Node, offset: number): void {
  const range = document.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  document.getSelection()!.removeAllRanges();
  document.getSelection()!.addRange(range);
}

function keydown(target: HTMLElement, key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

// SPEC-122: with the view's cell editor and a grid sheet on show the content
// is editable and holds the same draft as the in-cell editor — typing in
// either updates both; a click opens Edit mode at the clicked character;
// Enter, Tab and Escape behave as in the cell, and a commit from the bar
// returns focus to the in-cell editor. Without an editor (or on a foreign
// sheet) it stays read-only.
describe("FormulaBar editing (SPEC-122)", () => {
  it("is an editable textbox with an editor on a grid sheet, read-only without", async () => {
    const { wrapper, bar, editable } = mountEditing();
    expect(bar().getAttribute("contenteditable")).toBe("true");
    expect(bar().getAttribute("aria-readonly")).toBeNull();
    expect(bar().textContent).toBe("Revenue");
    editable.value = false;
    await nextTick();
    expect(bar().getAttribute("contenteditable")).toBeNull();
    expect(bar().getAttribute("aria-readonly")).toBe("true");
    wrapper.unmount();
    const plain = mount(FormulaBar, { props: { address: "A1", content: "x" } });
    expect(plain.get('[data-testid="formula-bar-content"]').attributes("contenteditable")).toBeUndefined();
  });

  it("SPEC-127: fx asks for Insert Function, and on a cell protection keeps from editing is disabled and titled \"This cell is protected\"", async () => {
    const { wrapper, canInsertFunction } = mountEditing();
    const fx = () => wrapper.get('[data-testid="formula-bar-fx"]');
    const asks = () => wrapper.findComponent(FormulaBar).emitted("insert-function") ?? [];
    expect(fx().attributes("aria-label")).toBe("Insert function");
    expect(fx().attributes("title")).toBe("Insert function");
    expect(fx().attributes("disabled")).toBeUndefined();
    await fx().trigger("click");
    expect(asks()).toHaveLength(1);
    canInsertFunction.value = false;
    await nextTick();
    expect(fx().attributes("disabled")).toBeDefined();
    expect(fx().attributes("title")).toBe("This cell is protected");
    await fx().trigger("click");
    expect(asks()).toHaveLength(1);
    wrapper.unmount();
  });

  it("a click opens Edit mode at the clicked character; typing there updates the in-cell editor's draft too", async () => {
    const { wrapper, editing, bar, cell } = mountEditing();
    bar().dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    bar().focus();
    caret(bar().firstChild!, 3);
    bar().dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await nextTick();
    expect([editing().mode.value, editing().text.value, editing().caret.value]).toEqual(["edit", "Revenue", 3]);
    const text = bar().firstChild as Text;
    text.data = "Rev!enue";
    caret(text, 4);
    bar().dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: "!" }));
    await nextTick();
    expect([editing().text.value, editing().caret.value]).toEqual(["Rev!enue", 4]);
    expect(cell().textContent).toBe("Rev!enue");
    wrapper.unmount();
  });

  it("typing in the cell shows in the bar, with the formula's references coloured there too", async () => {
    const { wrapper, cell, bar } = mountEditing();
    cell().focus();
    cell().append(document.createTextNode("=A1+B2"));
    caret(cell().firstChild!, 6);
    cell().dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: "=" }));
    await nextTick();
    expect(bar().textContent).toBe("=A1+B2");
    expect([...bar().querySelectorAll("span.sheet-ref")].map((span) => span.textContent)).toEqual(["A1", "B2"]);
    wrapper.unmount();
  });

  it("Enter, Tab and Escape behave as in the cell, and a commit hands focus back to the in-cell editor", async () => {
    const { wrapper, model, editing, bar, cell, active } = mountEditing();
    bar().focus();
    await nextTick();
    // Keyboard focus opens Edit mode with the caret at the end.
    expect([editing().mode.value, editing().caret.value]).toEqual(["edit", 7]);
    const text = bar().firstChild as Text;
    text.data = "Revenues";
    caret(text, 8);
    bar().dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: "s" }));
    expect(keydown(bar(), "Enter").defaultPrevented).toBe(true);
    await nextTick();
    expect(model.sheets[0]!.rows[0]!.cells[0]!.value).toBe("Revenues");
    expect([editing().mode.value, active()]).toEqual(["ready", "A2"]);
    expect(document.activeElement).toBe(cell());
    expect(bar().textContent).toBe("=A1");

    bar().focus();
    await nextTick();
    expect(keydown(bar(), "Tab").defaultPrevented).toBe(true);
    await nextTick();
    expect([editing().mode.value, active()]).toEqual(["ready", "B2"]);

    bar().focus();
    await nextTick();
    bar().replaceChildren(document.createTextNode("changed"));
    caret(bar().firstChild!, 7);
    bar().dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: "d" }));
    expect(keydown(bar(), "Escape").defaultPrevented).toBe(true);
    await nextTick();
    expect([editing().mode.value, active(), bar().textContent]).toEqual(["ready", "B2", ""]);
    wrapper.unmount();
  });
});
