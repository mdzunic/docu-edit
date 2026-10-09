// SPEC-117's read-only formula bar: the Name Box is a read-only input named
// "Name Box" showing the active cell's address (or the drag size), and the
// content is a read-only textbox showing the active cell's text with its line
// breaks preserved.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { computed, defineComponent, h, nextTick, shallowRef } from "vue";
import type { Cell, CellStyle, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../lib/spreadsheet/numberFormat/locale";
import { useFormulaDisplay } from "../../views/spreadsheet/useFormulaDisplay";
import { useSheetMetrics } from "../../views/spreadsheet/useSheetMetrics";
import { useSheetSelection } from "../../views/spreadsheet/useSheetSelection";
import type { SheetViewFields } from "../../views/spreadsheet/useSheetViewState";
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
