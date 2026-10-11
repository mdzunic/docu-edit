// SPEC-117 "Aggregates" and "Zoom" in the workbook's status bar: the Average,
// Count and Sum texts it is handed (each hidden when empty — a single cell or
// a selection of text shows less), and SPEC-063's ZoomControl at its right
// with its fit entries hidden.
import { beforeEach, describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { computed, defineComponent, h, nextTick, shallowRef } from "vue";
import { loadPreferences, resetPreferences } from "../../lib/preferences";
import type { CellStyle, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../lib/spreadsheet/numberFormat/locale";
import { useSheetMetrics } from "../../views/spreadsheet/useSheetMetrics";
import { useSheetSelection } from "../../views/spreadsheet/useSheetSelection";
import type { SheetViewFields } from "../../views/spreadsheet/useSheetViewState";
import SpreadsheetStatusBar from "./SpreadsheetStatusBar.vue";
import ZoomControl from "../ZoomControl.vue";

describe("SpreadsheetStatusBar", () => {
  it("shows Average, Count and Sum for a multi-cell numeric selection", () => {
    const wrapper = mount(SpreadsheetStatusBar, {
      props: { average: "Average: 20", count: "Count: 3", sum: "Sum: 60", zoomPercent: 85 },
    });
    expect(wrapper.find('[data-testid="spreadsheet-status-bar"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="sheet-status-average"]').text()).toBe("Average: 20");
    expect(wrapper.get('[data-testid="sheet-status-count"]').text()).toBe("Count: 3");
    expect(wrapper.get('[data-testid="sheet-status-sum"]').text()).toBe("Sum: 60");
  });

  it("shows only Count when nothing numeric is selected, and nothing for a single cell", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 2" } });
    expect(wrapper.find('[data-testid="sheet-status-average"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="sheet-status-sum"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="sheet-status-count"]').text()).toBe("Count: 2");
    await wrapper.setProps({ count: "" });
    expect(wrapper.find('[data-testid="sheet-status-count"]').exists()).toBe(false);
  });

  it("puts the zoom control at its right, at the sheet's zoom, without the fit entries", () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 3", zoomPercent: 85 }, attachTo: document.body });
    const bar = wrapper.get('[data-testid="spreadsheet-status-bar"]').element;
    const zoom = wrapper.get('[data-testid="zoom-control"]').element;
    // The zoom control is the bar's last child: aggregates left, zoom right.
    expect(bar.lastElementChild).toBe(zoom);
    expect((wrapper.get('[data-testid="zoom-percent"]').element as HTMLInputElement).value).toBe("85%");
    expect(wrapper.findComponent(ZoomControl).props("showFit")).toBe(false);
    expect(wrapper.find('[data-testid="zoom-fit-width"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="zoom-fit-page"]').exists()).toBe(false);
    wrapper.unmount();
  });

  it("forwards the zoom control's steps, presets and typed value", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { zoomPercent: 100 } });
    await wrapper.get('[data-testid="zoom-in"]').trigger("click");
    await wrapper.get('[data-testid="zoom-out"]').trigger("click");
    const zoom = wrapper.findComponent(ZoomControl);
    zoom.vm.$emit("set-preset", 150);
    zoom.vm.$emit("apply-input", "85%");
    expect(wrapper.emitted("zoom-in")).toHaveLength(1);
    expect(wrapper.emitted("zoom-out")).toHaveLength(1);
    expect(wrapper.emitted("set-zoom")).toEqual([[150]]);
    expect(wrapper.emitted("apply-zoom")).toEqual([["85%"]]);
  });

  it("SPEC-122: shows the cell editor's mode in sheet-status-mode at its left; empty hides it", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { mode: "Ready", count: "Count: 3" } });
    const bar = wrapper.get('[data-testid="spreadsheet-status-bar"]').element;
    expect(bar.firstElementChild).toBe(wrapper.get('[data-testid="sheet-status-mode"]').element);
    for (const mode of ["Ready", "Enter", "Edit", "Point"]) {
      await wrapper.setProps({ mode });
      expect(wrapper.get('[data-testid="sheet-status-mode"]').text()).toBe(mode);
    }
    await wrapper.setProps({ mode: "" });
    expect(wrapper.find('[data-testid="sheet-status-mode"]').exists()).toBe(false);
  });

  it("SPEC-128: shows the filter's record count in sheet-status-filter after the mode; empty hides it", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { mode: "Ready", filter: "8 of 12 records found" } });
    const filter = wrapper.get('[data-testid="sheet-status-filter"]');
    expect(filter.text()).toBe("8 of 12 records found");
    expect(filter.attributes("role")).toBe("status");
    expect(wrapper.get('[data-testid="sheet-status-mode"]').element.nextElementSibling).toBe(filter.element);
    await wrapper.setProps({ filter: "" });
    expect(wrapper.find('[data-testid="sheet-status-filter"]').exists()).toBe(false);
  });

  it("SPEC-120: shows the circular-reference warning and Calculating… only while they apply, after the aggregates", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 3" } });
    expect(wrapper.find('[data-testid="sheet-status-circular"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="sheet-status-calculating"]').exists()).toBe(false);
    await wrapper.setProps({ circular: "Circular reference: Loop!A1", calculating: true });
    expect(wrapper.get('[data-testid="sheet-status-circular"]').text()).toBe("Circular reference: Loop!A1");
    expect(wrapper.get('[data-testid="sheet-status-calculating"]').text()).toBe("Calculating…");
    const regions = wrapper.findAll('[role="status"]');
    expect(regions).toHaveLength(2);
    expect(regions[1]!.attributes("aria-live")).toBe("polite");
    await wrapper.setProps({ circular: "", calculating: false });
    expect(wrapper.find('[data-testid="sheet-status-circular"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="sheet-status-calculating"]').exists()).toBe(false);
  });
});

// SPEC-118: Sum and Average take the number format of the selection's first
// numeric cell in row-major order, in the view's locale; Count stays a plain
// integer. And the sheet-locale control: the tag, a menu of one
// menuitemradio per locale, a choice written as the preference.
describe("SpreadsheetStatusBar formatted aggregates and the locale control (SPEC-118)", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  const styles: CellStyle[] = [{}, { numberFormat: "#,##0.00" }, { numberFormat: "0.0%" }, { builtinFormatId: 14, numberFormat: "m/d/yyyy" }];
  /** B2 "Label", C2 1234.5 (#,##0.00), B3 0.5 (0.0%), C3 2 (General); B5 46303 and C5 1 (dates). */
  const sheet: Sheet = {
    id: "s1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [
      { index: 1, cells: [{ col: 1, value: "Label" }, { col: 2, value: 1234.5, styleId: 1 }] },
      { index: 2, cells: [{ col: 1, value: 0.5, styleId: 2 }, { col: 2, value: 2 }] },
      { index: 4, cells: [{ col: 1, value: 46303, styleId: 3 }, { col: 2, value: 1, styleId: 3 }] },
    ],
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

  function mountBar() {
    const locale = shallowRef<SpreadsheetLocaleTag>("en-US");
    const fields = shallowRef<SheetViewFields>({});
    let selectRange!: (top: number, left: number, bottom: number, right: number) => void;
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
          selectRange = (top, left, bottom, right) => {
            selection.press({ row: top, col: left }, { shift: false, mod: false });
            selection.press({ row: bottom, col: right }, { shift: true, mod: false });
            selection.release();
          };
          return () =>
            h(SpreadsheetStatusBar, {
              average: selection.statusTexts.value.average ?? "",
              count: selection.statusTexts.value.count ?? "",
              sum: selection.statusTexts.value.sum ?? "",
            });
        },
      }),
      { attachTo: document.body },
    );
    const texts = async (range: [number, number, number, number], tag: SpreadsheetLocaleTag = "en-US") => {
      locale.value = tag;
      selectRange(...range);
      await nextTick();
      const text = (id: string) => wrapper.find(`[data-testid="${id}"]`);
      return {
        average: text("sheet-status-average").exists() ? text("sheet-status-average").text() : undefined,
        count: text("sheet-status-count").text(),
        sum: text("sheet-status-sum").exists() ? text("sheet-status-sum").text() : undefined,
      };
    };
    return { wrapper, texts };
  }

  it("formats Sum and Average with the first numeric cell's format, row-major", async () => {
    const { wrapper, texts } = mountBar();
    // B2:C3 row-major: "Label", 1234.5 (#,##0.00) first, then 0.5, 2.
    expect(await texts([1, 1, 2, 2])).toEqual({ average: "Average: 412.33", count: "Count: 4", sum: "Sum: 1,237.00" });
    // B3:C3: 0.5 (0.0%) first.
    expect(await texts([2, 1, 2, 2])).toEqual({ average: "Average: 125.0%", count: "Count: 2", sum: "Sum: 250.0%" });
    // Dates sum as dates, as Excel's bar shows them.
    expect(await texts([4, 1, 4, 2])).toEqual({ average: "Average: 5/21/1963", count: "Count: 2", sum: "Sum: 10/9/2026" });
    wrapper.unmount();
  });

  it("formats them in the view's locale; Count stays a plain integer", async () => {
    const { wrapper, texts } = mountBar();
    expect(await texts([1, 1, 2, 2], "de-DE")).toEqual({ average: "Average: 412,33", count: "Count: 4", sum: "Sum: 1.237,00" });
    expect(await texts([4, 1, 4, 2], "de-DE")).toMatchObject({ sum: "Sum: 09.10.2026", count: "Count: 2" });
    wrapper.unmount();
  });

  it("shows the locale's tag and offers one menuitemradio per locale; a choice is the preference", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 3" }, attachTo: document.body });
    const button = wrapper.get('[data-testid="sheet-locale"]');
    expect(button.text()).toBe("en-US");
    expect(button.attributes("aria-haspopup")).toBe("menu");
    await button.trigger("click");
    const items = wrapper.findAll('[role="menuitemradio"]');
    expect(items.map((item) => item.attributes("data-testid"))).toEqual([
      "sheet-locale-en-US",
      "sheet-locale-en-GB",
      "sheet-locale-de-DE",
      "sheet-locale-fr-FR",
      "sheet-locale-es-ES",
      "sheet-locale-cs-CZ",
    ]);
    expect(items[0]!.attributes("aria-checked")).toBe("true");
    expect(items[2]!.text()).toBe("Deutsch (Deutschland) — 1.234,56 · 08.10.2026");
    await wrapper.get('[data-testid="sheet-locale-de-DE"]').trigger("click");
    expect(loadPreferences().spreadsheetLocale).toBe("de-DE");
    await nextTick();
    expect(wrapper.get('[data-testid="sheet-locale"]').text()).toBe("de-DE");
    // The zoom control stays the bar's last child.
    expect(wrapper.get('[data-testid="spreadsheet-status-bar"]').element.lastElementChild).toBe(wrapper.get('[data-testid="zoom-control"]').element);
    wrapper.unmount();
  });
});
