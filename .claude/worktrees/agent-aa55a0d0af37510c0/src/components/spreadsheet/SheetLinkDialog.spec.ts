// SPEC-135's SheetLinkDialog over a real useCellLinks: every test id the AC
// lists, the text prefilled from the active cell, the kind radios switching
// between the web address and the place fields, the inline refusal of a
// dangerous scheme with Save disabled, a place saved through the dialog, and
// Remove shown only for an existing link.
import { mount } from "@vue/test-utils";
import { afterEach, describe, expect, it } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { singleCell, type SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { LINK_SCHEME_NOT_ALLOWED_MESSAGE } from "../../views/spreadsheet/cellLinks";
import { useCellLinks } from "../../views/spreadsheet/useCellLinks";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import SheetLinkDialog from "./SheetLinkDialog.vue";

let scope: EffectScope | null = null;
afterEach(() => {
  scope?.stop();
  document.body.innerHTML = "";
});

function controller(active: SheetSelection) {
  scope = effectScope();
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  const summary: Sheet = {
    ...base,
    id: "s1",
    name: "Summary",
    rows: [{ index: 6, cells: [{ col: 0, value: "Example site" }, { col: 2, value: "Details" }] }],
    hyperlinks: [{ linkId: "h-A7", range: { top: 6, left: 0, bottom: 6, right: 0 }, target: { kind: "url", href: "https://example.com/" } }],
  };
  model.sheets = [summary, { ...base, id: "data", name: "Data", rows: [] }];
  model.names = [{ name: "Target", formula: "Data!$A$10" }];
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const links = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: wb, recalc: (_changed, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    return useCellLinks({
      book: { workbook: wb, activeIndex, activeSheet: computed(() => wb.value?.sheets[activeIndex.value] ?? null), selectSheet: () => {} },
      selection: { selection: computed(() => active), formatContext: computed(() => ({ locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" as const })) },
      viewState: { update: () => {}, syncInto: () => {} },
      undo,
      cell: { editing: computed(() => false), focusCell: () => {}, revealTick: ref(0), openEdit: () => true },
      locale: () => SPREADSHEET_LOCALES["en-US"],
      toast: () => {},
      measurer: fixedMeasurer(7, 20),
    });
  })!;
  links.openDialog();
  return { links, wb };
}

describe("SheetLinkDialog (SPEC-135)", () => {
  it("shows every field with the cell's text, the kind radios switching between an address and a place", async () => {
    const { links } = controller(singleCell({ row: 6, col: 2 }));
    const wrapper = mount(SheetLinkDialog, { props: { links }, attachTo: document.body });
    expect(wrapper.get('[data-testid="sheet-link-dialog"]').text()).toContain("Insert link");
    expect(wrapper.get<HTMLInputElement>('[data-testid="sheet-link-text"]').element.value).toBe("Details");
    expect(wrapper.get<HTMLInputElement>('[data-testid="sheet-link-kind-web"]').element.checked).toBe(true);
    for (const testid of ["sheet-link-address", "sheet-link-tooltip", "sheet-link-error", "sheet-link-save", "sheet-link-cancel"]) {
      expect(wrapper.find(`[data-testid="${testid}"]`).exists(), testid).toBe(true);
    }
    expect(wrapper.find('[data-testid="sheet-link-remove"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="sheet-link-save"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="sheet-link-kind-place"]').setValue(true);
    expect(wrapper.find('[data-testid="sheet-link-address"]').exists()).toBe(false);
    const sheet = wrapper.get<HTMLSelectElement>('[data-testid="sheet-link-sheet"]');
    expect(Array.from(sheet.element.options).map((option) => option.value)).toEqual(["Summary", "Data"]);
    expect(wrapper.get<HTMLInputElement>('[data-testid="sheet-link-ref"]').element.value).toBe("A1");
    const names = wrapper.get<HTMLSelectElement>('[data-testid="sheet-link-names"]');
    expect(Array.from(names.element.options).map((option) => option.value)).toEqual(["", "Target"]);
    wrapper.unmount();
  });

  it("refuses a dangerous scheme inline with Save disabled", async () => {
    const { links } = controller(singleCell({ row: 6, col: 2 }));
    const wrapper = mount(SheetLinkDialog, { props: { links }, attachTo: document.body });
    await wrapper.get('[data-testid="sheet-link-address"]').setValue("javascript:alert(1)");
    expect(wrapper.get('[data-testid="sheet-link-error"]').text()).toBe(LINK_SCHEME_NOT_ALLOWED_MESSAGE);
    expect(wrapper.get('[data-testid="sheet-link-save"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="sheet-link-address"]').setValue("https://example.com/");
    expect(wrapper.get('[data-testid="sheet-link-error"]').text()).toBe("");
    expect(wrapper.get('[data-testid="sheet-link-save"]').attributes("disabled")).toBeUndefined();
    wrapper.unmount();
  });

  it("saves a place picked from the sheet list and a reference", async () => {
    const { links, wb } = controller(singleCell({ row: 6, col: 2 }));
    const wrapper = mount(SheetLinkDialog, { props: { links }, attachTo: document.body });
    await wrapper.get('[data-testid="sheet-link-kind-place"]').setValue(true);
    await wrapper.get('[data-testid="sheet-link-sheet"]').setValue("Data");
    await wrapper.get('[data-testid="sheet-link-ref"]').setValue("A10");
    await wrapper.get('[data-testid="sheet-link-tooltip"]').setValue("Go to the target");
    await wrapper.get('[data-testid="sheet-link-save"]').trigger("click");
    expect(links.dialog.value).toBeNull();
    expect(wb.value!.sheets[0]!.hyperlinks!.find((link) => link.range.left === 2)).toMatchObject({ target: { kind: "place", location: "Data!A10" }, tooltip: "Go to the target" });
    wrapper.unmount();
  });

  it("edits an existing link, offering Remove", async () => {
    const { links, wb } = controller(singleCell({ row: 6, col: 0 }));
    const wrapper = mount(SheetLinkDialog, { props: { links }, attachTo: document.body });
    expect(wrapper.get('[data-testid="sheet-link-dialog"]').text()).toContain("Edit link");
    expect(wrapper.get<HTMLInputElement>('[data-testid="sheet-link-address"]').element.value).toBe("https://example.com/");
    await wrapper.get('[data-testid="sheet-link-remove"]').trigger("click");
    expect(wb.value!.sheets[0]!.hyperlinks).toBeUndefined();
    expect(links.dialog.value).toBeNull();
    wrapper.unmount();
  });
});
