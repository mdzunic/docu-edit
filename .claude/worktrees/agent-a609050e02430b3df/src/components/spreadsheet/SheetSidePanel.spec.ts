// SPEC-129 "The side panel" — SheetSidePanel renders SidePanel.vue from
// useSheetSidePanel with the active pane (the rules manager) in its default
// slot: nothing while closed; the tab, title and `cf-manager` once opened;
// the close button, the side toggle and the width handle drive the panel's
// state.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { computed, effectScope, nextTick, shallowRef } from "vue";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { useConditionalFormatManager } from "../../views/spreadsheet/useConditionalFormatManager";
import { useSheetComments } from "../../views/spreadsheet/useSheetComments";
import { useSheetSidePanel } from "../../views/spreadsheet/useSheetSidePanel";
import SheetSidePanel from "./SheetSidePanel.vue";

const scopes: (() => void)[] = [];
const wrappers: VueWrapper[] = [];
beforeEach(() => {
  localStorage.clear();
});
afterEach(() => {
  wrappers.splice(0).forEach((wrapper) => wrapper.unmount());
  scopes.splice(0).forEach((stop) => stop());
});

function setup() {
  const book = shallowRef<SpreadsheetModel | null>(createBlankWorkbook());
  const scope = effectScope();
  scopes.push(() => scope.stop());
  const editing = scope.run(() => ({
    sidePanel: useSheetSidePanel(),
    cfManager: useConditionalFormatManager({
      model: () => book.value,
      sheet: () => book.value?.sheets[0] ?? null,
      selection: () => ({ ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], active: { row: 0, col: 0 } }),
      undo: { record: () => undefined, commit: () => undefined },
      locale: () => SPREADSHEET_LOCALES["en-US"],
    }),
  }))!;
  const comments = scope.run(() =>
    useSheetComments({
      book: { workbook: book, activeSheet: shallowRef(book.value?.sheets[0] ?? null), activeIndex: shallowRef(0) },
      undo: { record: () => undefined, commit: () => undefined },
      selection: { selection: computed(() => ({ ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], active: { row: 0, col: 0 }, activeRangeIndex: 0 })) },
      sidePanel: editing.sidePanel,
    }),
  )!;
  Object.assign(editing, { comments });
  const wrapper = mount(SheetSidePanel, { props: { editing }, attachTo: document.body });
  wrappers.push(wrapper);
  return { editing, wrapper };
}

describe("SheetSidePanel (SPEC-129)", () => {
  it("SPEC-135: hosts the comments pane — its tab and SheetCommentsPanel, empty for a workbook without threads", async () => {
    const { editing, wrapper } = setup();
    editing.sidePanel.toggle("comments");
    await nextTick();
    expect(wrapper.get('[data-testid="side-panel-tab-comments"]').text()).toContain("Comments");
    expect(wrapper.get('[data-testid="sheet-comments-panel"]').attributes("aria-label")).toBe("Comments");
    expect(wrapper.find('[data-testid="sheet-comment-empty"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="cf-manager"]').exists()).toBe(false);
  });

  it("renders nothing while closed, and the rules manager in SidePanel once opened", async () => {
    const { editing, wrapper } = setup();
    expect(wrapper.find('[data-testid="side-panel"]').exists()).toBe(false);
    editing.sidePanel.open("conditionalFormats");
    await nextTick();
    expect(wrapper.find('[data-testid="side-panel"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="side-panel-tab-conditionalFormats"]').text()).toContain("Conditional Formatting");
    expect(wrapper.find('[data-testid="cf-manager"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="cf-empty"]').exists()).toBe(true);
  });

  it("closes from the panel's close button and flips sides from its toggle", async () => {
    const { editing, wrapper } = setup();
    editing.sidePanel.open("conditionalFormats");
    await nextTick();
    await wrapper.get('[data-testid="side-panel-side-toggle"]').trigger("click");
    expect(editing.sidePanel.side.value).toBe("left");
    await wrapper.get('[data-testid="side-panel-close"]').trigger("click");
    expect(editing.sidePanel.activePane.value).toBeNull();
    expect(wrapper.find('[data-testid="side-panel"]').exists()).toBe(false);
    expect(JSON.parse(localStorage.getItem("docu-edit:sheet-side-panel") ?? "{}")).toMatchObject({ side: "left", activePane: null });
  });
});
