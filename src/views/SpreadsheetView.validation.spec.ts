// SPEC-130 through the real workbook view at /sheets/:id: the input message
// (role="tooltip", the active cell's aria-describedby), the in-cell dropdown
// (Alt+Down from the cell editor, the list's keys, a pick entering the value
// and returning focus to the grid), the entry gate's alert (stop → Retry with
// the entered text selected in the focused cell editor; warning → "Continue?"
// → No, then Escape cancelling the edit), and a list rule with showDropdown
// false showing no arrow — the e2e's path, on a hand-built book (the readers
// are SPEC-130's other half). SpreadsheetView.spec.ts's harness: fresh
// modules per test, idb mocked, the view attached so focus is real, the grid
// given a 1280 × 720 viewport.
import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMemoryHistory, createRouter } from "vue-router";
import { parseSqref } from "../lib/spreadsheet/address";
import type { Sheet, SpreadsheetModel } from "../lib/spreadsheet/model";
import { SPREADSHEET_READER_VERSION } from "../lib/spreadsheet/open";
import { defaultValidationSettings, type DataValidation } from "../lib/spreadsheet/validation";

const cleanups: Array<() => void> = [];

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.spyOn(Element.prototype, "clientWidth", "get").mockReturnValue(1280);
  vi.spyOn(Element.prototype, "clientHeight", "get").mockReturnValue(720);
  // jsdom has no canvas: row autofit's measurer falls back to its defaults.
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

afterEach(() => {
  for (const cleanup of cleanups.splice(0)) {
    cleanup();
  }
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

function rule(sqref: string, settings: Partial<DataValidation>): DataValidation {
  return { id: crypto.randomUUID(), ranges: parseSqref(sqref), ...defaultValidationSettings(), ...settings };
}

/** Sheet "Form" as the e2e fixture has it: B3 holds 5; B2 a stop list with
 * an input message, B3 a whole 1–10 warning, B8 a list hiding its arrow. */
function formBook(): SpreadsheetModel {
  const sheet: Sheet = {
    conditionalFormats: [],
    id: crypto.randomUUID(),
    name: "Form",
    visibility: "visible",
    kind: "grid",
    rows: [{ index: 2, cells: [{ col: 1, value: 5 }] }],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    validations: [
      rule("B2", { type: "list", formula1: '"Low,Medium,High"', input: { show: true, title: "Priority", text: "Pick one" } }),
      rule("B3", { type: "whole", formula1: "1", formula2: "10", error: { show: true, style: "warning", title: "", text: "" } }),
      rule("B8", { type: "list", formula1: '"Yes,No"', showDropdown: false }),
    ],
  };
  return {
    names: [],
    dxfs: [],
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: "Form",
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

async function mountForm() {
  const store = await import("../lib/documentStore");
  const { routes } = await import("../router");
  const { default: SpreadsheetView } = await import("./SpreadsheetView.vue");
  const book = formBook();
  store.storeSpreadsheet(book);
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/sheets/${book.id}`);
  const wrapper = mount(SpreadsheetView, { global: { plugins: [router] }, attachTo: document.body });
  cleanups.push(() => wrapper.unmount());
  await flushPromises();
  return wrapper;
}

type View = Awaited<ReturnType<typeof mountForm>>;

const editorHost = (wrapper: View): HTMLElement => wrapper.get('[data-testid="cell-editor"]').element as HTMLElement;

/** A key on `target` as the browser delivers it. */
async function key(target: Element, init: KeyboardEventInit): Promise<void> {
  target.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }));
  await flushPromises();
}

/** The browser's native edit of the cell editor to `text` (typing over its
 * selection), then its input event. */
async function typeInEditor(wrapper: View, text: string): Promise<void> {
  const host = editorHost(wrapper);
  host.textContent = text;
  const range = document.createRange();
  range.setStart(host.firstChild!, text.length);
  range.collapse(true);
  document.getSelection()!.removeAllRanges();
  document.getSelection()!.addRange(range);
  host.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
  await flushPromises();
}

const cellText = (wrapper: View, ref: string): string => wrapper.get(`[data-ref="${ref}"]`).text();

describe("SPEC-130 in the workbook view", () => {
  it("walks the e2e's path: input message, Alt+Down, a pick, a stop alert's Retry, a warning's No and Escape, a hidden arrow", async () => {
    const wrapper = await mountForm();
    // A1 → B2.
    await key(editorHost(wrapper), { key: "ArrowDown" });
    await key(editorHost(wrapper), { key: "ArrowRight" });
    expect(wrapper.get('[data-testid="name-box"]').attributes("value")).toBe("B2");

    // The input message, describing the active cell.
    const message = wrapper.get('[data-testid="validation-input-message"]');
    expect(message.attributes("role")).toBe("tooltip");
    expect(message.text()).toContain("Priority");
    expect(message.text()).toContain("Pick one");
    expect(wrapper.get('[data-ref="B2"]').attributes("aria-describedby")).toBe(message.attributes("id"));
    expect(wrapper.find('.sheet-overlays [data-testid="validation-input-message"]').exists()).toBe(false);

    // Alt+Down from the cell editor opens the list, focused; the message hides.
    expect(wrapper.get('[data-testid="validation-dropdown"]').attributes("aria-expanded")).toBe("false");
    await key(editorHost(wrapper), { key: "ArrowDown", altKey: true });
    const list = wrapper.get('[data-testid="validation-list"]');
    expect(wrapper.findAll('[data-testid="validation-option"]').map((option) => option.text())).toEqual(["Low", "Medium", "High"]);
    expect(document.activeElement).toBe(list.element);
    expect(wrapper.find('[data-testid="validation-input-message"]').exists()).toBe(false);
    expect(wrapper.get('[data-ref="B2"]').attributes("aria-describedby")).toBeUndefined();

    // ArrowDown, ArrowDown, Enter: Medium, and focus back on the grid's editor.
    await key(list.element, { key: "ArrowDown" });
    await key(list.element, { key: "ArrowDown" });
    await key(wrapper.get('[data-testid="validation-list"]').element, { key: "Enter" });
    expect(wrapper.find('[data-testid="validation-list"]').exists()).toBe(false);
    expect(cellText(wrapper, "B2")).toBe("Medium");
    expect(wrapper.get('[data-testid="name-box"]').attributes("value")).toBe("B2");
    expect(document.activeElement).toBe(editorHost(wrapper));

    // "Urgent" + Enter: the stop alert; nothing written.
    await typeInEditor(wrapper, "Urgent");
    await key(editorHost(wrapper), { key: "Enter" });
    const alert = wrapper.get('[data-testid="validation-alert"]');
    expect([alert.attributes("role"), alert.attributes("data-style")]).toEqual(["alertdialog", "stop"]);
    expect(alert.text()).toContain("Invalid entry");
    expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-retry"]').element);
    expect(cellText(wrapper, "B2")).toBe("Medium");

    // Retry: the cell editor, focused, holds "Urgent" selected.
    await wrapper.get('[data-testid="validation-alert-retry"]').trigger("click");
    await flushPromises();
    expect(wrapper.find('[data-testid="validation-alert"]').exists()).toBe(false);
    expect(document.activeElement).toBe(editorHost(wrapper));
    expect(editorHost(wrapper).textContent).toBe("Urgent");
    expect(document.getSelection()!.toString()).toBe("Urgent");

    // "High" typed over it + Enter: written, the selection moves on to B3.
    await typeInEditor(wrapper, "High");
    await key(editorHost(wrapper), { key: "Enter" });
    expect(cellText(wrapper, "B2")).toBe("High");
    expect(wrapper.get('[data-testid="name-box"]').attributes("value")).toBe("B3");

    // 42 into B3: the warning asks "Continue?"; No goes back to editing with
    // the text selected, and Escape cancels the edit.
    await typeInEditor(wrapper, "42");
    await key(editorHost(wrapper), { key: "Enter" });
    const warning = wrapper.get('[data-testid="validation-alert"]');
    expect(warning.attributes("data-style")).toBe("warning");
    expect(warning.text()).toContain("Continue?");
    await wrapper.get('[data-testid="validation-alert-no"]').trigger("click");
    await flushPromises();
    expect(document.activeElement).toBe(editorHost(wrapper));
    expect(document.getSelection()!.toString()).toBe("42");
    await key(editorHost(wrapper), { key: "Escape" });
    expect(cellText(wrapper, "B3")).toBe("5");
    expect(wrapper.get('[data-testid="cell-editor"]').attributes("data-mode")).toBe("ready");

    // B8: a list whose arrow is hidden.
    for (let i = 0; i < 5; i += 1) {
      await key(editorHost(wrapper), { key: "ArrowDown" });
    }
    expect(wrapper.get('[data-testid="name-box"]').attributes("value")).toBe("B8");
    expect(wrapper.find('[data-testid="validation-dropdown"]').exists()).toBe(false);
    await key(editorHost(wrapper), { key: "ArrowDown", altKey: true });
    expect(wrapper.find('[data-testid="validation-list"]').exists()).toBe(false);
  });

  it("the arrow button opens and closes the list; Escape returns focus to the grid", async () => {
    const wrapper = await mountForm();
    await key(editorHost(wrapper), { key: "ArrowDown" });
    await key(editorHost(wrapper), { key: "ArrowRight" });
    await wrapper.get('[data-testid="validation-dropdown"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="validation-dropdown"]').attributes("aria-expanded")).toBe("true");
    await key(wrapper.get('[data-testid="validation-list"]').element, { key: "Escape" });
    expect(wrapper.find('[data-testid="validation-list"]').exists()).toBe(false);
    expect(document.activeElement).toBe(editorHost(wrapper));
    expect(wrapper.get('[data-testid="name-box"]').attributes("value")).toBe("B2");
    // A click on an option picks it.
    await wrapper.get('[data-testid="validation-dropdown"]').trigger("click");
    await flushPromises();
    await wrapper.findAll('[data-testid="validation-option"]')[2]!.trigger("click");
    await flushPromises();
    expect(cellText(wrapper, "B2")).toBe("High");
  });
});
