// SPEC-121 — the workbook app bar: the undo/redo pair under its own test ids
// (the text editor's defaults unchanged), the split Download with its copy
// item (disabled with a reason), and Close. SPEC-126 — the current sheet as
// CSV/TSV, and a csv/tsv workbook's .ods sibling copy.
import { afterEach, describe, expect, it } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import SpreadsheetAppBar from "./SpreadsheetAppBar.vue";
import SpreadsheetDownloadButton from "./SpreadsheetDownloadButton.vue";
import UndoRedoControls from "../UndoRedoControls.vue";
import { shortcutLabel } from "../../lib/shortcuts";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

function mountBar(props: Record<string, unknown> = {}) {
  const wrapper = mount(SpreadsheetAppBar, {
    props: { formatLabel: "Excel workbook (.xlsx)", copyLabel: "Download a copy as OpenDocument Spreadsheet (.ods)…", ...props },
    attachTo: document.body,
  });
  mounted.push(wrapper);
  return wrapper;
}

describe("SpreadsheetAppBar (SPEC-121)", () => {
  it("renders undo/redo as spreadsheet-undo/-redo with the step in the tooltip, and emits them", async () => {
    const wrapper = mountBar({ canUndo: true, undoLabel: "Rename sheet", canRedo: false });
    const undo = wrapper.get('[data-testid="spreadsheet-undo"]');
    expect(undo.attributes("title")).toBe(`Undo Rename sheet (${shortcutLabel("edit.undo")})`);
    expect(wrapper.get('[data-testid="spreadsheet-redo"]').attributes("disabled")).toBeDefined();
    await undo.trigger("click");
    expect(wrapper.emitted("undo")).toHaveLength(1);
    await wrapper.setProps({ canRedo: true, redoLabel: "Add sheet" });
    await wrapper.get('[data-testid="spreadsheet-redo"]').trigger("click");
    expect(wrapper.emitted("redo")).toHaveLength(1);
    expect(wrapper.get('[data-testid="spreadsheet-format"]').text()).toBe("Excel workbook (.xlsx)");
  });

  it("keeps the text editor's undo-button/redo-button ids by default", () => {
    const wrapper = mount(UndoRedoControls, { props: { canUndo: true } });
    expect(wrapper.find('[data-testid="undo-button"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="redo-button"]').exists()).toBe(true);
  });

  it("splits Download: the primary downloads, the caret's menu downloads a copy and closes", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="spreadsheet-download"]').trigger("click");
    expect(wrapper.emitted("download")).toHaveLength(1);
    const caret = wrapper.get('[data-testid="spreadsheet-download-menu"]');
    expect(caret.attributes("aria-haspopup")).toBe("menu");
    expect(caret.attributes("aria-label")).toBe("More download options");
    expect(caret.attributes("aria-expanded")).toBe("false");
    await caret.trigger("click");
    await flushPromises();
    expect(caret.attributes("aria-expanded")).toBe("true");
    const item = wrapper.get('[data-testid="spreadsheet-download-copy"]');
    expect(item.attributes("role")).toBe("menuitem");
    expect(item.text()).toBe("Download a copy as OpenDocument Spreadsheet (.ods)…");
    expect(document.activeElement).toBe(item.element);
    await item.trigger("click");
    expect(wrapper.emitted("download-copy")).toHaveLength(1);
    expect(wrapper.find('[data-testid="spreadsheet-download-copy"]').exists()).toBe(false);
    await wrapper.get('[data-testid="close-document"]').trigger("click");
    expect(wrapper.emitted("close")).toHaveLength(1);
  });

  it("disables the copy with its reason as the title, and closes the menu on Escape", async () => {
    const wrapper = mount(SpreadsheetDownloadButton, {
      props: { copyLabel: "Download a copy as Excel workbook (.xlsx)…", copyDisabledReason: "Not available for this workbook yet" },
      attachTo: document.body,
    });
    mounted.push(wrapper);
    await wrapper.get('[data-testid="spreadsheet-download-menu"]').trigger("click");
    const item = wrapper.get('[data-testid="spreadsheet-download-copy"]');
    expect(item.attributes("disabled")).toBeDefined();
    expect(item.attributes("title")).toBe("Not available for this workbook yet");
    expect(item.text()).toBe("Download a copy as Excel workbook (.xlsx)…");
    await item.trigger("click");
    expect(wrapper.emitted("download-copy")).toBeUndefined();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    await flushPromises();
    expect(wrapper.find('[data-testid="spreadsheet-download-copy"]').exists()).toBe(false);
  });

  it("offers the current sheet as CSV and TSV in every workbook's menu (SPEC-126)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="spreadsheet-download-menu"]').trigger("click");
    await flushPromises();
    expect(wrapper.find('[data-testid="spreadsheet-download-copy-ods"]').exists()).toBe(false);
    const items = wrapper.findAll('[role="menuitem"]').map((item) => item.attributes("data-testid"));
    // SPEC-136: "PDF (.pdf)…" sits before the sheet copies.
    expect(items).toEqual(["spreadsheet-download-copy", "spreadsheet-export-pdf", "spreadsheet-download-copy-csv", "spreadsheet-download-copy-tsv"]);
    expect(wrapper.get('[data-testid="spreadsheet-download-copy-csv"]').text()).toBe("Download a copy as CSV (current sheet)");
    expect(wrapper.get('[data-testid="spreadsheet-download-copy-tsv"]').text()).toBe("Download a copy as TSV (current sheet)");
    await wrapper.get('[data-testid="spreadsheet-download-copy-csv"]').trigger("click");
    expect(wrapper.emitted("download-copy-csv")).toHaveLength(1);
    expect(wrapper.find('[data-testid="spreadsheet-download-copy-csv"]').exists()).toBe(false);
    await wrapper.get('[data-testid="spreadsheet-download-menu"]').trigger("click");
    await wrapper.get('[data-testid="spreadsheet-download-copy-tsv"]').trigger("click");
    expect(wrapper.emitted("download-copy-tsv")).toHaveLength(1);
  });

  it("offers a csv/tsv workbook both workbook copies, .xlsx then .ods, and walks the items with the arrow keys (SPEC-126)", async () => {
    const wrapper = mountBar({
      formatLabel: "Comma-separated values (.csv)",
      copyLabel: "Download a copy as Excel workbook (.xlsx)…",
      copyOdsLabel: "Download a copy as OpenDocument Spreadsheet (.ods)…",
    });
    await wrapper.get('[data-testid="spreadsheet-download-menu"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="spreadsheet-download-copy"]').text()).toBe("Download a copy as Excel workbook (.xlsx)…");
    const ods = wrapper.get('[data-testid="spreadsheet-download-copy-ods"]');
    expect(ods.text()).toBe("Download a copy as OpenDocument Spreadsheet (.ods)…");
    expect(document.activeElement).toBe(wrapper.get('[data-testid="spreadsheet-download-copy"]').element);
    await wrapper.get('[data-testid="spreadsheet-download-menu-popup"]').trigger("keydown", { key: "ArrowDown" });
    expect(document.activeElement).toBe(ods.element);
    await wrapper.get('[data-testid="spreadsheet-download-menu-popup"]').trigger("keydown", { key: "ArrowUp" });
    await wrapper.get('[data-testid="spreadsheet-download-menu-popup"]').trigger("keydown", { key: "ArrowUp" });
    expect(document.activeElement).toBe(wrapper.get('[data-testid="spreadsheet-download-copy-tsv"]').element);
    await ods.trigger("click");
    expect(wrapper.emitted("download-copy-ods")).toHaveLength(1);
    expect(wrapper.emitted("download-copy")).toBeUndefined();
  });

  it("shows only Close for a missing workbook", () => {
    const wrapper = mountBar({ formatLabel: "" });
    expect(wrapper.find('[data-testid="spreadsheet-undo"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="spreadsheet-download"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="close-document"]').exists()).toBe(true);
  });
});
