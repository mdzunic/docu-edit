// SPEC-127: the Insert Function dialog — the search field, "All" then the
// categories holding a function in FUNCTION_CATEGORY_ORDER, the list ranked
// by functionHelp's search, the detail pane (signature, summary, every
// argument's description), Up/Down and Enter in the search field, Insert.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { FUNCTION_CATEGORY_ORDER, getFunction } from "../../lib/spreadsheet/formula/registry";
import InsertFunctionDialog from "./InsertFunctionDialog.vue";

describe("InsertFunctionDialog", () => {
  it("offers All, then the categories in the dialog's order", () => {
    const wrapper = mount(InsertFunctionDialog, { attachTo: document.body });
    const options = wrapper.findAll('[data-testid="insert-function-category"] option').map((option) => option.text());
    expect(options[0]).toBe("All");
    expect(options.slice(1)).toEqual(FUNCTION_CATEGORY_ORDER.filter((category) => options.includes(category)));
    expect(options).toContain("Financial");
    wrapper.unmount();
  });

  it("searches within a category and shows the selected function's detail, then inserts it", async () => {
    const wrapper = mount(InsertFunctionDialog, { attachTo: document.body });
    await wrapper.get('[data-testid="insert-function-category"]').setValue("Financial");
    await wrapper.get('[data-testid="insert-function-search"]').setValue("payment");
    const names = wrapper.findAll('[data-testid="insert-function-item"]').map((item) => item.attributes("data-name"));
    expect(names).toContain("PMT");
    expect(names.every((name) => getFunction(name!)?.category === "Financial")).toBe(true);
    await wrapper.get('[data-name="PMT"]').trigger("click");
    const detail = wrapper.get('[data-testid="insert-function-detail"]').text();
    expect(detail).toContain("PMT(rate, nper, pv, [fv], [type])");
    for (const arg of getFunction("PMT")!.args) {
      expect(detail).toContain(arg.description);
    }
    await wrapper.get('[data-testid="insert-function-insert"]').trigger("click");
    expect(wrapper.emitted("insert")).toEqual([["PMT"]]);
    wrapper.unmount();
  });

  it("moves with Up/Down and inserts with Enter from the search field", async () => {
    const wrapper = mount(InsertFunctionDialog, { attachTo: document.body });
    const search = wrapper.get('[data-testid="insert-function-search"]');
    await search.setValue("ROUND");
    await search.trigger("keydown", { key: "ArrowDown" });
    await search.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("insert")).toEqual([["ROUNDDOWN"]]);
    wrapper.unmount();
  });
});
