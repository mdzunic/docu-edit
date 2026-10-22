// SPEC-128's custom-filter and Top 10 dialogs and the sheet-alert dialog:
// their test ids, pre-set values, validation and what OK emits.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import FilterCustomDialog from "./FilterCustomDialog.vue";
import FilterTop10Dialog from "./FilterTop10Dialog.vue";
import SheetAlertDialog from "./SheetAlertDialog.vue";

describe("FilterCustomDialog (SPEC-128)", () => {
  it("opens pre-set, words the operators for a date column, and emits both conditions joined as chosen", async () => {
    const wrapper = mount(FilterCustomDialog, {
      props: { title: "Date", kind: "date", draft: { op1: "greaterThanOrEqual", value1: "", join: "and", op2: "lessThanOrEqual", value2: "" } },
    });
    expect(wrapper.find('[data-testid="filter-custom-dialog"]').exists()).toBe(true);
    const op1 = wrapper.get<HTMLSelectElement>('[data-testid="filter-custom-op-1"]');
    expect(op1.element.value).toBe("greaterThanOrEqual");
    expect(op1.findAll("option").map((option) => option.text())).toContain("is after or equal to");
    expect(wrapper.get<HTMLSelectElement>('[data-testid="filter-custom-op-2"]').element.value).toBe("lessThanOrEqual");
    expect(wrapper.get('[data-testid="filter-custom-ok"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="filter-custom-value-1"]').setValue("1/1/2024");
    await wrapper.get('[data-testid="filter-custom-value-2"]').setValue("12/31/2024");
    await wrapper.get('[data-testid="filter-custom-join"]').setValue("or");
    await wrapper.get('[data-testid="filter-custom-ok"]').trigger("click");
    expect(wrapper.emitted("confirm")).toEqual([[{ op1: "greaterThanOrEqual", value1: "1/1/2024", join: "or", op2: "lessThanOrEqual", value2: "12/31/2024" }]]);
    await wrapper.get('[data-testid="filter-custom-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});

describe("FilterTop10Dialog (SPEC-128)", () => {
  it("offers Top/Bottom, a value and Items/Percent, refusing values out of range", async () => {
    const wrapper = mount(FilterTop10Dialog, { props: { title: "Amount" } });
    expect(wrapper.find('[data-testid="filter-top10-dialog"]').exists()).toBe(true);
    expect(wrapper.get<HTMLInputElement>('[data-testid="filter-top10-value"]').element.value).toBe("10");
    await wrapper.get('[data-testid="filter-top10-value"]').setValue("200");
    await wrapper.get('[data-testid="filter-top10-unit"]').setValue("percent");
    expect(wrapper.get('[data-testid="filter-top10-ok"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="filter-top10-hint"]').text()).toBe("Enter a whole number from 1 to 100");
    await wrapper.get('[data-testid="filter-top10-value"]').setValue("25");
    await wrapper.get('[data-testid="filter-top10-direction"]').setValue("bottom");
    await wrapper.get('[data-testid="filter-top10-ok"]').trigger("click");
    expect(wrapper.emitted("confirm")).toEqual([[{ top: false, percent: true, value: "25" }]]);
  });
});

describe("SheetAlertDialog (SPEC-128)", () => {
  it("is a modal alertdialog named by its message with one OK; OK and Escape close it", async () => {
    const wrapper = mount(SheetAlertDialog, { props: { message: "Select a cell inside the data you want to filter" }, attachTo: document.body });
    const alert = wrapper.get('[data-testid="sheet-alert"]');
    expect(alert.attributes("role")).toBe("alertdialog");
    expect(alert.attributes("aria-modal")).toBe("true");
    expect(document.getElementById(alert.attributes("aria-labelledby")!)?.textContent?.trim()).toBe("Select a cell inside the data you want to filter");
    expect(wrapper.findAll("button")).toHaveLength(1);
    await wrapper.get('[data-testid="sheet-alert-ok"]').trigger("click");
    await alert.trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("close")).toHaveLength(2);
    wrapper.unmount();
  });

  it("keeps its keys from the window's shortcut listeners: Ctrl/Cmd+Z pressed in it never undoes behind it", async () => {
    const wrapper = mount(SheetAlertDialog, { props: { message: "Cannot sort" }, attachTo: document.body });
    const seen: string[] = [];
    const listener = (event: KeyboardEvent): void => {
      seen.push(event.key);
    };
    window.addEventListener("keydown", listener);
    await wrapper.get('[data-testid="sheet-alert-ok"]').trigger("keydown", { key: "z", ctrlKey: true, metaKey: true });
    await wrapper.get('[data-testid="sheet-alert"]').trigger("keydown", { key: "Escape" });
    window.removeEventListener("keydown", listener);
    expect(seen).toEqual([]);
    expect(wrapper.emitted("close")).toHaveLength(1);
    wrapper.unmount();
  });
});
