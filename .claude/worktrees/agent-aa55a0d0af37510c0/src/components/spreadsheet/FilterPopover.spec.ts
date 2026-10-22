// SPEC-128's filter popover: its test ids and labels, "(Select All)"'s
// three states, the search narrowing the list (and what OK keeps while it
// does), OK disabled with nothing checked, "Not all items showing", the
// conditions submenu, the preserved-criterion message, and Escape / Cancel /
// a press outside.
import { afterEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { PRESERVED_FILTER_MESSAGE } from "../../views/spreadsheet/useFilterPopover";
import FilterPopover from "./FilterPopover.vue";

const ENTRIES = [
  { key: "v0", text: "East", checked: true },
  { key: "v1", text: "North", checked: true },
  { key: "v2", text: "West", checked: true },
  { key: "blanks", text: "(Blanks)", checked: true },
];

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

function mountPopover(props: Record<string, unknown> = {}) {
  const wrapper = mount(FilterPopover, {
    attachTo: document.body,
    props: {
      title: "Region",
      conditions: [
        { id: "equals", label: "Equals…" },
        { id: "custom", label: "Custom Filter…" },
      ],
      entries: ENTRIES,
      ...props,
    },
  });
  mounted.push(wrapper);
  return wrapper;
}

const values = (wrapper: VueWrapper) => wrapper.findAll('[data-testid="filter-value"]');

describe("FilterPopover (SPEC-128)", () => {
  it("is a dialog named for its column with the sorts, Clear filter, the conditions, search, (Select All), one checkbox per entry, OK and Cancel", () => {
    const wrapper = mountPopover({ sortAscLabel: "Sort Smallest to Largest", sortDescLabel: "Sort Largest to Smallest", conditionsLabel: "Number Filters" });
    const root = wrapper.get('[data-testid="filter-popover"]');
    expect(root.attributes("role")).toBe("dialog");
    expect(root.attributes("aria-label")).toBe('Filter "Region"');
    expect(wrapper.get('[data-testid="filter-sort-asc"]').text()).toBe("Sort Smallest to Largest");
    expect(wrapper.get('[data-testid="filter-sort-desc"]').text()).toBe("Sort Largest to Smallest");
    const clear = wrapper.get('[data-testid="filter-clear"]');
    expect(clear.text()).toBe('Clear filter from "Region"');
    expect(clear.attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="filter-conditions"]').text()).toContain("Number Filters");
    expect(wrapper.find('[data-testid="filter-search"]').exists()).toBe(true);
    const all = wrapper.get('[data-testid="filter-select-all"]');
    expect(all.attributes("role")).toBe("checkbox");
    expect(all.text()).toContain("(Select All)");
    expect(all.attributes("aria-checked")).toBe("true");
    expect(values(wrapper).map((value) => value.text())).toEqual(["East", "North", "West", "(Blanks)"]);
    expect(values(wrapper).every((value) => value.attributes("role") === "checkbox" && value.attributes("aria-checked") === "true")).toBe(true);
    expect(wrapper.find('[data-testid="filter-not-all"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="filter-ok"]').text()).toBe("OK");
    expect(wrapper.get('[data-testid="filter-cancel"]').text()).toBe("Cancel");
  });

  it("(Select All) is tri-state: unchecking one item makes it mixed, a click checks all, another unchecks all and disables OK", async () => {
    const wrapper = mountPopover({ hasCriteria: true });
    expect(wrapper.get('[data-testid="filter-clear"]').attributes("disabled")).toBeUndefined();
    await values(wrapper)[2]!.trigger("click");
    const all = () => wrapper.get('[data-testid="filter-select-all"]');
    expect(values(wrapper)[2]!.attributes("aria-checked")).toBe("false");
    expect(all().attributes("aria-checked")).toBe("mixed");
    await all().trigger("click");
    expect(all().attributes("aria-checked")).toBe("true");
    await all().trigger("click");
    expect(all().attributes("aria-checked")).toBe("false");
    expect(wrapper.get('[data-testid="filter-ok"]').attributes("disabled")).toBeDefined();
    await values(wrapper)[0]!.trigger("click");
    await wrapper.get('[data-testid="filter-ok"]').trigger("click");
    expect(wrapper.emitted("apply")).toEqual([[["v0"]]]);
  });

  it("the search narrows the list case-insensitively, (Select All) reads (Select All Search Results), and OK keeps only the checked matches", async () => {
    const wrapper = mountPopover();
    await wrapper.get('[data-testid="filter-search"]').setValue("ST");
    expect(values(wrapper).map((value) => value.text())).toEqual(["East", "West"]);
    expect(wrapper.get('[data-testid="filter-select-all"]').text()).toContain("(Select All Search Results)");
    await values(wrapper)[1]!.trigger("click");
    await wrapper.get('[data-testid="filter-ok"]').trigger("click");
    expect(wrapper.emitted("apply")).toEqual([[["v0"]]]);
    await wrapper.get('[data-testid="filter-search"]').setValue("");
    expect(wrapper.get('[data-testid="filter-select-all"]').text()).toContain("(Select All)");
    expect(values(wrapper)).toHaveLength(4);
  });

  it("says 'Not all items showing' when the list was cut", () => {
    expect(mountPopover({ truncated: true }).get('[data-testid="filter-not-all"]').text()).toBe("Not all items showing");
  });

  it("emits the sorts, Clear filter and a condition; the conditions submenu lists its items", async () => {
    const wrapper = mountPopover({ hasCriteria: true });
    await wrapper.get('[data-testid="filter-sort-asc"]').trigger("click");
    await wrapper.get('[data-testid="filter-sort-desc"]').trigger("click");
    expect(wrapper.emitted("sort")).toEqual([[false], [true]]);
    await wrapper.get('[data-testid="filter-clear"]').trigger("click");
    expect(wrapper.emitted("clear")).toHaveLength(1);
    const trigger = wrapper.get('[data-testid="filter-conditions"]');
    expect(trigger.attributes("aria-expanded")).toBe("false");
    await trigger.trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("true");
    const items = wrapper.findAll('[data-testid="filter-condition"]');
    expect(items.map((item) => [item.attributes("role"), item.text()])).toEqual([
      ["menuitem", "Equals…"],
      ["menuitem", "Custom Filter…"],
    ]);
    await items[1]!.trigger("click");
    expect(wrapper.emitted("condition")).toEqual([["custom"]]);
  });

  it("a column filtered by colour or icon shows only its message and Clear filter", async () => {
    const wrapper = mountPopover({ preserved: true, hasCriteria: true, entries: [] });
    expect(wrapper.get('[data-testid="filter-preserved"]').text()).toBe(PRESERVED_FILTER_MESSAGE);
    expect(PRESERVED_FILTER_MESSAGE).toBe("Filtered by colour or icon — this filter can't be edited here");
    for (const testid of ["filter-sort-asc", "filter-conditions", "filter-search", "filter-select-all", "filter-ok"]) {
      expect(wrapper.find(`[data-testid="${testid}"]`).exists()).toBe(false);
    }
    await wrapper.get('[data-testid="filter-clear"]').trigger("click");
    expect(wrapper.emitted("clear")).toHaveLength(1);
  });

  it("Escape, Cancel and a press outside it cancel — saying how; a press on a filter button does not", async () => {
    const wrapper = mountPopover();
    await wrapper.get('[data-testid="filter-search"]').trigger("keydown", { key: "Escape" });
    await wrapper.get('[data-testid="filter-cancel"]').trigger("click");
    const button = document.createElement("button");
    button.dataset.testid = "filter-button";
    document.body.append(button);
    button.dispatchEvent(new Event("pointerdown", { bubbles: true }));
    document.body.dispatchEvent(new Event("pointerdown", { bubbles: true }));
    button.remove();
    expect(wrapper.emitted("cancel")).toEqual([["escape"], ["done"], ["outside"]]);
  });
});
