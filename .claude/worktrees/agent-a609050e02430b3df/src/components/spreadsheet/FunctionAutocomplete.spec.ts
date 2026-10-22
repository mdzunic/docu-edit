// SPEC-127: the function autocomplete — a listbox of function-autocomplete-
// items with data-name, the highlighted one selected, its summary below, and
// a press choosing a name.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import FunctionAutocomplete from "./FunctionAutocomplete.vue";

describe("FunctionAutocomplete", () => {
  it("lists the names as options, marks the highlighted one and shows its summary", () => {
    const wrapper = mount(FunctionAutocomplete, { props: { names: ["ROUND", "ROUNDDOWN"], highlighted: "ROUNDDOWN", summary: "Rounds toward zero." } });
    const list = wrapper.get('[data-testid="function-autocomplete"]');
    expect(list.attributes("role")).toBe("listbox");
    const items = wrapper.findAll('[data-testid="function-autocomplete-item"]');
    expect(items.map((item) => item.attributes("data-name"))).toEqual(["ROUND", "ROUNDDOWN"]);
    expect(items.map((item) => item.attributes("aria-selected"))).toEqual(["false", "true"]);
    expect(items[1]!.attributes("id")).toBe("function-autocomplete-ROUNDDOWN");
    expect(wrapper.get('[data-testid="function-autocomplete-summary"]').text()).toBe("Rounds toward zero.");
  });

  it("chooses a name on a press", async () => {
    const wrapper = mount(FunctionAutocomplete, { props: { names: ["SUM"], highlighted: "SUM" } });
    await wrapper.get('[data-name="SUM"]').trigger("mousedown");
    expect(wrapper.emitted("choose")).toEqual([["SUM"]]);
  });
});
