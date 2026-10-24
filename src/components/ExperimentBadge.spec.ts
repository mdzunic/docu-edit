// SPEC-113 — the "Experimental" chip: what it says, and that its tooltip
// links to Labs.
import { mount } from "@vue/test-utils";
import { afterEach, describe, expect, it } from "vitest";
import ExperimentBadge from "./ExperimentBadge.vue";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";

afterEach(() => {
  const dialog = usePreferencesDialog();
  dialog.close();
  dialog.labsRequested.value = false;
});

describe("ExperimentBadge", () => {
  it("is an “Experimental” chip described by its tooltip", () => {
    const wrapper = mount(ExperimentBadge, { props: { experiment: "deck-designer" } });
    const chip = wrapper.get('[data-testid="experiment-badge"]');
    expect(chip.text()).toBe("Experimental");
    expect(chip.attributes("tabindex")).toBe("0");
    const tooltip = wrapper.get('[data-testid="experiment-badge-tooltip"]');
    expect(tooltip.attributes("role")).toBe("tooltip");
    expect(chip.attributes("aria-describedby")).toBe(tooltip.attributes("id"));
    expect(tooltip.text()).toContain("Deck designer is experimental");
    expect(tooltip.text()).toContain("may change or be removed");
  });

  it("links to Labs: the Preferences dialog opens at its Labs section", async () => {
    const wrapper = mount(ExperimentBadge, { props: { experiment: "library-search" } });
    const dialog = usePreferencesDialog();
    await wrapper.get('[data-testid="experiment-badge-labs"]').trigger("click");
    expect(dialog.isOpen.value).toBe(true);
    expect(dialog.labsRequested.value).toBe(true);
  });
});
