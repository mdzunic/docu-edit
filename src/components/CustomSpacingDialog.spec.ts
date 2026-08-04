import { describe, it, expect } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import CustomSpacingDialog from "./CustomSpacingDialog.vue";

/** A form control's current value — test-utils types get() by selector as a
 * plain Element, so the read goes through one cast here. */
function fieldValue(wrapper: VueWrapper, selector: string): string {
  return (wrapper.get(selector).element as HTMLInputElement).value;
}

const MULTIPLE_INITIAL = { mode: "multiple", value: 1.15, beforePt: 6, afterPt: 8 };

describe("CustomSpacingDialog (SPEC-014 AC-12..14)", () => {
  it("offers the three line-spacing modes plus value and before/after inputs (AC-12)", () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    const options = wrapper
      .get('[data-testid="spacing-mode"]')
      .findAll("option")
      .map((option) => option.text());
    expect(options).toEqual(["Multiple", "Exactly", "At least"]);
    expect(wrapper.find('[data-testid="spacing-value"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="spacing-before"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="spacing-after"]').exists()).toBe(true);
  });

  it("prefills from the first touched paragraph's values", () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    expect(fieldValue(wrapper, '[data-testid="spacing-mode"]')).toBe("multiple");
    expect(fieldValue(wrapper, '[data-testid="spacing-value"]')).toBe("1.15");
    expect(fieldValue(wrapper, '[data-testid="spacing-before"]')).toBe("6");
    expect(fieldValue(wrapper, '[data-testid="spacing-after"]')).toBe("8");
  });

  it("OK emits the edited values (AC-13)", async () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    await wrapper.get('[data-testid="spacing-mode"]').setValue("exact");
    await wrapper.get('[data-testid="spacing-value"]').setValue("14");
    await wrapper.get('[data-testid="spacing-before"]').setValue("12");
    await wrapper.get('[data-testid="spacing-after"]').setValue("0");
    await wrapper.get('[data-testid="spacing-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toEqual([
      [{ mode: "exact", value: 14, beforePt: 12, afterPt: 0 }],
    ]);
  });

  it("Cancel emits cancel and never apply (AC-14)", async () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    await wrapper.get('[data-testid="spacing-value"]').setValue("3");
    await wrapper.get('[data-testid="spacing-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("blocks OK on an out-of-range value with an error message", async () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    await wrapper.get('[data-testid="spacing-value"]').setValue("500");
    expect(wrapper.get('[data-testid="spacing-error"]').text()).toContain("between 0.05 and 132");
    expect(wrapper.get('[data-testid="spacing-apply"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="spacing-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("blocks a negative space before/after", async () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    await wrapper.get('[data-testid="spacing-before"]').setValue("-4");
    expect(wrapper.get('[data-testid="spacing-error"]').text()).toContain("Space before/after");
    expect(wrapper.get('[data-testid="spacing-apply"]').attributes("disabled")).toBeDefined();
  });

  it("resets the value when the mode's unit changes meaning", async () => {
    const wrapper = mount(CustomSpacingDialog, { props: { initial: MULTIPLE_INITIAL } });
    // 1.15 lines is meaningless as an absolute height — flips to 12 pt.
    await wrapper.get('[data-testid="spacing-mode"]').setValue("atLeast");
    expect(fieldValue(wrapper, '[data-testid="spacing-value"]')).toBe("12");
  });
});
