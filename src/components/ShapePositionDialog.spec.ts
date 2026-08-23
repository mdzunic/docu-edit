import { describe, it, expect } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import ShapePositionDialog from "./ShapePositionDialog.vue";

// SPEC-068 AC-42 — the numeric "Position & size" dialog: the coarse-pointer
// fallback for dragging a drawing handle. Width/height always editable;
// horizontal/vertical offset editable only for a floating shape.

function fieldValue(wrapper: VueWrapper, selector: string): string {
  return (wrapper.get(selector).element as HTMLInputElement).value;
}

const FLOATING_INITIAL = { offsetXCm: 2, offsetYCm: 3, widthCm: 5, heightCm: 4, inline: false };

describe("ShapePositionDialog (SPEC-068 AC-42)", () => {
  it("prefills width, height and position from the selected shape", () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    expect(fieldValue(wrapper, '[data-testid="shape-position-width"]')).toBe("5");
    expect(fieldValue(wrapper, '[data-testid="shape-position-height"]')).toBe("4");
    expect(fieldValue(wrapper, '[data-testid="shape-position-x"]')).toBe("2");
    expect(fieldValue(wrapper, '[data-testid="shape-position-y"]')).toBe("3");
  });

  it("OK emits the edited size and offset in cm", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-width"]').setValue("8");
    await wrapper.get('[data-testid="shape-position-height"]').setValue("6");
    await wrapper.get('[data-testid="shape-position-x"]').setValue("1.5");
    await wrapper.get('[data-testid="shape-position-y"]').setValue("0.25");
    await wrapper.get('[data-testid="shape-position-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toEqual([
      [{ offsetXCm: 1.5, offsetYCm: 0.25, widthCm: 8, heightCm: 6 }],
    ]);
  });

  it("Cancel emits cancel and never apply", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-width"]').setValue("9");
    await wrapper.get('[data-testid="shape-position-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("blocks OK on an out-of-range width with an error", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-width"]').setValue("999");
    expect(wrapper.get('[data-testid="shape-position-error"]').text()).toContain("Width");
    expect(wrapper.get('[data-testid="shape-position-apply"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="shape-position-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("allows a zero height (a Line is a real horizontal rule)", async () => {
    const wrapper = mount(ShapePositionDialog, {
      props: { initial: { ...FLOATING_INITIAL, heightCm: 0 }, unit: "cm" },
    });
    await wrapper.get('[data-testid="shape-position-height"]').setValue("0");
    expect(wrapper.find('[data-testid="shape-position-error"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="shape-position-apply"]').attributes("disabled")).toBeUndefined();
  });

  it("disables the position fields for an inline shape and emits zero offsets", async () => {
    const wrapper = mount(ShapePositionDialog, {
      props: { initial: { offsetXCm: 0, offsetYCm: 0, widthCm: 5, heightCm: 4, inline: true }, unit: "cm" },
    });
    expect(wrapper.get('[data-testid="shape-position-x"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="shape-position-y"]').attributes("disabled")).toBeDefined();
    expect(wrapper.find('[data-testid="shape-position-inline-note"]').exists()).toBe(true);
    await wrapper.get('[data-testid="shape-position-width"]').setValue("7");
    await wrapper.get('[data-testid="shape-position-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toEqual([
      [{ offsetXCm: 0, offsetYCm: 0, widthCm: 7, heightCm: 4 }],
    ]);
  });
});
