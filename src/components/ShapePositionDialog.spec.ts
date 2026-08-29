import { describe, it, expect } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import ShapePositionDialog from "./ShapePositionDialog.vue";

// SPEC-072 — the drawing "Position & size" dialog: the numeric-entry path for
// every geometric property a drag adjusts (width, height, horizontal/vertical
// position, rotation), each in the current measurement unit, applied LIVE as
// fields change (one `apply` per committed field => one undo entry). Width and
// height validate against MIN_SHAPE_SIZE_CM (0.2cm). Expands SPEC-068 AC-42.

function fieldValue(wrapper: VueWrapper, selector: string): string {
  return (wrapper.get(selector).element as HTMLInputElement).value;
}

const FLOATING_INITIAL = {
  offsetXCm: 2,
  offsetYCm: 3,
  widthCm: 5,
  heightCm: 4,
  rotationDeg: 0,
  relativeToH: "margin",
  relativeToV: "paragraph",
  inline: false,
  isLine: false,
};

function lastApply(wrapper: VueWrapper) {
  const events = wrapper.emitted("apply");
  return events ? events[events.length - 1][0] : undefined;
}

describe("ShapePositionDialog (SPEC-072)", () => {
  it("prefills width, height, position and rotation from the selected shape", () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    expect(fieldValue(wrapper, '[data-testid="shape-position-width"]')).toBe("5");
    expect(fieldValue(wrapper, '[data-testid="shape-position-height"]')).toBe("4");
    expect(fieldValue(wrapper, '[data-testid="shape-position-x"]')).toBe("2");
    expect(fieldValue(wrapper, '[data-testid="shape-position-y"]')).toBe("3");
    expect(fieldValue(wrapper, '[data-testid="shape-position-rotation"]')).toBe("0");
  });

  it("applies a size change live, in cm (AC-43)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-width"]').setValue("8");
    await wrapper.get('[data-testid="shape-position-width"]').trigger("change");
    expect(lastApply(wrapper)).toMatchObject({ widthCm: 8, heightCm: 4, offsetXCm: 2, offsetYCm: 3 });
  });

  it("applies a position change live and carries the relative-to reference (AC-40/41)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-y"]').setValue("1.25");
    await wrapper.get('[data-testid="shape-position-y"]').trigger("change");
    expect(lastApply(wrapper)).toMatchObject({ offsetYCm: 1.25, relativeToH: "margin", relativeToV: "paragraph" });
  });

  it("applies a rotation change (AC-39)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-rotation"]').setValue("45");
    await wrapper.get('[data-testid="shape-position-rotation"]').trigger("change");
    expect(lastApply(wrapper)).toMatchObject({ rotationDeg: 45 });
  });

  it("keeps proportions when aspect ratio is locked (AC-42)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-lock-aspect"]').setValue(true);
    // Original ratio h/w = 4/5; doubling width to 10 => height 8.
    await wrapper.get('[data-testid="shape-position-width"]').setValue("10");
    await wrapper.get('[data-testid="shape-position-width"]').trigger("change");
    expect(fieldValue(wrapper, '[data-testid="shape-position-height"]')).toBe("8");
    expect(lastApply(wrapper)).toMatchObject({ widthCm: 10, heightCm: 8 });
  });

  it("rejects a width below MIN_SHAPE_SIZE_CM with an error and no apply (AC-44)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-width"]').setValue("0.1");
    await wrapper.get('[data-testid="shape-position-width"]').trigger("change");
    expect(wrapper.text()).toContain("Width must be at least");
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("allows a zero height for a Line (a horizontal rule)", async () => {
    const wrapper = mount(ShapePositionDialog, {
      props: { initial: { ...FLOATING_INITIAL, heightCm: 0, isLine: true }, unit: "cm" },
    });
    await wrapper.get('[data-testid="shape-position-height"]').setValue("0");
    await wrapper.get('[data-testid="shape-position-height"]').trigger("change");
    expect(lastApply(wrapper)).toMatchObject({ heightCm: 0 });
  });

  it("disables the position fields for an inline shape", () => {
    const wrapper = mount(ShapePositionDialog, {
      props: { initial: { ...FLOATING_INITIAL, inline: true }, unit: "cm" },
    });
    expect(wrapper.get('[data-testid="shape-position-x"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="shape-position-y"]').attributes("disabled")).toBeDefined();
    expect(wrapper.find('[data-testid="shape-position-inline-note"]').exists()).toBe(true);
  });

  it("dismiss closes without applying", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.get('[data-testid="shape-position-done"]').trigger("click");
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
  });

  it("re-seeds fields when the selection changes under an open dialog (AC-51)", async () => {
    const wrapper = mount(ShapePositionDialog, { props: { initial: FLOATING_INITIAL, unit: "cm" } });
    await wrapper.setProps({ initial: { ...FLOATING_INITIAL, widthCm: 12, offsetXCm: 1 } });
    expect(fieldValue(wrapper, '[data-testid="shape-position-width"]')).toBe("12");
    expect(fieldValue(wrapper, '[data-testid="shape-position-x"]')).toBe("1");
  });
});
