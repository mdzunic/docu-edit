import { describe, it, expect, afterEach } from "vitest";
import { mount } from "@vue/test-utils";
import ShapeToolbar from "./ShapeToolbar.vue";
import { setTierWidthForTest, FULL_MIN_WIDTH } from "../composables/useResponsiveTier";

// SPEC-068 AC-12/13 & AC-42 — the shape toolbar carries the "Position & size…"
// button, the numeric fallback for a handle drag. That precise editor is one
// of the commands disabled (with a reason) at the compact tier; the shape can
// still be moved by drag and deleted there.

const BASE_PROPS = { wrap: "square" };

describe("ShapeToolbar — position dialog entry point (SPEC-068 AC-42, AC-12/13)", () => {
  afterEach(() => setTierWidthForTest(FULL_MIN_WIDTH));

  it('exposes a "Position & size…" button that emits options', async () => {
    setTierWidthForTest(FULL_MIN_WIDTH);
    const wrapper = mount(ShapeToolbar, { props: BASE_PROPS });
    const button = wrapper.get('[data-testid="shape-position-button"]');
    expect(button.attributes("disabled")).toBeUndefined();
    await button.trigger("click");
    expect(wrapper.emitted("options")).toHaveLength(1);
  });

  it("disables the position button with a stated reason at the compact tier", () => {
    setTierWidthForTest(390);
    const wrapper = mount(ShapeToolbar, { props: BASE_PROPS });
    const button = wrapper.get('[data-testid="shape-position-button"]');
    expect(button.attributes("disabled")).toBeDefined();
    expect(button.attributes("title")).toBe("not available on this screen size");
  });

  it("keeps delete/duplicate enabled at compact — a shape stays movable and deletable", () => {
    setTierWidthForTest(390);
    const wrapper = mount(ShapeToolbar, { props: BASE_PROPS });
    expect(wrapper.get('[data-testid="shape-delete-button"]').attributes("disabled")).toBeUndefined();
    expect(wrapper.get('[data-testid="shape-duplicate-button"]').attributes("disabled")).toBeUndefined();
  });
});
