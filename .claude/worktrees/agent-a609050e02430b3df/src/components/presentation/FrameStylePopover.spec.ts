// View spec: the style popover's control surface and commit events (AC-37/40).
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import FrameStylePopover from "./FrameStylePopover.vue";
import type { ImageFrame, ShapeFrame } from "../../lib/presentation/model";

const shape: ShapeFrame = {
  id: "s1",
  xCm: 0,
  yCm: 0,
  widthCm: 3,
  heightCm: 2,
  kind: "shape",
  geometry: "rect",
  fill: { type: "solid", color: "#4472c4" },
  stroke: { color: "#2f528f", widthPt: 1 },
};

const image: ImageFrame = { id: "i1", xCm: 0, yCm: 0, widthCm: 3, heightCm: 2, kind: "image", mediaId: "ppt/media/image1.png" };

describe("FrameStylePopover", () => {
  it("shows fill + stroke for a shape and emits on change", async () => {
    const wrapper = mount(FrameStylePopover, { props: { frame: shape } });
    expect(wrapper.find('[data-testid="style-fill-color"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="style-stroke-color"]').exists()).toBe(true);

    await wrapper.find('[data-testid="style-no-fill"]').trigger("click");
    expect(wrapper.emitted("set-fill")?.[0]).toEqual([{ type: "none" }]);

    const width = wrapper.find('[data-testid="style-stroke-width"]');
    (width.element as HTMLInputElement).value = "4";
    await width.trigger("change");
    expect(wrapper.emitted("set-stroke")?.[0]).toEqual([{ color: "#2f528f", widthPt: 4 }]);
  });

  it("exposes only stroke controls for an image (AC-40)", () => {
    const wrapper = mount(FrameStylePopover, { props: { frame: image } });
    expect(wrapper.find('[data-testid="style-fill-color"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="style-no-fill"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="style-stroke-color"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="style-stroke-width"]').exists()).toBe(true);
  });

  it("toggles stroke off and back on", async () => {
    const wrapper = mount(FrameStylePopover, { props: { frame: shape } });
    await wrapper.find('[data-testid="style-stroke-toggle"]').trigger("click");
    expect(wrapper.emitted("set-stroke")?.[0]).toEqual([null]);
  });
});
