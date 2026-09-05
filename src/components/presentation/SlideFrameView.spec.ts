// SPEC-096 — group-frame rendering geometry: the group box occupies its REAL
// slide-space rectangle (a rotate transform must spin about the group's own
// centre, AC-45/AC-31) while its children, stored in absolute slide space,
// re-express through the origin shim. A regression here once pinned the box
// to a 0×0 corner, which rotated children about the STAGE origin and made
// verbatim groups unreachable to a click (QA round 1).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SlideFrameView from "./SlideFrameView.vue";
import { cmPx } from "./slideRender";

const child = { id: "c1", xCm: 3, yCm: 4, widthCm: 2, heightCm: 1, kind: "shape", geometry: "rect" };

function mountGroup(rotationDeg?: number) {
  const frame = {
    id: "g1",
    xCm: 2,
    yCm: 3,
    widthCm: 6,
    heightCm: 4,
    kind: "group",
    children: [child],
    ...(rotationDeg === undefined ? {} : { rotationDeg }),
  };
  return mount(SlideFrameView, { props: { frame, resolveMedia: () => "" } });
}

describe("SlideFrameView group frame (SPEC-096)", () => {
  it("the group box carries its real slide-space geometry, not a zeroed one", () => {
    const wrapper = mountGroup();
    const box = wrapper.get(".group-frame").element as HTMLElement;
    expect(box.style.left).toBe(cmPx(2));
    expect(box.style.top).toBe(cmPx(3));
    expect(box.style.width).toBe(cmPx(6));
    expect(box.style.height).toBe(cmPx(4));
  });

  it("a rotated group rotates about its own centre (transform on the real box)", () => {
    const wrapper = mountGroup(37);
    const box = wrapper.get(".group-frame").element as HTMLElement;
    expect(box.style.transform).toBe("rotate(37deg)");
    expect(box.style.transformOrigin).toBe("center");
    expect(box.style.width).toBe(cmPx(6));
  });

  it("the origin shim shifts children's slide-space coordinates into the box", () => {
    const wrapper = mountGroup();
    const shim = wrapper.get(".group-origin").element as HTMLElement;
    expect(shim.style.left).toBe(cmPx(-2));
    expect(shim.style.top).toBe(cmPx(-3));
    // The child keeps its own absolute slide-space position inside the shim,
    // so it lands at (3-2, 4-3)cm within the group box — unchanged on screen.
    const childEl = wrapper.get(".group-origin .frame").element as HTMLElement;
    expect(childEl.style.left).toBe(cmPx(3));
    expect(childEl.style.top).toBe(cmPx(4));
  });
});

const mediaFrame = (over = {}) => ({
  id: "m1",
  xCm: 1,
  yCm: 1,
  widthCm: 16,
  heightCm: 9,
  kind: "media",
  media: "video",
  ...over,
});

describe("SlideFrameView media frame (SPEC-095)", () => {
  it("renders a slide-media-frame with the poster when one exists (AC-18)", () => {
    const wrapper = mount(SlideFrameView, {
      props: {
        frame: mediaFrame({ mediaId: "m.mp4", posterMediaId: "p", contentType: "video/mp4" }),
        resolveMedia: (id: string) => `data:image/png;base64,${id}`,
      },
    });
    expect(wrapper.find("[data-testid='slide-media-frame']").exists()).toBe(true);
    expect(wrapper.find("[data-testid='media-poster']").exists()).toBe(true);
  });

  it("shows the link placeholder for a linked clip (AC-32)", () => {
    const wrapper = mount(SlideFrameView, {
      props: {
        frame: mediaFrame({ linkTarget: "../far/remote.mp4", contentType: "video/mp4" }),
        resolveMedia: () => "",
      },
    });
    expect(wrapper.find("[data-testid='media-link-placeholder']").exists()).toBe(true);
    expect(wrapper.text()).toContain("remote.mp4");
  });

  it("shows the oversize chip for an oversize clip (AC-11)", () => {
    const wrapper = mount(SlideFrameView, {
      props: {
        frame: mediaFrame({ mediaId: "big.mp4", oversize: true, contentType: "video/mp4" }),
        resolveMedia: () => "",
      },
    });
    expect(wrapper.find("[data-testid='media-oversize']").text()).toContain("Not loaded");
  });

  it("shows media-unsupported when the browser can't play the codec (AC-33)", () => {
    // jsdom's canPlayType returns "" for everything, so any embedded clip is
    // unsupported here — the poster still shows.
    const wrapper = mount(SlideFrameView, {
      props: {
        frame: mediaFrame({ mediaId: "m.mp4", posterMediaId: "p", contentType: "video/mp4" }),
        resolveMedia: (id: string) => `data:image/png;base64,${id}`,
      },
    });
    expect(wrapper.find("[data-testid='media-unsupported']").exists()).toBe(true);
    expect(wrapper.find("[data-testid='media-poster']").exists()).toBe(true);
  });
});
