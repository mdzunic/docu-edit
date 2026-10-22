// SPEC-130: the input message — a role="tooltip" with the rule's title above
// its text, placed from the size it measures: its natural one (max-content,
// wrapped at 16rem or the cell area's width), so a message on a cell at the
// right edge flips to the cell's left rather than being measured squeezed
// into the strip left there, and re-measured whenever the box resizes.
// jsdom has no layout: the box's offset size is stubbed.
import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, describe, expect, it, vi } from "vitest";
import ValidationInputMessage from "./ValidationInputMessage.vue";
import { INPUT_MESSAGE_GAP_PX, VALIDATION_INPUT_MESSAGE_ID } from "../../views/spreadsheet/validationChrome";

const BOUNDS = { width: 600, height: 400 };
/** A cell at the area's right edge. */
const EDGE = { left: 520, top: 40, width: 64, height: 20 };

const mounted: Array<() => void> = [];
afterEach(() => {
  for (const unmount of mounted.splice(0)) {
    unmount();
  }
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

/** Stubs every element's offset size (jsdom lays nothing out). */
function boxSize(width: number, height: number) {
  vi.spyOn(HTMLElement.prototype, "offsetWidth", "get").mockReturnValue(width);
  vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockReturnValue(height);
}

async function messageOf(props: Record<string, unknown> = {}) {
  const host = document.createElement("div");
  document.body.append(host);
  const wrapper = mount(ValidationInputMessage, {
    props: { title: "Score", text: "Enter a whole number\nbetween 1 and 10", cellRect: EDGE, bounds: BOUNDS, ...props },
    attachTo: host,
  });
  mounted.push(() => {
    wrapper.unmount();
    host.remove();
  });
  await flushPromises();
  return wrapper;
}

const styleOf = (wrapper: Awaited<ReturnType<typeof messageOf>>) => (wrapper.element as HTMLElement).style;

describe("ValidationInputMessage", () => {
  it("is a tooltip with the title above the text, line breaks kept", async () => {
    boxSize(120, 40);
    const wrapper = await messageOf();
    expect(wrapper.attributes("role")).toBe("tooltip");
    expect(wrapper.attributes("id")).toBe(VALIDATION_INPUT_MESSAGE_ID);
    expect(wrapper.get('[data-testid="validation-input-title"]').text()).toBe("Score");
    expect(wrapper.get('[data-testid="validation-input-text"]').element.textContent).toContain("Enter a whole number\nbetween 1 and 10");
  });

  it("wraps at no more than the cell area's width", async () => {
    boxSize(120, 40);
    const wrapper = await messageOf({ bounds: { width: 180, height: 400 } });
    expect(styleOf(wrapper).getPropertyValue("--validation-input-room")).toBe("180px");
  });

  it("measured at its natural width on a right-edge cell, flips to the cell's left and stays below", async () => {
    boxSize(221, 40);
    const wrapper = await messageOf();
    expect(styleOf(wrapper).left).toBe(`${520 - 221}px`);
    expect(styleOf(wrapper).top).toBe(`${60 + INPUT_MESSAGE_GAP_PX}px`);
  });

  it("is re-measured when the box resizes", async () => {
    const observed: Array<() => void> = [];
    vi.stubGlobal(
      "ResizeObserver",
      class {
        constructor(callback: () => void) {
          observed.push(callback);
        }
        observe() {}
        disconnect() {}
      },
    );
    boxSize(40, 40);
    const wrapper = await messageOf({ bounds: { width: 800, height: 400 } });
    expect(styleOf(wrapper).left).toBe(`${520 + 64}px`);
    // Wider now (its text or its font changed): no longer room on the right.
    boxSize(300, 40);
    for (const callback of observed) {
      callback();
    }
    await flushPromises();
    expect(styleOf(wrapper).left).toBe(`${520 - 300}px`);
  });
});
