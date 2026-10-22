import { describe, expect, it } from "vitest";
import {
  transitionAnimates,
  transitionDirection,
  transitionDurationMs,
  transitionStagePair,
} from "./transitionStyles";

describe("transitionStagePair", () => {
  it("fades the incoming stage in over a static outgoing", () => {
    expect(transitionStagePair("fade", "left", "enter").incoming.opacity).toBe("0");
    expect(transitionStagePair("fade", "left", "active").incoming.opacity).toBe("1");
    expect(transitionStagePair("fade", "left", "active").outgoing).toEqual({});
    expect(transitionStagePair("fade", "left", "active").incomingOnTop).toBe(true);
  });

  it("pushes both stages for the direction", () => {
    const enter = transitionStagePair("push", "left", "enter");
    const active = transitionStagePair("push", "left", "active");
    expect(enter.incoming.transform).toBe("translateX(100%)"); // enters from the right
    expect(active.incoming.transform).toBe("translateX(0)");
    expect(enter.outgoing.transform).toBe("translateX(0)");
    expect(active.outgoing.transform).toBe("translateX(-100%)"); // exits left
  });

  it("moves only the incoming stage for cover, only the outgoing for uncover", () => {
    expect(transitionStagePair("cover", "left", "active").outgoing).toEqual({});
    expect(transitionStagePair("cover", "left", "active").incoming.transform).toBe("translateX(0)");
    const uncover = transitionStagePair("uncover", "left", "active");
    expect(uncover.incoming).toEqual({});
    expect(uncover.outgoing.transform).toBe("translateX(-100%)");
    expect(uncover.incomingOnTop).toBe(false); // outgoing slides away over a static incoming
  });

  it("reveals wipe via a shrinking clip inset per direction", () => {
    expect(transitionStagePair("wipe", "left", "enter").incoming.clipPath).toBe("inset(0 0 0 100%)");
    expect(transitionStagePair("wipe", "left", "active").incoming.clipPath).toBe("inset(0 0 0 0)");
    expect(transitionStagePair("wipe", "up", "enter").incoming.clipPath).toBe("inset(100% 0 0 0)");
  });

  it("opens split from the centre outward", () => {
    expect(transitionStagePair("split", "left", "enter").incoming.clipPath).toBe("inset(0 50% 0 50%)");
    expect(transitionStagePair("split", "left", "active").incoming.clipPath).toBe("inset(0 0 0 0)");
  });

  it("is static for none", () => {
    const pair = transitionStagePair("none", "left", "active");
    expect(pair.incoming).toEqual({});
    expect(pair.outgoing).toEqual({});
  });
});

describe("transitionAnimates", () => {
  it("animates a real effect with an outgoing slide and motion allowed", () => {
    expect(transitionAnimates({ type: "fade" }, false, true)).toBe(true);
  });
  it("does not animate none, reduced motion, or a missing outgoing", () => {
    expect(transitionAnimates({ type: "none" }, false, true)).toBe(false);
    expect(transitionAnimates({ type: "fade" }, true, true)).toBe(false);
    expect(transitionAnimates({ type: "fade" }, false, false)).toBe(false);
    expect(transitionAnimates(null, false, true)).toBe(false);
  });
});

describe("transitionDirection / transitionDurationMs", () => {
  it("defaults direction per type and duration to 600 with clamping", () => {
    expect(transitionDirection({ type: "push" })).toBe("left");
    expect(transitionDirection({ type: "push", direction: "down" })).toBe("down");
    expect(transitionDurationMs({})).toBe(600);
    expect(transitionDurationMs({ durationMs: 2000 })).toBe(2000);
    expect(transitionDurationMs({ durationMs: 10 })).toBe(100);
    expect(transitionDurationMs({ durationMs: 9999 })).toBe(5000);
  });
});
