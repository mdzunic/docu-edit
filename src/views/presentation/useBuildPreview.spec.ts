// SPEC-086 build engine behaviours: hidden-until-built, group playback settling,
// reduced motion, and mid-build instant-complete. Uses fake timers so the
// scheduled animation offsets resolve deterministically.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { Slide, SlideAnimation } from "../../lib/presentation/model";
import { useBuildPreview } from "./useBuildPreview";

function slide(animations: SlideAnimation[]): Slide {
  return {
    id: "s",
    frames: [
      { id: "a", kind: "text", xCm: 0, yCm: 0, widthCm: 5, heightCm: 2, paragraphs: [] },
      { id: "b", kind: "shape", xCm: 6, yCm: 0, widthCm: 5, heightCm: 2, geometry: "rect" },
      { id: "c", kind: "text", xCm: 0, yCm: 5, widthCm: 5, heightCm: 2, paragraphs: [] },
    ],
    animations,
  };
}

const ANIMS: SlideAnimation[] = [
  { frameId: "a", effect: "fadeIn", trigger: "click" },
  { frameId: "b", effect: "flyIn", trigger: "afterPrevious" },
];

beforeEach(() => {
  vi.useFakeTimers();
  // Deterministic rAF for the from→to cross-fade.
  vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => setTimeout(() => cb(0), 0) as unknown as number);
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

function hidden(styles: Record<string, Record<string, string>>, id: string): boolean {
  return styles[id]?.visibility === "hidden";
}

describe("useBuildPreview", () => {
  it("settles unbuilt entrance frames hidden (AC-20)", () => {
    const p = useBuildPreview();
    p.settle(slide(ANIMS), 0);
    expect(hidden(p.frameStyles.value, "a")).toBe(true);
    expect(hidden(p.frameStyles.value, "b")).toBe(true);
    expect(hidden(p.frameStyles.value, "c")).toBe(false); // no animation → visible
  });

  it("plays a group and settles its frames visible", () => {
    const p = useBuildPreview();
    const s = slide(ANIMS);
    p.playGroup(s, 0, false);
    vi.runAllTimers();
    expect(hidden(p.frameStyles.value, "a")).toBe(false);
    expect(hidden(p.frameStyles.value, "b")).toBe(false);
  });

  it("reduced motion settles instantly to the after-state (AC-23)", () => {
    const p = useBuildPreview();
    const s = slide(ANIMS);
    let done = false;
    p.playGroup(s, 0, true, () => (done = true));
    expect(done).toBe(true); // no timers scheduled
    expect(hidden(p.frameStyles.value, "a")).toBe(false);
    expect(hidden(p.frameStyles.value, "b")).toBe(false);
  });

  it("mid-build input completes the group instantly (AC-24)", () => {
    const p = useBuildPreview();
    const s = slide([
      { frameId: "a", effect: "fadeIn", trigger: "click" },
      { frameId: "b", effect: "fadeIn", trigger: "click" },
    ]);
    p.playGroup(s, 0, false);
    vi.advanceTimersByTime(10); // group 0 still animating
    // A second advance interrupts: playing group 1 first snaps group 0 done.
    p.playGroup(s, 1, false);
    // Frame a (group 0) is settled visible immediately, not left mid-fade.
    expect(hidden(p.frameStyles.value, "a")).toBe(false);
    vi.runAllTimers();
    expect(hidden(p.frameStyles.value, "b")).toBe(false);
  });
});
