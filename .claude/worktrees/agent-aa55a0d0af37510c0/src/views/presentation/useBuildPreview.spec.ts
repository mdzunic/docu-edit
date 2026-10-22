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

  it("complete() snaps the in-flight group done and holds — the last group does not advance (AC-24)", () => {
    const p = useBuildPreview();
    const s = slide(ANIMS); // a single click group (fadeIn + afterPrevious)
    p.playGroup(s, 0, false);
    vi.advanceTimersByTime(10); // still animating: b (afterPrevious) not shown yet
    expect(p.playing.value).toBe(true);
    expect(hidden(p.frameStyles.value, "b")).toBe(true);

    // A mid-build advance completes the group instantly and holds on the slide.
    expect(p.complete()).toBe(true);
    expect(p.playing.value).toBe(false);
    expect(hidden(p.frameStyles.value, "a")).toBe(false);
    expect(hidden(p.frameStyles.value, "b")).toBe(false);

    // No pending timers remain to fire a further settle/advance.
    vi.runAllTimers();
    expect(hidden(p.frameStyles.value, "b")).toBe(false);
  });

  it("holds the start style through a rendered frame before flipping to the end style", () => {
    // In one frame the browser would style from and to together, so the
    // transition would start from the settled hidden state (a fly-in plays as a
    // fade). The flip waits for the frame after the one that renders `from`.
    const queue: FrameRequestCallback[] = [];
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => queue.push(cb));
    const frame = (): void => queue.splice(0).forEach((cb) => cb(0));
    const p = useBuildPreview();
    p.playGroup(slide([{ frameId: "b", effect: "flyIn", trigger: "click" }]), 0, false);
    vi.advanceTimersByTime(0);
    expect(p.frameStyles.value.b).toMatchObject({ transform: "translateX(-120%)", transition: "none" });
    frame();
    expect(p.frameStyles.value.b).toMatchObject({ transform: "translateX(-120%)", transition: "none" });
    frame();
    expect(p.frameStyles.value.b).toMatchObject({ transform: "none" });
    expect(p.frameStyles.value.b.transition).toContain("transform");
  });

  it("drops a pending flip when the group completes first, so it cannot undo the settle", () => {
    const queue: FrameRequestCallback[] = [];
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => queue.push(cb));
    const frame = (): void => queue.splice(0).forEach((cb) => cb(0));
    const p = useBuildPreview();
    p.playGroup(slide([{ frameId: "b", effect: "flyIn", trigger: "click" }]), 0, false);
    vi.advanceTimersByTime(0);
    expect(p.complete()).toBe(true);
    frame();
    frame();
    expect(p.frameStyles.value.b).toEqual({}); // settled visible, no stale transition
  });

  it("complete() is a no-op with nothing in flight — a settled slide advances normally (AC-24)", () => {
    const p = useBuildPreview();
    const s = slide(ANIMS);
    p.playGroup(s, 0, false);
    vi.runAllTimers(); // group fully settled
    expect(p.playing.value).toBe(false);
    // No group in flight → the input is not consumed by the build engine.
    expect(p.complete()).toBe(false);
  });
});
