import { describe, expect, it } from "vitest";
import type { SlideAnimation } from "./model";
import { groupTimeline, isFrameHiddenAt, slideBuildCount, slideBuildGroups } from "./buildPlayback";

function slide(animations: SlideAnimation[]) {
  return { id: "s", frames: [], animations };
}

describe("build playback timing", () => {
  const anims: SlideAnimation[] = [
    { frameId: "a", effect: "fadeIn", trigger: "click" },
    { frameId: "b", effect: "flyIn", trigger: "afterPrevious" },
    { frameId: "c", effect: "fadeIn", trigger: "click" },
    { frameId: "d", effect: "fadeOut", trigger: "withPrevious" },
  ];

  it("groups by click and counts builds", () => {
    expect(slideBuildCount(slide(anims))).toBe(2);
    expect(slideBuildGroups(slide(anims)).map((g) => g.map((a) => a.frameId))).toEqual([["a", "b"], ["c", "d"]]);
    expect(slideBuildCount(slide([]))).toBe(0);
  });

  it("hides entrance-targeted frames until their build plays (AC-20)", () => {
    const groups = slideBuildGroups(slide(anims));
    // Before any build: a, b, c hidden (entrances); d visible (its exit not played).
    expect(isFrameHiddenAt(groups, "a", 0)).toBe(true);
    expect(isFrameHiddenAt(groups, "b", 0)).toBe(true);
    expect(isFrameHiddenAt(groups, "c", 0)).toBe(true);
    expect(isFrameHiddenAt(groups, "d", 0)).toBe(false);
    // After group 1: a, b shown; c still hidden.
    expect(isFrameHiddenAt(groups, "a", 1)).toBe(false);
    expect(isFrameHiddenAt(groups, "b", 1)).toBe(false);
    expect(isFrameHiddenAt(groups, "c", 1)).toBe(true);
    // After group 2: c shown, d exited (hidden).
    expect(isFrameHiddenAt(groups, "c", 2)).toBe(false);
    expect(isFrameHiddenAt(groups, "d", 2)).toBe(true);
  });

  it("frames with no animation are always visible", () => {
    const groups = slideBuildGroups(slide(anims));
    expect(isFrameHiddenAt(groups, "unanimated", 0)).toBe(false);
    expect(isFrameHiddenAt(groups, "unanimated", 2)).toBe(false);
  });

  it("chains afterPrevious after the predecessor, withPrevious alongside it", () => {
    const group: SlideAnimation[] = [
      { frameId: "a", effect: "fadeIn", trigger: "click", durationMs: 500 },
      { frameId: "b", effect: "fadeIn", trigger: "withPrevious", durationMs: 300 },
      { frameId: "c", effect: "flyIn", trigger: "afterPrevious", durationMs: 400, delayMs: 100 },
    ];
    const { entries, totalMs } = groupTimeline(group);
    expect(entries[0]).toMatchObject({ startMs: 0, durationMs: 500 });
    expect(entries[1]).toMatchObject({ startMs: 0, durationMs: 300 }); // with previous
    expect(entries[2]).toMatchObject({ startMs: 400, durationMs: 400 }); // after b ends (300) + 100 delay
    expect(totalMs).toBe(800);
  });

  it("treats appear as instant so an afterPrevious does not wait", () => {
    const group: SlideAnimation[] = [
      { frameId: "a", effect: "appear", trigger: "click" },
      { frameId: "b", effect: "fadeIn", trigger: "afterPrevious", durationMs: 500 },
    ];
    const { entries } = groupTimeline(group);
    expect(entries[0].durationMs).toBe(0);
    expect(entries[1].startMs).toBe(0);
  });
});
