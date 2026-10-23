// SPEC-105 timeline geometry: 60px per second, 1s ticks, bars offset by their
// stepTimeline start and sized by playback duration (8px minimum), edge edits
// in 100ms steps clamped to SPEC-086's bounds, and the timeline model.
import { describe, expect, it } from "vitest";
import type { Slide, SlideFrame } from "../../lib/presentation/model";
import {
  barRect,
  draggedDelayMs,
  draggedDurationMs,
  MIN_BAR_PX,
  msFromDx,
  PX_PER_SECOND,
  rulerTicks,
  timelineModel,
} from "./animationTimeline";

const body: SlideFrame = {
  id: "body",
  kind: "text",
  xCm: 0,
  yCm: 0,
  widthCm: 10,
  heightCm: 5,
  paragraphs: ["a", "b", "c", "d"].map((text) => ({ runs: [{ text }] })),
};
const box: SlideFrame = { id: "box", kind: "shape", geometry: "rect", xCm: 12, yCm: 0, widthCm: 4, heightCm: 2 };

describe("timeline geometry", () => {
  it("places bars at 60px per second, at least 8px wide", () => {
    expect(PX_PER_SECOND).toBe(60);
    expect(barRect(500, 1000)).toEqual({ left: 30, width: 60 });
    expect(barRect(0, 0)).toEqual({ left: 0, width: MIN_BAR_PX });
  });

  it("ticks every second through the last one reached", () => {
    expect(rulerTicks(0).map((t) => t.label)).toEqual(["0s", "1s"]);
    expect(rulerTicks(2100).map((t) => [t.seconds, t.x])).toEqual([[0, 0], [1, 60], [2, 120], [3, 180]]);
  });

  it("turns pointer travel into whole 100ms steps", () => {
    expect(msFromDx(6)).toBe(100);
    expect(msFromDx(-12)).toBe(-200);
    expect(msFromDx(2)).toBe(0);
  });

  it("clamps edge edits to 0–10000 delay and 100–5000 duration", () => {
    expect(draggedDelayMs(200, -1000)).toBe(0);
    expect(draggedDelayMs(9900, 500)).toBe(10000);
    expect(draggedDurationMs(500, -900)).toBe(100);
    expect(draggedDurationMs(4900, 300)).toBe(5000);
    // A repeated bar's edge moves each pass by its share.
    expect(draggedDurationMs(500, 600, 3)).toBe(700);
    // Steps move from where the value was: an off-grid delay keeps its offset.
    expect(draggedDelayMs(250, 100)).toBe(350);
    expect(draggedDelayMs(250, 0)).toBe(250);
    expect(draggedDurationMs(450, -100)).toBe(350);
  });
});

describe("timelineModel", () => {
  it("a by-paragraph build shows four bars in four groups", () => {
    const slide: Slide = { id: "s", frames: [body], animations: [{ frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" }] };
    const model = timelineModel(slide);
    expect(model.groups.map((g) => g.bars.length)).toEqual([1, 1, 1, 1]);
    expect(model.groups.flatMap((g) => g.bars.map((b) => [b.barIndex, b.animationIndex, b.paragraphIndex]))).toEqual([
      [0, 0, 0],
      [1, 0, 1],
      [2, 0, 2],
      [3, 0, 3],
    ]);
  });

  it("offsets bars by their start and sizes them by playback duration, so overlaps show", () => {
    const slide: Slide = {
      id: "s",
      frames: [body, box],
      animations: [
        { frameId: "body", effect: "fadeIn", trigger: "click", durationMs: 1000 },
        { frameId: "box", effect: "spin", trigger: "withPrevious", delayMs: 500, repeatCount: 2 },
        { frameId: "box", effect: "pulse", trigger: "afterPrevious" },
        { frameId: "body", effect: "spin", trigger: "onClickOf", triggerFrameId: "box" },
      ],
    };
    const model = timelineModel(slide);
    expect(model.groups).toHaveLength(2);
    const [main, interactive] = model.groups;
    expect(main.bars.map((b) => [b.left, b.width])).toEqual([
      [0, 60],
      [30, 60],
      [90, 30],
    ]);
    expect(interactive.triggerFrameId).toBe("box");
    expect(interactive.bars.map((b) => b.animationIndex)).toEqual([3]);
    expect(model.totalMs).toBe(2000);
    expect(model.ticks.at(-1)?.seconds).toBe(2);
  });
});
