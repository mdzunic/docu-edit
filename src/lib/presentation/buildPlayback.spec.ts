import { describe, expect, it } from "vitest";
import type { Slide, SlideAnimation, SlideFrame, SlideParagraph } from "./model";
import {
  expandAnimation,
  groupTimeline,
  interactiveTriggerAt,
  isBuildTargetHiddenAt,
  isFrameHiddenAt,
  playbackDurationMs,
  settledMotion,
  slideBuildCount,
  slideBuildGroups,
  slideBuildSteps,
  slideInteractiveBuilds,
  stepTimeline,
} from "./buildPlayback";

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

// ---------------------------------------------------------------------------
// SPEC-105: build steps, interactive sequences, emphasis/path visibility
// ---------------------------------------------------------------------------

function para(text: string, level?: number): SlideParagraph {
  return level ? { runs: [{ text }], level } : { runs: [{ text }] };
}
function textFrame(id: string, paragraphs: SlideParagraph[]): SlideFrame {
  return { id, kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 5, paragraphs };
}
const FOUR = textFrame("body", [para("a"), para("b"), para("c"), para("d")]);
const steps = (a: SlideAnimation, frame: SlideFrame | undefined) =>
  expandAnimation(a, frame).map((s) => [s.key, s.paragraphIndex, s.paragraphEnd]);

describe("expandAnimation — the three build modes", () => {
  it("no build, no paragraphs, fewer than two, or an image → one whole-frame step", () => {
    const whole = [["body", undefined, undefined]];
    expect(steps({ frameId: "body", effect: "fadeIn", trigger: "click" }, FOUR)).toEqual(whole);
    const byP: SlideAnimation = { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" };
    expect(steps(byP, textFrame("body", []))).toEqual(whole);
    expect(steps(byP, textFrame("body", [para("only")]))).toEqual(whole);
    expect(steps(byP, { id: "body", kind: "image", mediaId: "m", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 })).toEqual(whole);
    expect(steps(byP, undefined)).toEqual(whole);
  });

  it("by paragraph: one step per paragraph, in order", () => {
    expect(steps({ frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" }, FOUR)).toEqual([
      ["body#0", 0, 0],
      ["body#1", 1, 1],
      ["body#2", 2, 2],
      ["body#3", 3, 3],
    ]);
  });

  it("by first-level paragraph: a step per level-0 paragraph carrying its deeper ones", () => {
    const a: SlideAnimation = { frameId: "f", effect: "fadeIn", trigger: "click", buildBy: "byFirstLevelParagraph" };
    expect(steps(a, textFrame("f", [para("A"), para("a1", 1), para("a2", 2), para("B"), para("b1", 1)]))).toEqual([
      ["f#0", 0, 2],
      ["f#3", 3, 4],
    ]);
    // A leading deeper paragraph joins the first step.
    expect(steps(a, textFrame("f", [para("x", 1), para("A"), para("a1", 1), para("B")]))).toEqual([
      ["f#0", 0, 2],
      ["f#3", 3, 3],
    ]);
    // No level-0 paragraph at all: one step over everything.
    expect(steps(a, textFrame("f", [para("x", 1), para("y", 2)]))).toEqual([["f#0", 0, 1]]);
  });

  it("is derived, never stored: a shrinking body loses steps and repairs nothing", () => {
    const a: SlideAnimation = { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" };
    const body = textFrame("body", [para("a"), para("b"), para("c")]);
    const slideOf = (): Slide => ({ id: "s", frames: [body], animations: [a] });
    expect(slideBuildCount(slideOf())).toBe(3);
    (body as { paragraphs: SlideParagraph[] }).paragraphs.splice(1, 1);
    expect(slideBuildCount(slideOf())).toBe(2);
    (body as { paragraphs: SlideParagraph[] }).paragraphs.splice(1, 1);
    expect(slideBuildCount(slideOf())).toBe(1);
    expect(a).toEqual({ frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" });
  });
});

describe("slideBuildSteps / slideInteractiveBuilds", () => {
  const title = textFrame("title", [para("T")]);
  const shape: SlideFrame = { id: "btn", kind: "shape", geometry: "rect", xCm: 20, yCm: 10, widthCm: 4, heightCm: 2 };

  it("a row's trigger opens its first step's group; later steps open on click or join on with/after", () => {
    const s: Slide = {
      id: "s",
      frames: [title, FOUR],
      animations: [
        { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
        { frameId: "title", effect: "pulse", trigger: "withPrevious" },
      ],
    };
    expect(slideBuildSteps(s).map((g) => g.map((x) => x.key))).toEqual([["body#0"], ["body#1"], ["body#2"], ["body#3", "title"]]);
    s.animations![0].buildTrigger = "afterPrevious";
    expect(slideBuildSteps(s).map((g) => g.map((x) => x.key))).toEqual([["body#0", "body#1", "body#2", "body#3", "title"]]);
    expect(slideBuildGroups(s)).toEqual([s.animations]);
  });

  it("onClickOf rows leave the main sequence and never count towards slideBuildCount", () => {
    const s: Slide = {
      id: "s",
      frames: [title, FOUR, shape],
      animations: [
        { frameId: "title", effect: "spin", trigger: "onClickOf", triggerFrameId: "btn" },
        { frameId: "body", effect: "fadeIn", trigger: "afterPrevious" },
        { frameId: "body", effect: "pulse", trigger: "onClickOf", triggerFrameId: "btn", buildBy: "byParagraph", buildTrigger: "withPrevious" },
      ],
    };
    expect(slideBuildCount(s)).toBe(1);
    expect(slideBuildSteps(s).map((g) => g.map((x) => x.key))).toEqual([["body"]]);
    const interactive = slideInteractiveBuilds(s);
    expect(interactive.map((i) => i.triggerFrameId)).toEqual(["btn"]);
    expect(interactive[0].groups.map((g) => g.map((x) => x.key))).toEqual([["title", "body#0", "body#1", "body#2", "body#3"]]);
    // A by-paragraph interactive row stepping on click waits for further clicks.
    s.animations![2].buildTrigger = undefined;
    expect(slideInteractiveBuilds(s)[0].groups.map((g) => g.map((x) => x.key))).toEqual([
      ["title", "body#0"],
      ["body#1"],
      ["body#2"],
      ["body#3"],
    ]);
  });

  it("finds the topmost trigger frame under a point", () => {
    const s: Slide = {
      id: "s",
      frames: [title, shape],
      animations: [{ frameId: "title", effect: "spin", trigger: "onClickOf", triggerFrameId: "btn" }],
    };
    expect(interactiveTriggerAt(s, 21, 11)).toBe("btn");
    expect(interactiveTriggerAt(s, 1, 1)).toBeUndefined(); // the title is no trigger
    expect(interactiveTriggerAt(s, 30, 30)).toBeUndefined();
  });
});

/** Where `rows`, all played on target "a", leave it. */
function playedMotion(rows: SlideAnimation[]): ReturnType<typeof settledMotion> {
  return settledMotion(rows.map((r) => ({ animation: r, frameId: "a", key: "a" })), "a");
}

describe("SPEC-105 visibility and durations", () => {
  it("isFrameHiddenAt no longer treats a played emphasis or path effect as an exit", () => {
    const groups: SlideAnimation[][] = [
      [{ frameId: "a", effect: "spin", trigger: "click" }],
      [{ frameId: "a", effect: "motionPath", trigger: "click", path: { preset: "line" } }],
      [{ frameId: "a", effect: "fadeOut", trigger: "click" }],
    ];
    expect([0, 1, 2, 3].map((n) => isFrameHiddenAt(groups, "a", n))).toEqual([false, false, false, true]);
  });

  it("isBuildTargetHiddenAt tracks paragraphs and frames by key; a rewind restores what it found", () => {
    const s: Slide = {
      id: "s",
      frames: [FOUR],
      animations: [
        { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
        { frameId: "body", effect: "fadeOut", trigger: "click", rewindOnComplete: true },
      ],
    };
    const groups = slideBuildSteps(s);
    expect(groups).toHaveLength(5);
    expect([0, 1, 2, 4, 5].map((n) => isBuildTargetHiddenAt(groups, "body#1", n))).toEqual([true, true, false, false, false]);
    // The frame itself is never hidden by its paragraphs' build; the rewound exit leaves it visible.
    expect([0, 4, 5].map((n) => isBuildTargetHiddenAt(groups, "body", n))).toEqual([false, false, false]);
    // An entrance with rewind hides its target again.
    const rewound: Slide = { id: "s", frames: [FOUR], animations: [{ frameId: "body", effect: "fadeIn", trigger: "click", rewindOnComplete: true }] };
    expect([0, 1].map((n) => isBuildTargetHiddenAt(slideBuildSteps(rewound), "body", n))).toEqual([true, true]);
  });

  it("playbackDurationMs = single × passes × (autoReverse ? 2 : 1); until-click counts one pass", () => {
    const base: SlideAnimation = { frameId: "a", effect: "spin", trigger: "click", durationMs: 400 };
    expect(playbackDurationMs(base)).toBe(400);
    expect(playbackDurationMs({ ...base, repeatCount: 3 })).toBe(1200);
    expect(playbackDurationMs({ ...base, repeatCount: 3, autoReverse: true })).toBe(2400);
    expect(playbackDurationMs({ ...base, repeatUntilClick: true, autoReverse: true })).toBe(800);
    expect(playbackDurationMs({ ...base, effect: "appear", repeatCount: 5 })).toBe(0);
    // An afterPrevious chained behind an until-click repeat never stalls.
    const { entries } = groupTimeline([{ ...base, repeatUntilClick: true }, { frameId: "b", effect: "fadeIn", trigger: "afterPrevious" }]);
    expect(entries[1].startMs).toBe(400);
  });

  it("stepTimeline chains a by-paragraph row's steps on its build trigger", () => {
    const a: SlideAnimation = { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph", buildTrigger: "afterPrevious", delayMs: 100 };
    const { entries, totalMs } = stepTimeline(expandAnimation(a, FOUR));
    expect(entries.map((e) => e.startMs)).toEqual([100, 700, 1300, 1900]);
    expect(totalMs).toBe(2400);
    const together = stepTimeline(expandAnimation({ ...a, buildTrigger: "withPrevious" }, FOUR));
    expect(together.entries.map((e) => e.startMs)).toEqual([100, 200, 300, 400]);
  });

  it("settledMotion: where played spins and paths leave a target", () => {
    const spin: SlideAnimation = { frameId: "a", effect: "spin", trigger: "click", spinDegrees: 90, repeatCount: 2 };
    const path: SlideAnimation = { frameId: "a", effect: "motionPath", trigger: "click", path: { preset: "line", direction: "down" } };

    expect(playedMotion([spin, path])).toEqual({ rotateDeg: 180, offset: { x: 0, y: 0.25 } });
    expect(playedMotion([{ ...spin, autoReverse: true }])).toEqual({ rotateDeg: 0, offset: { x: 0, y: 0 } });
    expect(playedMotion([{ ...path, rewindOnComplete: true }])).toEqual({ rotateDeg: 0, offset: { x: 0, y: 0 } });
  });
});
