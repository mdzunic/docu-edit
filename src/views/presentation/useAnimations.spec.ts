import { beforeEach, describe, expect, it, vi } from "vitest";
import { computed, ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import { useAnimations, type AnimationsDeps } from "./useAnimations";

function frame(id: string, kind: SlideFrame["kind"] = "text"): SlideFrame {
  const base = { id, xCm: 0, yCm: 0, widthCm: 5, heightCm: 2 };
  if (kind === "shape") return { ...base, kind: "shape", geometry: "rect" };
  if (kind === "image") return { ...base, kind: "image", mediaId: "m" };
  if (kind === "group") return { ...base, kind: "group", children: [], xml: "<g/>" };
  return { ...base, kind: "text", paragraphs: [] };
}

function harness(slide: Slide, format: "pptx" | "odp" = "pptx") {
  const model = ref<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides: [slide],
    slideWidthCm: 25,
    slideHeightCm: 14,
  });
  const currentIndex = ref(0);
  const currentSlide = computed(() => model.value?.slides[currentIndex.value] ?? null);
  const selectedFrameId = ref<string | null>(null);
  const record = vi.fn();
  const commit = vi.fn();
  const selectFrame = vi.fn((id: string | null) => (selectedFrameId.value = id));
  const deps: AnimationsDeps = { model, currentSlide, currentIndex, record, commit, selectedFrameId, selectFrame };
  return { deps, model, selectedFrameId, record, commit, selectFrame, c: useAnimations(deps) };
}

beforeEach(() => localStorage.clear());

describe("useAnimations", () => {
  it("Add is disabled with helper text when no frame is selected (AC-8)", () => {
    const { c } = harness({ id: "s", frames: [frame("f1")] });
    expect(c.canAdd.value).toBe(false);
    expect(c.addHelp.value).toContain("Select a frame");
  });

  it("Add is disabled for a non-animatable frame kind", () => {
    const h = harness({ id: "s", frames: [frame("g", "group")] });
    h.selectedFrameId.value = "g";
    expect(h.c.canAdd.value).toBe(false);
    expect(h.c.addHelp.value).toContain("can't be animated");
  });

  it("adds a fade-in/click animation, records + commits, selects the new row (AC-16)", () => {
    const h = harness({ id: "s", frames: [frame("f1"), frame("f2", "shape")] });
    h.selectedFrameId.value = "f1";
    h.c.add();
    expect(h.model.value!.slides[0].animations).toEqual([{ frameId: "f1", effect: "fadeIn", trigger: "click" }]);
    expect(h.record).toHaveBeenCalledTimes(1);
    expect(h.commit).toHaveBeenCalledTimes(1);
    expect(h.c.selectedIndex.value).toBe(0);
  });

  it("normalises row 1 trigger to click and locks its trigger control (AC-3)", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1")],
      animations: [{ frameId: "f1", effect: "flyIn", trigger: "afterPrevious" }],
    });
    expect(h.c.rows.value[0].trigger).toBe("click");
    expect(h.c.triggerLocked.value).toBe(true);
    h.c.setTrigger("withPrevious"); // ignored while locked
    expect(h.model.value!.slides[0].animations![0].trigger).toBe("afterPrevious");
  });

  it("selecting a row highlights its target frame (AC-14)", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1"), frame("f2", "shape")],
      animations: [
        { frameId: "f1", effect: "fadeIn", trigger: "click" },
        { frameId: "f2", effect: "flyIn", trigger: "click" },
      ],
    });
    h.c.selectRow(1);
    expect(h.selectFrame).toHaveBeenCalledWith("f2");
    expect(h.c.selectedIndex.value).toBe(1);
  });

  it("reorders and removes with one gesture each", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1"), frame("f2", "shape")],
      animations: [
        { frameId: "f1", effect: "fadeIn", trigger: "click" },
        { frameId: "f2", effect: "flyIn", trigger: "click" },
      ],
    });
    h.c.selectRow(1);
    h.c.moveUp();
    expect(h.model.value!.slides[0].animations!.map((a) => a.frameId)).toEqual(["f2", "f1"]);
    expect(h.c.selectedIndex.value).toBe(0);
    h.c.remove();
    expect(h.model.value!.slides[0].animations!.map((a) => a.frameId)).toEqual(["f1"]);
  });

  it("gates Add behind confirm when the slide has preserved timing (AC-17)", () => {
    const h = harness({ id: "s", frames: [frame("f1")], preserved: { "p:timing": "<p:timing/>" } });
    h.selectedFrameId.value = "f1";
    h.c.add();
    // Nothing added yet — the confirm is pending.
    expect(h.c.pendingReplaceFrameId.value).toBe("f1");
    expect(h.model.value!.slides[0].animations).toBeUndefined();
    h.c.confirmReplace();
    expect(h.model.value!.slides[0].animations).toHaveLength(1);
    expect(h.model.value!.slides[0].preserved?.["p:timing"]).toBeUndefined();
  });

  it("cancel leaves preserved timing and adds nothing", () => {
    const h = harness({ id: "s", frames: [frame("f1")], preserved: { "p:timing": "<p:timing/>" } });
    h.selectedFrameId.value = "f1";
    h.c.add();
    h.c.cancelReplace();
    expect(h.c.pendingReplaceFrameId.value).toBeNull();
    expect(h.model.value!.slides[0].animations).toBeUndefined();
    expect(h.model.value!.slides[0].preserved?.["p:timing"]).toBe("<p:timing/>");
  });
});

// ---------------------------------------------------------------------------
// SPEC-105 controls
// ---------------------------------------------------------------------------

function textWith(id: string, lines: string[]): SlideFrame {
  return { id, kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 4, paragraphs: lines.map((text) => ({ runs: [{ text }] })) };
}
function depthHarness(animations: Slide["animations"]) {
  return harness({
    id: "s",
    frames: [textWith("title", ["T"]), textWith("body", ["a", "b", "c", "d"]), frame("box", "shape"), frame("pic", "image")],
    animations,
  });
}

function at(clientX: number): PointerEvent {
  return { clientX, pointerId: 1, target: null } as unknown as PointerEvent;
}
function key(k: string): KeyboardEvent {
  return { key: k, preventDefault: () => {} } as unknown as KeyboardEvent;
}

describe("useAnimations — SPEC-105", () => {
  it("every new control records exactly one undo entry in SPEC-080's voice, dirties and commits", () => {
    const h = depthHarness([
      { frameId: "body", effect: "spin", trigger: "click" },
      { frameId: "box", effect: "motionPath", trigger: "click", path: { preset: "line" } },
    ]);
    const steps: [string, () => void][] = [
      ["Change animation build", () => h.c.setBuildBy("byParagraph")],
      ["Change animation build", () => h.c.setBuildTrigger("afterPrevious")],
      ["Change animation repeat", () => h.c.setRepeatCount(3)],
      ["Change animation repeat", () => h.c.setRepeatUntilClick(true)],
      ["Change animation auto-reverse", () => h.c.setAutoReverse(true)],
      ["Change animation rewind", () => h.c.setRewind(true)],
      ["Change animation trigger", () => h.c.setTrigger("onClickOf")],
      ["Change animation trigger frame", () => h.c.setTriggerFrame("title")],
      ["Change animation spin", () => h.c.setSpinDegrees(90)],
      ["Change motion path", () => (h.c.selectRow(1), h.c.setPathPreset("arc"))],
      ["Change motion path", () => h.c.setPathDirection("down")],
      ["Change motion path", () => h.c.setPathDistancePercent(40)],
    ];
    for (const [label, run] of steps) {
      h.record.mockClear();
      h.commit.mockClear();
      run();
      expect(h.record, label).toHaveBeenCalledTimes(1);
      expect(h.record).toHaveBeenCalledWith(label);
      expect(h.commit).toHaveBeenCalledTimes(1);
    }
    const [spin, path] = h.model.value!.slides[0].animations!;
    expect(spin).toEqual({
      frameId: "body",
      effect: "spin",
      trigger: "onClickOf",
      triggerFrameId: "title",
      buildBy: "byParagraph",
      buildTrigger: "afterPrevious",
      repeatUntilClick: true,
      autoReverse: true,
      rewindOnComplete: true,
      spinDegrees: 90,
    });
    expect(path.path).toEqual({ preset: "arc", direction: "down", distanceFraction: 0.4 });
    expect(h.model.value!.slides[0].dirty).toBe(true);
  });

  it("scale and colour record their own entries", () => {
    const h = depthHarness([{ frameId: "box", effect: "growShrink", trigger: "click" }]);
    h.c.setScalePercent(250);
    expect(h.record).toHaveBeenLastCalledWith("Change animation scale");
    h.c.setEffect("colorChange");
    h.c.setEmphasisColor("#00aa00");
    expect(h.record).toHaveBeenLastCalledWith("Change animation colour");
    expect(h.model.value!.slides[0].animations![0]).toMatchObject({ effect: "colorChange", emphasisColor: "#00aa00" });
  });

  it("build-by is disabled with helper text under two paragraphs or on an image (AC-9)", () => {
    const h = depthHarness([
      { frameId: "body", effect: "fadeIn", trigger: "click" },
      { frameId: "title", effect: "fadeIn", trigger: "click" },
      { frameId: "pic", effect: "fadeIn", trigger: "click" },
    ]);
    expect([h.c.buildByEnabled.value, h.c.buildByHelp.value]).toEqual([true, ""]);
    expect(h.c.buildTriggerShown.value).toBe(false);
    h.c.setBuildBy("byParagraph");
    expect(h.c.buildTriggerShown.value).toBe(true);
    h.c.selectRow(1);
    expect(h.c.buildByEnabled.value).toBe(false);
    expect(h.c.buildByHelp.value).toContain("two or more paragraphs");
    h.record.mockClear();
    h.c.setBuildBy("byParagraph"); // refused
    expect(h.record).not.toHaveBeenCalled();
    h.c.selectRow(2);
    expect(h.c.buildByHelp.value).toContain("Only text");
  });

  it("shows each control only where the row uses it", () => {
    const h = depthHarness([
      { frameId: "body", effect: "fadeIn", trigger: "click" },
      { frameId: "box", effect: "spin", trigger: "click" },
      { frameId: "box", effect: "appear", trigger: "click" },
      { frameId: "box", effect: "motionPath", trigger: "click", path: { preset: "custom", d: "M 0 0 L 0.1 0" } },
    ]);
    expect([h.c.repeatShown.value, h.c.rewindShown.value]).toEqual([false, true]);
    h.c.selectRow(1);
    expect([h.c.repeatShown.value, h.c.rewindShown.value, h.c.selectedSpinDegrees.value]).toEqual([true, true, 360]);
    h.c.selectRow(2);
    expect(h.c.rewindShown.value).toBe(false);
    h.c.selectRow(3);
    expect([h.c.selectedPathPreset.value, h.c.pathDirectionShown.value]).toEqual(["custom", false]);
    // A custom path changes only by choosing another preset.
    h.c.setPathDistancePercent(50);
    h.c.setPathPreset("custom");
    expect(h.model.value!.slides[0].animations![3].path).toEqual({ preset: "custom", d: "M 0 0 L 0.1 0" });
    h.c.setPathPreset("loop");
    expect(h.model.value!.slides[0].animations![3].path).toEqual({ preset: "loop" });
  });

  it("lists the slide's other animatable frames as trigger frames, enabled only for On click of", () => {
    const h = depthHarness([{ frameId: "title", effect: "spin", trigger: "click" }]);
    expect(h.c.triggerFrameOptions.value.map((o) => [o.value, o.label])).toEqual([
      ["body", "Text box"],
      ["box", "Rectangle"],
      ["pic", "Image"],
    ]);
    expect(h.c.triggerFrameEnabled.value).toBe(false);
    h.record.mockClear();
    h.c.setTriggerFrame("box"); // refused while not "On click of"
    expect(h.record).not.toHaveBeenCalled();
    // The first main row's with/after are locked, but it may become On click of.
    expect(h.c.triggerLocked.value).toBe(true);
    h.c.setTrigger("onClickOf");
    expect(h.model.value!.slides[0].animations![0]).toMatchObject({ trigger: "onClickOf", triggerFrameId: "body" });
    expect(h.c.triggerFrameEnabled.value).toBe(true);
    expect(h.c.triggerLocked.value).toBe(false);
    expect(h.c.rows.value[0]).toMatchObject({ trigger: "onClickOf", triggerFrameName: "Text box" });
  });

  it("the timeline toggle persists as presentationAnimationTimelineOpen", () => {
    const h = depthHarness([]);
    expect(h.c.timelineOpen.value).toBe(false);
    h.c.toggleTimeline();
    expect(h.c.timelineOpen.value).toBe(true);
    expect(JSON.stringify(localStorage)).toContain("presentationAnimationTimelineOpen");
  });

  it("a bar selects its row; an edge drag records once on pointer-up; arrows step 100ms per key", () => {
    const h = depthHarness([
      { frameId: "body", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
      { frameId: "box", effect: "spin", trigger: "click" },
    ]);
    const bars = h.c.timeline.value!.groups.flatMap((g) => g.bars);
    expect(bars).toHaveLength(5);
    h.c.selectBar(bars[4]);
    expect(h.c.selectedIndex.value).toBe(1);
    // Drag the delay edge 0.5s right: live preview, nothing recorded yet.
    h.record.mockClear();
    h.c.startEdgeDrag(bars[4], "delay", at(100));
    h.c.moveEdgeDrag(at(130));
    expect(h.c.timeline.value!.groups.at(-1)!.bars[0].left).toBe(30);
    expect(h.record).not.toHaveBeenCalled();
    h.c.endEdgeDrag();
    expect(h.record).toHaveBeenCalledTimes(1);
    expect(h.record).toHaveBeenCalledWith("Change animation delay");
    expect(h.model.value!.slides[0].animations![1].delayMs).toBe(500);
    // Arrows: one entry per key press, clamped.
    const bar = h.c.timeline.value!.groups.at(-1)!.bars[0];
    h.c.edgeKeydown(bar, "duration", key("ArrowRight"));
    expect(h.model.value!.slides[0].animations![1].durationMs).toBe(600);
    expect(h.record).toHaveBeenLastCalledWith("Change animation duration");
    expect(h.record).toHaveBeenCalledTimes(2);
    h.c.edgeKeydown(h.c.timeline.value!.groups.at(-1)!.bars[0], "delay", key("ArrowLeft"));
    expect(h.model.value!.slides[0].animations![1].delayMs).toBe(400);
  });

  it("creates the motion-path editor itself", () => {
    const h = depthHarness([{ frameId: "box", effect: "motionPath", trigger: "click", path: { preset: "line" } }]);
    expect(h.c.pathEdit.overlay.value).toBeNull();
    h.c.pathEdit.toggle();
    expect(h.c.pathEdit.overlay.value).not.toBeNull();
  });
});
