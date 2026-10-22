// SPEC-084 — the playback step machine. Covers advance/back/goTo, the end
// screen, index clamping, and the reduced-motion plumbing that downstream specs
// (SPEC-085/086) rely on being present. Pure state, no component mount.
import { describe, expect, it } from "vitest";
import { ref } from "vue";
import { usePlayback } from "./usePlayback";

function machine(count: number, start = 0) {
  return usePlayback({ slideCount: ref(count), start });
}

describe("usePlayback step machine", () => {
  it("starts at the given (clamped) slide with buildIndex 0", () => {
    expect(machine(5, 2).current.value).toEqual({ slideIndex: 2, buildIndex: 0 });
    expect(machine(5, 99).current.value.slideIndex).toBe(4); // clamps high
    expect(machine(5, -3).current.value.slideIndex).toBe(0); // clamps low
    expect(machine(0, 0).current.value.slideIndex).toBe(0); // empty deck safe
  });

  it("advance walks to the next slide, then to the end screen", () => {
    const p = machine(3, 0);
    p.advance();
    expect(p.current.value.slideIndex).toBe(1);
    p.advance();
    expect(p.current.value.slideIndex).toBe(2);
    expect(p.atEnd.value).toBe(false);
    p.advance(); // past the last slide
    expect(p.atEnd.value).toBe(true);
    expect(p.current.value.slideIndex).toBe(2); // stays on the last slide
    p.advance(); // idempotent once at end (the view turns this into an exit)
    expect(p.atEnd.value).toBe(true);
  });

  it("back steps to the previous slide and stops at the first", () => {
    const p = machine(3, 2);
    p.back();
    expect(p.current.value.slideIndex).toBe(1);
    p.back();
    expect(p.current.value.slideIndex).toBe(0);
    p.back(); // no wraparound
    expect(p.current.value.slideIndex).toBe(0);
  });

  it("back from the end screen returns to the last slide", () => {
    const p = machine(3, 2);
    p.advance();
    expect(p.atEnd.value).toBe(true);
    p.back();
    expect(p.atEnd.value).toBe(false);
    expect(p.current.value.slideIndex).toBe(2);
  });

  it("goTo jumps and clears the end screen, clamping out-of-range targets", () => {
    const p = machine(5, 0);
    p.goTo(3);
    expect(p.current.value.slideIndex).toBe(3);
    p.goTo(100);
    expect(p.current.value.slideIndex).toBe(4);
    p.advance();
    expect(p.atEnd.value).toBe(true);
    p.goTo(1);
    expect(p.atEnd.value).toBe(false);
    expect(p.current.value.slideIndex).toBe(1);
  });

  it("advance walks a slide's build groups before moving on (SPEC-086)", () => {
    // slide 0 has 2 build groups, slide 1 has 1, slide 2 has none.
    const counts = [2, 1, 0];
    const p = usePlayback({ slideCount: ref(3), buildCountAt: (i) => counts[i], start: 0 });
    p.advance(); // build 1 of slide 0
    expect(p.current.value).toEqual({ slideIndex: 0, buildIndex: 1 });
    p.advance(); // build 2 of slide 0
    expect(p.current.value).toEqual({ slideIndex: 0, buildIndex: 2 });
    p.advance(); // builds done → slide 1
    expect(p.current.value).toEqual({ slideIndex: 1, buildIndex: 0 });
    p.advance(); // build 1 of slide 1
    expect(p.current.value).toEqual({ slideIndex: 1, buildIndex: 1 });
    p.advance(); // → slide 2 (no builds)
    expect(p.current.value).toEqual({ slideIndex: 2, buildIndex: 0 });
    p.advance(); // → end
    expect(p.atEnd.value).toBe(true);
  });

  it("back lands the previous slide fully-built (AC-22)", () => {
    const counts = [3, 2];
    const p = usePlayback({ slideCount: ref(2), buildCountAt: (i) => counts[i], start: 1 });
    p.back();
    expect(p.current.value).toEqual({ slideIndex: 0, buildIndex: 3 }); // fully built
    // From the end screen, back lands the last slide fully built.
    const q = usePlayback({ slideCount: ref(2), buildCountAt: (i) => counts[i], start: 1 });
    q.advance();
    q.advance();
    q.advance(); // exhaust slide 1's 2 builds then → end
    expect(q.atEnd.value).toBe(true);
    q.back();
    expect(q.current.value).toEqual({ slideIndex: 1, buildIndex: 2 });
  });

  it("exposes reduced-motion reactively (SPEC-085/086 plumbing)", () => {
    const reducedMotion = ref(false);
    const p = usePlayback({ slideCount: ref(3), reducedMotion });
    expect(p.reducedMotion.value).toBe(false);
    reducedMotion.value = true;
    expect(p.reducedMotion.value).toBe(true);
  });

  it("defaults reduced-motion to a never-reduced ref when none is passed", () => {
    expect(machine(3).reducedMotion.value).toBe(false);
  });
});

// SPEC-104 — the optional isHidden dep: hidden slides are stepped over by
// advance/back/goTo, the end screen follows the last visible slide, and a
// deep-linked start on a hidden slide still shows it.
function hiddenMachine(count: number, hidden: number[], start = 0, builds: Record<number, number> = {}) {
  return usePlayback({
    slideCount: ref(count),
    start,
    isHidden: (index) => hidden.includes(index),
    buildCountAt: (index) => builds[index] ?? 0,
  });
}

describe("usePlayback hidden slides (SPEC-104)", () => {
  it("absent isHidden hides nothing", () => {
    const p = machine(3, 0);
    p.advance();
    expect(p.current.value.slideIndex).toBe(1);
  });

  it("advance steps over a hidden slide", () => {
    const p = hiddenMachine(3, [1]);
    p.advance();
    expect(p.current.value).toEqual({ slideIndex: 2, buildIndex: 0 });
  });

  it("advance walks a slide's builds before stepping over the hidden one", () => {
    const p = hiddenMachine(3, [1], 0, { 0: 1 });
    p.advance();
    expect(p.current.value).toEqual({ slideIndex: 0, buildIndex: 1 });
    p.advance();
    expect(p.current.value).toEqual({ slideIndex: 2, buildIndex: 0 });
  });

  it("atEnd is reached when no later visible slide exists (a hidden tail)", () => {
    const p = hiddenMachine(4, [2, 3]);
    p.advance();
    expect(p.current.value.slideIndex).toBe(1);
    p.advance();
    expect(p.atEnd.value).toBe(true);
    expect(p.current.value.slideIndex).toBe(1);
  });

  it("back steps over a hidden slide and stops at the first visible one", () => {
    const p = hiddenMachine(4, [0, 2], 3);
    p.back();
    expect(p.current.value.slideIndex).toBe(1);
    p.back(); // slide 0 is hidden: nowhere further back to go
    expect(p.current.value.slideIndex).toBe(1);
  });

  it("back from the end screen lands on the last visible slide, fully built", () => {
    const p = hiddenMachine(4, [3], 2, { 2: 2 });
    p.advance();
    p.advance();
    p.advance();
    expect(p.atEnd.value).toBe(true);
    p.back();
    expect(p.atEnd.value).toBe(false);
    expect(p.current.value).toEqual({ slideIndex: 2, buildIndex: 2 });
  });

  it("goTo steps over a hidden target — forward, else back", () => {
    const p = hiddenMachine(5, [0, 2, 4]);
    p.goTo(0);
    expect(p.current.value.slideIndex).toBe(1);
    p.goTo(2);
    expect(p.current.value.slideIndex).toBe(3);
    p.goTo(4); // nothing visible after the last: the previous visible one
    expect(p.current.value.slideIndex).toBe(3);
    p.goTo(100); // clamps to 4, which is hidden
    expect(p.current.value.slideIndex).toBe(3);
  });

  it("a deep-linked start on a hidden slide still shows it; advancing then steps over hidden ones", () => {
    const p = hiddenMachine(4, [1, 2], 1);
    expect(p.current.value).toEqual({ slideIndex: 1, buildIndex: 0 });
    p.advance();
    expect(p.current.value.slideIndex).toBe(3);
  });
});
