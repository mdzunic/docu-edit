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
