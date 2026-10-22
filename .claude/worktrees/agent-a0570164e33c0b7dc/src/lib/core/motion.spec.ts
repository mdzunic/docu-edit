import { describe, it, expect, vi, afterEach } from "vitest";
import { motionBaseMs, prefersReducedMotion } from "./motion";

function mockMotionToken(value = "") {
  vi.spyOn(window, "getComputedStyle").mockReturnValue({
    getPropertyValue: () => value,
  } as unknown as CSSStyleDeclaration);
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("prefersReducedMotion (SPEC-060)", () => {
  it("returns true when the OS reports reduced motion", () => {
    vi.stubGlobal(
      "matchMedia",
      vi.fn(() => ({ matches: true })),
    );
    expect(prefersReducedMotion()).toBe(true);
  });

  it("returns false when motion is allowed", () => {
    vi.stubGlobal(
      "matchMedia",
      vi.fn(() => ({ matches: false })),
    );
    expect(prefersReducedMotion()).toBe(false);
  });

  it("is safe when matchMedia is unavailable", () => {
    vi.stubGlobal("matchMedia", undefined);
    expect(prefersReducedMotion()).toBe(false);
  });
});

describe("motionBaseMs (SPEC-060)", () => {
  it("reads a millisecond token", () => {
    mockMotionToken("200ms");
    expect(motionBaseMs()).toBe(200);
  });

  it("treats a bare number as seconds (legacy tokens.css units)", () => {
    mockMotionToken("0.2");
    expect(motionBaseMs()).toBe(200);
  });

  it("falls back to 200ms for a missing or unparseable token", () => {
    mockMotionToken("");
    expect(motionBaseMs()).toBe(200);
    mockMotionToken("abc");
    expect(motionBaseMs()).toBe(200);
  });
});
