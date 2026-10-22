import { describe, it, expect } from "vitest";
import { isKeyboardOpen, keyboardScrollAdjustment } from "./useVisualViewportKeyboard";

describe("keyboardScrollAdjustment (SPEC-068 AC-29/30)", () => {
  // Visible region 0..500 (keyboard occupies 500..844), pad 8.
  it("scrolls down when the caret is hidden behind the keyboard", () => {
    // Caret bottom at 540 is below the visible bottom (500) → scroll down to
    // bring it 8px above.
    expect(keyboardScrollAdjustment(540, 520, 0, 500)).toBe(540 - (500 - 8));
  });

  it("does nothing when the caret is comfortably inside the visible region", () => {
    expect(keyboardScrollAdjustment(300, 280, 0, 500)).toBe(0);
  });

  it("scrolls up when the caret sits above the visible top", () => {
    // Visible region 100..600; caret top at 90 is above it → negative delta.
    expect(keyboardScrollAdjustment(110, 90, 100, 600)).toBe(90 - (100 + 8));
  });
});

describe("isKeyboardOpen", () => {
  it("is true when the visual viewport is much shorter than the layout", () => {
    expect(isKeyboardOpen(844, 500)).toBe(true);
  });

  it("ignores small differences (collapsing browser chrome)", () => {
    expect(isKeyboardOpen(844, 800)).toBe(false);
    expect(isKeyboardOpen(844, 844)).toBe(false);
  });
});
