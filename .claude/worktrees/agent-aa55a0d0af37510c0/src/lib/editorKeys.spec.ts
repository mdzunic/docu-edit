import { describe, it, expect } from "vitest";
import {
  isContentKey,
  isDeleteKey,
  isDuplicateCombo,
  isPlainPasteCombo,
  isUnmodifiedArrowKey,
  nudgeDirectionForKey,
} from "./editorKeys";

const key = (init: KeyboardEventInit & { key: string }) => new KeyboardEvent("keydown", init);

describe("isContentKey", () => {
  it("is true for a printable character and for Enter/Backspace/Delete", () => {
    for (const k of ["a", "Z", "7", " ", "é", "Enter", "Backspace", "Delete"]) {
      expect(isContentKey(key({ key: k })), k).toBe(true);
    }
  });

  it("is false for navigation and Tab", () => {
    for (const k of ["ArrowLeft", "ArrowRight", "Home", "End", "PageUp", "Tab", "Escape", "F2"]) {
      expect(isContentKey(key({ key: k })), k).toBe(false);
    }
  });

  it("is false for any modifier combo, so copy still works over a selected atom", () => {
    expect(isContentKey(key({ key: "c", ctrlKey: true }))).toBe(false);
    expect(isContentKey(key({ key: "c", metaKey: true }))).toBe(false);
    expect(isContentKey(key({ key: "a", altKey: true }))).toBe(false);
  });

  it("still counts a shifted capital as content", () => {
    expect(isContentKey(key({ key: "A", shiftKey: true }))).toBe(true);
  });
});

describe("isDeleteKey", () => {
  it("matches both delete keys and nothing else", () => {
    expect(isDeleteKey(key({ key: "Delete" }))).toBe(true);
    expect(isDeleteKey(key({ key: "Backspace" }))).toBe(true);
    expect(isDeleteKey(key({ key: "d" }))).toBe(false);
  });
});

describe("isDuplicateCombo", () => {
  it("matches Ctrl+D and Cmd+D, upper or lower case", () => {
    expect(isDuplicateCombo(key({ key: "d", ctrlKey: true }))).toBe(true);
    expect(isDuplicateCombo(key({ key: "D", metaKey: true }))).toBe(true);
  });

  it("does not match a bare D, or Ctrl+Shift+D", () => {
    expect(isDuplicateCombo(key({ key: "d" }))).toBe(false);
    expect(isDuplicateCombo(key({ key: "d", ctrlKey: true, shiftKey: true }))).toBe(false);
  });
});

describe("isPlainPasteCombo", () => {
  it("matches a plain Ctrl/Cmd+V", () => {
    expect(isPlainPasteCombo(key({ key: "v", ctrlKey: true }))).toBe(true);
    expect(isPlainPasteCombo(key({ key: "V", metaKey: true }))).toBe(true);
  });

  it("leaves Ctrl/Cmd+Shift+V to the registry (SPEC-016 AC-27)", () => {
    expect(isPlainPasteCombo(key({ key: "v", ctrlKey: true, shiftKey: true }))).toBe(false);
  });
});

describe("nudgeDirectionForKey", () => {
  it("maps the four arrows and nothing else", () => {
    expect(nudgeDirectionForKey("ArrowUp")).toBe("up");
    expect(nudgeDirectionForKey("ArrowDown")).toBe("down");
    expect(nudgeDirectionForKey("ArrowLeft")).toBe("left");
    expect(nudgeDirectionForKey("ArrowRight")).toBe("right");
    expect(nudgeDirectionForKey("Home")).toBeNull();
    expect(nudgeDirectionForKey("a")).toBeNull();
  });
});

describe("isUnmodifiedArrowKey", () => {
  it("is true only for a bare arrow key", () => {
    expect(isUnmodifiedArrowKey(key({ key: "ArrowLeft" }))).toBe(true);
    expect(isUnmodifiedArrowKey(key({ key: "ArrowDown" }))).toBe(true);
  });

  it("is false once any modifier is held — those are native browser moves", () => {
    for (const mod of ["ctrlKey", "metaKey", "shiftKey", "altKey"] as const) {
      expect(isUnmodifiedArrowKey(key({ key: "ArrowLeft", [mod]: true })), mod).toBe(false);
    }
  });

  it("is false for a non-arrow key", () => {
    expect(isUnmodifiedArrowKey(key({ key: "Home" }))).toBe(false);
  });
});
