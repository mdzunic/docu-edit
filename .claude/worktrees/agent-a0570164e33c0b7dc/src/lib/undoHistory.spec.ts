import { afterEach, describe, expect, it, vi } from "vitest";
import {
  breakGroup,
  canRedo,
  canUndo,
  clearHistory,
  MAX_HISTORY_STEPS,
  recordEdit,
  redo,
  redoLabel,
  TYPING_GROUP_PAUSE_MS,
  undo,
  undoLabel,
  type HistoryEntry,
  type HistorySnapshot,
} from "./undoHistory";
import type { DocumentModel } from "./documentModel";
import type { CaretPosition } from "./textEditing";

// SPEC-015 — the undo/redo history stack. DocumentView.vue owns the DOM
// wiring (keyboard shortcuts, toolbar buttons, detecting a genuine caret
// move); this only exercises the framework-free grouping/bounding/per-
// document logic itself.

function caret(offset: number): CaretPosition {
  return { section: 0, block: 0, offset };
}

// A snapshot only needs to be distinguishable and carry a selection — its
// `model` is never interpreted by this module, so a tagged stub stands in
// for a real DocumentModel.
function snapshot(tag: string, offset = 0): HistorySnapshot {
  return {
    model: { id: tag } as unknown as DocumentModel,
    selection: { anchor: caret(offset), focus: caret(offset) },
  };
}

function entry(
  label: string,
  groupable: boolean,
  beforeTag: string,
  afterTag: string,
): HistoryEntry {
  return { before: snapshot(beforeTag), after: snapshot(afterTag), label, groupable };
}

const DOC = "doc-1";

afterEach(() => {
  clearHistory(DOC);
  clearHistory("doc-2");
});

describe("undo/redo basics (AC-1, AC-2)", () => {
  it("is empty for a document with no recorded edits", () => {
    expect(canUndo(DOC)).toBe(false);
    expect(canRedo(DOC)).toBe(false);
    expect(undoLabel(DOC)).toBeNull();
    expect(redoLabel(DOC)).toBeNull();
    expect(undo(DOC)).toBeNull();
    expect(redo(DOC)).toBeNull();
  });

  it("undo restores the before snapshot and moves the step to redo; redo restores after", () => {
    recordEdit(DOC, entry("bold", false, "s0", "s1"));
    expect(canUndo(DOC)).toBe(true);
    expect(undoLabel(DOC)).toBe("bold");

    const restored = undo(DOC);
    expect(restored?.model).toEqual({ id: "s0" });
    expect(canUndo(DOC)).toBe(false);
    expect(canRedo(DOC)).toBe(true);
    expect(redoLabel(DOC)).toBe("bold");

    const replayed = redo(DOC);
    expect(replayed?.model).toEqual({ id: "s1" });
    expect(canUndo(DOC)).toBe(true);
    expect(canRedo(DOC)).toBe(false);
  });
});

describe("typing grouping (AC-9, AC-12, AC-13)", () => {
  it("merges consecutive groupable edits with the same label into one step", () => {
    recordEdit(DOC, entry("typing", true, "s0", "s1"));
    recordEdit(DOC, entry("typing", true, "s1", "s2"));
    recordEdit(DOC, entry("typing", true, "s2", "s3"));

    expect(canRedo(DOC)).toBe(false);
    const restored = undo(DOC);
    // One undo unwinds the whole run, back to the state before the first
    // keystroke — the group's original `before`, not the last edit's.
    expect(restored?.model).toEqual({ id: "s0" });
    expect(canUndo(DOC)).toBe(false);
  });

  it("a non-typing (non-groupable) command is always its own step, never merged", () => {
    recordEdit(DOC, entry("typing", true, "s0", "s1"));
    recordEdit(DOC, entry("bold", false, "s1", "s2"));
    recordEdit(DOC, entry("typing", true, "s2", "s3"));

    expect(undo(DOC)?.model).toEqual({ id: "s2" });
    expect(undo(DOC)?.model).toEqual({ id: "s1" });
    expect(undo(DOC)?.model).toEqual({ id: "s0" });
    expect(canUndo(DOC)).toBe(false);
  });

  it("does not merge across two different groupable labels", () => {
    recordEdit(DOC, entry("typing", true, "s0", "s1"));
    recordEdit(DOC, entry("special-paste", true, "s1", "s2"));

    expect(undo(DOC)?.model).toEqual({ id: "s1" });
    expect(undo(DOC)?.model).toEqual({ id: "s0" });
  });
});

describe("breaking the typing group (AC-10, AC-11)", () => {
  it("breakGroup (a caret move) stops the next groupable edit from merging", () => {
    recordEdit(DOC, entry("typing", true, "s0", "s1"));
    breakGroup(DOC);
    recordEdit(DOC, entry("typing", true, "s1", "s2"));

    expect(undo(DOC)?.model).toEqual({ id: "s1" });
    expect(undo(DOC)?.model).toEqual({ id: "s0" });
  });

  it("breakGroup on a document with no history yet is a no-op", () => {
    expect(() => breakGroup("never-touched")).not.toThrow();
  });

  it("a 2 s pause breaks the group so the next keystroke starts a new step", () => {
    vi.useFakeTimers();
    try {
      recordEdit(DOC, entry("typing", true, "s0", "s1"));
      vi.advanceTimersByTime(TYPING_GROUP_PAUSE_MS - 1);
      recordEdit(DOC, entry("typing", true, "s1", "s2"));
      vi.advanceTimersByTime(TYPING_GROUP_PAUSE_MS);
      recordEdit(DOC, entry("typing", true, "s2", "s3"));

      // The pause landed between the second and third edit, so undo peels
      // off "s2..s3" first, then the merged "s0..s2" run.
      expect(undo(DOC)?.model).toEqual({ id: "s2" });
      expect(undo(DOC)?.model).toEqual({ id: "s0" });
      expect(canUndo(DOC)).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });

  it("resets the pause timer on every groupable edit", () => {
    vi.useFakeTimers();
    try {
      recordEdit(DOC, entry("typing", true, "s0", "s1"));
      vi.advanceTimersByTime(TYPING_GROUP_PAUSE_MS - 1);
      recordEdit(DOC, entry("typing", true, "s1", "s2"));
      vi.advanceTimersByTime(TYPING_GROUP_PAUSE_MS - 1);
      recordEdit(DOC, entry("typing", true, "s2", "s3"));

      // Never idle for a full 2 s, so the whole run stays one step.
      expect(undo(DOC)?.model).toEqual({ id: "s0" });
      expect(canUndo(DOC)).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("redo stack (AC-15)", () => {
  it("a new edit after undo clears the redo stack", () => {
    recordEdit(DOC, entry("bold", false, "s0", "s1"));
    undo(DOC);
    expect(canRedo(DOC)).toBe(true);

    recordEdit(DOC, entry("italic", false, "s0", "s2"));
    expect(canRedo(DOC)).toBe(false);
    expect(redo(DOC)).toBeNull();
  });
});

describe("bounded history (AC-16, AC-17)", () => {
  it("holds at least 200 steps", () => {
    expect(MAX_HISTORY_STEPS).toBeGreaterThanOrEqual(200);
    for (let i = 0; i < MAX_HISTORY_STEPS; i += 1) {
      recordEdit(DOC, entry("bold", false, `s${i}`, `s${i + 1}`));
    }
    let steps = 0;
    while (undo(DOC)) {
      steps += 1;
    }
    expect(steps).toBe(MAX_HISTORY_STEPS);
  });

  it("evicts the oldest step once the cap is exceeded, bounding memory", () => {
    for (let i = 0; i < MAX_HISTORY_STEPS + 10; i += 1) {
      recordEdit(DOC, entry("bold", false, `s${i}`, `s${i + 1}`));
    }
    let steps = 0;
    let last: HistorySnapshot | null = null;
    let restored: HistorySnapshot | null;
    while ((restored = undo(DOC))) {
      steps += 1;
      last = restored;
    }
    expect(steps).toBe(MAX_HISTORY_STEPS);
    // The oldest 10 steps (before "s10") were evicted.
    expect(last?.model).toEqual({ id: "s10" });
  });
});

describe("per-document isolation and clearing (AC-18, AC-19)", () => {
  it("keeps separate stacks per document id", () => {
    recordEdit(DOC, entry("bold", false, "a0", "a1"));
    recordEdit("doc-2", entry("italic", false, "b0", "b1"));

    expect(undoLabel(DOC)).toBe("bold");
    expect(undoLabel("doc-2")).toBe("italic");

    undo(DOC);
    expect(canUndo(DOC)).toBe(false);
    expect(canUndo("doc-2")).toBe(true);
  });

  it("clearHistory discards both stacks for that document only", () => {
    recordEdit(DOC, entry("bold", false, "a0", "a1"));
    recordEdit("doc-2", entry("italic", false, "b0", "b1"));
    undo("doc-2");

    clearHistory(DOC);

    expect(canUndo(DOC)).toBe(false);
    expect(canRedo(DOC)).toBe(false);
    expect(canRedo("doc-2")).toBe(true);
  });
});
