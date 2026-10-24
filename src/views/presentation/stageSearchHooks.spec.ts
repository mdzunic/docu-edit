// @vitest-environment jsdom
// SPEC-098 — the stage hooks a deck search drives: a hit opens a session with
// its range (a cell's in its cell), a notes hit selects its range in the
// notes textarea, and closing the bar ends a session it opened before handing
// focus back to the stage.
import { nextTick, ref, shallowRef } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel } from "../../lib/presentation/model";
import type { DeckHit } from "../../lib/presentation/deckSearch";
import { notesOffset, stageSearchHooks } from "./stageSearchHooks";

function hooks(editingFrameId: string | null = null) {
  const editing = { editingFrameId: ref(editingFrameId), stopEditing: vi.fn(), startEditing: vi.fn(), startEditingCell: vi.fn() };
  const selection = { select: vi.fn() };
  const notes = { open: ref(false), toggle: vi.fn(() => void (notes.open.value = true)) };
  const model = shallowRef({ slides: [{ id: "s", frames: [], notes: ["first line", "Open here"] }] } as unknown as PresentationModel);
  const stageEdit = { editing, selection, surface: () => null } as unknown as Parameters<typeof stageSearchHooks>[0];
  return { h: stageSearchHooks(stageEdit, notes, model), editing, selection, notes };
}

const hit = (extra: Partial<DeckHit>): DeckHit => ({ slideIndex: 0, para: 1, start: 0, end: 4, scope: "frame", snippet: "Open", snippetOffset: 0, ...extra });

afterEach(() => {
  document.body.innerHTML = "";
});

describe("stageSearchHooks", () => {
  it("opens a session on a frame hit, or on a cell hit's cell, with the range selected", () => {
    const { h, editing } = hooks();
    const range = { start: { para: 1, offset: 0 }, end: { para: 1, offset: 4 } };
    h.startEditing(hit({ frameId: "f" }));
    expect(editing.startEditing).toHaveBeenCalledWith("f", false, undefined, range);
    h.startEditing(hit({ frameId: "t", scope: "cell", row: 1, column: 2 }));
    expect(editing.startEditingCell).toHaveBeenCalledWith("t", 1, 2, undefined, range);
  });

  it("opens the notes panel and selects a notes hit in the textarea", async () => {
    const area = document.createElement("textarea");
    area.dataset.testid = "slides-notes-editor";
    area.value = "first line\nOpen here";
    document.body.appendChild(area);
    const { h, notes } = hooks();
    await h.openNotes(hit({ scope: "notes" }));
    expect(notes.toggle).toHaveBeenCalledTimes(1);
    expect(area.value.slice(area.selectionStart, area.selectionEnd)).toBe("Open");
    expect(notesOffset(["ab", "cd"], 1, 1)).toBe(4);
  });

  it("closing the bar ends a session it opened, then focuses the selected frame", async () => {
    const frame = document.createElement("div");
    frame.dataset.testid = "slide-frame";
    frame.className = "is-selected";
    frame.tabIndex = 0;
    document.body.appendChild(frame);
    const { h, editing } = hooks("f");
    h.focusStage();
    await nextTick();
    expect(editing.stopEditing).toHaveBeenCalledTimes(1);
    expect(document.activeElement).toBe(frame);
  });
});
